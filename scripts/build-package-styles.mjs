import less from 'less';
import {
  access,
  readdir,
  readFile,
  stat,
  unlink,
  writeFile,
} from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';

/** 发布样式只允许处理仓库构建生成的固定 dist 目录。 */
const DIST_ROOT = resolve('dist');

/** 仅改写 ESM 中带引号的相对 Less 路径，避免误伤文本和包名。 */
const LESS_SPECIFIER_PATTERN = /(['"])(\.{1,2}\/[^'"\r\n]+)\.less\1/g;

/** 用于确认改写后的每个相对 CSS 引用都有对应产物。 */
const CSS_SPECIFIER_PATTERN = /(['"])(\.{1,2}\/[^'"\r\n]+\.css)\1/g;

/** 递归收集目录内的文件，供编译、改写和最终复检共用。 */
async function collectFiles(directory) {
  const entries = await readdir(directory, { withFileTypes: true });
  const nested = await Promise.all(
    entries.map((entry) => {
      const target = join(directory, entry.name);
      return entry.isDirectory() ? collectFiles(target) : [target];
    }),
  );
  return nested.flat();
}

/** 编译单个 Less 文件，并在 CSS 成功写入后移除发布源文件。 */
async function compileLessFile(lessFile) {
  try {
    const source = await readFile(lessFile, 'utf8');
    const result = await less.render(source, { filename: lessFile });
    const cssFile = lessFile.slice(0, -'.less'.length) + '.css';
    await writeFile(cssFile, result.css, 'utf8');
    await unlink(lessFile);
  } catch (error) {
    const reason = error instanceof Error ? error.message : String(error);
    throw new Error(`编译发布样式失败：${lessFile}\n${reason}`);
  }
}

/** 把构建 JS 中的相对 Less 引用精确替换为同路径 CSS。 */
function replaceLessSpecifiers(source) {
  return source.replace(
    LESS_SPECIFIER_PATTERN,
    (_match, quote, specifier) => `${quote}${specifier}.css${quote}`,
  );
}

/** 检查最终产物不再包含 Less，并验证所有相对 CSS 引用可解析。 */
async function validatePackageStyles(distRoot) {
  const outputFiles = await collectFiles(distRoot);
  const remainingLessFiles = outputFiles.filter((file) =>
    file.endsWith('.less'),
  );
  const jsFiles = outputFiles.filter((file) => file.endsWith('.js'));
  const jsSources = await Promise.all(
    jsFiles.map(async (file) => ({
      file,
      source: await readFile(file, 'utf8'),
    })),
  );
  const remainingLessReferences = jsSources.filter(({ source }) =>
    source.includes('.less'),
  );

  if (remainingLessFiles.length || remainingLessReferences.length) {
    const details = [
      ...remainingLessFiles,
      ...remainingLessReferences.map(({ file }) => file),
    ];
    throw new Error(`发布产物仍包含 Less 文件或引用：\n${details.join('\n')}`);
  }

  await Promise.all(
    jsSources.flatMap(({ file, source }) =>
      Array.from(source.matchAll(CSS_SPECIFIER_PATTERN), async (match) => {
        const cssFile = resolve(dirname(file), match[2]);
        try {
          await access(cssFile);
        } catch {
          throw new Error(`发布 JS 引用了不存在的 CSS：${file} -> ${match[2]}`);
        }
      }),
    ),
  );
}

/** 编译并校验 Father 输出的包样式。 */
async function compilePackageStyles(distRoot) {
  const distStats = await stat(distRoot).catch(() => null);
  if (!distStats?.isDirectory()) {
    throw new Error(`发布目录不存在：${distRoot}`);
  }

  const files = await collectFiles(distRoot);
  const lessFiles = files.filter((file) => file.endsWith('.less'));
  for (const lessFile of lessFiles) {
    await compileLessFile(lessFile);
  }

  const jsFiles = files.filter((file) => file.endsWith('.js'));
  for (const jsFile of jsFiles) {
    const source = await readFile(jsFile, 'utf8');
    const nextSource = replaceLessSpecifiers(source);
    if (nextSource !== source) await writeFile(jsFile, nextSource, 'utf8');
  }

  await validatePackageStyles(distRoot);
  console.log(`已生成并校验 ${lessFiles.length} 个发布 CSS 文件。`);
}

try {
  await compilePackageStyles(DIST_ROOT);
} catch (error) {
  console.error(error instanceof Error ? error.message : error);
  process.exitCode = 1;
}
