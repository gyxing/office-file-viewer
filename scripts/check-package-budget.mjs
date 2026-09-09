import { readdir, stat } from 'node:fs/promises';
import { join, resolve } from 'node:path';

/** 发布目录允许的最大文件数量，避免模块输出无意膨胀。 */
const MAX_DIST_FILE_COUNT = 1300;
/** 发布目录允许的最大未压缩体积。 */
const MAX_DIST_BYTES = 6 * 1024 * 1024;
/** 单文件解析 Worker 允许的最大体积。 */
const MAX_WORKER_BYTES = 900 * 1024;
/** Father、样式和 Worker 构建共同写入的发布目录。 */
const DIST_ROOT = resolve('dist');
/** 所有格式共享的单文件 Worker 产物。 */
const WORKER_FILE = resolve(
  'dist/office-file-viewer/services/parsing/runtime/worker/entry.js',
);

/** 递归统计发布目录文件数量和总字节数。 */
async function measureDirectory(directory) {
  const entries = await readdir(directory, { withFileTypes: true });
  const measurements = await Promise.all(
    entries.map(async (entry) => {
      const target = join(directory, entry.name);
      if (entry.isDirectory()) return measureDirectory(target);
      const fileStat = await stat(target);
      return { files: 1, bytes: fileStat.size };
    }),
  );
  return measurements.reduce(
    (total, current) => ({
      files: total.files + current.files,
      bytes: total.bytes + current.bytes,
    }),
    { files: 0, bytes: 0 },
  );
}

/** 输出稳定入口及 Worker 分组的文件数和字节数，便于定位体积变化。 */
async function logEntryGroups() {
  const groups = {
    root: ['dist/index.js', 'dist/index.d.ts'],
    core: ['dist/core.js', 'dist/core.d.ts'],
    layout: ['dist/layout.js', 'dist/layout.d.ts'],
    plugins: ['dist/plugins.js', 'dist/plugins.d.ts'],
    export: ['dist/export.js', 'dist/export.d.ts'],
    worker: [WORKER_FILE],
  };
  const entries = await Promise.all(
    Object.entries(groups).map(async ([name, paths]) => {
      const existing = await Promise.all(
        paths.map(async (path) => {
          const fileStat = await stat(resolve(path)).catch(() => null);
          return fileStat ? { files: 1, bytes: fileStat.size } : { files: 0, bytes: 0 };
        }),
      );
      return [
        name,
        existing.reduce(
          (total, current) => ({ files: total.files + current.files, bytes: total.bytes + current.bytes }),
          { files: 0, bytes: 0 },
        ),
      ];
    }),
  );
  for (const [name, measurement] of entries) {
    console.log(`  ${name}: ${measurement.files} 个文件，${measurement.bytes} B`);
  }
}

/** 超出预算时终止构建，并提供实际值和上限供定位。 */
function assertBudget(name, actual, limit, unit = '') {
  if (actual <= limit) return;
  throw new Error(`${name} 超出预算：${actual}${unit} > ${limit}${unit}`);
}

/** 校验发布文件数量、总体积和 Worker 体积。 */
async function checkPackageBudget() {
  const [dist, worker] = await Promise.all([
    measureDirectory(DIST_ROOT),
    stat(WORKER_FILE),
  ]);
  assertBudget('发布文件数量', dist.files, MAX_DIST_FILE_COUNT);
  assertBudget('发布目录字节数', dist.bytes, MAX_DIST_BYTES, ' B');
  assertBudget('解析 Worker 字节数', worker.size, MAX_WORKER_BYTES, ' B');
  console.log(
    `发布预算通过：${dist.files} 个文件，${dist.bytes} B，Worker ${worker.size} B。`,
  );
  await logEntryGroups();
}

await checkPackageBudget();
