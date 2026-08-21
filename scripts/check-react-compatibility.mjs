import { execFile } from 'node:child_process';
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
/** Windows 需要显式调用 Yarn 的 cmd 包装器。 */
const YARN_COMMAND = process.platform === 'win32' ? 'yarn.cmd' : 'yarn';
/** React 运行时与类型声明按公开 peer 范围覆盖三个主版本。 */
const REACT_TARGETS = [
  {
    name: 'react-16',
    react: '16.9.0',
    reactTypes: '16.9.56',
    reactDomTypes: '16.9.25',
  },
  {
    name: 'react-17',
    react: '17.0.2',
    reactTypes: '17.0.80',
    reactDomTypes: '17.0.26',
  },
  {
    name: 'react-18',
    react: '18.3.1',
    reactTypes: '18.3.18',
    reactDomTypes: '18.3.7',
  },
];

/** 执行消费者安装或构建，并保留子进程输出供 CI 诊断。 */
async function runYarn(arguments_, cwd) {
  const result = await execFileAsync(YARN_COMMAND, arguments_, {
    cwd,
    env: process.env,
    shell: process.platform === 'win32',
    maxBuffer: 10 * 1024 * 1024,
  });
  if (result.stdout) process.stdout.write(result.stdout);
  if (result.stderr) process.stderr.write(result.stderr);
}

/** 创建只使用根入口的最小 React 消费项目。 */
async function createConsumer(directory, target, tarball) {
  await mkdir(directory, { recursive: true });
  await writeFile(
    join(directory, 'package.json'),
    JSON.stringify(
      {
        private: true,
        type: 'module',
        dependencies: {
          'office-file-viewer': `file:${tarball}`,
          react: target.react,
          'react-dom': target.react,
        },
        devDependencies: {
          '@types/react': target.reactTypes,
          '@types/react-dom': target.reactDomTypes,
          esbuild: '~0.18.20',
          typescript: '^4.8.4',
        },
      },
      null,
      2,
    ),
    'utf8',
  );
  await writeFile(
    join(directory, 'tsconfig.json'),
    JSON.stringify(
      {
        compilerOptions: {
          target: 'ES2015',
          module: 'ESNext',
          moduleResolution: 'Node',
          strict: true,
          jsx: 'react',
          skipLibCheck: false,
          esModuleInterop: true,
        },
        include: ['entry.tsx'],
      },
      null,
      2,
    ),
    'utf8',
  );
  await writeFile(
    join(directory, 'entry.tsx'),
    `import React from 'react';\nimport { OfficeFileViewer } from 'office-file-viewer';\nexport const preview = <OfficeFileViewer height="80vh" toolbar={{ openFile: false }} />;\n`,
    'utf8',
  );
}

/** 打包当前产物并验证 React 16.9、17 和 18 消费端类型与浏览器构建。 */
async function checkReactCompatibility() {
  const temporaryRoot = await mkdtemp(join(tmpdir(), 'office-react-compat-'));
  const tarball = join(temporaryRoot, 'office-file-viewer.tgz');
  // Yarn v1 会把 Windows 反斜杠 file 依赖误判为注册表版本。
  const tarballDependency = tarball.replace(/\\/g, '/');
  const packageMetadata = JSON.parse(
    await readFile(resolve('package.json'), 'utf8'),
  );
  try {
    await runYarn(['pack', '--filename', tarball], resolve('.'));
    for (const target of REACT_TARGETS) {
      const consumer = join(temporaryRoot, target.name);
      await createConsumer(consumer, target, tarballDependency);
      await runYarn(
        [
          'install',
          '--ignore-scripts',
          '--non-interactive',
          '--cache-folder',
          join(consumer, '.yarn-cache'),
        ],
        consumer,
      );
      const installedMetadata = JSON.parse(
        await readFile(
          join(consumer, 'node_modules/office-file-viewer/package.json'),
          'utf8',
        ),
      );
      if (installedMetadata.version !== packageMetadata.version) {
        throw new Error(
          `React ${target.react} 安装了错误包版本：${installedMetadata.version}，期望 ${packageMetadata.version}`,
        );
      }
      await runYarn(['run', 'tsc', '--noEmit'], consumer);
      await runYarn(
        [
          'run',
          'esbuild',
          'entry.tsx',
          '--bundle',
          '--format=esm',
          '--splitting',
          '--platform=browser',
          '--outdir=dist',
        ],
        consumer,
      );
      console.log(`React ${target.react} 消费端验证通过。`);
    }
  } finally {
    await rm(temporaryRoot, { recursive: true, force: true });
  }
}

await checkReactCompatibility();
