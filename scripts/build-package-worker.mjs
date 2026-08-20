import { build } from 'esbuild';
import { access, mkdir, readFile, stat } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';

/** Worker 源入口与 Father 产物路径均固定在当前包内，避免构建脚本越界写入。 */
const WORKER_SOURCE = resolve(
  'src/office-file-viewer/services/parsing/runtime/worker/entry.js',
);
const WORKER_OUTPUT = resolve(
  'dist/office-file-viewer/services/parsing/runtime/worker/entry.js',
);

/**
 * 把包含动态解析分块的 Worker 收敛为单文件，兼容默认生成 IIFE Worker 的消费端构建器。
 */
async function buildPackageWorker() {
  await access(WORKER_SOURCE);
  await mkdir(dirname(WORKER_OUTPUT), { recursive: true });
  await build({
    entryPoints: [WORKER_SOURCE],
    outfile: WORKER_OUTPUT,
    bundle: true,
    format: 'iife',
    platform: 'browser',
    target: ['es2020'],
    splitting: false,
    minify: true,
    legalComments: 'none',
  });

  const source = await readFile(WORKER_OUTPUT, 'utf8');
  if (/\bimport(?:\s+[{*\w]|\s*\()/.test(source)) {
    throw new Error('发布 Worker 仍包含模块导入，无法保证消费端单文件构建');
  }
  const { size } = await stat(WORKER_OUTPUT);
  console.log(`已生成单文件解析 Worker（${size} 字节）。`);
}

await buildPackageWorker();
