import { defineConfig } from 'dumi';
import { resolve as resolvePath } from 'node:path';

export default defineConfig({
  base: '/office-file-viewer/',
  publicPath: '/office-file-viewer/',
  outputPath: 'homepage',
  locales: [
    { id: 'en-US', name: 'English' },
    { id: 'zh-CN', name: '简体中文', base: '/zh-CN' },
  ],
  // 只注册公开首页，避免本地 smoke-test 和样例文件进入 Pages 构建产物。
  routes: [
    {
      path: '/',
      component: resolvePath(__dirname, 'docs/dev/index.md'),
      layout: false,
    },
    {
      path: '/index/zh-CN',
      component: resolvePath(__dirname, 'docs/dev/index.zh-CN.md'),
      layout: false,
    },
  ],
  themeConfig: {
    name: 'Office File Viewer',
  },
  resolve: {
    docDirs: [],
  },
});
