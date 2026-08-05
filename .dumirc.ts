import { defineConfig } from 'dumi';
import { resolve as resolvePath } from 'node:path';

export default defineConfig({
  base: '/office-file-viewer/',
  publicPath: '/office-file-viewer/',
  outputPath: 'docs-dist',
  locales: [
    { id: 'en-US', name: 'English' },
    { id: 'zh-CN', name: '简体中文', base: '/zh-CN' },
  ],
  // 公开页面全部通过路由白名单注册，避免本地 smoke-test 和样例文件进入 Pages 构建产物。
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
    {
      path: '/docs',
      component: resolvePath(__dirname, 'docs/dev/docs.md'),
    },
    {
      path: '/docs/zh-CN',
      component: resolvePath(__dirname, 'docs/dev/docs.zh-CN.md'),
    },
  ],
  themeConfig: {
    name: 'Office File Viewer',
    prefersColor: {
      default: 'light',
      switch: false,
    },
    logo: false,
    nav: {
      'en-US': [
        { title: 'Home', link: '/' },
        { title: 'Docs', link: '/docs' },
      ],
      'zh-CN': [
        { title: '首页', link: '/zh-CN/' },
        { title: '文档', link: '/zh-CN/docs' },
      ],
    },
    socialLinks: {
      github: 'https://github.com/gyxing/office-file-viewer',
    },
    sidebar: {
      '/docs': [],
      '/zh-CN/docs': [],
    },
  },
  resolve: {
    docDirs: [],
  },
});
