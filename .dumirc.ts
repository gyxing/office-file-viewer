import { defineConfig } from 'dumi';

export default defineConfig({
  outputPath: 'docs-dist',
  themeConfig: {
    name: 'Office File Viewer',
  },
  resolve: {
    docDirs: ['./docs/dev'],
  },
});
