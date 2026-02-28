import { defineConfig } from 'vite';
import { resolve } from 'path';

export default defineConfig({
  resolve: {
    alias: {
      '@': resolve(__dirname, 'src'),
      '@docs': resolve(__dirname, 'docs')
    }
  },
  build: {
    lib: {
      entry: resolve(__dirname, 'src/index.js'),
      name: 'SheetNext',
      formats: ['es', 'umd'],
      fileName: (format) => `sheetnext.${format}.js`
    },
    cssCodeSplit: false,
    emptyOutDir: true
  }
});
