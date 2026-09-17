import { defineConfig } from 'vite';

// Em GitHub Pages o site fica em /cubo/; localmente e em outros hosts usa a raiz.
export default defineConfig({
  base: process.env.GITHUB_PAGES === 'true' ? '/cubo/' : '/',
  build: {
    outDir: 'dist',
    assetsDir: 'assets',
  },
});
