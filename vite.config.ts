import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { viteSingleFile } from 'vite-plugin-singlefile';
import path from 'node:path';

export default defineConfig({
  root: path.resolve(__dirname, 'ui'),
  plugins: [react(), viteSingleFile()],
  build: {
    outDir: path.resolve(__dirname, 'dist/ui'),
    emptyOutDir: true,
    target: 'es2022',
    assetsInlineLimit: 100000000,
    cssCodeSplit: false,
  },
});
