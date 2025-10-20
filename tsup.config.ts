import { defineConfig } from 'tsup'

export default defineConfig(
  ['src/docx-preview.ts'].map(entry => ({
    entry: [entry],
    tsconfig: './tsconfig.json',
    outDir: `dist${entry.replace('src', '').split('/').slice(0, -1).join('/')}`,
    dts: true,
    sourcemap: true,
    format: ['esm'],
    //minify: true,
    clean: true,
    //external: [/^[^./]/],
    //esbuildPlugins: [],
  })),
)
