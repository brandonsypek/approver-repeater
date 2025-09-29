import resolve from '@rollup/plugin-node-resolve';
import typescript from '@rollup/plugin-typescript';
import terser from '@rollup/plugin-terser'; // Correct Rollup plugin

export default {
  input: 'approvers-repeater.ts',
  output: {
    file: 'dist/approvers-repeater.min.js',
    format: 'iife',
    name: 'ApproversRepeater',
    sourcemap: false // Smaller bundle
  },
  plugins: [
    resolve({
      moduleDirectories: ['node_modules']
    }),
    typescript({
      tsconfig: './tsconfig.json',
      include: ['*.ts'],
      exclude: ['node_modules', 'dist']
    }),
    terser() // Minify output
  ]
};