import resolve from '@rollup/plugin-node-resolve';
import commonjs from '@rollup/plugin-commonjs';

export default {
  input: 'src/index.js',
  output: [
    {
      file: 'dist/index.cjs.js',
      format: 'cjs',
      exports: 'named',
      sourcemap: true
    },
    {
      file: 'dist/index.esm.js',
      format: 'esm',
      sourcemap: true
    },
    {
      file: 'dist/index.umd.js',
      format: 'umd',
      name: 'OfficeWordDiff',
      exports: 'named',
      sourcemap: true,
      globals: {
        '@microsoft/office-js': 'Office'
      }
    }
  ],
  external: ['@microsoft/office-js'],
  plugins: [
    resolve(),
    commonjs()
  ]
};
