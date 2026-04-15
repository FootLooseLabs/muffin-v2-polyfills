import terser from '@rollup/plugin-terser';
import replace from '@rollup/plugin-replace';

const isProd = process.env.NODE_ENV === 'production';

const banner = `/*! muffin-v2-polyfills v${process.env.npm_package_version} | Footloose Labs */`;

export default {
    input: 'src/index.js',
    output: {
        file: 'dist/v2-polyfills.min.js',
        format: 'iife',
        banner,
        sourcemap: isProd ? false : 'inline'
    },
    plugins: [
        replace({
            'process.env.NODE_ENV': JSON.stringify(process.env.NODE_ENV || 'development'),
            preventAssignment: true
        }),
        isProd && terser()
    ].filter(Boolean)
};
