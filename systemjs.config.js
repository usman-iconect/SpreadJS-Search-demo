(function (global) {
    System.config({
        transpiler: 'plugin-babel',
        babelOptions: {
            es2015: true,
            react: true
        },
        meta: {
            '*.css': { loader: 'css' }
        },
        paths: {
            // paths serve as alias
            'npm:': 'node_modules/'
        },
        // map tells the System loader where to look for things
        map: {
            '@mescius/spread-sheets': 'npm:@mescius/spread-sheets/index.js',
            '@mescius/spread-sheets-print': 'npm:@mescius/spread-sheets-print/index.js',
            '@mescius/spread-sheets-react': 'npm:@mescius/spread-sheets-react/index.js',
            '@mescius/spread-sheets-io': 'npm:@mescius/spread-sheets-io/index.js',
            '@mescius/spread-sheets-charts': 'npm:@mescius/spread-sheets-charts/index.js',
            '@mescius/spread-sheets-shapes': 'npm:@mescius/spread-sheets-shapes/index.js',
            '@mescius/spread-sheets-slicers': 'npm:@mescius/spread-sheets-slicers/index.js',
            '@mescius/spread-sheets-pivot-addon': 'npm:@mescius/spread-sheets-pivot-addon/index.js',
            '@mescius/spread-sheets-reportsheet-addon': 'npm:@mescius/spread-sheets-reportsheet-addon/index.js',
            '@mescius/spread-sheets-tablesheet': 'npm:@mescius/spread-sheets-tablesheet/index.js',
            '@mescius/spread-sheets-ganttsheet': 'npm:@mescius/spread-sheets-ganttsheet/index.js',
            '@grapecity/jsob-test-dependency-package/react-components': 'npm:@grapecity/jsob-test-dependency-package/react-components/index.js',
            'react': 'npm:react/umd/react.production.min.js',
            'react-dom': 'npm:react-dom/umd/react-dom.production.min.js',
            'css': 'npm:systemjs-plugin-css/css.js',
            'plugin-babel': 'npm:systemjs-plugin-babel/plugin-babel.js',
            'systemjs-babel-build':'npm:systemjs-plugin-babel/systemjs-babel-browser.js'
        },
        // packages tells the System loader how to load when no filename and/or no extension
        packages: {
            src: {
                defaultExtension: 'jsx'
            },
            "node_modules": {
                defaultExtension: 'js'
            },
        }
    });
})(this);
