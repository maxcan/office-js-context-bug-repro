const fs = require('fs');
const path = require('path');
const webpack = require('webpack');
const webpackMerge = require('webpack-merge');
const commonConfig = require('./webpack.common.js');
const BrowserSyncPlugin = require('browser-sync-webpack-plugin');

const devHost = `localshim`

module.exports = webpackMerge(commonConfig, {
    devtool: 'eval-source-map',

    plugins: [
        new BrowserSyncPlugin(
            {

                https: {
                    key: "./certs/server.key",
                    cert: "./certs/localshim.crt",
                    // ca: fs.readFileSync("/path/to/ca.pem"),
                },
                host: devHost,
                port: 3000,
                proxy: 'https://' + devHost + ':3100/'
            },
            {
                reload: false
            }
        )
    ],

    devServer: {
        publicPath: '/',
        contentBase: path.resolve('dist'),
        https: {
            key: fs.readFileSync("./certs/server.key"),
            cert: fs.readFileSync("./certs/localshim.crt"),
            // ca: fs.readFileSync("/path/to/ca.pem"),
        },
        host: devHost,
        compress: true,
        overlay: {
            warnings: false,
            errors: true
        },
        port: 3100,
        historyApiFallback: true
    }
});
