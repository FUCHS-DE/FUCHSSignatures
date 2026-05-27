const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyPlugin = require('copy-webpack-plugin');

module.exports = {
    entry: {
        commands: './src/commands/commands.js'
    },
    output: {
        path: path.resolve(__dirname, 'dist'),
        filename: '[name].bundle.js',
        clean: true
    },
    plugins: [
        new HtmlWebpackPlugin({
            template: './src/commands/commands.html',
            filename: 'commands.html',
            chunks: ['commands'],
            // office.js muss VOR dem Bundle geladen werden
            inject: false
        }),
        new CopyPlugin({
            patterns: [
                { from: 'manifest.xml', to: 'manifest.xml' },
                { from: 'assets', to: 'assets', noErrorOnMissing: true }
            ]
        })
    ],
    devServer: {
        port: 3000,
        // HTTPS ist Pflicht für Office Add-ins
        server: 'https',
        static: './dist'
    }
};
