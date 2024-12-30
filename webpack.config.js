const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const MiniCssExtractPlugin = require('mini-css-extract-plugin');

module.exports = {
    mode: 'production',
    entry: {
        popup: './src/popup.js',
        background: './src/background.js',
    },
    output: {
        path: path.resolve(__dirname, 'dist'),
        filename: '[name].js',
        clean: true,
    },
    module: {
        rules: [
            {
                test: /\.css$/i, // Для обработки CSS-файлов
                use: [MiniCssExtractPlugin.loader, 'css-loader'],
            },
            // {
            //     test: /\.js$/,
            //     exclude: /node_modules/,
            //     use: {
            //         loader: 'babel-loader'
            //     }
            // }
        ],
    },
    // resolve: {
    //     modules: [path.resolve(__dirname, 'src'), 'node_modules'], // Авто-поиск модулей
    // },
    plugins: [
        new HtmlWebpackPlugin({
            template: './src/popup.html',
            filename: 'popup.html',
            chunks: ['popup'],
        }),
        new CopyWebpackPlugin({
            patterns: [
                { from: './src/manifest.json', to: 'manifest.json' },
                { from: 'src/notification.html', to: 'notification.html' }, // Копируем HTML
                { from: 'src/images', to: 'images' }, // Копируем папку с изображениями
                { from: 'src/notification.js', to: 'notification.js' },
                { from: 'src/assets', to: 'assets' },
            ],
        }),
        new MiniCssExtractPlugin({
            filename: 'styles.css', // Генерация файлов CSS
        })
    ],
};
