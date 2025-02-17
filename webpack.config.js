const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const MiniCssExtractPlugin = require('mini-css-extract-plugin');

module.exports = {
    mode: 'production', // Режим сборки (production/development)
    entry: {
        popup: './src/popup.js', // Точка входа для popup.js
        background: './src/background.js', // Точка входа для background.js
    },
    output: {
        path: path.resolve(__dirname, 'dist'), // Папка для собранных файлов
        filename: '[name].js', // Имя выходного файла ([name] будет заменено на "popup" или "background")
        clean: true, // Очистка папки dist перед каждой сборкой
    },
    module: {
        rules: [
            {
                test: /\.css$/i, // Обработка CSS-файлов
                use: [MiniCssExtractPlugin.loader, 'css-loader'],
            },
            {
                test: /\.js$/, // Обработка JavaScript-файлов
                exclude: /node_modules/, // Исключаем node_modules
                use: {
                    loader: 'babel-loader', // Используем babel-loader для поддержки ES-модулей
                    options: {
                        presets: ['@babel/preset-env'], // Используем @babel/preset-env для поддержки современного JavaScript
                    },
                },
            },
        ],
    },
    resolve: {
        extensions: ['.js'], // Автоматическое разрешение расширений файлов
    },
    plugins: [
        // Генерация popup.html
        new HtmlWebpackPlugin({
            template: './src/popup.html', // Шаблон HTML
            filename: 'popup.html', // Имя выходного файла
            chunks: ['popup'], // Подключаем только popup.js
        }),
        // Копирование статических файлов
        new CopyWebpackPlugin({
            patterns: [
                { from: './src/manifest.json', to: 'manifest.json' }, // Копируем manifest.json
                { from: 'src/notification.html', to: 'notification.html' }, // Копируем HTML
                { from: 'src/images', to: 'images' }, // Копируем папку с изображениями
                { from: 'src/notification.js', to: 'notification.js' }, // Копируем notification.js
                { from: 'src/assets', to: 'assets' }, // Копируем папку с ассетами
                { from: 'src/create_event.html', to: 'create_event.html' }, // Копируем модальное окно создания события
            ],
        }),
        // Извлечение CSS в отдельный файл
        new MiniCssExtractPlugin({
            filename: 'styles.css', // Имя выходного CSS-файла
        }),
    ],
};