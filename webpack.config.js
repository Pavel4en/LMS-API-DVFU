// webpack.config.js
const path = require('path');

module.exports = {
  entry: './src/renderer.js',  // исходный код renderer
  output: {
    path: path.resolve(__dirname, 'dist'),
    filename: 'renderer.bundle.js'
  },
  target: 'electron-renderer',  // сборка для renderer-процесса Electron
  module: {
    rules: [
      {
        test: /\.js$/,
        exclude: /node_modules/,
        use: {
          loader: 'babel-loader',
          options: {
            presets: ['@babel/preset-env']  // используем Babel для транспиляции современного кода
          }
        }
      }
    ]
  },
  mode: 'development'
};
