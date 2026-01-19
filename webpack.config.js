const path = require('path');
const webpack = require('webpack');
const HtmlWebpackPlugin = require('html-webpack-plugin');

/** @type {import('webpack').Configuration} */
module.exports = {
  mode: process.env.NODE_ENV === 'production' ? 'production' : 'development',
  // Renderer runs with nodeIntegration=false, so bundle for a browser-like environment.
  target: 'web',
  entry: ['./src/renderer/polyfills.ts', './src/renderer/index.tsx'],
  devtool: 'source-map',
  output: {
    path: path.resolve(__dirname, 'dist'),
    filename: 'bundle.js',
    clean: true,
    // Ensure webpack runtime uses a browser-compatible global object in Electron renderer.
    globalObject: 'globalThis',
  },
  resolve: {
    extensions: ['.ts', '.tsx', '.js', '.jsx'],
    // Polyfill a small subset of Node core modules that some deps may pull in.
    fallback: {
      events: require.resolve('events/'),
      buffer: require.resolve('buffer/'),
      process: require.resolve('process/browser'),
    },
  },
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: 'ts-loader',
        exclude: /node_modules/,
      },
      {
        test: /\.css$/,
        use: ['style-loader', 'css-loader'],
      },
    ],
  },
  plugins: [
    new HtmlWebpackPlugin({
      template: path.resolve(__dirname, 'public', 'index.html'),
    }),
    new webpack.ProvidePlugin({
      Buffer: ['buffer', 'Buffer'],
      process: ['process'],
    }),
  ],
  devServer: {
    static: {
      directory: path.join(__dirname, 'dist'),
    },
    hot: true,
    port: 3000,
  },
};
