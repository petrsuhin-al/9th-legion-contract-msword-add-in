import * as devCerts from 'office-addin-dev-certs';
import { CleanWebpackPlugin } from 'clean-webpack-plugin';
import * as CopyWebpackPlugin from 'copy-webpack-plugin';
import * as ExtractTextPlugin from 'extract-text-webpack-plugin';
import * as HtmlWebpackPlugin from 'html-webpack-plugin';
import * as webpack from 'webpack';
import { ServerOptions } from 'https';

interface Options {
  mode: string;
  https: ServerOptions;
}

module.exports = async (_: any, options: Options): Promise<webpack.Configuration> => ({
  devtool: 'source-map',
  entry: {
    vendor: [
      'react',
      'react-dom',
      'core-js',
      'office-ui-fabric-react'
    ],
    taskpane: [
      'react-hot-loader/patch',
      './src/taskpane/index.tsx'
    ],
    commands: './src/commands/commands.ts'
  },
  resolve: {
    extensions: ['.ts', '.tsx', '.html', '.js']
  },
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: [
          'react-hot-loader/webpack',
          'ts-loader'
        ],
        exclude: /node_modules/
      },
      {
        test: /\.css$/,
        use: ['style-loader', 'css-loader']
      },
      {
        test: /\.(png|jpe?g|gif|svg|woff|woff2|ttf|eot|ico)$/,
        use: {
          loader: 'file-loader',
          query: {
            name: 'assets/[name].[ext]'
          }
        }
      }
    ]
  },
  plugins: [
    new CleanWebpackPlugin(),
    new CopyWebpackPlugin({
      patterns: [
        {
          from: './src/taskpane/taskpane.css',
          to: 'taskpane.css'
        }
      ]
    }),
    new ExtractTextPlugin('[name].[hash].css'),
    new HtmlWebpackPlugin({
      filename: 'taskpane.html',
      template: './src/taskpane/taskpane.html',
      chunks: ['taskpane', 'vendor', 'polyfills']
    }),
    new HtmlWebpackPlugin({
      filename: 'commands.html',
      template: './src/commands/commands.html',
      chunks: ['commands']
    }),
    new CopyWebpackPlugin({
      patterns: [
        {
          from: './assets',
          to: 'assets'
        }
      ]
    }),
    new webpack.ProvidePlugin({
      Promise: ['es6-promise', 'Promise']
    })
  ],
  devServer: {
    hot: true,
    headers: {
      'Access-Control-Allow-Origin': '*'
    },
    https: (options.https !== undefined) ? options.https : await devCerts.getHttpsServerOptions(),
    port: Number(process.env.npm_package_config_dev_server_port) || 3000
  }
});
