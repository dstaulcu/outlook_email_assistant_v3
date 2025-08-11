const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = {
  entry: {
    taskpane: './src/taskpane/taskpane.js',
    commands: './src/commands/commands.js'
  },
  output: {
    filename: '[name].bundle.js',
    path: path.resolve(__dirname, 'public'),
    clean: true
  },
  module: {
    rules: [
      {
        test: /\.css$/i,
        use: ['style-loader', 'css-loader']
      },
      {
        test: /\.(png|svg|jpg|jpeg|gif)$/i,
        type: 'asset/resource',
        generator: {
          filename: 'icons/[name][ext]'
        }
      }
    ]
  },
  plugins: [
    new HtmlWebpackPlugin({
      template: './src/taskpane/taskpane.html',
      filename: 'taskpane.html',
      chunks: ['taskpane']
    }),
    new HtmlWebpackPlugin({
      template: './src/commands/commands.html',
      filename: 'index.html',
      chunks: ['commands']
    }),
    new CopyWebpackPlugin({
      patterns: [
        {
          from: './src/assets/icons',
          to: 'icons'
        },
        {
          from: './src/assets/css',
          to: '.',
          globOptions: {
            ignore: ['**/*.scss']
          }
        },
        {
          from: './src/default-providers.json',
          to: 'default-providers.json'
        },
        {
          from: './src/default-models.json',
          to: 'default-models.json'
        }
      ]
    })
  ],
  resolve: {
    extensions: ['.js', '.css']
  },
  devtool: 'source-map',
  optimization: {
    minimize: true
  }
};
