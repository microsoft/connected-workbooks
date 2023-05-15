const path = require('path');
const nodeExternals = require('webpack-node-externals');

module.exports = {
  target: 'node',
  mode: 'production',
  entry: './dist/cli.js', // make sure this matches the main root of your code 
  output: {
    path: path.join(__dirname, 'bundle'), // this can be any path and directory you want
    filename: 'extract-mquery.js',
  },
  optimization: {
    minimize: false, // enabling this reduces file size and readability
  },
};