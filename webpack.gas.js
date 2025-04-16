const path = require('path');
const CopyPlugin = require('copy-webpack-plugin');
const GasPlugin = require('gas-webpack-plugin');
const fs = require('fs');

// Get a list of all JavaScript files in the root directory
const jsFiles = fs.readdirSync(__dirname)
  .filter(file => file.endsWith('.js') && !file.startsWith('webpack'))
  .filter(file => file !== 'index.js' && !file.includes('.bundle.js'));

// Get a list of all HTML files in the root directory
const htmlFiles = fs.readdirSync(__dirname)
  .filter(file => file.endsWith('.html'));

// Create copy patterns for all files
const copyPatterns = [
  { from: 'appsscript.json', to: 'appsscript.json' },
  { from: '.clasp.json', to: '.clasp.json' },
  { from: 'src/api.js', to: 'api.js' }
];

// Add all JavaScript files to copy patterns
jsFiles.forEach(file => {
  copyPatterns.push({ from: file, to: file });
});

// Add all HTML files to copy patterns
htmlFiles.forEach(file => {
  copyPatterns.push({ from: file, to: file });
});

module.exports = {
  mode: 'development',
  entry: {
    main: './src/index.js'
  },
  output: {
    filename: '[name].bundle.js',
    path: path.resolve(__dirname, 'dist'),
    libraryTarget: 'var',
    library: 'AppLib'
  },
  resolve: {
    fallback: {
      "path": require.resolve("path-browserify"),
      "fs": false,
      "os": false,
      "util": false,
      "stream": false,
      "buffer": false,
      "http": false,
      "https": false,
      "url": false,
      "net": false,
      "tls": false,
      "crypto": false,
      "zlib": false
    }
  },
  module: {
    rules: [
      {
        test: /\.js$/,
        exclude: /node_modules/,
        use: {
          loader: 'babel-loader',
          options: {
            presets: ['@babel/preset-env']
          }
        }
      }
    ]
  },
  plugins: [
    new GasPlugin(),
    new CopyPlugin(copyPatterns)
  ]
}; 