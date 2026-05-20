/*
 * Use this file to extend the webpack configuration for local development.
 * This configuration is merged with the default fast-serve webpack configuration.
 *
 * For more details, see: https://github.com/s-KaiNet/spfx-fast-serve
 */

const path = require('path');

// Option 1: Merging via webpackConfig object
const webpackConfig = {
  resolve: {
    alias: {
      // Add your custom path aliases here, for example:
      // '@src': path.resolve(__dirname, '..', 'src')
    }
  },
  plugins: [
    // Add custom plugins here
  ]
};

// Option 2: Advanced control via transformConfig function
const transformConfig = (initialWebpackConfig) => {
  // Transform initialWebpackConfig directly if needed
  return initialWebpackConfig;
};

module.exports = {
  webpackConfig,
  transformConfig
};
