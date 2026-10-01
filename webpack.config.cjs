const path = require("path");
const MiniCssExtractPlugin = require("mini-css-extract-plugin");

module.exports = {
  entry: "./plugin/src/index.tsx",
  output: {
    path: path.resolve(__dirname, "plugin/dist"),
    filename: "plugin.bundle.js",
  },
  resolve: {
    extensions: [".ts", ".tsx", ".js"],
  },
  plugins: [new MiniCssExtractPlugin({ filename: "plugin.css" })],
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        loader: "ts-loader",
        exclude: /node_modules/,
        options: { transpileOnly: true },
      },
      {
        test: /\.css$/,
        use: [MiniCssExtractPlugin.loader, "css-loader", "postcss-loader"],
      },
    ],
  },
  devtool: "source-map",
  performance: { hints: false },
};
