/* eslint-disable */
const path = require("path");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

module.exports = (env, argv) => {
  const dev = argv.mode !== "production";

  return {
    entry: {
      taskpane: "./src/index.tsx",
    },

    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].bundle.js",
      clean: true,
      publicPath: "/", // ensure correct asset resolution
    },

    resolve: {
      extensions: [".ts", ".tsx", ".js", ".jsx"],
    },

    devtool: dev ? "inline-source-map" : false,

    devServer: {
      static: path.resolve(__dirname, "dist"),
      hot: true,
      port: 3000,
      allowedHosts: "all",
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
    },

    module: {
      rules: [
        {
          test: /\.tsx?$/,
          use: "ts-loader",
          exclude: /node_modules/,
        },
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader"],
        },
        {
          test: /\.(png|jpg|jpeg|gif|svg)$/,
          type: "asset/resource",
        }
      ],
    },

    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane.html",
        chunks: ["taskpane"],
      }),

      // Copies manifest + assets
      new CopyWebpackPlugin({
        patterns: [
          { from: "../manifest.xml", to: "../manifest.xml" }, // up one dir
        ],
      }),
    ],
  };
};
