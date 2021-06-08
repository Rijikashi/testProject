// const devCerts = require("office-addin-dev-certs");
const { CleanWebpackPlugin } = require("clean-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const MiniCssExtractPlugin = require('mini-css-extract-plugin');
const precss = require('precss');
const autoprefixer = require('autoprefixer');
const path = require('path');
const fs = require("fs");
const webpack = require("webpack");

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: "@babel/polyfill",
      taskpane: "./src/taskpane/taskpane.js",
      commands: "./src/commands/commands.js"
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"]
    },
    // output:{
    //   path: path.join(__dirname,'build'),
    //   filename: '[name].bundle.js',
    // },
    module: {
      rules: [
        {
          test: /\.(scss)$/,
          use:[{
            loader: MiniCssExtractPlugin.loader,
          },{
            loader: 'css-loader',
          }, {
            loader: 'postcss-loader',
            options:{
              plugins(){
                return [precss, autoprefixer,]
              },
            },
          },{
            loader:'sass-loader',
          }],
        },
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
            options: {
              presets: ["@babel/preset-env"]
            }
          }
        },
        {
          test:/\.css$/,
          use:[{
            loader:MiniCssExtractPlugin.loader,
          },'css-loader']
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader"
        },
        {
          test: /\.(png|jpg|jpeg|gif)$/,
          use: "file-loader",
        }
      ]
    },
    plugins: [
      new CleanWebpackPlugin(),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"]
      }),
      new webpack.ProvidePlugin({
        $:'jquery',
        jQuery:'jquery',
        jquery:'jquery',
        moment:'moment'
      }),
      new CopyWebpackPlugin([
        {
          to: "assets/",
          from: "./assets/"
        },
        {
          to:"template-response-example.json",
          from:"template-response-example.json"
        }
      ]),
      new MiniCssExtractPlugin({
        filename: '[name].css',
        chunkFileName:'[id].css',
        ignoreOrder: false,
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      })
    ]
    // devServer: {
    //   headers: {
    //     "Access-Control-Allow-Origin": "*"
    //   },      
    //   https: (options.https !== undefined) ? options.https : await devCerts.getHttpsServerOptions(),
    //   port: process.env.npm_package_config_dev_server_port || 3000
    // }
  };

  return config;
};
