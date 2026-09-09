/* eslint-disable no-undef */
const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const devOrigin = "https://localhost:3000";
const prodOrigin = (process.env.ADDIN_PUBLIC_URL || "https://YOUR-ADDIN-HOST.example").replace(/\/$/, "");

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const activeOrigin = dev ? devOrigin : prodOrigin;

  return {
    devtool: dev ? "source-map" : false,
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: "./src/taskpane/taskpane.ts",
      commands: "./src/commands/commands.ts"
    },
    output: { clean: true },
    resolve: { extensions: [".ts", ".js"] },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: { loader: "babel-loader", options: { presets: ["@babel/preset-typescript"] } }
        }
      ]
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"]
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      }),
      new HtmlWebpackPlugin({
        filename: "auth.html",
        template: "./src/taskpane/auth.html",
        chunks: []
      }),
      new HtmlWebpackPlugin({
        filename: "support.html",
        template: "./src/taskpane/support.html",
        chunks: []
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets", to: "assets" },
          { from: "src/taskpane/taskpane.css", to: "taskpane.css" },
          {
            from: "manifest.xml",
            to: "manifest.xml",
            transform(content) {
              return content.toString().replaceAll(devOrigin, activeOrigin);
            }
          }
        ]
      })
    ],
    devServer: {
      headers: { "Access-Control-Allow-Origin": "*" },
      server: { type: "https", options: await getHttpsOptions() },
      port: 3000
    }
  };
};
