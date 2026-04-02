/* eslint-disable @typescript-eslint/no-var-requires */
const path = require("path");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const officeAddinConfig = require("./office-addin.config.cjs");

/** Local dev origin (must match manifest/manifest.xml placeholders). */
const urlDevOrigin = "https://localhost:3000";

/**
 * Production manifest URLs (no trailing slash).
 * 1) OFFICE_ADDIN_ORIGIN — set manually (e.g. custom domain).
 * 2) VERCEL_URL — set automatically on Vercel builds (so Office loads the deployed host, not localhost).
 * 3) office-addin.config.cjs productionOrigin (often localhost for local `npm run build` tests).
 */
function resolveProductionOrigin() {
  const explicit = process.env.OFFICE_ADDIN_ORIGIN;
  if (explicit) return String(explicit).replace(/\/$/, "");
  const vercel = process.env.VERCEL_URL;
  if (vercel) return `https://${String(vercel).replace(/\/$/, "")}`;
  return String(officeAddinConfig.productionOrigin).replace(/\/$/, "");
}

module.exports = async (env, options) => {
  const isProd = options.mode === "production";
  const urlProdOrigin = resolveProductionOrigin();

  // Dev HTTPS certs only for `webpack serve` — never load on production CI (no sudo, no cert install).
  const devServerHttpsOptions = isProd ? null : await require("office-addin-dev-certs").getHttpsServerOptions();

  return {
    devtool: "source-map",
    entry: {
      taskpane: "./src/taskpane/taskpane.tsx"
    },
    resolve: {
      extensions: [".ts", ".tsx", ".js"]
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: {
            loader: "ts-loader",
            // tsconfig has noEmit:true for `tsc --noEmit` typechecks; webpack needs emit for bundling.
            options: { compilerOptions: { noEmit: false } }
          }
        },
        {
          test: /\.css$/i,
          use: ["style-loader", "css-loader"]
        },
        {
          enforce: "pre",
          test: /\.js$/,
          loader: "source-map-loader"
        }
      ]
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"]
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "manifest/manifest.xml",
            to: "manifest.xml",
            transform: (content) => {
              const src = content.toString();
              if (!isProd) return src;
              // Replace every manifest URL that uses the local dev origin.
              return src.split(urlDevOrigin).join(urlProdOrigin);
            }
          }
        ]
      })
    ],
    devServer: isProd
      ? undefined
      : {
          port: 3000,
          // Office’s embedded WebView (esp. on Mac) often cannot use the dev server’s WebSocket/HMR client.
          // That shows up as a generic "Script error" from webpack’s overlay (handleError). Disable it.
          hot: false,
          liveReload: false,
          client: false,
          headers: {
            "Access-Control-Allow-Origin": "*"
          },
          server: {
            type: "https",
            options: devServerHttpsOptions
          },
          static: {
            directory: path.join(__dirname, "dist")
          },
          // Makes it obvious the dev server is running (it does not return to a shell prompt until you Ctrl+C).
          onListening(devServer) {
            const addr = devServer.server && devServer.server.address();
            const port = addr && typeof addr === "object" ? addr.port : 3000;
            // eslint-disable-next-line no-console
            console.log(
              `\n[Lower Third Builder] Dev server is running. Leave this terminal open.\n` +
                `  Task pane: https://localhost:${port}/taskpane.html\n` +
                `  Then sideload manifest/manifest.xml in PowerPoint (Insert → Add-ins → Add from file).\n`
            );
          }
        },
    output: {
      clean: true,
      path: path.resolve(__dirname, "dist"),
      filename: "[name].js"
    }
  };
};

