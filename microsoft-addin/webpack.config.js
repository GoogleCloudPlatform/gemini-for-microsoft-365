/**
 * Copyright 2026 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * Webpack Configuration for Gemini for Microsoft 365
 * 
 * @author Sathya AG, Principal Architect, Google
 */

/* eslint-disable no-undef */

const fs = require("fs");
const path = require("path");
const webpack = require("webpack");
const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const pkg = require("./package.json");

function loadEnv() {
  const candidates = [
    path.resolve(__dirname, "..", ".env"),
    path.resolve(__dirname, ".env")
  ];
  for (const envPath of candidates) {
    if (fs.existsSync(envPath)) {
      const envContent = fs.readFileSync(envPath, "utf-8");
      for (const line of envContent.split(/\r?\n/)) {
        const trimmed = line.trim();
        if (trimmed && !trimmed.startsWith("#")) {
          const eqIdx = trimmed.indexOf("=");
          if (eqIdx !== -1) {
            const key = trimmed.slice(0, eqIdx).trim();
            const val = trimmed.slice(eqIdx + 1).trim().replace(/^["']|["']$/g, "");
            if (key) {
              process.env[key] = val;
            }
          }
        }
      }
    }
  }
}

loadEnv();

const urlDev = "https://localhost:3000/";
const defaultProdProject = process.env.PROJECT_ID || "";
const defaultProdRegion = process.env.REGION || "us-central1";
const urlProd = (process.env.FRONTEND_URL || (defaultProdProject ? `https://gemini-frontend-${defaultProdProject}.${defaultProdRegion}.run.app/` : "https://localhost:3000/"));

function getFrontendUrl(dev) {
  if (dev) return urlDev;
  let feUrl = (process.env.FRONTEND_URL || "").trim();
  if (feUrl) {
    if (!feUrl.endsWith("/")) feUrl += "/";
    return feUrl;
  }
  return urlProd;
}


async function getHttpsOptions() {
  try {
    const certPath = path.resolve(process.env.HOME || "", ".office-addin-dev-certs");
    const keyFile = path.join(certPath, "localhost.key");
    const certFile = path.join(certPath, "localhost.crt");
    const caFile = path.join(certPath, "ca.crt");
    if (fs.existsSync(keyFile) && fs.existsSync(certFile)) {
      return {
        key: fs.readFileSync(keyFile),
        cert: fs.readFileSync(certFile),
        ca: fs.existsSync(caFile) ? fs.readFileSync(caFile) : undefined
      };
    }
  } catch (e) {}
  return {};
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const configuredProxyUrl = (process.env.GEMINI_PROXY_URL || "").trim();
  const authProxyBase = configuredProxyUrl
    ? configuredProxyUrl.replace(/\/askGeminiEnterprise.*$/, "")
    : (defaultProdProject ? `https://auth-proxy-${defaultProdProject}.${defaultProdRegion}.run.app` : "http://localhost:8080");
  const devProxyTarget = process.env.BACKEND_PROXY_TARGET || authProxyBase;
  const proxyUrl = dev
    ? (process.env.GEMINI_PROXY_DEV_URL || (configuredProxyUrl.includes("localhost") ? configuredProxyUrl : "http://localhost:8080/askGeminiEnterprise"))
    : configuredProxyUrl;

  const copyPatterns = [
    {
      from: "Dockerfile",
      to: "Dockerfile",
      toType: "file",
    },
    {
      from: "nginx.conf",
      to: "nginx.conf",
      toType: "file",
    },
    {
      from: "assets/*",
      to: "assets/[name][ext][query]",
    },
    {
      from: "src/taskpane/taskpane.css",
      to: "taskpane.css",
      toType: "file",
      noErrorOnMissing: true,
    },
  ];

  const manifestDevSrc = path.resolve(__dirname, "manifest-dev.xml");
  if (fs.existsSync(manifestDevSrc)) {
    copyPatterns.push({
      from: manifestDevSrc,
      to: "manifest-dev.xml",
      toType: "file",
      noErrorOnMissing: true
    });
  }
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: "./src/taskpane/taskpane.js",
      commands: "./src/commands/commands.js",
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: { loader: "html-loader", options: { sources: false } },
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new webpack.DefinePlugin({
        "process.env.GEMINI_PROXY_URL": JSON.stringify(proxyUrl),
        "process.env.PROJECT_ID": JSON.stringify(process.env.PROJECT_ID || ""),
        "process.env.REGION": JSON.stringify(process.env.REGION || "us-central1"),
        "process.env.SERVICE_ACCOUNT": JSON.stringify(process.env.SERVICE_ACCOUNT || ""),
        "process.env.GE_GCP_PROJECT_ID": JSON.stringify(process.env.GE_GCP_PROJECT_ID || process.env.PROJECT_ID || ""),
        "process.env.GE_GCP_LOCATION": JSON.stringify(process.env.GE_GCP_LOCATION || "us"),
        "process.env.STREAM_ASSIST_ENDPOINT_LOCATION": JSON.stringify(process.env.STREAM_ASSIST_ENDPOINT_LOCATION || "us"),
        "process.env.GEMINI_ENTERPRISE_APP_ID": JSON.stringify(process.env.GEMINI_ENTERPRISE_APP_ID || ""),
        "process.env.ENTERPRISE_COLLECTION_ID": JSON.stringify(process.env.ENTERPRISE_COLLECTION_ID || "default_collection"),
        "process.env.ENTERPRISE_ASSISTANT_ID": JSON.stringify(process.env.ENTERPRISE_ASSISTANT_ID || "default_assistant"),
        "process.env.MICROSOFT_ENTRA_APP_ID": JSON.stringify(process.env.MICROSOFT_ENTRA_APP_ID || ""),
        "process.env.MICROSOFT_ENTRA_TENANT_ID": JSON.stringify(process.env.MICROSOFT_ENTRA_TENANT_ID || ""),
        "process.env.GOOGLE_OAUTH_CLIENT_ID": JSON.stringify(process.env.GOOGLE_OAUTH_CLIENT_ID || ""),
        "process.env.APP_VERSION": JSON.stringify(pkg.version || "1.2.0"),
        "process.env.BUILD_MODE": JSON.stringify(dev ? "dev" : "prod"),
        "process.env.BUILD_TIMESTAMP": JSON.stringify(new Date().toISOString())
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
        hash: true,
      }),
      new CopyWebpackPlugin({
        patterns: copyPatterns,
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
      new HtmlWebpackPlugin({
        filename: "google-auth.html",
        template: "./src/auth/google-auth.html",
        chunks: [],
      }),
      new HtmlWebpackPlugin({
        filename: "google-callback.html",
        template: "./src/auth/google-callback.html",
        chunks: [],
      }),
    ],
    devServer: {
      client: {
        overlay: false,
      },
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
      proxy: [
        {
          context: ["/askGeminiEnterprise", "/askGeminiStream", "/askGemini", "/sessions", "/deleteSession", "/api", "/auth", "/health"],
          target: devProxyTarget || "http://localhost:8080",
          secure: false,
          changeOrigin: true,
        },
      ],
    },
  };

  return config;
};
