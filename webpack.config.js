/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require("webpack");
const { spawn } = require("child_process");
const path = require("path");
const fs = require("fs").promises;
const fsSync = require("fs");
const https = require("https");

// Read version from version.json for DefinePlugin injection
const versionFile = path.join(__dirname, "version.json");
let APP_VERSION = "1.0.0";
try {
  APP_VERSION = JSON.parse(fsSync.readFileSync(versionFile, "utf8")).version || "1.0.0";
} catch { /* use default */ }

const urlDev = "https://localhost:3000/";
const urlProd = "https://www.contoso.com/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.js", "./src/taskpane/taskpane.html"],
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
          use: "html-loader",
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
        "process.env.APP_VERSION": JSON.stringify(APP_VERSION),
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "assets/templates/**/*",
            to: ({ absoluteFilename }) => {
              const rel = absoluteFilename.split("assets\\")[1] || absoluteFilename.split("assets/")[1];
              return `assets/${rel}`;
            },
          },
          {
            from: "src/taskpane/taskpane.css",
            to: "taskpane.css",
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
      setupMiddlewares(middlewares, devServer) {
        // Spawn a Python pipe process using the user's python (try common launchers)
        const cwd = process.cwd();
        const pipePath = path.join(cwd, "mhr_pipe.py");
        const pricePipePath = path.join(cwd, "price_pipe.py");
        const candidates = [process.env.PYTHON, process.env.python, "py", "python", "python3"].filter(Boolean);
        let py;
        let pyPrice;
        for (const exe of candidates) {
          try {
            py = spawn(exe, ["-u", pipePath], { stdio: ["pipe", "pipe", "inherit"], cwd });
            break;
          } catch (e) {
            // try next
          }
        }
        for (const exe of candidates) {
          try {
            pyPrice = spawn(exe, ["-u", pricePipePath], { stdio: ["pipe", "pipe", "inherit"], cwd });
            break;
          } catch (e) {
            // try next
          }
        }
        if (!py) {
          console.error("Failed to spawn any Python executable for mhr_pipe. Set PYTHON env var to your interpreter.");
        }
        if (!pyPrice) {
          console.error("Failed to spawn any Python executable for price_pipe. Set PYTHON env var to your interpreter.");
        }

        // JSONL request/response map
        let nextId = 1;
        const pending = new Map();
        if (py && py.stdout) {
          let buffer = "";
          py.stdout.on("data", (chunk) => {
            buffer += chunk.toString();
            let idx;
            while ((idx = buffer.indexOf("\n")) >= 0) {
              const line = buffer.slice(0, idx).trim();
              buffer = buffer.slice(idx + 1);
              if (!line) continue;
              try {
                const msg = JSON.parse(line);
                const id = msg && msg.id;
                const res = pending.get(id);
                if (res) {
                  pending.delete(id);
                  res.json(msg);
                }
              } catch {
                // ignore
              }
            }
          });
        }

        // JSONL request/response map for price
        let nextPriceId = 1;
        const pendingPrice = new Map();
        if (pyPrice && pyPrice.stdout) {
          let buffer = "";
          pyPrice.stdout.on("data", (chunk) => {
            buffer += chunk.toString();
            let idx;
            while ((idx = buffer.indexOf("\n")) >= 0) {
              const line = buffer.slice(0, idx).trim();
              buffer = buffer.slice(idx + 1);
              if (!line) continue;
              try {
                const msg = JSON.parse(line);
                const id = msg && msg.id;
                const res = pendingPrice.get(id);
                if (res) {
                  pendingPrice.delete(id);
                  res.json({ candidates: msg.candidates || [] });
                }
              } catch {
                // ignore
              }
            }
          });
        }

        devServer.app.post("/api/mhr", expressJson(), (req, res) => {
          const description = (req.body && req.body.description) || "";
          // Fix: pass limit
          const limit = (req.body && req.body.limit);

          if (!py || !py.stdin) {
            return res.status(500).json({ error: "Python pipe not running" });
          }
          const id = nextId++;
          pending.set(id, res);
          const payload = JSON.stringify({ id, description, limit }) + "\n";
          try {
            py.stdin.write(payload);
          } catch (e) {
            pending.delete(id);
            return res.status(500).json({ error: "Failed to write to Python pipe" });
          }
        });

        devServer.app.post("/api/mhr/learn", expressJson(), (req, res) => {
          const description = (req.body && req.body.description) || "";
          const value = (req.body && req.body.value);

          if (!py || !py.stdin) {
            return res.status(500).json({ error: "Python pipe not running" });
          }

          const id = nextId++;
          pending.set(id, res);
          // mode: 'learn'
          const payload = JSON.stringify({ id, mode: "learn", description, value }) + "\n";

          try {
            py.stdin.write(payload);
          } catch (e) {
            pending.delete(id);
            return res.status(500).json({ error: "Failed to write to Python pipe" });
          }
        });

        devServer.app.post("/api/price", expressJson(), (req, res) => {
          const description = (req.body && req.body.description) || "";
          const limit = (req.body && req.body.limit);

          if (!pyPrice || !pyPrice.stdin) {
            return res.status(500).json({ error: "Price Python pipe not running" });
          }
          const id = nextPriceId++;
          pendingPrice.set(id, res);
          const payload = JSON.stringify({ id, description, limit }) + "\n";
          try {
            pyPrice.stdin.write(payload);
          } catch (e) {
            pendingPrice.delete(id);
            return res.status(500).json({ error: "Failed to write to Price Python pipe" });
          }
        });

        devServer.app.get("/api/browse-folder", async (req, res) => {
          try {
            const { execSync } = require("child_process");
            const cmd = `powershell -Command "$app = New-Object -ComObject Shell.Application; $f = $app.BrowseForFolder(0, 'Select Folder', 0, 0); if ($f) { Write-Output $f.Self.Path }"`;
            const result = execSync(cmd, { encoding: "utf8" }).trim();
            res.json({ folderPath: result });
          } catch (err) {
            console.error("Error in /api/browse-folder:", err);
            res.status(500).json({ error: err.message });
          }
        });

        devServer.app.post("/api/reference-search", expressJson(), async (req, res) => {
          try {
            const { folderPath, searchTerms } = req.body;
            if (!folderPath || !searchTerms || !Array.isArray(searchTerms)) {
              return res.status(400).json({ error: "Missing folderPath or searchTerms (array)" });
            }

            const results = {}; // map of term -> [bestFilePath]
            const bestScores = {}; // map of term -> highest score
            const queue = [folderPath];
            const maxDepth = 5;
            let currentDepth = 0;
            const validSearchTerms = searchTerms.filter(t => t && t.toString().trim());

            if (validSearchTerms.length === 0) {
              return res.json({ results: {} });
            }

            // Scoring helper
            const diceCoefficient = (t, c) => {
              if (t === c) return 1;
              if (t.length < 2 || c.length < 2) return 0;
              let tBigrams = new Map();
              for (let i = 0; i < t.length - 1; i++) {
                const b = t.substring(i, i + 2);
                tBigrams.set(b, (tBigrams.get(b) || 0) + 1);
              }
              let inter = 0;
              for (let i = 0; i < c.length - 1; i++) {
                const b = c.substring(i, i + 2);
                const count = tBigrams.get(b) || 0;
                if (count > 0) {
                  tBigrams.set(b, count - 1);
                  inter++;
                }
              }
              return (2.0 * inter) / (t.length - 1 + c.length - 1);
            };

            const scoreMatch = (term, filename) => {
              const termLower = term.toLowerCase().trim();
              const nameLower = filename.toLowerCase();
              if (nameLower.includes(termLower)) return 100; // Exact substring is best

              const termTokens = termLower.replace(/[^a-z0-9]/g, ' ').split(/\s+/).filter(Boolean);
              const nameTokens = nameLower.replace(/[^a-z0-9]/g, ' ').split(/\s+/).filter(Boolean);

              if (termTokens.length === 0) return 0;

              let matchedTokens = 0;
              for (const tt of termTokens) {
                if (nameTokens.some(nt => nt.includes(tt) || tt.includes(nt))) {
                  matchedTokens++;
                } else {
                  for (const nt of nameTokens) {
                    if (nt.length > 2 && tt.length > 2 && diceCoefficient(tt, nt) > 0.6) {
                      matchedTokens += 0.8;
                      break;
                    }
                  }
                }
              }
              return (matchedTokens / termTokens.length) * 50;
            };

            // Simple BFS for directory traversal to avoid deep recursion issues
            while (queue.length > 0 && currentDepth <= maxDepth) {
              const levelSize = queue.length;
              for (let i = 0; i < levelSize; i++) {
                const currentDir = queue.shift();
                try {
                  const entries = await fs.readdir(currentDir, { withFileTypes: true });
                  for (const entry of entries) {
                    const fullPath = path.join(currentDir, entry.name);
                    if (entry.isDirectory()) {
                      queue.push(fullPath);
                    } else if (entry.isFile()) {
                      const nameLower = entry.name.toLowerCase();
                      if (!nameLower.endsWith('.pdf') && !nameLower.endsWith('.xlsx') && !nameLower.endsWith('.docx')) {
                        continue; // Skip files that are not pdf, xlsx, or docx
                      }

                      for (const term of validSearchTerms) {
                        const score = scoreMatch(term, entry.name);
                        // Require a minimum score of 15 (e.g. at least ~30% token match)
                        if (score >= 15) {
                          if (!bestScores[term] || score > bestScores[term]) {
                            bestScores[term] = score;
                            results[term] = [fullPath];
                          }
                        }
                      }
                    }
                  }
                } catch (err) {
                  console.error(`Error reading dir ${currentDir}:`, err.message);
                  continue; // Skip folders with permission issues, etc.
                }
              }
              currentDepth++;
            }

            res.json({ results });
          } catch (err) {
            console.error("Error in /api/reference-search:", err);
            res.status(500).json({ error: err.message });
          }
        });

        // ── Version / Update endpoints ──
        devServer.app.get("/api/version", (req, res) => {
          try {
            const ver = JSON.parse(fsSync.readFileSync(versionFile, "utf8"));
            res.json({ version: ver.version || "1.0.0" });
          } catch {
            res.json({ version: APP_VERSION });
          }
        });

        devServer.app.get("/api/check-update", (req, res) => {
          const ghUrl = "https://api.github.com/repos/ibilbisi123-ship-it/excel-system/releases/latest";
          const opts = {
            headers: {
              "User-Agent": "ExcelAddinAutoUpdater/1.0",
              Accept: "application/vnd.github.v3+json",
            },
          };
          https.get(ghUrl, opts, (ghRes) => {
            if ([301, 302, 307].includes(ghRes.statusCode) && ghRes.headers.location) {
              https.get(ghRes.headers.location, opts, (r2) => {
                let body = "";
                r2.on("data", (c) => (body += c));
                r2.on("end", () => handleGhResponse(body, res));
              });
              return;
            }
            let body = "";
            ghRes.on("data", (c) => (body += c));
            ghRes.on("end", () => handleGhResponse(body, res));
          }).on("error", (err) => {
            res.json({ updateAvailable: false, error: err.message });
          });

          function handleGhResponse(body, res) {
            try {
              const data = JSON.parse(body);
              const remote = (data.tag_name || "").replace(/^v/i, "");
              const local = APP_VERSION;
              const semCmp = (a, b) => {
                const pa = a.split(".").map(Number);
                const pb = b.split(".").map(Number);
                for (let i = 0; i < 3; i++) {
                  if ((pa[i] || 0) > (pb[i] || 0)) return 1;
                  if ((pa[i] || 0) < (pb[i] || 0)) return -1;
                }
                return 0;
              };
              res.json({
                updateAvailable: semCmp(remote, local) > 0,
                currentVersion: local,
                latestVersion: remote,
                releaseUrl: data.html_url || "",
                releaseNotes: data.body || "",
              });
            } catch {
              res.json({ updateAvailable: false, error: "Failed to parse GitHub response" });
            }
          }
        });

        return middlewares;
      },
    },
  };

  return config;
};

function expressJson() {
  // Tiny JSON parser middleware to avoid adding express dependency explicitly
  return (req, res, next) => {
    if (req.headers["content-type"] && req.headers["content-type"].includes("application/json")) {
      let body = "";
      req.on("data", (chunk) => (body += chunk));
      req.on("end", () => {
        try {
          req.body = body ? JSON.parse(body) : {};
        } catch {
          req.body = {};
        }
        next();
      });
    } else {
      next();
    }
  };
}
