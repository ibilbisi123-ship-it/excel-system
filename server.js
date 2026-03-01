/**
 * ExcelAddin Production Server
 *
 * Standalone HTTPS server that replaces webpack-dev-server for production.
 * Compiled into .exe using pkg — no Node.js or source code needed on user machines.
 *
 * Features:
 *   - HTTPS with auto-generated self-signed certs (trusted on first run)
 *   - Serves built web assets from dist/
 *   - All API endpoints (mhr, price, browse-folder, reference-search, version, check-update)
 *   - Python pipe management
 *   - Auto-update from GitHub releases on startup
 */

/* eslint-disable no-undef */
const httpsModule = require("https");
const http = require("http");
const fs = require("fs");
const path = require("path");
const url = require("url");
const { spawn, execSync } = require("child_process");
const os = require("os");

// ─── Determine root directory ────────────────────────────────────────
// When compiled with pkg: use the directory where the .exe lives
// When running as plain Node: use the script's directory
const ROOT_DIR = process.pkg
    ? path.dirname(process.execPath)
    : path.join(__dirname);

const DIST_DIR = path.join(ROOT_DIR, "dist");
const VERSION_FILE = path.join(ROOT_DIR, "version.json");
const CERT_DIR = path.join(ROOT_DIR, "certs");

// ─── GitHub Configuration ────────────────────────────────────────────
const GITHUB_OWNER = "ibilbisi123-ship-it";
const GITHUB_REPO = "excel-system";
const PORT = 3000;

// Files/folders preserved during updates
const PRESERVE = new Set([
    "node_modules", ".git", "version.json",
    "my2.db", "my22.db", "my_database.db",
    "__pycache__", ".vscode", "certs",
    "package-lock.json",
]);

// ─── Version helpers ─────────────────────────────────────────────────
function readVersion() {
    try { return JSON.parse(fs.readFileSync(VERSION_FILE, "utf8")).version || "0.0.0"; }
    catch { return "0.0.0"; }
}
function writeVersion(v) {
    fs.writeFileSync(VERSION_FILE, JSON.stringify({ version: v }, null, 2) + "\n", "utf8");
}
function semverCmp(a, b) {
    const pa = a.replace(/^v/i, "").split(".").map(Number);
    const pb = b.replace(/^v/i, "").split(".").map(Number);
    for (let i = 0; i < 3; i++) {
        if ((pa[i] || 0) > (pb[i] || 0)) return 1;
        if ((pa[i] || 0) < (pb[i] || 0)) return -1;
    }
    return 0;
}

// ─── HTTP(S) fetch with redirect following ───────────────────────────
function httpGet(targetUrl, headers = {}) {
    return new Promise((resolve, reject) => {
        const opts = { headers: { "User-Agent": "ExcelAddinServer/1.0", ...headers } };
        const handler = (res) => {
            if ([301, 302, 307, 308].includes(res.statusCode) && res.headers.location) {
                return httpGet(res.headers.location, headers).then(resolve).catch(reject);
            }
            if (res.statusCode < 200 || res.statusCode >= 300) {
                let body = "";
                res.on("data", (c) => (body += c));
                res.on("end", () => reject(new Error(`HTTP ${res.statusCode}: ${body.slice(0, 200)}`)));
                return;
            }
            const chunks = [];
            res.on("data", (c) => chunks.push(c));
            res.on("end", () => resolve(Buffer.concat(chunks)));
            res.on("error", reject);
        };
        const scheme = targetUrl.startsWith("https") ? httpsModule : http;
        scheme.get(targetUrl, opts, handler).on("error", reject);
    });
}

// ─── Auto-Update ─────────────────────────────────────────────────────
async function runAutoUpdate() {
    const local = readVersion();
    console.log("============================================");
    console.log("  Auto-Update Check");
    console.log(`  Current version: v${local}`);
    console.log("============================================\n");

    try {
        console.log("[Update] Checking for updates...");
        const buf = await httpGet(
            `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/releases/latest`,
            { Accept: "application/vnd.github.v3+json" }
        );
        const release = JSON.parse(buf.toString("utf8"));
        const remote = (release.tag_name || "").replace(/^v/i, "");
        if (!remote) { console.log("[Update] Could not determine remote version.\n"); return; }

        console.log(`[Update] Latest release: v${remote}`);
        if (semverCmp(remote, local) <= 0) { console.log("[Update] You are up to date!\n"); return; }

        console.log(`[Update] New version available! v${local} -> v${remote}\n`);

        // Find .zip asset
        const zipAsset = (release.assets || []).find(a => a.name && a.name.endsWith(".zip"));
        const dlUrl = zipAsset
            ? zipAsset.browser_download_url
            : release.zipball_url || `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/zipball/${release.tag_name}`;

        console.log("[Update] Downloading...");
        const tempDir = path.join(ROOT_DIR, ".update-temp");
        rmrf(tempDir);
        fs.mkdirSync(tempDir, { recursive: true });

        const zipPath = path.join(tempDir, "update.zip");
        const zipBuf = await httpGet(dlUrl, { Accept: "application/octet-stream" });
        fs.writeFileSync(zipPath, zipBuf);
        console.log("[Update] Download complete. Extracting...");

        // Extract using PowerShell
        const extractDir = path.join(tempDir, "extracted");
        fs.mkdirSync(extractDir, { recursive: true });
        execSync(
            `powershell -NoProfile -Command "Expand-Archive -Path '${zipPath}' -DestinationPath '${extractDir}' -Force"`,
            { stdio: "inherit" }
        );

        // Copy files (handle GitHub's top-level folder)
        let srcDir = extractDir;
        const items = fs.readdirSync(extractDir);
        if (items.length === 1 && fs.statSync(path.join(extractDir, items[0])).isDirectory()) {
            srcDir = path.join(extractDir, items[0]);
        }
        copyRecursive(srcDir, ROOT_DIR);
        writeVersion(remote);
        rmrf(tempDir);

        console.log(`\n  Update applied: v${local} -> v${remote}\n`);
    } catch (err) {
        if (err.message && err.message.includes("404")) {
            console.log("[Update] No releases found on GitHub yet.\n");
        } else {
            console.log(`[Update] Update check failed: ${err.message}\n`);
        }
    }
}

function copyRecursive(src, dest) {
    for (const entry of fs.readdirSync(src, { withFileTypes: true })) {
        if (PRESERVE.has(entry.name)) continue;
        const s = path.join(src, entry.name);
        const d = path.join(dest, entry.name);
        if (entry.isDirectory()) {
            if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true });
            copyRecursive(s, d);
        } else {
            fs.copyFileSync(s, d);
        }
    }
}
function rmrf(dir) {
    if (!fs.existsSync(dir)) return;
    try { fs.rmSync(dir, { recursive: true, force: true }); }
    catch { try { execSync(`rmdir /s /q "${dir}"`, { stdio: "ignore" }); } catch { } }
}

// ─── HTTPS Certificate Management ───────────────────────────────────
function ensureCerts() {
    const keyPath = path.join(CERT_DIR, "localhost.key");
    const certPath = path.join(CERT_DIR, "localhost.crt");

    // 1. Try existing certs from office-addin-dev-certs (already trusted)
    const devCertDir = path.join(os.homedir(), ".office-addin-dev-certs");
    const devKey = path.join(devCertDir, "localhost.key");
    const devCert = path.join(devCertDir, "localhost.crt");
    const devCa = path.join(devCertDir, "ca.crt");

    if (fs.existsSync(devKey) && fs.existsSync(devCert)) {
        console.log("[Certs] Using existing Office dev certificates.");
        const opts = { key: fs.readFileSync(devKey), cert: fs.readFileSync(devCert) };
        if (fs.existsSync(devCa)) opts.ca = fs.readFileSync(devCa);
        return opts;
    }

    // 2. Try local certs directory
    if (fs.existsSync(keyPath) && fs.existsSync(certPath)) {
        console.log("[Certs] Using existing local certificates.");
        return { key: fs.readFileSync(keyPath), cert: fs.readFileSync(certPath) };
    }

    // 3. Generate self-signed certs
    console.log("[Certs] Generating self-signed HTTPS certificates for localhost...");
    try {
        const selfsigned = require("selfsigned");
        const attrs = [{ name: "commonName", value: "localhost" }];
        const pems = selfsigned.generate(attrs, {
            algorithm: "sha256",
            days: 3650,
            keySize: 2048,
            extensions: [
                {
                    name: "subjectAltName", altNames: [
                        { type: 2, value: "localhost" },
                        { type: 7, ip: "127.0.0.1" },
                    ]
                },
            ],
        });

        fs.mkdirSync(CERT_DIR, { recursive: true });
        fs.writeFileSync(keyPath, pems.private);
        fs.writeFileSync(certPath, pems.cert);

        // Try to trust the cert on Windows
        try {
            execSync(`certutil -addstore -user -f "Root" "${certPath}"`, { stdio: "ignore" });
            console.log("[Certs] Certificate added to trusted store.");
        } catch {
            console.log("[Certs] Could not auto-trust cert. You may need to trust it manually.");
        }

        return { key: pems.private, cert: pems.cert };
    } catch (e) {
        console.error("[Certs] Failed to generate certificates:", e.message);
        process.exit(1);
    }
}

// ─── Static File Server ─────────────────────────────────────────────
const MIME_TYPES = {
    ".html": "text/html", ".js": "application/javascript",
    ".css": "text/css", ".json": "application/json",
    ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
    ".gif": "image/gif", ".ico": "image/x-icon",
    ".svg": "image/svg+xml", ".xml": "application/xml",
    ".woff": "font/woff", ".woff2": "font/woff2",
    ".map": "application/json", ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
};

function serveStatic(req, res) {
    let pathname = url.parse(req.url).pathname;
    if (pathname === "/") pathname = "/taskpane.html";

    const filePath = path.join(DIST_DIR, pathname);

    // Security: prevent path traversal
    if (!filePath.startsWith(DIST_DIR)) {
        res.writeHead(403); res.end("Forbidden"); return;
    }

    try {
        const data = fs.readFileSync(filePath);
        const ext = path.extname(filePath).toLowerCase();
        res.writeHead(200, {
            "Content-Type": MIME_TYPES[ext] || "application/octet-stream",
            "Access-Control-Allow-Origin": "*",
        });
        res.end(data);
    } catch {
        res.writeHead(404);
        res.end("Not found");
    }
}

// ─── JSON Body Parser ───────────────────────────────────────────────
function parseJsonBody(req) {
    return new Promise((resolve) => {
        if (!req.headers["content-type"] || !req.headers["content-type"].includes("application/json")) {
            return resolve({});
        }
        let body = "";
        req.on("data", (c) => (body += c));
        req.on("end", () => {
            try { resolve(body ? JSON.parse(body) : {}); }
            catch { resolve({}); }
        });
    });
}

// ─── Python Pipe Management ─────────────────────────────────────────
function spawnPython(scriptName) {
    const scriptPath = path.join(ROOT_DIR, scriptName);
    const candidates = [process.env.PYTHON, process.env.python, "py", "python", "python3"].filter(Boolean);
    for (const exe of candidates) {
        try {
            const proc = spawn(exe, ["-u", scriptPath], { stdio: ["pipe", "pipe", "inherit"], cwd: ROOT_DIR });
            return proc;
        } catch { /* try next */ }
    }
    console.error(`Failed to spawn Python for ${scriptName}. Set PYTHON env var.`);
    return null;
}

function createPipeHandler(proc) {
    let nextId = 1;
    const pending = new Map();

    if (proc && proc.stdout) {
        let buffer = "";
        proc.stdout.on("data", (chunk) => {
            buffer += chunk.toString();
            let idx;
            while ((idx = buffer.indexOf("\n")) >= 0) {
                const line = buffer.slice(0, idx).trim();
                buffer = buffer.slice(idx + 1);
                if (!line) continue;
                try {
                    const msg = JSON.parse(line);
                    const id = msg && msg.id;
                    const cb = pending.get(id);
                    if (cb) { pending.delete(id); cb(msg); }
                } catch { /* ignore */ }
            }
        });
    }

    return {
        send(payload) {
            return new Promise((resolve, reject) => {
                if (!proc || !proc.stdin) return reject(new Error("Python pipe not running"));
                const id = nextId++;
                pending.set(id, resolve);
                try {
                    proc.stdin.write(JSON.stringify({ id, ...payload }) + "\n");
                } catch (e) {
                    pending.delete(id);
                    reject(e);
                }
            });
        },
    };
}

// ─── Reference Search (Dice Coefficient) ────────────────────────────
function diceCoefficient(t, c) {
    if (t === c) return 1;
    if (t.length < 2 || c.length < 2) return 0;
    const tBigrams = new Map();
    for (let i = 0; i < t.length - 1; i++) {
        const b = t.substring(i, i + 2);
        tBigrams.set(b, (tBigrams.get(b) || 0) + 1);
    }
    let inter = 0;
    for (let i = 0; i < c.length - 1; i++) {
        const b = c.substring(i, i + 2);
        const count = tBigrams.get(b) || 0;
        if (count > 0) { tBigrams.set(b, count - 1); inter++; }
    }
    return (2.0 * inter) / (t.length - 1 + c.length - 1);
}

function scoreMatch(term, filename) {
    const tl = term.toLowerCase().trim();
    const nl = filename.toLowerCase();
    if (nl.includes(tl)) return 100;
    const tt = tl.replace(/[^a-z0-9]/g, " ").split(/\s+/).filter(Boolean);
    const nt = nl.replace(/[^a-z0-9]/g, " ").split(/\s+/).filter(Boolean);
    if (!tt.length) return 0;
    let matched = 0;
    for (const tok of tt) {
        if (nt.some(n => n.includes(tok) || tok.includes(n))) { matched++; }
        else {
            for (const n of nt) {
                if (n.length > 2 && tok.length > 2 && diceCoefficient(tok, n) > 0.6) { matched += 0.8; break; }
            }
        }
    }
    return (matched / tt.length) * 50;
}

async function referenceSearch(folderPath, searchTerms) {
    const results = {};
    const bestScores = {};
    const fsPromises = require("fs").promises;
    const valid = searchTerms.filter(t => t && t.toString().trim());
    if (!valid.length) return {};

    const queue = [folderPath];
    let depth = 0;
    while (queue.length > 0 && depth <= 5) {
        const len = queue.length;
        for (let i = 0; i < len; i++) {
            const dir = queue.shift();
            try {
                const entries = await fsPromises.readdir(dir, { withFileTypes: true });
                for (const entry of entries) {
                    const full = path.join(dir, entry.name);
                    if (entry.isDirectory()) { queue.push(full); }
                    else if (entry.isFile()) {
                        const nm = entry.name.toLowerCase();
                        if (!nm.endsWith(".pdf") && !nm.endsWith(".xlsx") && !nm.endsWith(".docx")) continue;
                        for (const term of valid) {
                            const score = scoreMatch(term, entry.name);
                            if (score >= 15 && (!bestScores[term] || score > bestScores[term])) {
                                bestScores[term] = score;
                                results[term] = [full];
                            }
                        }
                    }
                }
            } catch { continue; }
        }
        depth++;
    }
    return results;
}

// ─── Request Router ─────────────────────────────────────────────────
function createRouter(mhrPipe, pricePipe) {
    const APP_VERSION = readVersion();

    return async function handleRequest(req, res) {
        const parsed = url.parse(req.url, true);
        const pathname = parsed.pathname;

        // CORS headers
        res.setHeader("Access-Control-Allow-Origin", "*");
        res.setHeader("Access-Control-Allow-Headers", "Content-Type");
        if (req.method === "OPTIONS") { res.writeHead(204); res.end(); return; }

        // ── API Routes ──
        if (pathname === "/api/mhr" && req.method === "POST") {
            const body = await parseJsonBody(req);
            try {
                const result = await mhrPipe.send({ description: body.description || "", limit: body.limit });
                sendJson(res, result);
            } catch (e) { sendJson(res, { error: e.message }, 500); }
            return;
        }

        if (pathname === "/api/mhr/learn" && req.method === "POST") {
            const body = await parseJsonBody(req);
            try {
                const result = await mhrPipe.send({ mode: "learn", description: body.description || "", value: body.value });
                sendJson(res, result);
            } catch (e) { sendJson(res, { error: e.message }, 500); }
            return;
        }

        if (pathname === "/api/price" && req.method === "POST") {
            const body = await parseJsonBody(req);
            try {
                const result = await pricePipe.send({ description: body.description || "", limit: body.limit });
                sendJson(res, { candidates: result.candidates || [] });
            } catch (e) { sendJson(res, { error: e.message }, 500); }
            return;
        }

        if (pathname === "/api/browse-folder" && req.method === "GET") {
            try {
                const cmd = `powershell -Command "$app = New-Object -ComObject Shell.Application; $f = $app.BrowseForFolder(0, 'Select Folder', 0, 0); if ($f) { Write-Output $f.Self.Path }"`;
                const result = execSync(cmd, { encoding: "utf8" }).trim();
                sendJson(res, { folderPath: result });
            } catch (e) { sendJson(res, { error: e.message }, 500); }
            return;
        }

        if (pathname === "/api/reference-search" && req.method === "POST") {
            const body = await parseJsonBody(req);
            if (!body.folderPath || !body.searchTerms || !Array.isArray(body.searchTerms)) {
                sendJson(res, { error: "Missing folderPath or searchTerms" }, 400);
                return;
            }
            try {
                const results = await referenceSearch(body.folderPath, body.searchTerms);
                sendJson(res, { results });
            } catch (e) { sendJson(res, { error: e.message }, 500); }
            return;
        }

        if (pathname === "/api/version" && req.method === "GET") {
            sendJson(res, { version: APP_VERSION });
            return;
        }

        if (pathname === "/api/check-update" && req.method === "GET") {
            try {
                const buf = await httpGet(
                    `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/releases/latest`,
                    { Accept: "application/vnd.github.v3+json" }
                );
                const data = JSON.parse(buf.toString("utf8"));
                const remote = (data.tag_name || "").replace(/^v/i, "");
                sendJson(res, {
                    updateAvailable: semverCmp(remote, APP_VERSION) > 0,
                    currentVersion: APP_VERSION,
                    latestVersion: remote,
                    releaseUrl: data.html_url || "",
                    releaseNotes: data.body || "",
                });
            } catch (e) {
                sendJson(res, { updateAvailable: false, error: e.message });
            }
            return;
        }

        // ── Static files ──
        serveStatic(req, res);
    };
}

function sendJson(res, data, status = 200) {
    res.writeHead(status, { "Content-Type": "application/json" });
    res.end(JSON.stringify(data));
}

// ─── Main ────────────────────────────────────────────────────────────
async function main() {
    console.log("\n  Excel Add-in Production Server");
    console.log("  ================================\n");

    // 1. Auto-update
    await runAutoUpdate();

    // 2. Spawn Python pipes
    const pyMhr = spawnPython("mhr_pipe.py");
    const pyPrice = spawnPython("price_pipe.py");
    const mhrPipe = createPipeHandler(pyMhr);
    const pricePipe = createPipeHandler(pyPrice);

    // 3. HTTPS certs
    const httpsOpts = ensureCerts();

    // 4. Start HTTPS server
    const handler = createRouter(mhrPipe, pricePipe);
    const server = httpsModule.createServer(httpsOpts, handler);

    server.listen(PORT, () => {
        console.log("============================================");
        console.log(`  Server running: https://localhost:${PORT}/`);
        console.log("  Keep this window open while using Excel!");
        console.log("============================================\n");
    });

    // Graceful shutdown
    process.on("SIGINT", () => {
        console.log("\nShutting down...");
        if (pyMhr) try { pyMhr.kill(); } catch { }
        if (pyPrice) try { pyPrice.kill(); } catch { }
        server.close();
        process.exit(0);
    });
}

main().catch((err) => {
    console.error("Fatal error:", err);
    process.exit(1);
});
