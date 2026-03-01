/**
 * Auto-Update Script
 *
 * Runs before the dev server starts. Checks the GitHub Releases API for a
 * newer version and, if found, downloads the release .zip asset and replaces
 * local project files automatically.
 *
 * Usage:  node scripts/autoUpdate.js
 *
 * GitHub Repo: ibilbisi123-ship-it/excel-system
 */

/* eslint-disable no-undef */
const https = require("https");
const http = require("http");
const fs = require("fs");
const path = require("path");
const { execSync } = require("child_process");

// ─── Configuration ───────────────────────────────────────────────────
const GITHUB_OWNER = "ibilbisi123-ship-it";
const GITHUB_REPO = "excel-system";
const VERSION_FILE = path.join(__dirname, "..", "version.json");

// Files/folders that should NEVER be replaced by an update
const PRESERVE = new Set([
    "node_modules",
    ".git",
    "version.json",      // we update this ourselves after success
    "my2.db",
    "my22.db",
    "my_database.db",
    "__pycache__",
    ".vscode",
    "dist",
    "package-lock.json",  // keep the user's lockfile
]);

// ─── Helpers ─────────────────────────────────────────────────────────

/** Read the current local version from version.json */
function readLocalVersion() {
    try {
        const data = JSON.parse(fs.readFileSync(VERSION_FILE, "utf8"));
        return data.version || "0.0.0";
    } catch {
        return "0.0.0";
    }
}

/** Write a new version string to version.json */
function writeLocalVersion(ver) {
    fs.writeFileSync(VERSION_FILE, JSON.stringify({ version: ver }, null, 2) + "\n", "utf8");
}

/**
 * Simple semver compare: returns 1 if a > b, -1 if a < b, 0 if equal.
 * Supports tags like "v1.2.3" or "1.2.3".
 */
function compareSemver(a, b) {
    const parse = (v) => v.replace(/^v/i, "").split(".").map(Number);
    const pa = parse(a);
    const pb = parse(b);
    for (let i = 0; i < 3; i++) {
        const na = pa[i] || 0;
        const nb = pb[i] || 0;
        if (na > nb) return 1;
        if (na < nb) return -1;
    }
    return 0;
}

/** HTTPS GET that follows redirects and returns a Buffer. */
function httpGet(url, headers = {}) {
    return new Promise((resolve, reject) => {
        const opts = {
            headers: {
                "User-Agent": "ExcelAddinAutoUpdater/1.0",
                ...headers,
            },
        };

        const handler = (res) => {
            // Follow redirects (GitHub sends 302 for asset downloads)
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
            res.on("data", (chunk) => chunks.push(chunk));
            res.on("end", () => resolve(Buffer.concat(chunks)));
            res.on("error", reject);
        };

        const scheme = url.startsWith("https") ? https : http;
        scheme.get(url, opts, handler).on("error", reject);
    });
}

/** Fetch the latest release info from GitHub. */
async function fetchLatestRelease() {
    const url = `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/releases/latest`;
    const buf = await httpGet(url, { Accept: "application/vnd.github.v3+json" });
    return JSON.parse(buf.toString("utf8"));
}

/**
 * Download a file from a URL to a local path.
 * Shows a simple progress indicator in the console.
 */
async function downloadFile(url, destPath) {
    const buf = await httpGet(url, { Accept: "application/octet-stream" });
    fs.writeFileSync(destPath, buf);
    return destPath;
}

/**
 * Extract a .zip to a destination directory using PowerShell
 * (available on all modern Windows systems).
 */
function extractZip(zipPath, destDir) {
    // Ensure destination exists
    if (!fs.existsSync(destDir)) {
        fs.mkdirSync(destDir, { recursive: true });
    }
    // Use PowerShell's Expand-Archive
    const cmd = `powershell -NoProfile -Command "Expand-Archive -Path '${zipPath}' -DestinationPath '${destDir}' -Force"`;
    execSync(cmd, { stdio: "inherit" });
}

/**
 * Copy files from src directory to dest directory, skipping PRESERVE entries.
 * This handles the case where the zip might contain a top-level folder.
 */
function copyUpdatedFiles(srcDir, destDir) {
    // Check if the zip extracted into a single top-level folder
    const entries = fs.readdirSync(srcDir);
    let actualSrc = srcDir;
    if (entries.length === 1) {
        const single = path.join(srcDir, entries[0]);
        if (fs.statSync(single).isDirectory()) {
            actualSrc = single; // GitHub zips usually have a top-level folder like "repo-main"
        }
    }

    copyRecursive(actualSrc, destDir);
}

function copyRecursive(src, dest) {
    const entries = fs.readdirSync(src, { withFileTypes: true });
    for (const entry of entries) {
        if (PRESERVE.has(entry.name)) {
            continue; // Skip preserved files/folders
        }

        const srcPath = path.join(src, entry.name);
        const destPath = path.join(dest, entry.name);

        if (entry.isDirectory()) {
            if (!fs.existsSync(destPath)) {
                fs.mkdirSync(destPath, { recursive: true });
            }
            copyRecursive(srcPath, destPath);
        } else {
            fs.copyFileSync(srcPath, destPath);
        }
    }
}

/** Remove a directory recursively (like rm -rf). */
function rmrf(dir) {
    if (!fs.existsSync(dir)) return;
    try {
        fs.rmSync(dir, { recursive: true, force: true });
    } catch {
        // Fallback for older Node versions
        try {
            execSync(`rmdir /s /q "${dir}"`, { stdio: "ignore" });
        } catch {
            // ignore
        }
    }
}

// ─── Main ────────────────────────────────────────────────────────────

async function main() {
    const projectRoot = path.join(__dirname, "..");
    const localVersion = readLocalVersion();

    console.log("============================================");
    console.log("  Auto-Update Check");
    console.log(`  Current version: v${localVersion}`);
    console.log("============================================");
    console.log();

    try {
        // 1. Fetch the latest release from GitHub
        console.log("[Update] Checking for updates...");
        const release = await fetchLatestRelease();
        const remoteVersion = (release.tag_name || "").replace(/^v/i, "");

        if (!remoteVersion) {
            console.log("[Update] Could not determine remote version. Skipping.");
            return;
        }

        console.log(`[Update] Latest release: v${remoteVersion}`);

        // 2. Compare versions
        if (compareSemver(remoteVersion, localVersion) <= 0) {
            console.log("[Update] You are up to date!\n");
            return;
        }

        console.log(`[Update] New version available! v${localVersion} -> v${remoteVersion}`);
        console.log();

        // 3. Find the .zip asset in the release
        const zipAsset = (release.assets || []).find(
            (a) => a.name && a.name.endsWith(".zip")
        );

        let downloadUrl;
        if (zipAsset) {
            // Use the browser_download_url for the attached .zip asset
            downloadUrl = zipAsset.browser_download_url;
            console.log(`[Update] Downloading release asset: ${zipAsset.name}`);
        } else {
            // Fallback: download the source zipball from GitHub
            downloadUrl = release.zipball_url || `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/zipball/${release.tag_name}`;
            console.log("[Update] No .zip asset found — downloading source archive...");
        }

        // 4. Download to a temp file
        const tempDir = path.join(projectRoot, ".update-temp");
        rmrf(tempDir);
        fs.mkdirSync(tempDir, { recursive: true });

        const zipPath = path.join(tempDir, "update.zip");
        console.log("[Update] Downloading...");
        await downloadFile(downloadUrl, zipPath);
        console.log("[Update] Download complete.");

        // 5. Extract the zip
        const extractDir = path.join(tempDir, "extracted");
        console.log("[Update] Extracting...");
        extractZip(zipPath, extractDir);
        console.log("[Update] Extraction complete.");

        // 6. Copy updated files to the project root
        console.log("[Update] Applying update...");
        copyUpdatedFiles(extractDir, projectRoot);

        // 7. Update local version
        writeLocalVersion(remoteVersion);

        // 8. Clean up temp files
        rmrf(tempDir);

        console.log();
        console.log("============================================");
        console.log(`  Update applied: v${localVersion} -> v${remoteVersion}`);
        console.log("============================================");
        console.log();

        // 9. Check if package.json changed — need npm install
        console.log("[Update] Running npm install to sync dependencies...");
        try {
            execSync("npm install --production=false", {
                cwd: projectRoot,
                stdio: "inherit",
            });
        } catch {
            console.log("[Update] npm install had issues — you may need to run it manually.");
        }

        console.log("[Update] Update complete! Starting server...\n");
    } catch (err) {
        // Non-fatal: if update check fails, just start the server anyway
        if (err.message && err.message.includes("404")) {
            console.log("[Update] No releases found on GitHub yet. Skipping update check.\n");
        } else {
            console.log(`[Update] Update check failed: ${err.message}`);
            console.log("[Update] Continuing with current version...\n");
        }
    }
}

main();
