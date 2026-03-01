/**
 * Build Release Script
 *
 * Assembles a clean `release/` folder containing ONLY the files
 * that users need — no source code, no build scripts, no node_modules.
 *
 * Usage:  node scripts/build-release.js
 *
 * What it does:
 *   1. Runs webpack production build      → dist/
 *   2. Compiles server.js into .exe       → ExcelAddinServer.exe
 *   3. Copies everything into release/    → ready to distribute
 *   4. Creates a release.zip              → upload to GitHub Releases
 */

const fs = require("fs");
const path = require("path");
const { execSync } = require("child_process");

const ROOT = path.join(__dirname, "..");
const RELEASE_DIR = path.join(ROOT, "release");
const ZIP_NAME = "excel-addin-release.zip";

// Files/folders to copy into release/
const COPY_FILES = [
    "mhr_pipe.py",
    "price_pipe.py",
    "mhrsqlsystem.py",
    "pricesqlsystem.py",
    "manifest.xml",
    "version.json",
];

const COPY_DIRS = [
    "assets",
];

// Database files to include as empty templates (users bring their own data)
const DB_FILES = [
    "my_database.db",
    "my2.db",
    "my22.db",
];

// ─── Helpers ─────────────────────────────────────────────────────────
function rmrf(dir) {
    if (fs.existsSync(dir)) {
        fs.rmSync(dir, { recursive: true, force: true });
    }
}

function copyDir(src, dest) {
    fs.mkdirSync(dest, { recursive: true });
    for (const entry of fs.readdirSync(src, { withFileTypes: true })) {
        const s = path.join(src, entry.name);
        const d = path.join(dest, entry.name);
        if (entry.isDirectory()) {
            copyDir(s, d);
        } else {
            fs.copyFileSync(s, d);
        }
    }
}

function run(cmd, label) {
    console.log(`\n[Build] ${label}...`);
    execSync(cmd, { cwd: ROOT, stdio: "inherit" });
}

// ─── Main ────────────────────────────────────────────────────────────
function main() {
    console.log("============================================");
    console.log("  Building Release Package");
    console.log("============================================");

    // 1. Clean release folder
    console.log("\n[Build] Cleaning release/ folder...");
    rmrf(RELEASE_DIR);
    fs.mkdirSync(RELEASE_DIR, { recursive: true });

    // 2. Webpack production build
    run("npx webpack --mode production", "Building webpack (production)");

    // 3. Compile .exe with pkg
    const exePath = path.join(RELEASE_DIR, "ExcelAddinServer.exe");
    run(
        `npx @yao-pkg/pkg server.js --targets node20-win-x64 --output "${exePath}"`,
        "Compiling server.js to .exe"
    );

    // 4. Copy dist/ folder
    console.log("\n[Build] Copying dist/ folder...");
    copyDir(path.join(ROOT, "dist"), path.join(RELEASE_DIR, "dist"));

    // 5. Copy individual files
    console.log("[Build] Copying project files...");
    for (const file of COPY_FILES) {
        const src = path.join(ROOT, file);
        if (fs.existsSync(src)) {
            fs.copyFileSync(src, path.join(RELEASE_DIR, file));
        }
    }

    // 6. Copy directories
    for (const dir of COPY_DIRS) {
        const src = path.join(ROOT, dir);
        if (fs.existsSync(src)) {
            copyDir(src, path.join(RELEASE_DIR, dir));
        }
    }

    // 7. Copy database files
    console.log("[Build] Copying database files...");
    for (const db of DB_FILES) {
        const src = path.join(ROOT, db);
        if (fs.existsSync(src)) {
            fs.copyFileSync(src, path.join(RELEASE_DIR, db));
        }
    }

    // 8. Create a simple launcher batch file for users
    console.log("[Build] Creating launcher...");
    const launcherContent = `@echo off
title Excel Add-in Server
echo ============================================
echo  Excel Add-in Server - localhost:3000
echo  Keep this window open while using Excel!
echo ============================================
echo.
cd /d "%~dp0"
ExcelAddinServer.exe
pause
`;
    fs.writeFileSync(path.join(RELEASE_DIR, "START SERVER.bat"), launcherContent);

    // 9. Create a README for users
    const readmeContent = `# Excel AI Assistant

## Quick Start
1. Double-click "START SERVER.bat" to launch the server
2. Keep the server window open
3. Open Excel and use the add-in

## Requirements
- Python 3.x installed and on PATH (for MHR/Price calculators)
- Excel with the add-in manifest loaded

## Files
- ExcelAddinServer.exe  - The server (don't delete!)
- dist/                 - Web interface files
- manifest.xml          - Office add-in manifest
- *.py                  - Python calculation scripts
- *.db                  - Database files
- version.json          - Current version (auto-updated)

## Auto-Update
The server checks for updates on startup automatically.
`;
    fs.writeFileSync(path.join(RELEASE_DIR, "README.txt"), readmeContent);

    // 10. Create a .zip using PowerShell
    console.log("\n[Build] Creating release zip...");
    const zipPath = path.join(ROOT, ZIP_NAME);
    if (fs.existsSync(zipPath)) fs.unlinkSync(zipPath);
    try {
        execSync(
            `powershell -NoProfile -Command "Compress-Archive -Path '${RELEASE_DIR}\\*' -DestinationPath '${zipPath}' -Force"`,
            { stdio: "inherit" }
        );
        console.log(`[Build] Created: ${ZIP_NAME}`);
    } catch {
        console.log("[Build] Warning: Could not create zip. You can zip release/ manually.");
    }

    // Summary
    const files = fs.readdirSync(RELEASE_DIR);
    console.log("\n============================================");
    console.log("  Release package ready!");
    console.log("============================================");
    console.log(`\n  Folder: release/`);
    console.log(`  Files:  ${files.length} items\n`);
    files.forEach((f) => {
        const stat = fs.statSync(path.join(RELEASE_DIR, f));
        const size = stat.isDirectory()
            ? "(folder)"
            : `${(stat.size / 1024 / 1024).toFixed(1)} MB`;
        console.log(`    ${f.padEnd(30)} ${size}`);
    });
    console.log(`\n  Zip:    ${ZIP_NAME}`);
    console.log("\n  Upload the zip to GitHub Releases!");
    console.log("  Or distribute the release/ folder directly.\n");
}

main();
