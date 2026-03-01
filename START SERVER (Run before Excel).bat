@echo off
title Excel Add-in Dev Server
echo ============================================
echo  Excel Add-in Server - localhost:3000
echo  Keep this window open while using Excel!
echo ============================================
echo.
cd /d "%~dp0"
echo [Startup] Checking for updates...
node scripts/autoUpdate.js
echo.
npm run dev-server
pause
