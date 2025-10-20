@echo off
:: Check if Chrome is running with remote debugging on port 9222
tasklist /FI "IMAGENAME eq chrome.exe" /FO CSV | findstr /I "chrome.exe" > nul
if %errorlevel% NEQ 0 (
    :: Chrome is not running, start it with remote debugging
    start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222  --incognito --remote-allow-origins=http://localhost:9222
    echo "Chrome started with remote debugging on port 9222."
) else (
    echo "Chrome is already running with remote debugging on port 9222."
)