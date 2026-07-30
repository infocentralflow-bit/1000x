@echo off
REM Expose the (password-protected) dashboard to the public internet via ngrok.
REM First, in a SEPARATE window, run:   python excel_bridge.py
REM (plain — no --lan needed; ngrok talks to 127.0.0.1 directly, so the bridge
REM  can stay on loopback-only, which is the smaller/safer surface).
setlocal

set "NGROK=%LOCALAPPDATA%\Microsoft\WinGet\Packages\Ngrok.Ngrok_Microsoft.Winget.Source_8wekyb3d8bbwe\ngrok.exe"
if not exist "%NGROK%" set "NGROK=ngrok"

echo.
echo If this is your first time: sign up free at https://dashboard.ngrok.com/signup,
echo copy your authtoken from https://dashboard.ngrok.com/get-started/your-authtoken,
echo then run once:
echo     "%NGROK%" config add-authtoken YOUR_TOKEN_HERE
echo.
echo Starting tunnel to port 8765...
echo Use the "https://...ngrok-free.app" URL it prints (not the http:// one).
echo You will be asked for the username/password printed by the bridge.
echo.

"%NGROK%" http 8765

endlocal
