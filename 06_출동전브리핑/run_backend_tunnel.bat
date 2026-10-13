@echo off
rem ons-brief backend + Cloudflare Quick Tunnel (trycloudflare)
rem Usage: run_backend_tunnel.bat [port]
chcp 65001 >nul
setlocal
cd /d "%~dp0"
set "PORT=%~1"
if "%PORT%"=="" set "PORT=8000"

where python >nul 2>nul || (echo [ERROR] python not found & pause & exit /b 1)

rem cloudflared: PATH -> local tools\ -> download
set "CF=cloudflared"
where cloudflared >nul 2>nul || set "CF=%~dp0tools\cloudflared.exe"
if not "%CF%"=="cloudflared" if not exist "%CF%" (
  echo [INFO] cloudflared not found, downloading...
  if not exist "%~dp0tools" mkdir "%~dp0tools"
  curl -L -o "%CF%" https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe || (echo [ERROR] download failed & pause & exit /b 1)
)

rem Build briefing data if missing (names/phones masked for public use)
if not exist "dist\briefing.json" (
  echo [INFO] building dist ...
  python prepare.py --mask-names || (echo [ERROR] prepare.py failed & pause & exit /b 1)
)

echo [INFO] backend: http://127.0.0.1:%PORT%  (new window)
start "ons-brief backend" cmd /k python server.py --port %PORT%
timeout /t 2 /nobreak >nul

echo.
echo [INFO] Starting tunnel. Copy the https://xxxx.trycloudflare.com URL below
echo        into the app's backend field, then press 연결 테스트.
echo        Ctrl+C to stop the tunnel (close the backend window separately).
echo.
"%CF%" tunnel --no-autoupdate --url http://127.0.0.1:%PORT%
endlocal
