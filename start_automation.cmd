@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Khong tim thay moi truong Python tai .venv\Scripts\python.exe
  echo Hay chay lai buoc cai dat trong README.md.
  pause
  exit /b 1
)
start "OrderBridge DN2" /min ".venv\Scripts\python.exe" app.py
timeout /t 2 /nobreak >nul
set "EDGE=%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"
if not exist "%EDGE%" set "EDGE=%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"
if exist "%EDGE%" (
  start "" "%EDGE%" "http://127.0.0.1:8765"
) else (
  start "" "http://127.0.0.1:8765"
)
echo Dashboard dang chay tai http://127.0.0.1:8765
endlocal
