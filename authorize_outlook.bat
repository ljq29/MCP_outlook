@echo off
setlocal EnableExtensions
chcp 65001 > nul

set "PROJECT_DIR=%~dp0"
if "%PROJECT_DIR:~-1%"=="\" set "PROJECT_DIR=%PROJECT_DIR:~0,-1%"
cd /d "%PROJECT_DIR%"

if not exist "%PROJECT_DIR%\logs" mkdir "%PROJECT_DIR%\logs"

for /f %%I in ('powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-Date -Format yyyyMMdd-HHmmss"') do set "AUTH_TS=%%I"

set "OUTLOOK_CONFIG_FILE=%PROJECT_DIR%\outlook-mcp.json"
set "AUTH_LOG_FILE=%PROJECT_DIR%\logs\outlook-auth-%AUTH_TS%.log"
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"

echo Outlook MCP Graph authorization
echo Project: %PROJECT_DIR%
echo Config : %OUTLOOK_CONFIG_FILE%
echo Log    : %AUTH_LOG_FILE%
echo.

uv run python -m outlook_mcp.auth_setup --debug --config-file "%OUTLOOK_CONFIG_FILE%" --log-file "%AUTH_LOG_FILE%"
set "AUTH_EXIT=%ERRORLEVEL%"

echo.
if "%AUTH_EXIT%"=="0" (
  echo Authorization completed successfully.
) else (
  echo Authorization failed. Check the log file:
  echo %AUTH_LOG_FILE%
)

exit /b %AUTH_EXIT%
