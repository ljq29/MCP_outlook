@echo off
setlocal EnableExtensions

chcp 65001 > nul

set "ROOT=%~dp0"
set "PORT=%~1"
if "%PORT%"=="" set "PORT=8000"

set "HOST=%~2"
if "%HOST%"=="" set "HOST=127.0.0.1"

if not defined OUTLOOK_CONFIG_FILE set "OUTLOOK_CONFIG_FILE=%ROOT%outlook-mcp.json"
if not defined OUTLOOK_MCP_DB set "OUTLOOK_MCP_DB=%USERPROFILE%\.outlook-mcp\outlook-mcp.sqlite3"
if not defined OUTLOOK_MCP_METHOD set "OUTLOOK_MCP_METHOD=graph"
if not defined OUTLOOK_MCP_MODE set "OUTLOOK_MCP_MODE=full"
if not defined OUTLOOK_MCP_SCHEME set "OUTLOOK_MCP_SCHEME=http"
if not defined OUTLOOK_MCP_SERVER_NAME set "OUTLOOK_MCP_SERVER_NAME=outlook-http"

set "OUTLOOK_MCP_HTTP_PORT=%PORT%"
set "OUTLOOK_MCP_HTTP_HOST=%HOST%"
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"

if not defined OUTLOOK_MCP_ACCOUNT (
  echo OUTLOOK_MCP_ACCOUNT is required, for example you@outlook.com.
  exit /b 2
)
if not defined OUTLOOK_MCP_BEARER_TOKEN (
  echo OUTLOOK_MCP_BEARER_TOKEN is required.
  exit /b 2
)

echo Note: the first uv run will automatically install dependencies (uv sync).
echo Registering default debug credential: %OUTLOOK_MCP_METHOD%/%OUTLOOK_MCP_ACCOUNT%
uv run python -m outlook_mcp.credential_store register "%OUTLOOK_MCP_METHOD%" "%OUTLOOK_MCP_ACCOUNT%" "%OUTLOOK_MCP_BEARER_TOKEN%"
if errorlevel 1 exit /b %errorlevel%

echo.
echo Full MCP JSON:
if defined OUTLOOK_MCP_TOOLS (
  uv run python -m outlook_mcp.launcher_json --host "%HOST%" --port "%PORT%" --method "%OUTLOOK_MCP_METHOD%" --account "%OUTLOOK_MCP_ACCOUNT%" --bearer-token "%OUTLOOK_MCP_BEARER_TOKEN%" --mode "%OUTLOOK_MCP_MODE%" --tools "%OUTLOOK_MCP_TOOLS%" --scheme "%OUTLOOK_MCP_SCHEME%" --server-name "%OUTLOOK_MCP_SERVER_NAME%"
) else (
  uv run python -m outlook_mcp.launcher_json --host "%HOST%" --port "%PORT%" --method "%OUTLOOK_MCP_METHOD%" --account "%OUTLOOK_MCP_ACCOUNT%" --bearer-token "%OUTLOOK_MCP_BEARER_TOKEN%" --mode "%OUTLOOK_MCP_MODE%" --scheme "%OUTLOOK_MCP_SCHEME%" --server-name "%OUTLOOK_MCP_SERVER_NAME%"
)
if errorlevel 1 exit /b %errorlevel%

echo.
echo Service base URL: %OUTLOOK_MCP_SCHEME%://%HOST%:%PORT%
echo MCP URL:          %OUTLOOK_MCP_SCHEME%://%HOST%:%PORT%/%OUTLOOK_MCP_METHOD%/%OUTLOOK_MCP_ACCOUNT%
echo Auth/debug URL:   %OUTLOOK_MCP_SCHEME%://%HOST%:%PORT%/%OUTLOOK_MCP_METHOD%/%OUTLOOK_MCP_ACCOUNT%/%OUTLOOK_MCP_BEARER_TOKEN%
echo.
if not "%OUTLOOK_MCP_SKIP_PROMPT%"=="1" (
  set /p "_START=Copy the MCP JSON above, then press Enter to start Outlook MCP HTTP..."
)

uv run outlook-mcp-http
