@echo off
setlocal EnableExtensions EnableDelayedExpansion
set "ROOT=%~dp0"
set "AGENTIC_DOCS_ROOT=%ROOT%"
set "PYTHONPATH=%ROOT%src"

if defined AGENTIC_DOCS_PYTHON (
  set "PYTHON_EXE=%AGENTIC_DOCS_PYTHON%"
  goto launch
)

if exist "%ROOT%.venv\Scripts\python.exe" (
  set "PYTHON_EXE=%ROOT%.venv\Scripts\python.exe"
  goto launch
)

for %%P in (py.exe python.exe) do (
  where %%P >nul 2>nul
  if not errorlevel 1 (
    set "PYTHON_EXE=%%P"
    goto launch
  )
)

echo ERROR: Python 3.12 or newer was not found. Run setup.ps1 or set AGENTIC_DOCS_PYTHON. 1>&2
exit /b 2

:launch
"%PYTHON_EXE%" -c "import lxml, openpyxl, PIL, pypdf, docx, pydantic" >nul 2>nul
if errorlevel 1 (
  echo ERROR: Required Python packages are missing. Run setup.ps1. 1>&2
  exit /b 2
)
"%PYTHON_EXE%" -m agentic_docs.cli %*
exit /b !errorlevel!
