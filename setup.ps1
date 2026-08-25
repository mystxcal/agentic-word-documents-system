[CmdletBinding()]
param(
    [string]$Python = "py"
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Venv = Join-Path $Root ".venv"

if (-not (Test-Path -LiteralPath (Join-Path $Venv "Scripts\python.exe"))) {
    $PythonName = [System.IO.Path]::GetFileNameWithoutExtension($Python)
    if ($PythonName -eq "py") {
        & $Python -3.12 -m venv $Venv
    } else {
        & $Python -m venv $Venv
    }
    if ($LASTEXITCODE -ne 0) { throw "Could not create the Python virtual environment." }
}

$VenvPython = Join-Path $Venv "Scripts\python.exe"
& $VenvPython -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) { throw "Could not upgrade pip in $Venv." }
& $VenvPython -m pip install -e $Root
if ($LASTEXITCODE -ne 0) { throw "Could not install the document system." }
& (Join-Path $Root "Document-System.cmd") doctor
if ($LASTEXITCODE -ne 0) { throw "Installation completed, but the runtime check failed. See the doctor output above." }
