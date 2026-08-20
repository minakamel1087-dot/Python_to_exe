# Build the portable app folder and prove it works before deploying it.
#
#   .\build.ps1            build, self-check, and deploy beside the workbook
#   .\build.ps1 -NoDeploy  build and self-check only
#
# The output is a folder, not a single exe. Copy that folder anywhere and run
# it — nothing needs compiling again to deploy, and it starts faster than a
# one-file build, which unpacks itself to a temp folder on every launch.
#
# The self-check matters more than the build succeeding. PyInstaller cheerfully
# produces an app that dies on first use because a data file was left behind or
# a module was only imported lazily, and that surfaces when someone presses a
# button rather than at build time.

param([switch]$NoDeploy)

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

# PyInstaller and pip write progress to stderr, and under ErrorActionPreference
# 'Stop' PowerShell turns any stderr line from a native command into a
# terminating error. So native calls run with it relaxed and are judged on
# their exit code, which is the only thing that actually means failure.
function Invoke-Native {
    param([string]$What, [scriptblock]$Command)
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try { & $Command } finally { $ErrorActionPreference = $previous }
    if ($LASTEXITCODE -ne 0) { throw "$What failed (exit $LASTEXITCODE)" }
}

$AppName  = 'ClearanceWIRLogValidation'
$DeployTo = Split-Path $PSScriptRoot -Parent   # beside the workbook

Write-Host "== dependencies ==" -ForegroundColor Cyan
Invoke-Native 'pip install' { python -m pip install -q -r requirements.txt }
Invoke-Native 'pip install' { python -m pip install -q pyinstaller }

Write-Host "== tests ==" -ForegroundColor Cyan
Invoke-Native 'tests' { python -m pytest -q }

Write-Host "== build ==" -ForegroundColor Cyan
Invoke-Native 'PyInstaller' { python -m PyInstaller --clean --noconfirm log_validation.spec }

$built = Join-Path $PSScriptRoot "dist\$AppName"
$exe   = Join-Path $built "$AppName.exe"
if (-not (Test-Path $exe)) { throw "no exe at $exe" }
$mb = (Get-ChildItem $built -Recurse -File | Measure-Object Length -Sum).Sum / 1MB
Write-Host ("built dist\{0}\  ({1:N0} MB)" -f $AppName, $mb) -ForegroundColor Green

Write-Host "== self-check ==" -ForegroundColor Cyan
$log = Join-Path $built 'selftest.log'
if (Test-Path $log) { Remove-Item -LiteralPath $log -Force }
$run = Start-Process -FilePath $exe -ArgumentList '--selftest' -Wait -PassThru
if (Test-Path $log) { Get-Content $log -Encoding UTF8 | ForEach-Object { "  $_" } }
if ($run.ExitCode -ne 0) { throw "the built app failed its self-check (exit $($run.ExitCode))" }

if ($NoDeploy) {
    Write-Host ""
    Write-Host "Ready: dist\$AppName\" -ForegroundColor Green
    return
}

Write-Host "== deploy ==" -ForegroundColor Cyan
$target = Join-Path $DeployTo $AppName
if (Test-Path $target) { Remove-Item -LiteralPath $target -Recurse -Force }
Copy-Item $built $target -Recurse
Write-Host ""
Write-Host "Deployed to $target" -ForegroundColor Green
Write-Host "Run $AppName\$AppName.exe" -ForegroundColor Green
