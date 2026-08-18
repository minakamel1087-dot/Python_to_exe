<#
.SYNOPSIS
    Compile a Python script into a Windows .exe.

.DESCRIPTION
    The whole tool lives here. Both GitHub Actions workflows call this same
    script with -NoVenv, so what runs in CI is what runs on your machine — there
    is no second copy of the logic to drift out of sync.

    Dependencies are resolved in this order:
      1. -Requirements <path>   use exactly that file
      2. -Requirements none     install nothing (stdlib-only script)
      3. (default)              nearest requirements.txt, walking up from the
                                script's folder to the repo root
      4. nothing found          pipreqs derives one from the script's imports

.EXAMPLE
    .\build-exe.ps1 -Script scripts\hello.py

.EXAMPLE
    .\build-exe.ps1 -Script scripts\gui.py -Console windowed -Icon assets\app.ico

.EXAMPLE
    .\build-exe.ps1 -Script scripts\tool.py -ExtraArgs '--add-data "assets;assets" --hidden-import pkg.mod'
#>

param(
    [Parameter(Mandatory = $true)][string]$Script,
    [string]$Requirements = "",
    [string]$Name = "",
    [ValidateSet("onefile", "onedir")][string]$Mode = "onefile",
    [ValidateSet("console", "windowed")][string]$Console = "console",
    [string]$Icon = "",
    [string]$ExtraArgs = "",
    [string]$SmokeArgs = "",
    [string]$OutDir = "",
    [switch]$NoVenv          # CI: use the Python already on PATH instead of a venv
)

$ErrorActionPreference = "Stop"

# PowerShell 7.4+ (what the GitHub runners use) turns any non-zero exit from a
# native command into a terminating error while ErrorActionPreference is Stop.
# That would kill the build on the first pip install this script deliberately
# tolerates — the per-package pipreqs retry below. Exit codes are checked
# explicitly everywhere it matters instead.
if (Get-Variable PSNativeCommandUseErrorActionPreference -ErrorAction SilentlyContinue) {
    $PSNativeCommandUseErrorActionPreference = $false
}

$repo = $PSScriptRoot
if (-not $OutDir) { $OutDir = Join-Path $repo "dist" }
$workDir = Join-Path $repo "build"

if (-not (Test-Path -LiteralPath $Script)) { throw "Script not found: $Script" }
$scriptPath = (Resolve-Path -LiteralPath $Script).Path
$scriptDir = Split-Path -Parent $scriptPath
$isSpec = $scriptPath.EndsWith(".spec")

if (-not $Name) {
    $stem = [IO.Path]::GetFileNameWithoutExtension($scriptPath)
    # scripts\mytool\main.py is "mytool", not "main"
    $Name = if ($stem -eq "main") { Split-Path -Leaf $scriptDir } else { $stem }
}

# --- Python ------------------------------------------------------------------
if ($NoVenv) {
    $python = (Get-Command python -ErrorAction Stop).Source
}
else {
    $venv = Join-Path $repo ".venv-build"
    $python = Join-Path $venv "Scripts\python.exe"
    if (-not (Test-Path $python)) {
        $launcher = if (Get-Command py -ErrorAction SilentlyContinue) { "py" }
                    elseif (Get-Command python -ErrorAction SilentlyContinue) { "python" }
                    else { $null }
        if (-not $launcher) {
            throw "Python not found. Install it from python.org (tick 'Add to PATH'), or push the script and let the GitHub Actions workflow build it — that needs nothing installed here."
        }
        Write-Host "Creating build virtualenv at $venv"
        if ($launcher -eq "py") { py -3 -m venv $venv } else { python -m venv $venv }
    }
}
Write-Host "Python:  $python"
Write-Host "Script:  $scriptPath"
Write-Host "Name:    $Name"

& $python -m pip install --upgrade pip | Out-Null
& $python -m pip install pyinstaller==6.11.1
if ($LASTEXITCODE -ne 0) { throw "Could not install PyInstaller" }

# --- requirements ------------------------------------------------------------
$reqs = ""
$derive = $false

if ($Requirements -eq "none") {
    Write-Host "Requirements: none (stdlib only)"
}
elseif ($Requirements) {
    if (-not (Test-Path -LiteralPath $Requirements)) { throw "requirements.txt not found: $Requirements" }
    $reqs = (Resolve-Path -LiteralPath $Requirements).Path
}
else {
    # Walk up from the script, not a repo-wide glob: scripts\mytool\main.py must
    # get scripts\mytool\requirements.txt, never a neighbouring tool's.
    $dir = $scriptDir
    while ($dir -and $dir.StartsWith($repo)) {
        $candidate = Join-Path $dir "requirements.txt"
        if (Test-Path -LiteralPath $candidate) { $reqs = $candidate; break }
        if ($dir -eq $repo) { break }
        $dir = Split-Path -Parent $dir
    }
    if (-not $reqs) { $derive = $true }
}

if ($reqs) {
    Write-Host "Requirements: $reqs"
    & $python -m pip install -r $reqs
    if ($LASTEXITCODE -ne 0) { throw "pip install -r $reqs failed" }
}
elseif ($derive) {
    Write-Host "Requirements: none found — deriving from the script's imports (pipreqs)"
    & $python -m pip install pipreqs
    $generated = Join-Path $repo "generated-requirements.txt"
    & $python -m pipreqs.pipreqs $scriptDir --savepath $generated --force --mode no-pin
    if (Test-Path $generated) {
        Write-Host "--- generated-requirements.txt ---"
        Get-Content $generated
        Write-Host "----------------------------------"
        # One package at a time: pipreqs infers package names from import names
        # and gets some wrong (a local module, or cv2 -> opencv-python). One bad
        # line shouldn't stop the rest from installing.
        foreach ($line in Get-Content $generated) {
            $pkg = $line.Trim()
            if ($pkg -and -not $pkg.StartsWith("#")) {
                & $python -m pip install $pkg
                if ($LASTEXITCODE -ne 0) { Write-Warning "Could not install '$pkg' — skipped. If the build fails on it, add a requirements.txt next to the script." }
            }
        }
    }
}

# --- build -------------------------------------------------------------------
if ($isSpec) {
    # A .spec carries its own name, mode, console setting and data files.
    $cmd = "& `"$python`" -m PyInstaller `"$scriptPath`" --noconfirm --clean --distpath `"$OutDir`" --workpath `"$workDir`""
}
else {
    $cmd = "& `"$python`" -m PyInstaller `"$scriptPath`" --noconfirm --clean " +
           "--distpath `"$OutDir`" --workpath `"$workDir`" " +
           "--specpath `"$workDir`" " +
           "--name `"$Name`" --$Mode --paths `"$scriptDir`""
    if ($Console -eq "windowed") { $cmd += " --windowed" }
    if ($Icon) {
        if (-not (Test-Path -LiteralPath $Icon)) { throw "Icon not found: $Icon" }
        $cmd += " --icon `"$((Resolve-Path -LiteralPath $Icon).Path)`""
    }
}
# A build.args file beside the script carries per-script flags (--windowed,
# --hidden-import, --add-data). Without it the drop folder could only ever
# produce default console builds, and anything with a GUI or a runtime import
# would need the manual workflow every time.
$argsFile = Join-Path $scriptDir "build.args"
if (Test-Path -LiteralPath $argsFile) {
    $fileArgs = (Get-Content $argsFile |
                 ForEach-Object { $_.Trim() } |
                 Where-Object { $_ -and -not $_.StartsWith("#") }) -join " "
    if ($fileArgs) {
        Write-Host "build.args: $fileArgs"
        $cmd += " $fileArgs"
    }
}

# Invoke-Expression rather than an argument array, so quoted extras like
# --add-data "assets;assets" stay a single argument.
if ($ExtraArgs) { $cmd += " $ExtraArgs" }

Write-Host "`n> $cmd`n"
Invoke-Expression $cmd
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed" }

$exe = Get-ChildItem $OutDir -Recurse -Filter *.exe | Sort-Object Length -Descending | Select-Object -First 1
if (-not $exe) { throw "No .exe produced under $OutDir" }

if ($SmokeArgs) {
    # Worth doing: a bundled exe usually breaks on a missing hidden import, and
    # that only shows when it actually runs — never at build time.
    Write-Host "`nSmoke test: $($exe.Name) $SmokeArgs"
    Invoke-Expression "& `"$($exe.FullName)`" $SmokeArgs"
    if ($LASTEXITCODE -ne 0) { throw "Smoke test failed with exit code $LASTEXITCODE" }
    Write-Host "Smoke test passed"
}

$mb = [math]::Round($exe.Length / 1MB, 1)
Write-Host "`nBuilt $($exe.FullName) ($mb MB)"
