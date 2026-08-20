$ErrorActionPreference = "Stop"
$finder = Join-Path $PSScriptRoot "find-build-targets.ps1"
$fixture = Join-Path ([IO.Path]::GetTempPath()) ("py-to-exe-discovery-" + [guid]::NewGuid())

try {
    New-Item -ItemType Directory -Path "$fixture/scripts/flat", "$fixture/scripts/custom", "$fixture/scripts/spec" | Out-Null
    Set-Content "$fixture/scripts/top.py" "print('top')"
    Set-Content "$fixture/scripts/flat/main.py" "print('main')"
    Set-Content "$fixture/scripts/custom/z_entry.py" "print('z')"
    Set-Content "$fixture/scripts/custom/a_entry.py" "print('a')"
    Set-Content "$fixture/scripts/spec/main.py" "print('ignored')"
    Set-Content "$fixture/scripts/spec/app.spec" "# spec wins"

    Push-Location $fixture
    try {
        $lines = @(& $finder -ScriptsRoot scripts)
        $actual = $lines[-1] | ConvertFrom-Json
    }
    finally {
        Pop-Location
    }

    $expected = @(
        "scripts/top.py",
        "scripts/custom/a_entry.py",
        "scripts/flat/main.py",
        "scripts/spec/app.spec"
    )
    if ((ConvertTo-Json @($actual) -Compress) -ne (ConvertTo-Json $expected -Compress)) {
        throw "Discovery mismatch. Expected $($expected -join ', '); got $($actual -join ', ')"
    }

    Write-Host "Discovery tests passed."
}
finally {
    Remove-Item -LiteralPath $fixture -Recurse -Force -ErrorAction SilentlyContinue
}
