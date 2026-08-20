param(
    [string]$ScriptsRoot = "scripts",
    [string]$GitHubOutput = ""
)

$ErrorActionPreference = "Stop"
$preferredNames = @("main.py", "__main__.py", "app.py", "run.py", "cli.py", "gui.py", "start.py")
$found = [System.Collections.Generic.List[string]]::new()

function Get-RepoRelativePath([string]$Path) {
    return (($Path | Resolve-Path -Relative) -replace '^\.\\', '' -replace '\\', '/')
}

if (Test-Path -LiteralPath $ScriptsRoot) {
    Get-ChildItem -LiteralPath $ScriptsRoot -File -Filter *.py |
        Where-Object Name -ne "__init__.py" |
        Sort-Object Name |
        ForEach-Object { $found.Add((Get-RepoRelativePath $_.FullName)) }

    Get-ChildItem -LiteralPath $ScriptsRoot -Directory | Sort-Object Name | ForEach-Object {
        $project = $_
        $target = Get-ChildItem -LiteralPath $project.FullName -File -Filter *.spec |
                  Sort-Object Name | Select-Object -First 1

        if (-not $target) {
            foreach ($name in $preferredNames) {
                $candidate = Join-Path $project.FullName $name
                if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                    $target = Get-Item -LiteralPath $candidate
                    break
                }
            }
        }

        # A project does not have to rename its entry point to main.py.
        # If no conventional name exists, build the first root-level .py file.
        if (-not $target) {
            $target = Get-ChildItem -LiteralPath $project.FullName -File -Filter *.py |
                      Where-Object Name -ne "__init__.py" |
                      Sort-Object Name | Select-Object -First 1
        }

        if ($target) {
            $relative = Get-RepoRelativePath $target.FullName
            $found.Add($relative)
        }
        else {
            Write-Warning "No .spec or root-level Python entry point found in '$($project.Name)' — skipped."
        }
    }
}

if ($found.Count -eq 0) {
    throw "No buildable Python files found under '$ScriptsRoot'. Add a .py file, a project-root .py entry point, or a .spec file."
}

$found | ForEach-Object { Write-Host "Build target: $_" }
$json = ConvertTo-Json -Compress -InputObject @($found)
Write-Host "Matrix: $json"

if ($GitHubOutput) {
    "scripts=$json" | Out-File -FilePath $GitHubOutput -Append -Encoding utf8
}
else {
    $json
}
