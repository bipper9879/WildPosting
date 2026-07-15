$ErrorActionPreference = 'Stop'

# Legacy compatibility shim:
# Keep old launchers/shortcuts working, but run the single canonical script.
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$canonicalScript = Join-Path $scriptDir 'Create_poster_Pics.ps1'

if (-not (Test-Path -LiteralPath $canonicalScript)) {
    throw "Canonical script not found: $canonicalScript"
}

& $canonicalScript
exit $LASTEXITCODE
