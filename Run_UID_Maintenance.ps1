param(
    [string[]]$Cities = @("philly", "dc", "boston"),
    [string[]]$SkipPatterns = @("7-22", "07-22"),
    [switch]$LatestOnly
)

$ErrorActionPreference = "Stop"

function Invoke-PythonFile {
    param(
        [string]$ScriptPath,
        [string[]]$Arguments = @()
    )

    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($python) {
        & python $ScriptPath @Arguments
        return
    }

    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($py) {
        & py -3 $ScriptPath @Arguments
        return
    }

    throw "Python was not found in PATH."
}

function Normalize-CityName {
    param([string]$City)

    if ([string]::IsNullOrWhiteSpace($City)) { return $null }
    switch ($City.Trim().ToLowerInvariant()) {
        "phl" { return "philly" }
        "philly" { return "philly" }
        "philadelphia" { return "philly" }
        "dc" { return "dc" }
        "d.c." { return "dc" }
        "washington" { return "dc" }
        "bos" { return "boston" }
        "boston" { return "boston" }
        default { return $null }
    }
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

$normalizedCities = New-Object System.Collections.Generic.List[string]
foreach ($city in $Cities) {
    $normalized = Normalize-CityName -City $city
    if ($null -ne $normalized -and -not $normalizedCities.Contains($normalized)) {
        [void]$normalizedCities.Add($normalized)
    }
}
if ($normalizedCities.Count -eq 0) {
    $normalizedCities.Add("philly")
    $normalizedCities.Add("dc")
    $normalizedCities.Add("boston")
}

Write-Host "=== WildPosting UID Maintenance Run ===" -ForegroundColor Cyan
Write-Host ("Cities: {0}" -f ($normalizedCities -join ", "))
Write-Host ("Skip patterns: {0}" -f ($SkipPatterns -join ", "))
Write-Host ""

# Phase 0: Verify city master folders align with workbook Master locations.
foreach ($city in $normalizedCities) {
    Write-Host ("[0/3] Verifying folder vs workbook for {0}..." -f $city) -ForegroundColor Yellow
    try {
        & "$scriptDir\Compare-MasterList.ps1" -City $city
        if ($LASTEXITCODE -ne 0) {
            Write-Host ("  -> Mismatch detected for {0}. Review output before trusting GPS backfill." -f $city) -ForegroundColor Red
        }
    } catch {
        Write-Host ("  -> Compare step failed for {0}: {1}" -f $city, $_.Exception.Message) -ForegroundColor Red
    }
}

# Phase 1: Refresh Master street-view links per city from master folders.
foreach ($city in $normalizedCities) {
    Write-Host ("[1/3] Refreshing Master Street View links for {0}..." -f $city) -ForegroundColor Yellow
    & "$scriptDir\LinkGenerator.ps1" -City $city -Automation
}

# Phase 2: Rebuild registry + stamp each city Master sheet with UID/Lat/Lon.
Write-Host "[2/3] Refreshing central UID registry and stamping city Master sheets..." -ForegroundColor Yellow
Invoke-PythonFile -ScriptPath "$scriptDir\UID_Create.py"

# Phase 3: Backfill dated workbooks.
Write-Host "[3/3] Backfilling dated workbooks (UID + Lat/Lon)..." -ForegroundColor Yellow
$args = @("--cities")
$args += [string[]]$normalizedCities
if ($LatestOnly) {
    $args += "--latest-only"
} else {
    $args += "--all-dated"
}
if ($SkipPatterns.Count -gt 0) {
    $args += "--skip-patterns"
    $args += $SkipPatterns
}
Invoke-PythonFile -ScriptPath "$scriptDir\UID_Udate_sheets.py" -Arguments $args

Write-Host ""
Write-Host "Maintenance run complete." -ForegroundColor Green
