# Compare-MasterList.ps1
# Reconciles <City>_Master\ subfolder names against the Master sheet's column C
# in <City>_Workbook.xlsx. Reports three buckets:
#   - In folder but NOT in workbook   (need to add a row in Master, or rename the folder)
#   - In workbook but NOT in folder   (need to create the folder, or remove the row)
#   - Matched (same name in both)
#
# Default city: DC. Pass -City to override.
#
# Examples:
#   .\Compare-MasterList.ps1
#   .\Compare-MasterList.ps1 -City Philly
#   .\Compare-MasterList.ps1 -ExportCsv

param(
    [string] $City = "DC",
    [switch] $ExportCsv,
    [switch] $ShowMatched
)

$UserRootPath   = [System.Environment]::GetFolderPath("UserProfile")
$CityRootFolder = Join-Path $UserRootPath "OneDrive\Documents\posters\$City"
$MasterFolder   = Join-Path $CityRootFolder "${City}_Master"
$WorkbookPath   = Join-Path $CityRootFolder "${City}_Workbook.xlsx"

Write-Host "=== Compare ${City}_Master folder vs ${City}_Workbook.xlsx [Master] sheet ===" -ForegroundColor Cyan
Write-Host "  Folder    : $MasterFolder"
Write-Host "  Workbook  : $WorkbookPath"
Write-Host ""

if (-not (Test-Path $MasterFolder)) { Write-Error "Master folder not found: $MasterFolder"; exit 1 }
if (-not (Test-Path $WorkbookPath)) { Write-Error "Workbook not found: $WorkbookPath"; exit 1 }

# --- Folder side: directory names directly under <City>_Master --------------
Write-Progress -Activity "Compare ${City}_Master" -Status "Scanning folder..." -PercentComplete 5
$FolderNames = Get-ChildItem -Path $MasterFolder -Directory |
    ForEach-Object { $_.Name.Trim() } |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
    Sort-Object -Unique

# --- Workbook side: column C of the Master sheet (skip header row) ---------
Write-Progress -Activity "Compare ${City}_Master" -Status "Opening workbook..." -PercentComplete 10
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $false

try {
    $Workbook  = $Excel.Workbooks.Open($WorkbookPath, [Type]::Missing, $true)   # ReadOnly
    $Worksheet = $Workbook.Worksheets.Item("Master")
    $UsedRows  = $Worksheet.UsedRange.Rows.Count

    $WorkbookNames = New-Object System.Collections.Generic.List[string]
    $totalDataRows = [Math]::Max(1, $UsedRows - 1)
    for ($r = 2; $r -le $UsedRows; $r++) {
        if ((($r - 1) % 10) -eq 0 -or $r -eq $UsedRows) {
            $pct = [int](20 + (60 * (($r - 1) / $totalDataRows)))
            Write-Progress -Activity "Compare ${City}_Master" -Status ("Reading workbook row {0} of {1}" -f ($r - 1), $totalDataRows) -PercentComplete $pct
        }
        $val = $Worksheet.Cells.Item($r, 3).Value2
        if ($null -eq $val) { continue }
        $name = ([string]$val).Trim()
        if (-not [string]::IsNullOrWhiteSpace($name)) { [void]$WorkbookNames.Add($name) }
    }
} finally {
    Write-Progress -Activity "Compare ${City}_Master" -Status "Closing workbook..." -PercentComplete 85
    if ($Workbook) { $Workbook.Close($false) }
    if ($Excel)    { $Excel.Quit() }
    foreach ($obj in @($Worksheet, $Workbook, $Excel)) {
        if ($obj) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

$WorkbookUnique = $WorkbookNames | Sort-Object -Unique

# --- Detect duplicates inside the workbook column C (case-insensitive) ------
Write-Progress -Activity "Compare ${City}_Master" -Status "Comparing lists..." -PercentComplete 92
$WorkbookDupes = $WorkbookNames |
    Group-Object -Property { $_.ToLower() } |
    Where-Object { $_.Count -gt 1 } |
    ForEach-Object { [PSCustomObject]@{ Name = $_.Group[0]; Count = $_.Count } }

# --- Build a case-insensitive set so DC_Master matches "Dc_master" -> same bucket
$FolderSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($n in $FolderNames)   { [void]$FolderSet.Add($n) }

$WorkbookSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($n in $WorkbookUnique) { [void]$WorkbookSet.Add($n) }

$OnlyInFolder   = $FolderNames   | Where-Object { -not $WorkbookSet.Contains($_) }
$OnlyInWorkbook = $WorkbookUnique | Where-Object { -not $FolderSet.Contains($_) }
$Matched        = $FolderNames   | Where-Object {     $WorkbookSet.Contains($_) }
Write-Progress -Activity "Compare ${City}_Master" -Completed

# --- Output ------------------------------------------------------------------
Write-Host ("Folder entries      : {0}" -f $FolderNames.Count)
Write-Host ("Workbook entries    : {0}" -f $WorkbookUnique.Count)
Write-Host ("Matched (both)      : {0}" -f $Matched.Count)         -ForegroundColor Green
Write-Host ("Only in folder      : {0}" -f $OnlyInFolder.Count)    -ForegroundColor Yellow
Write-Host ("Only in workbook    : {0}" -f $OnlyInWorkbook.Count)  -ForegroundColor Yellow
if ($WorkbookDupes) {
    Write-Host ("Duplicate rows in C : {0}" -f $WorkbookDupes.Count) -ForegroundColor Magenta
}
Write-Host ""

if ($OnlyInFolder.Count -gt 0) {
    Write-Host "-- Only in ${City}_Master folder (no row in workbook) --" -ForegroundColor Yellow
    $OnlyInFolder | ForEach-Object { Write-Host "  $_" }
    Write-Host ""
}

if ($OnlyInWorkbook.Count -gt 0) {
    Write-Host "-- Only in workbook column C (no folder on disk) --" -ForegroundColor Yellow
    $OnlyInWorkbook | ForEach-Object { Write-Host "  $_" }
    Write-Host ""
}

if ($WorkbookDupes) {
    Write-Host "-- Duplicate street names inside workbook column C --" -ForegroundColor Magenta
    $WorkbookDupes | ForEach-Object { Write-Host ("  {0}  (x{1})" -f $_.Name, $_.Count) }
    Write-Host ""
}

if ($ShowMatched -and $Matched.Count -gt 0) {
    Write-Host "-- Matched in both --" -ForegroundColor Green
    $Matched | ForEach-Object { Write-Host "  $_" }
    Write-Host ""
}

# --- Optional CSV export for diffing later ----------------------------------
if ($ExportCsv) {
    $stamp   = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvPath = Join-Path $CityRootFolder "${City}_MasterDiff_${stamp}.csv"

    $rows = @()
    $rows += $OnlyInFolder   | ForEach-Object { [PSCustomObject]@{ Status = "OnlyInFolder";   Name = $_ } }
    $rows += $OnlyInWorkbook | ForEach-Object { [PSCustomObject]@{ Status = "OnlyInWorkbook"; Name = $_ } }
    $rows += $WorkbookDupes  | ForEach-Object { [PSCustomObject]@{ Status = "WorkbookDuplicate"; Name = ("{0} (x{1})" -f $_.Name, $_.Count) } }

    if ($rows.Count -gt 0) {
        $rows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Host "CSV diff written: $csvPath" -ForegroundColor Cyan
    } else {
        Write-Host "Lists are identical - no CSV written." -ForegroundColor Green
    }
}

# --- Final verdict -----------------------------------------------------------
if ($OnlyInFolder.Count -eq 0 -and $OnlyInWorkbook.Count -eq 0 -and -not $WorkbookDupes) {
    Write-Host "OK - Master folder and workbook column C are IDENTICAL." -ForegroundColor Green
    exit 0
} else {
    Write-Host "MISMATCH - review the lists above." -ForegroundColor Yellow
    exit 1
}

