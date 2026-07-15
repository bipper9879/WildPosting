# 1. SETUP SMART PARAMETER ENTRANCE GATE (MUST BE LINE 1)
Param(
    [string]$City,
    [switch]$Automation
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 2. POPUP DETECTOR: Only open form if $City was NOT passed in from master script
if ([string]::IsNullOrEmpty($City)) {
    
    $PromptForm = New-Object System.Windows.Forms.Form
    $PromptForm.Text = "Excel Street View Linker"
    $PromptForm.Size = New-Object System.Drawing.Size(320, 160)
    $PromptForm.StartPosition = "CenterScreen"
    $PromptForm.FormBorderStyle = "FixedDialog"
    $PromptForm.MaximizeBox = $false
    $PromptForm.TopMost = $true
    
    $PromptLabel = New-Object System.Windows.Forms.Label
    $PromptLabel.Text = "Enter the City code (e.g., DC, Miami):"
    $PromptLabel.Location = New-Object System.Drawing.Point(20, 15)
    $PromptLabel.Size = New-Object System.Drawing.Size(260, 20)
    $PromptForm.Controls.Add($PromptLabel)
    
    $PromptTextBox = New-Object System.Windows.Forms.TextBox
    $PromptTextBox.Location = New-Object System.Drawing.Point(20, 40)
    $PromptTextBox.Size = New-Object System.Drawing.Size(260, 20)
    $PromptForm.Controls.Add($PromptTextBox)
    
    $PromptButton = New-Object System.Windows.Forms.Button
    $PromptButton.Text = "OK"
    $PromptButton.Location = New-Object System.Drawing.Point(110, 80)
    $PromptButton.Size = New-Object System.Drawing.Size(80, 28)
    $PromptButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $PromptForm.AcceptButton = $PromptButton
    $PromptForm.Controls.Add($PromptButton)
    
    $PromptForm.Add_Shown({$PromptForm.Activate(); $PromptTextBox.Focus()})
    
    if ($PromptForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $City = $PromptTextBox.Text.Trim()
    }
    
    if ([string]::IsNullOrEmpty($City)) {
        [System.Windows.Forms.MessageBox]::Show("Operation cancelled. No city entered.", "Exiting")
        exit
    }
}

# 3. Fully Automated Path Generation based on city input
$PostersRoot = $env:WILDPOSTING_POSTERS_ROOT
if ([string]::IsNullOrWhiteSpace($PostersRoot)) { $PostersRoot = $env:POSTERS_ROOT }
if ([string]::IsNullOrWhiteSpace($PostersRoot)) {
    $PostersRoot = Join-Path -Path ([System.Environment]::GetFolderPath("UserProfile")) -ChildPath "OneDrive\Documents\posters"
}

$UserRootPath = [System.Environment]::GetFolderPath("UserProfile")
$CityRootFolder = Join-Path -Path $PostersRoot -ChildPath $City
$MasterFolder = Join-Path -Path $CityRootFolder -ChildPath "${City}_Master"
$ExistingExcelFile = Join-Path -Path $CityRootFolder -ChildPath "${City}_Workbook.xlsx"

Write-Host "=== RUNNING EXCEL INJECTOR FOR: $City ===" -ForegroundColor Cyan
Write-Host "Targeting Workbook: ${City}_Workbook.xlsx`n" -ForegroundColor Cyan

if (-not (Test-Path $MasterFolder)) { 
    if ($Automation) {
        Write-Error "Master folder not found at: $MasterFolder"
    } else {
        [System.Windows.Forms.MessageBox]::Show("Master folder not found at:`n$MasterFolder", "Error")
    }
    exit 
}
if (-not (Test-Path $ExistingExcelFile)) { 
    if ($Automation) {
        Write-Error "Excel workbook not found at: $ExistingExcelFile"
    } else {
        [System.Windows.Forms.MessageBox]::Show("Excel workbook not found at:`n$ExistingExcelFile", "Error")
    }
    exit 
}

function Convert-ExifToDecimal ($RawData) {
    if ($RawData.Count -lt 6) { return $null }
    $DegNum = [BitConverter]::ToUInt32($RawData, 0); $DegDen = [BitConverter]::ToUInt32($RawData, 4)
    $Degrees = if ($DegDen -ne 0) { $DegNum / $DegDen } else { $DegNum }
    $MinNum = [BitConverter]::ToUInt32($RawData, 8); $MinDen = [BitConverter]::ToUInt32($RawData, 12)
    $Minutes = if ($MinDen -ne 0) { $MinNum / $MinDen } else { $MinNum }
    $SecNum = [BitConverter]::ToUInt32($RawData, 16); $SecDen = [BitConverter]::ToUInt32($RawData, 20)
    $Seconds = if ($SecDen -ne 0) { $SecNum / $SecDen } else { $SecNum }
    return [Math]::Round(($Degrees + ($Minutes / 60) + ($Seconds / 3600)), 6)
}

function Get-ImageGPS ($FilePath) {
    try {
        $Bitmap = New-Object System.Drawing.Bitmap($FilePath)
        $LatRefProp = $Bitmap.PropertyItems | Where-Object { $_.Id -eq 1 }
        $LatProp = $Bitmap.PropertyItems | Where-Object { $_.Id -eq 2 }
        $LonRefProp = $Bitmap.PropertyItems | Where-Object { $_.Id -eq 3 }
        $LonProp = $Bitmap.PropertyItems | Where-Object { $_.Id -eq 4 }
        $Bitmap.Dispose()
        if ($LatProp -and $LonProp) {
            $Lat = Convert-ExifToDecimal -RawData $LatProp.Value
            $Lon = Convert-ExifToDecimal -RawData $LonProp.Value
            if ($LatRefProp -and $LatRefProp.Value -eq 83) { $Lat = -$Lat }
            if ($LonRefProp -and $LonRefProp.Value -eq 87) { $Lon = -$Lon }
            return [PSCustomObject]@{ Lat = $Lat; Lon = $Lon }
        }
    } catch { if ($Bitmap) { $Bitmap.Dispose() } }
    return $null
}
# Open Excel engine cleanly
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = (-not $Automation)
$Excel.DisplayAlerts = $false

$Workbook = $Excel.Workbooks.Open($ExistingExcelFile)
$Worksheet = $Workbook.Worksheets.Item("Master")
$UsedRows = $Worksheet.UsedRange.Rows.Count

# Write header title to Column F if it is empty
if ([string]::IsNullOrWhiteSpace($Worksheet.Cells.Item(1, 6).Value2)) {
    $Worksheet.Cells.Item(1, 6).Value2 = "Google Street View"
    $Worksheet.Cells.Item(1, 6).Font.Bold = $true
}

# Pre-load all available hard drive master folders for fuzzy checking
$DiskFolders = Get-ChildItem -Path $MasterFolder -Directory

# Loop through every row starting at row 2 (skipping titles)
for ($r = 2; $r -le $UsedRows; $r++) {
    $RawLocationName = $Worksheet.Cells.Item($r, 3).Value2
    if ($RawLocationName -eq $null) { continue }
    
    $LocationString = [string]$RawLocationName
    if ([string]::IsNullOrWhiteSpace($LocationString)) { continue }

    $LocationName = $LocationString.Trim()
    $MatchingFolder = Join-Path -Path $MasterFolder -ChildPath $LocationName
    $LinkCellRange = $Worksheet.Cells.Item($r, 6)
    
    $TargetFolderToOpen = $null
    
    # 1. Check for a perfect match first
    if (Test-Path $MatchingFolder) {
        $TargetFolderToOpen = $MatchingFolder
    } 
    # 2. FUZZY ENGINE ACTIVE: Scan disk folders using compressed text metrics
    else {
        $ExcelNorm = ($LocationName -replace '^\d{1,2}\s+', '' -replace '\b(NW|NE|SE|SW)\b', '' -replace '[^a-zA-Z0-9]', '').ToLower().Trim()
        
        foreach ($DiskDir in $DiskFolders) {
            $DiskNorm = ($DiskDir.Name -replace '^\d{1,2}\s+', '' -replace '\b(NW|NE|SE|SW)\b', '' -replace '[^a-zA-Z0-9]', '').ToLower().Trim()
            
            if ($ExcelNorm -eq $DiskNorm) {
                $TargetFolderToOpen = $DiskDir.FullName
                break
            }
        }
    }
    
    # EXECUTION: Process the folder if found via direct or fuzzy routing
    if ($TargetFolderToOpen -ne $null) {
        $Images = Get-ChildItem -Path $TargetFolderToOpen -File | Where-Object { $_.Extension -match '\.(jpg|jpeg|png)$' }
        $LatSum = 0; $LonSum = 0; $ValidGpsCount = 0
        
        foreach ($Img in $Images) {
            $GPS = Get-ImageGPS -FilePath $Img.FullName
            if ($GPS) { $LatSum += $GPS.Lat; $LonSum += $GPS.Lon; $ValidGpsCount++ }
        }
        
        if ($ValidGpsCount -gt 0) {
            $AvgLat = [Math]::Round(($LatSum / $ValidGpsCount), 6)
            $AvgLon = [Math]::Round(($LonSum / $ValidGpsCount), 6)
            
            # Build Google Maps Street View URL using the averaged GPS coordinates
            $StreetViewUrl = "https://www.google.com/maps/@?api=1&map_action=pano&viewpoint=$AvgLat,$AvgLon"
            
            # WIPE remnants completely before building the clean link object
            $LinkCellRange.Value2 = $null
            $LinkCellRange.Hyperlinks.Delete()
            
            # Directly invoke the macro injector via un-truncatable parameter blocks
            $DisplayTextString = [string]"Street View"
            $Worksheet.Hyperlinks.Add($LinkCellRange, [string]$StreetViewUrl, [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value, $DisplayTextString) | Out-Null
            
            if ($TargetFolderToOpen -eq $MatchingFolder) {
                Write-Host " -> Row $r [$LocationName]: Generated link in Column F." -ForegroundColor Green
            } else {
                Write-Host " -> Row $r [$LocationName]: Generated link via Fuzzy Match engine." -ForegroundColor Cyan
            }
        } else {
            $LinkCellRange.Value2 = "No GPS Photos Found"
            Write-Host " -> Row $r [$LocationName]: Skipped. No photos with GPS metadata." -ForegroundColor Yellow
        }
    } else {
        $LinkCellRange.Value2 = "Subfolder Missing"
        Write-Host " -> Row $r [$LocationName]: Skipped. Subfolder completely missing from hard drive." -ForegroundColor Red
    }
}

# Save, close, and clean out background processes
$Worksheet.UsedRange.Columns.AutoFit() | Out-Null
$Workbook.Save()
$Workbook.Close($true)
$Excel.Quit()

$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Worksheet)
$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook)
$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel)
[GC]::Collect()

Write-Host "`n🎉 COMPLETE! Column F has been fully updated in ${City}_Workbook.xlsx." -ForegroundColor Green
if (-not $Automation) {
    Invoke-Item -Path $ExistingExcelFile
}
