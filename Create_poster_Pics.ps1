Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$UserRootPath = [System.Environment]::GetFolderPath("UserProfile")

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Smart Deep-Anchor Sorter"
$Form.Size = New-Object System.Drawing.Size(350, 300)
$Form.StartPosition = "CenterScreen"
$Form.FormBorderStyle = "FixedDialog"
$Form.MaximizeBox = $false
$Form.TopMost = $true
$Form.Add_Shown({$Form.Activate()})

$CityLabel = New-Object System.Windows.Forms.Label
$CityLabel.Text = "Enter City:"
$CityLabel.Location = New-Object System.Drawing.Point(20, 20)
$CityLabel.Size = New-Object System.Drawing.Size(200, 20)
$Form.Controls.Add($CityLabel)

$CityTextBox = New-Object System.Windows.Forms.TextBox
$CityTextBox.Location = New-Object System.Drawing.Point(20, 45)
$CityTextBox.Size = New-Object System.Drawing.Size(280, 20)
$Form.Controls.Add($CityTextBox)

$DateLabel = New-Object System.Windows.Forms.Label
$DateLabel.Text = "Select Date:"
$DateLabel.Location = New-Object System.Drawing.Point(20, 85)
$DateLabel.Size = New-Object System.Drawing.Size(200, 20)
$Form.Controls.Add($DateLabel)

$DatePicker = New-Object System.Windows.Forms.DateTimePicker
$DatePicker.Location = New-Object System.Drawing.Point(20, 110)
$DatePicker.Size = New-Object System.Drawing.Size(280, 20)
$DatePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$DatePicker.CustomFormat = "MM-dd-yyyy"
$Form.Controls.Add($DatePicker)

$SubmitButton = New-Object System.Windows.Forms.Button
$SubmitButton.Text = "Generate & Smart Sort"
$SubmitButton.Location = New-Object System.Drawing.Point(60, 180)
$SubmitButton.Size = New-Object System.Drawing.Size(210, 35)
$SubmitButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$Form.AcceptButton = $SubmitButton
$Form.Controls.Add($SubmitButton)

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

function Get-DistanceInFeet ($Lat1, $Lon1, $Lat2, $Lon2) {
    $R = 6371000
    $RadLat1 = $Lat1 * [Math]::PI / 180; $RadLat2 = $Lat2 * [Math]::PI / 180
    $DeltaLat = ($Lat2 - $Lat1) * [Math]::PI / 180
    $DeltaLon = ($Lon2 - $Lon1) * [Math]::PI / 180
    $A = [Math]::Sin($DeltaLat/2) * [Math]::Sin($DeltaLat/2) + [Math]::Cos($RadLat1) * [Math]::Cos($RadLat2) * [Math]::Sin($DeltaLon/2) * [Math]::Sin($DeltaLon/2)
    return (6371000 * (2 * [Math]::Asin([Math]::Sqrt($A)))) * 3.28084
}
###   begine second part
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
    } catch {
        if ($Bitmap) { $Bitmap.Dispose() }
    }
    return $null
}

if ($Form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $City = $CityTextBox.Text.Trim()
    $SelectedDate = $DatePicker.Value.ToString("MM-dd-yyyy")
    if ([string]::IsNullOrEmpty($City)) { exit }
    
    $CityRootFolder = Join-Path -Path $UserRootPath -ChildPath "OneDrive\Documents\posters\$City"
    $WorkbookPath = Join-Path -Path $CityRootFolder -ChildPath "${City}_Workbook.xlsx"
    $TodayDirectoryPath = Join-Path -Path $CityRootFolder -ChildPath "${SelectedDate}_$City"
    $MasterFolder = Join-Path -Path $CityRootFolder -ChildPath "${City}_Master"
    $SourceFolder = Join-Path -Path $UserRootPath -ChildPath "OneDrive\Documents\posters\TEsting_auto"

    if (-not (Test-Path -Path $WorkbookPath)) {
        [System.Windows.Forms.MessageBox]::Show("Workbook not found at:`n$WorkbookPath", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        exit
    }

    # ── Read <City>_Workbook.xlsx Master sheet via Excel COM ──────────────────
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $Workbook = $Excel.Workbooks.Open($WorkbookPath)
    $Master = $Workbook.Worksheets.Item("Master")
    $UsedRows = $Master.UsedRange.Rows.Count
    $UsedCols = $Master.UsedRange.Columns.Count

    # Pick rows where col C has a street AND (col D or col E contains a number).
    # We record the SOURCE row number from Master (so we can Range.Copy it later),
    # the street name (for folder creation), and the raw D/E values (for the job<n> wipe).
    $FolderNames = New-Object System.Collections.Generic.List[string]
    $FilteredRowSpecs = New-Object 'System.Collections.Generic.List[object]'

    $totalDataRows = [Math]::Max(1, $UsedRows - 1)
    for ($r = 2; $r -le $UsedRows; $r++) {
        if ((($r - 1) % 10) -eq 0 -or $r -eq $UsedRows) {
            $pct = [int](100 * (($r - 1) / $totalDataRows))
            Write-Progress -Id 1 -Activity "Create Poster Pics ($City $SelectedDate)" -Status ("Filtering Master row {0}/{1}" -f ($r - 1), $totalDataRows) -PercentComplete $pct
        }
        $StreetRaw = $Master.Cells.Item($r, 3).Value2
        if ($null -eq $StreetRaw) { continue }
        $Street = ([string]$StreetRaw).Trim()
        if ([string]::IsNullOrWhiteSpace($Street)) { continue }

        $D = $Master.Cells.Item($r, 4).Value2
        $E = $Master.Cells.Item($r, 5).Value2
        $hasNumber = $false
        foreach ($cellVal in @($D, $E)) {
            if ($null -eq $cellVal) { continue }
            $dbl = 0.0
            if ([double]::TryParse([string]$cellVal, [ref]$dbl)) { $hasNumber = $true; break }
        }
        if (-not $hasNumber) { continue }

        $FolderNames.Add($Street) | Out-Null
        [void]$FilteredRowSpecs.Add([PSCustomObject]@{
            SourceRow = $r
            Street    = $Street
            DRaw      = $D
            ERaw      = $E
        })
    }

    if ($FolderNames.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No rows in Master have a number in column D or E. Nothing to do.", "Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $Workbook.Close($false)
        $Excel.Quit()
        $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Master)
        $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook)
        $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel)
        [GC]::Collect()
        exit
    }

    # ── Create new sheet named after the selected date (replicates Master layout) ──
    $NewSheetName = $SelectedDate
    $ExistingSheet = $null
    foreach ($ws in $Workbook.Worksheets) {
        if ($ws.Name -eq $NewSheetName) { $ExistingSheet = $ws; break }
    }
    if ($ExistingSheet) {
        $ExistingSheet.Cells.Clear() | Out-Null
        $NewSheet = $ExistingSheet
    } else {
        $NewSheet = $Workbook.Worksheets.Add([System.Reflection.Missing]::Value, $Workbook.Worksheets.Item($Workbook.Worksheets.Count))
        $NewSheet.Name = $NewSheetName
    }

    # Copy header row from Master in one shot (preserves text + formatting)
    $masterHeaderRange = $Master.Range($Master.Cells.Item(1, 1), $Master.Cells.Item(1, $UsedCols))
    $newHeaderRange    = $NewSheet.Range($NewSheet.Cells.Item(1, 1), $NewSheet.Cells.Item(1, $UsedCols))
    [void]$masterHeaderRange.Copy($newHeaderRange)

    # Add a "Google Street View" header in column F (column 6)
    $StreetViewCol = 6
    $NewSheet.Cells.Item(1, $StreetViewCol).Value2 = "Google Street View"
    $NewSheet.Cells.Item(1, $StreetViewCol).Font.Bold = $true

    # Pre-load <City>_Master subfolder list once for fuzzy matching
    $MasterDiskFolders = @()
    if (Test-Path $MasterFolder) {
        $MasterDiskFolders = Get-ChildItem -Path $MasterFolder -Directory
    }

    # Copy each filtered Master row into the new sheet, then wipe job<n> placeholders
    # in D/E and append Street View hyperlink in F.
    $rowIdx = 2
    $totalRows = $FilteredRowSpecs.Count
    $rowCounter = 0
    foreach ($spec in $FilteredRowSpecs) {
        $rowCounter++
        $pct = [int](100 * ($rowCounter / [Math]::Max(1, $totalRows)))
        Write-Progress -Id 1 -Activity "Create Poster Pics ($City $SelectedDate)" -Status ("Writing dated sheet + Street View links: {0}/{1} [{2}]" -f $rowCounter, $totalRows, $spec.Street) -PercentComplete $pct
        $srcRange = $Master.Range($Master.Cells.Item($spec.SourceRow, 1), $Master.Cells.Item($spec.SourceRow, $UsedCols))
        $dstRange = $NewSheet.Range($NewSheet.Cells.Item($rowIdx, 1), $NewSheet.Cells.Item($rowIdx, $UsedCols))
        [void]$srcRange.Copy($dstRange)

        # Wipe column D / E if they hold literal placeholder values like job1, Job 2, etc.
        foreach ($jobCol in 4, 5) {
            $rawVal = if ($jobCol -eq 4) { $spec.DRaw } else { $spec.ERaw }
            if ($null -ne $rawVal -and ([string]$rawVal).Trim() -match '^(?i)job\s*\d+$') {
                $NewSheet.Cells.Item($rowIdx, $jobCol).ClearContents() | Out-Null
            }
        }

        # Build Street View link from averaged GPS in <City>_Master\<Street>
        $LocationName = $spec.Street
        $TargetFolder = $null

        # 1. Exact subfolder match in Master
        $ExactPath = Join-Path -Path $MasterFolder -ChildPath $LocationName
        if (Test-Path $ExactPath) {
            $TargetFolder = $ExactPath
        } else {
            # 2. Fuzzy match (strip leading number, NW/NE/SE/SW, non-alphanumerics)
            $ExcelNorm = ($LocationName -replace '^\d{1,2}\s+', '' -replace '\b(NW|NE|SE|SW)\b', '' -replace '[^a-zA-Z0-9]', '').ToLower().Trim()
            foreach ($DiskDir in $MasterDiskFolders) {
                $DiskNorm = ($DiskDir.Name -replace '^\d{1,2}\s+', '' -replace '\b(NW|NE|SE|SW)\b', '' -replace '[^a-zA-Z0-9]', '').ToLower().Trim()
                if ($ExcelNorm -eq $DiskNorm) { $TargetFolder = $DiskDir.FullName; break }
            }
        }

        $LinkCell = $NewSheet.Cells.Item($rowIdx, $StreetViewCol)
        if ($TargetFolder) {
            $Pictures = Get-ChildItem -Path $TargetFolder -File | Where-Object { $_.Extension -match '\.(jpg|jpeg|png)$' }
            $LatSum = 0.0; $LonSum = 0.0; $GpsCount = 0
            foreach ($Pic in $Pictures) {
                $GPS = Get-ImageGPS -FilePath $Pic.FullName
                if ($GPS) { $LatSum += $GPS.Lat; $LonSum += $GPS.Lon; $GpsCount++ }
            }
            if ($GpsCount -gt 0) {
                $AvgLat = [Math]::Round(($LatSum / $GpsCount), 6)
                $AvgLon = [Math]::Round(($LonSum / $GpsCount), 6)
                $StreetViewUrl = "https://www.google.com/maps/@?api=1&map_action=pano&viewpoint=$AvgLat,$AvgLon"
                $LinkCell.Value2 = $null
                $LinkCell.Hyperlinks.Delete()
                $NewSheet.Hyperlinks.Add($LinkCell, [string]$StreetViewUrl, [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value, [string]"Street View") | Out-Null
            } else {
                $LinkCell.Value2 = "No GPS Photos Found"
            }
        } else {
            $LinkCell.Value2 = "Subfolder Missing"
        }

        $rowIdx++
    }
    # NOTE: We intentionally do NOT set $Excel.CutCopyMode here. The strongly-typed
    # Office.Interop.Excel typelib (loaded on machines with Visual Studio tools)
    # rejects every value for this setter except xlCopy/xlCut - there is no "off"
    # member. The marching-ants overlay only matters when Excel is visible, and we
    # run with $Excel.Visible = $false, so leaving it alone is harmless.
    $NewSheet.UsedRange.Columns.AutoFit() | Out-Null

    # ── Reorder tabs: today (leftmost) | Master | dated sheets newest->oldest | everything else ──
    # 1. Today's sheet goes to position 1
    $NewSheet.Move($Workbook.Worksheets.Item(1)) | Out-Null

    # 2. Master goes immediately after today's sheet (position 2)
    $MasterTab = $null
    foreach ($ws in $Workbook.Worksheets) {
        if ($ws.Name -eq "Master") { $MasterTab = $ws; break }
    }
    if ($MasterTab) {
        $MasterTab.Move([System.Reflection.Missing]::Value, $Workbook.Worksheets.Item($NewSheetName)) | Out-Null
    }

    # 3. All other MM-dd-yyyy sheets (excluding today) sorted newest -> oldest, placed after Master
    $DatedSheets = @()
    foreach ($ws in $Workbook.Worksheets) {
        if ($ws.Name -eq $NewSheetName) { continue }
        if ($ws.Name -match '^\d{2}-\d{2}-\d{4}$') {
            $parsed = [datetime]::MinValue
            if ([datetime]::TryParseExact($ws.Name, 'MM-dd-yyyy', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
                $DatedSheets += [PSCustomObject]@{ Sheet = $ws; Date = $parsed }
            }
        }
    }
    $AnchorName = if ($MasterTab) { "Master" } else { $NewSheetName }
    foreach ($entry in ($DatedSheets | Sort-Object Date -Descending)) {
        # Move each one immediately after the current anchor; next iteration anchors on this sheet
        $entry.Sheet.Move([System.Reflection.Missing]::Value, $Workbook.Worksheets.Item($AnchorName)) | Out-Null
        $AnchorName = $entry.Sheet.Name
    }

    # Make the new dated sheet the one that opens by default next time the workbook is launched
    $NewSheet.Activate() | Out-Null
    $NewSheet.Cells.Item(1, 1).Select() | Out-Null

    $Workbook.Save()
    $Workbook.Close($true)
    $Excel.Quit()
    $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Master)
    $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($NewSheet)
    $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook)
    $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel)
    [GC]::Collect()

    # ── Build folder tree from the filtered street list ───────────────────────
    if ($true) {
        if (-not (Test-Path -Path $TodayDirectoryPath)) {
            New-Item -ItemType Directory -Path $TodayDirectoryPath -Force | Out-Null
        }
        $UnsortedFolderPath = Join-Path -Path $TodayDirectoryPath -ChildPath "Needs_Manual_Sorting"
        if (-not (Test-Path -Path $UnsortedFolderPath)) {
            New-Item -ItemType Directory -Path $UnsortedFolderPath -Force | Out-Null
        }
        $TodaySubFolderPaths = @{}
        ForEach ($Name in $FolderNames) {
            $CleanedName = $Name.Trim()
            $SubPath = Join-Path -Path $TodayDirectoryPath -ChildPath $CleanedName
            if (-not (Test-Path -Path $SubPath)) {
                New-Item -ItemType Directory -Path $SubPath -Force | Out-Null
            }
            $TodaySubFolderPaths[$CleanedName] = $SubPath
        }

        $LogFilePath = Join-Path -Path $CityRootFolder -ChildPath "${City}_log.txt"
        $LogContent = New-Object System.Collections.Generic.List[string]
        $null = $LogContent.Add("=== SMART DEEP-ANCHOR SORTER RUN REPORT ===")
        $null = $LogContent.Add("Timestamp: $(Get-Date -Format 'MM-dd-yyyy HH:mm:ss')")
        $null = $LogContent.Add("Target City: $City")
        $null = $LogContent.Add("Source: ${City}_Workbook.xlsx [Master] -> rows where col D or E has a number")
        $null = $LogContent.Add("New sheet added/updated: $NewSheetName ($($FolderNames.Count) streets)")
        $null = $LogContent.Add("===========================================")
        
        $HistoricalDateFolders = Get-ChildItem -Path $CityRootFolder -Directory | Where-Object { $_.Name -match '^\d{2}-\d{2}-\d{4}_' } | Sort-Object CreationTime -Descending
        $GlobalGPSAnchors = @()
        $StreetRadius = @{}   # street name -> per-street override radius (ft); absent = use default
        $null = $LogContent.Add("`n[SYSTEM] Loading GPS anchors (Master folder = primary, historical = fallback)...")
        $anchorTotal = $TodaySubFolderPaths.Count
        $anchorIdx = 0

        # Pre-load Master subfolders once (drift-corrected, curated truth source)
        $MasterSubFolders = @()
        if (Test-Path $MasterFolder) {
            $MasterSubFolders = Get-ChildItem -Path $MasterFolder -Directory
            $null = $LogContent.Add("  [MASTER] Found $($MasterSubFolders.Count) curated street folders in ${City}_Master.")
        } else {
            $null = $LogContent.Add("  [MASTER] WARNING: ${City}_Master folder not found at $MasterFolder — falling back to history only.")
        }

        foreach ($LocationName in $TodaySubFolderPaths.Keys) {
            $anchorIdx++
            $pct = [int](100 * ($anchorIdx / [Math]::Max(1, $anchorTotal)))
            Write-Progress -Id 1 -Activity "Create Poster Pics ($City $SelectedDate)" -Status ("Loading GPS anchors: {0}/{1} [{2}]" -f $anchorIdx, $anchorTotal, $LocationName) -PercentComplete $pct
            $FoundAnchorForThisLocation = $false
            $CleanTodayName = $LocationName -replace '^\d{1,2}\s+', ''

            # 1. PRIMARY: Look in <City>_Master first (drift-corrected anchors)
            $MasterMatch = $null
            foreach ($MSub in $MasterSubFolders) {
                $CleanMasterName = $MSub.Name -replace '^\d{1,2}\s+', ''
                if ($MSub.Name -eq $LocationName -or $CleanMasterName -eq $CleanTodayName) {
                    $MasterMatch = $MSub.FullName
                    break
                }
            }

            if ($MasterMatch) {
                # Per-street radius override: presence of a `.radius` text file inside the
                # Master subfolder overrides the global default for this street only.
                $RadiusFile = Join-Path $MasterMatch ".radius"
                if (Test-Path $RadiusFile) {
                    try {
                        $rawRadius = (Get-Content -Path $RadiusFile -ErrorAction Stop -TotalCount 1).Trim()
                        $parsed = 0
                        if ([int]::TryParse($rawRadius, [ref]$parsed) -and $parsed -gt 0) {
                            $StreetRadius[$LocationName] = $parsed
                            $null = $LogContent.Add("    [RADIUS] [$LocationName] override = $parsed ft (from .radius file)")
                        } else {
                            $null = $LogContent.Add("    [RADIUS] [$LocationName] .radius file unreadable (value: '$rawRadius') - using default")
                        }
                    } catch {
                        $null = $LogContent.Add("    [RADIUS] [$LocationName] .radius file read failed: $($_.Exception.Message)")
                    }
                }

                $Pictures = Get-ChildItem -Path $MasterMatch -File | Where-Object { $_.Extension -match '\.(jpg|jpeg|png)$' }
                $MasterAnchorCount = 0
                foreach ($Pic in $Pictures) {
                    $GPS = Get-ImageGPS -FilePath $Pic.FullName
                    if ($GPS) {
                        $GlobalGPSAnchors += [PSCustomObject]@{ FolderName = $LocationName; Lat = $GPS.Lat; Lon = $GPS.Lon }
                        $FoundAnchorForThisLocation = $true
                        $MasterAnchorCount++
                    }
                }
                if ($FoundAnchorForThisLocation) {
                    $null = $LogContent.Add("  -> MASTER: [$LocationName] anchored from ${City}_Master\$(Split-Path $MasterMatch -Leaf) ($MasterAnchorCount GPS pics)")
                    continue   # Master is the source of truth — skip historical fallback
                }
            }

            # 2. FALLBACK: Scan historical dated folders (only if Master had nothing)
            foreach ($OldDateFolder in $HistoricalDateFolders) {
                $OldSubFolders = Get-ChildItem -Path $OldDateFolder.FullName -Directory
                $MatchingOldSubFolder = $null

                foreach ($Sub in $OldSubFolders) {
                    $CleanOldName = $Sub.Name -replace '^\d{1,2}\s+', ''
                    if ($CleanOldName -eq $CleanTodayName) {
                        $MatchingOldSubFolder = $Sub.FullName
                        break
                    }
                }

                if ($MatchingOldSubFolder -and (Test-Path $MatchingOldSubFolder)) {
                    $Pictures = Get-ChildItem -Path $MatchingOldSubFolder -File | Where-Object { $_.Extension -match '\.(jpg|jpeg|png)$' }
                    foreach ($Pic in $Pictures) {
                        $GPS = Get-ImageGPS -FilePath $Pic.FullName
                        if ($GPS) {
                            $GlobalGPSAnchors += [PSCustomObject]@{ FolderName = $LocationName; Lat = $GPS.Lat; Lon = $GPS.Lon }
                            $FoundAnchorForThisLocation = $true
                        }
                    }
                }
                if ($FoundAnchorForThisLocation) {
                    $null = $LogContent.Add("  -> HISTORY: [$LocationName] fallback-anchored from [$($Sub.Name)] in [$($OldDateFolder.Name)] (no Master folder for this street)")
                    break
                }
            }
            if (-not $FoundAnchorForThisLocation) {
                $null = $LogContent.Add("  -> WARNING: No GPS anchors found in Master or history for street [$CleanTodayName]")
            }
        }

        $null = $LogContent.Add("`nTotal GPS Anchors Loaded: $($GlobalGPSAnchors.Count)")
        $null = $LogContent.Add("===========================================`n")
        
        if (Test-Path $SourceFolder) {
            $FreshImages = Get-ChildItem -Path $SourceFolder -File | Where-Object { $_.Extension -match '\.(jpg|jpeg|png)$' }
            $MatchRadiusFeet = 35
            $ClusterWindowSeconds = 90    # +/- this many seconds for the cluster fallback
            $movedCount = 0; $manualCount = 0; $errorCount = 0; $clusterCount = 0
            $null = $LogContent.Add("[PROCESSING] Found $($FreshImages.Count) fresh images in source folder.")
            $null = $LogContent.Add("[PROCESSING] Default match radius: $MatchRadiusFeet ft   |   Cluster window: +/-$ClusterWindowSeconds s   |   Total anchors: $($GlobalGPSAnchors.Count)")
            if ($StreetRadius.Count -gt 0) {
                $overridesText = ($StreetRadius.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join "; "
                $null = $LogContent.Add("[PROCESSING] Per-street overrides: $overridesText")
            }
            $null = $LogContent.Add("-------------------------------------------")

            # ---------- PHASE 1: classify each photo (no moves yet) ----------
            $classifications = New-Object System.Collections.Generic.List[object]
            $imgIdx = 0
            $imgTotal = $FreshImages.Count
            foreach ($ImgFile in $FreshImages) {
                $imgIdx++
                if ($imgTotal -gt 0) {
                    $pct = [int](100 * ($imgIdx / [Math]::Max(1, $imgTotal)))
                    Write-Progress -Id 1 -Activity "Create Poster Pics ($City $SelectedDate)" -Status ("Classifying photos: {0}/{1} [{2}]" -f $imgIdx, $imgTotal, $ImgFile.Name) -PercentComplete $pct
                }

                # Parse timestamp from filename: IMG_yyyyMMdd_HHmmss_*  (used for cluster fallback)
                $imgTime = $null
                if ($ImgFile.BaseName -match '_(\d{8})_(\d{6})') {
                    $stamp = "$($Matches[1])_$($Matches[2])"
                    $tmp = [datetime]::MinValue
                    if ([datetime]::TryParseExact($stamp, 'yyyyMMdd_HHmmss', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$tmp)) {
                        $imgTime = $tmp
                    }
                }

                $ImgGPS = Get-ImageGPS -FilePath $ImgFile.FullName
                $entry = [PSCustomObject]@{
                    File         = $ImgFile
                    Time         = $imgTime
                    GPS          = $ImgGPS
                    Result       = 'PENDING'   # MOVED / MANUAL_RADIUS / MANUAL_NOGPS / MANUAL_NOANCHORS / CLUSTER
                    Street       = $null
                    Distance     = $null
                    Runner       = $null
                    RunnerDist   = $null
                    UsedRadius   = $null
                    InfoText     = $null      # cached log fragment
                    ClusterInfo  = $null      # populated only by Phase 2 if cluster fallback wins
                }

                if (-not $ImgGPS) {
                    $entry.Result = 'MANUAL_NOGPS'
                    $entry.InfoText = "$($ImgFile.Name) -> No GPS metadata in file."
                    [void]$classifications.Add($entry)
                    continue
                }
                if ($GlobalGPSAnchors.Count -eq 0) {
                    $entry.Result = 'MANUAL_NOANCHORS'
                    $entry.InfoText = "$($ImgFile.Name) -> Has GPS data, but script has zero anchors loaded."
                    [void]$classifications.Add($entry)
                    continue
                }

                $BestMatchName = $null;  $ClosestDistance = [double]::MaxValue
                $SecondName    = $null;  $SecondDistance  = [double]::MaxValue
                foreach ($Anchor in $GlobalGPSAnchors) {
                    $Distance = Get-DistanceInFeet -Lat1 $ImgGPS.Lat -Lon1 $ImgGPS.Lon -Lat2 $Anchor.Lat -Lon2 $Anchor.Lon
                    if ($Distance -lt $ClosestDistance) {
                        $ClosestDistance = $Distance; $BestMatchName = $Anchor.FolderName
                    }
                }
                foreach ($Anchor in $GlobalGPSAnchors) {
                    if ($Anchor.FolderName -eq $BestMatchName) { continue }
                    $Distance = Get-DistanceInFeet -Lat1 $ImgGPS.Lat -Lon1 $ImgGPS.Lon -Lat2 $Anchor.Lat -Lon2 $Anchor.Lon
                    if ($Distance -lt $SecondDistance) {
                        $SecondDistance = $Distance; $SecondName = $Anchor.FolderName
                    }
                }

                $entry.Street     = $BestMatchName
                $entry.Distance   = [Math]::Round($ClosestDistance, 2)
                $entry.Runner     = $SecondName
                $entry.RunnerDist = if ($SecondName) { [Math]::Round($SecondDistance, 2) } else { $null }

                $effectiveRadius = if ($BestMatchName -and $StreetRadius.ContainsKey($BestMatchName)) { $StreetRadius[$BestMatchName] } else { $MatchRadiusFeet }
                $entry.UsedRadius = $effectiveRadius

                if ($BestMatchName -and $ClosestDistance -le $effectiveRadius) {
                    $entry.Result = 'MOVED'
                } else {
                    $entry.Result = 'MANUAL_RADIUS'
                }
                [void]$classifications.Add($entry)
            }

            # ---------- PHASE 2: cluster fallback for the leftovers ----------
            # For each MANUAL_* photo with a valid timestamp, look at neighboring
            # confidently-MOVED photos within +/-$ClusterWindowSeconds. If they all
            # agree on a single street, infer this photo is at that street too.
            $movedWithTime = $classifications | Where-Object { $_.Result -eq 'MOVED' -and $_.Time }
            foreach ($entry in $classifications) {
                if ($entry.Result -notin @('MANUAL_RADIUS', 'MANUAL_NOGPS', 'MANUAL_NOANCHORS')) { continue }
                if (-not $entry.Time) { continue }

                $neighbors = $movedWithTime | Where-Object {
                    [Math]::Abs(($entry.Time - $_.Time).TotalSeconds) -le $ClusterWindowSeconds
                }
                if (-not $neighbors -or $neighbors.Count -eq 0) { continue }
                $uniqueStreets = ($neighbors | Select-Object -ExpandProperty Street -Unique)
                if ($uniqueStreets.Count -ne 1) { continue }   # ambiguous neighbors -> leave manual

                $inferredStreet = $uniqueStreets[0]
                $secondsList    = ($neighbors | ForEach-Object { [int][Math]::Abs(($entry.Time - $_.Time).TotalSeconds) } | Sort-Object) -join ','
                $entry.Result        = 'CLUSTER'
                $entry.Street        = $inferredStreet
                $entry.ClusterInfo   = "matched $($neighbors.Count) neighbor(s) within +/-${ClusterWindowSeconds}s ($secondsList s)"
            }

            # ---------- PHASE 3: apply moves and write log ----------
            $applyIdx = 0
            $applyTotal = $classifications.Count
            foreach ($entry in $classifications) {
                $applyIdx++
                if ($applyTotal -gt 0) {
                    $pct = [int](100 * ($applyIdx / [Math]::Max(1, $applyTotal)))
                    Write-Progress -Id 1 -Activity "Create Poster Pics ($City $SelectedDate)" -Status ("Moving photos: {0}/{1} [{2}]" -f $applyIdx, $applyTotal, $entry.File.Name) -PercentComplete $pct
                }

                $img = $entry.File
                $imgInfo = if ($entry.GPS) { "$($img.Name) [GPS $([Math]::Round($entry.GPS.Lat,6)),$([Math]::Round($entry.GPS.Lon,6))]" } else { $img.Name }
                $runnerUp = if ($entry.Runner) { "; runner-up [$($entry.Runner)] $($entry.RunnerDist) ft" } else { "" }

                switch ($entry.Result) {
                    'MOVED' {
                        $TargetMovePath = $TodaySubFolderPaths[$entry.Street]
                        if (-not $TargetMovePath -or -not (Test-Path $TargetMovePath)) {
                            $null = $LogContent.Add("[ERROR] $imgInfo -> Anchor matched [$($entry.Street)] @ $($entry.Distance) ft BUT today has no folder for that street. Sending to Needs_Manual_Sorting.")
                            try { Move-Item -Path $img.FullName -Destination $UnsortedFolderPath -Force -ErrorAction Stop; $manualCount++ }
                            catch { $null = $LogContent.Add("[ERROR] Move-Item failed: $($_.Exception.Message)"); $errorCount++ }
                            break
                        }
                        try {
                            Move-Item -Path $img.FullName -Destination $TargetMovePath -Force -ErrorAction Stop
                            $radiusNote = if ($StreetRadius.ContainsKey($entry.Street)) { "limit $($entry.UsedRadius) ft [override]" } else { "limit $($entry.UsedRadius) ft" }
                            $null = $LogContent.Add("[MOVED]  $imgInfo -> [$($entry.Street)] $($entry.Distance) ft ($radiusNote)$runnerUp -> $TargetMovePath")
                            $movedCount++
                        } catch {
                            $null = $LogContent.Add("[ERROR] $imgInfo -> Move to [$($entry.Street)] FAILED: $($_.Exception.Message). Trying Needs_Manual_Sorting.")
                            try { Move-Item -Path $img.FullName -Destination $UnsortedFolderPath -Force -ErrorAction Stop; $manualCount++ }
                            catch { $null = $LogContent.Add("[ERROR] Fallback move also failed: $($_.Exception.Message). File still in source: $($img.FullName)"); $errorCount++ }
                        }
                    }
                    'CLUSTER' {
                        $TargetMovePath = $TodaySubFolderPaths[$entry.Street]
                        if (-not $TargetMovePath -or -not (Test-Path $TargetMovePath)) {
                            $null = $LogContent.Add("[ERROR] $imgInfo -> Cluster inferred [$($entry.Street)] but no folder for that street. Sending to Needs_Manual_Sorting.")
                            try { Move-Item -Path $img.FullName -Destination $UnsortedFolderPath -Force -ErrorAction Stop; $manualCount++ }
                            catch { $null = $LogContent.Add("[ERROR] Move-Item failed: $($_.Exception.Message)"); $errorCount++ }
                            break
                        }
                        try {
                            Move-Item -Path $img.FullName -Destination $TargetMovePath -Force -ErrorAction Stop
                            $closestNote = if ($entry.Distance) { "GPS-closest [$($entry.Street)] was $($entry.Distance) ft (over limit)" } else { "no usable GPS" }
                            $null = $LogContent.Add("[CLUSTER] $imgInfo -> [$($entry.Street)] via timestamp cluster ($($entry.ClusterInfo)). $closestNote -> $TargetMovePath")
                            $clusterCount++
                        } catch {
                            $null = $LogContent.Add("[ERROR] $imgInfo -> Cluster move to [$($entry.Street)] FAILED: $($_.Exception.Message). Trying Needs_Manual_Sorting.")
                            try { Move-Item -Path $img.FullName -Destination $UnsortedFolderPath -Force -ErrorAction Stop; $manualCount++ }
                            catch { $null = $LogContent.Add("[ERROR] Fallback move also failed: $($_.Exception.Message). File still in source: $($img.FullName)"); $errorCount++ }
                        }
                    }
                    'MANUAL_RADIUS' {
                        $closestNote = if ($entry.Street) { "Closest [$($entry.Street)] $($entry.Distance) ft (limit $($entry.UsedRadius) ft)$runnerUp" } else { "No anchors compared" }
                        $null = $LogContent.Add("[MANUAL] $imgInfo -> $closestNote")
                        try { Move-Item -Path $img.FullName -Destination $UnsortedFolderPath -Force -ErrorAction Stop; $manualCount++ }
                        catch { $null = $LogContent.Add("[ERROR] Move to Needs_Manual_Sorting failed: $($_.Exception.Message). File still in source: $($img.FullName)"); $errorCount++ }
                    }
                    'MANUAL_NOGPS' {
                        $null = $LogContent.Add("[MANUAL] $($entry.InfoText)")
                        try { Move-Item -Path $img.FullName -Destination $UnsortedFolderPath -Force -ErrorAction Stop; $manualCount++ }
                        catch { $null = $LogContent.Add("[ERROR] Move to Needs_Manual_Sorting failed: $($_.Exception.Message). File still in source: $($img.FullName)"); $errorCount++ }
                    }
                    'MANUAL_NOANCHORS' {
                        $null = $LogContent.Add("[MANUAL] $($entry.InfoText)")
                        try { Move-Item -Path $img.FullName -Destination $UnsortedFolderPath -Force -ErrorAction Stop; $manualCount++ }
                        catch { $null = $LogContent.Add("[ERROR] Move to Needs_Manual_Sorting failed: $($_.Exception.Message). File still in source: $($img.FullName)"); $errorCount++ }
                    }
                }
            }

            $null = $LogContent.Add("-------------------------------------------")
            $null = $LogContent.Add("[SUMMARY] Total: $($FreshImages.Count)  |  GPS-moved: $movedCount  |  Cluster-moved: $clusterCount  |  Manual sort: $manualCount  |  Errors: $errorCount")
        } else {
            $null = $LogContent.Add("[ERROR] Source folder path does not exist: $SourceFolder")
        }
        Write-Progress -Id 1 -Activity "Create Poster Pics ($City $SelectedDate)" -Completed
        $LogContent | Out-File -FilePath $LogFilePath -Encoding utf8
        Invoke-Item -Path $TodayDirectoryPath
        Invoke-Item -Path $LogFilePath
    }
}
