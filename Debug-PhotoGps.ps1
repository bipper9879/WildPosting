# Debug-PhotoGps.ps1
# Dumps EXIF GPS for every image in a folder so you can see what the sorter sees.
# Default folder: posters\TEsting_auto
#
# Usage:
#   .\Debug-PhotoGps.ps1
#   .\Debug-PhotoGps.ps1 -Folder "C:\path\to\some\folder"
#   .\Debug-PhotoGps.ps1 -Folder "C:\path\to\folder" -Sample "5th btw Neil Pl & Morse NE"
#       (Sample = a Master subfolder name; prints distance from each photo to its centroid)

param(
    [string] $Folder = (Join-Path ([System.Environment]::GetFolderPath("UserProfile")) "OneDrive\Documents\posters\TEsting_auto"),
    [string] $City   = "DC",
    [string] $Sample
)

Add-Type -AssemblyName System.Drawing

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
        $LatProp    = $Bitmap.PropertyItems | Where-Object { $_.Id -eq 2 }
        $LonRefProp = $Bitmap.PropertyItems | Where-Object { $_.Id -eq 3 }
        $LonProp    = $Bitmap.PropertyItems | Where-Object { $_.Id -eq 4 }
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

function Get-DistanceInFeet ($Lat1, $Lon1, $Lat2, $Lon2) {
    $RadLat1 = $Lat1 * [Math]::PI / 180; $RadLat2 = $Lat2 * [Math]::PI / 180
    $DeltaLat = ($Lat2 - $Lat1) * [Math]::PI / 180
    $DeltaLon = ($Lon2 - $Lon1) * [Math]::PI / 180
    $A = [Math]::Sin($DeltaLat/2) * [Math]::Sin($DeltaLat/2) +
         [Math]::Cos($RadLat1) * [Math]::Cos($RadLat2) *
         [Math]::Sin($DeltaLon/2) * [Math]::Sin($DeltaLon/2)
    return (6371000 * (2 * [Math]::Asin([Math]::Sqrt($A)))) * 3.28084
}

if (-not (Test-Path $Folder)) { Write-Error "Folder not found: $Folder"; exit 1 }

$images = Get-ChildItem -Path $Folder -File | Where-Object { $_.Extension -match '\.(jpg|jpeg|png)$' }
Write-Host "Folder: $Folder" -ForegroundColor Cyan
Write-Host "Images: $($images.Count)"
Write-Host ""

$sampleCentroid = $null
if ($Sample) {
    $UserRootPath = [System.Environment]::GetFolderPath("UserProfile")
    $SamplePath   = Join-Path $UserRootPath "OneDrive\Documents\posters\$City\${City}_Master\$Sample"
    if (-not (Test-Path $SamplePath)) {
        Write-Warning "Sample folder not found: $SamplePath"
    } else {
        $sLat = 0.0; $sLon = 0.0; $sCount = 0
        foreach ($p in (Get-ChildItem -Path $SamplePath -File | Where-Object { $_.Extension -match '\.(jpg|jpeg|png)$' })) {
            $g = Get-ImageGPS -FilePath $p.FullName
            if ($g) { $sLat += $g.Lat; $sLon += $g.Lon; $sCount++ }
        }
        if ($sCount -gt 0) {
            $sampleCentroid = [PSCustomObject]@{ Lat = [Math]::Round($sLat / $sCount, 6); Lon = [Math]::Round($sLon / $sCount, 6); Count = $sCount }
            Write-Host ("Sample [{0}] centroid: {1},{2}  ({3} GPS pics)" -f $Sample, $sampleCentroid.Lat, $sampleCentroid.Lon, $sampleCentroid.Count) -ForegroundColor Yellow
            Write-Host ""
        } else {
            Write-Warning "No GPS pics in $SamplePath"
        }
    }
}

$rows = foreach ($img in $images) {
    $g = Get-ImageGPS -FilePath $img.FullName
    if (-not $g) {
        [PSCustomObject]@{ File = $img.Name; Lat = $null; Lon = $null; FtToSample = $null }
    } else {
        $ft = if ($sampleCentroid) { [Math]::Round((Get-DistanceInFeet -Lat1 $g.Lat -Lon1 $g.Lon -Lat2 $sampleCentroid.Lat -Lon2 $sampleCentroid.Lon), 2) } else { $null }
        [PSCustomObject]@{ File = $img.Name; Lat = $g.Lat; Lon = $g.Lon; FtToSample = $ft }
    }
}

$rows | Format-Table -AutoSize

if ($sampleCentroid) {
    Write-Host ""
    Write-Host "Photos within 100 ft of [$Sample]:" -ForegroundColor Green
    $rows | Where-Object { $_.FtToSample -ne $null -and $_.FtToSample -le 100 } | Format-Table -AutoSize
}
