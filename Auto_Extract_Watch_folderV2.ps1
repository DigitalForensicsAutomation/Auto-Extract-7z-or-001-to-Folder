# === CONFIGURE PATHS ===

$watchFolder     = "D:\Watchfolder\inbound"
$extractBase     = "D:\Watchfolder\Extracted"
$completedFolder = "D:\Watchfolder\completed"
$logFolder       = "D:\Watchfolder\log"
$sevenZip        = "C:\Program Files\7-Zip\7z.exe"
$intervalMinutes = 1  # check every 1 minute for demo purposes

# Ensure necessary folders exist

foreach ($folder in @($watchFolder, $extractBase, $completedFolder, $logFolder)) {
if (-not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder | Out-Null }
}

# === LOGGING FUNCTION WITH COLOR-CODED BACKGROUND ===

function Log-Message($message, $type="INFO") {
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$logLine = "[$timestamp] [$type] $message"

    
switch ($type) {
    "DEBUG"   { Write-Host $logLine -ForegroundColor Cyan -BackgroundColor Black }
    "INFO"    { Write-Host $logLine -ForegroundColor Yellow -BackgroundColor Black }
    "SUCCESS" { Write-Host $logLine -ForegroundColor Green -BackgroundColor Black }
    "ERROR"   { Write-Host $logLine -ForegroundColor White -BackgroundColor Red }
    "PROCESS" { Write-Host $logLine -ForegroundColor Black -BackgroundColor DarkYellow }
    default   { Write-Host $logLine -ForegroundColor White -BackgroundColor Black }
}

Add-Content -Path (Join-Path $logFolder ("log_" + (Get-Date -Format "yyyyMMdd") + ".txt")) -Value $logLine
    

}

# === WAIT UNTIL FILE IS FULLY WRITTEN ===

function Wait-ForFileComplete($filePath) {
$prevSize = -1
$stableCount = 0
while ($stableCount -lt 2) {
if (-not (Test-Path $filePath)) { return $false }
$size = (Get-Item $filePath).Length
if ($size -eq $prevSize) {
$stableCount++
} else {
$stableCount = 0
$prevSize = $size
}
Start-Sleep -Seconds 2
}
return $true
}

# === FUNCTION TO GET UNIQUE FOLDER NAME ===

function Get-UniqueFolderName($basePath, $desiredName) {
$uniqueName = $desiredName
$counter = 1
while (Test-Path (Join-Path $basePath $uniqueName)) {
$uniqueName = "$desiredName" + "_$counter"
$counter++
}
return $uniqueName
}

# === EXTRACT ARCHIVE FUNCTION WITH AUDIO/POPUP AND DUPLICATE HANDLING ===

function ExtractArchive($archiveBase) {
$firstPart = Join-Path $watchFolder "$archiveBase.7z.001"
$uniqueExtractFolder = Get-UniqueFolderName $extractBase $archiveBase
$destination = Join-Path $extractBase $uniqueExtractFolder

    
if (-not (Test-Path $firstPart)) {
    Log-Message "File $firstPart does not exist, skipping..." "INFO"
    return
}

if (-not (Wait-ForFileComplete $firstPart)) {
    Log-Message "File $firstPart is still being written, skipping..." "INFO"
    return
}

New-Item -ItemType Directory -Path $destination | Out-Null
Log-Message "Created folder $destination" "DEBUG"

Log-Message "Extracting $archiveBase into $destination ..." "PROCESS"
& "$sevenZip" x "$firstPart" -o"$destination" -y

if ($LASTEXITCODE -eq 0) {
    Log-Message "Extraction completed for $archiveBase" "SUCCESS"

    # Audio alert
    [System.Media.SystemSounds]::Exclamation.Play()

    # Popup alert
   # Add-Type -AssemblyName PresentationFramework
   # [System.Windows.MessageBox]::Show("Extraction completed for $archiveBase","Extraction Complete","OK","Information")

    # Move all archive parts to a unique subfolder in completed folder
    $uniqueCompletedFolder = Get-UniqueFolderName $completedFolder $archiveBase
    $completedSubfolder = Join-Path $completedFolder $uniqueCompletedFolder
    New-Item -ItemType Directory -Path $completedSubfolder | Out-Null

    $parts = Get-ChildItem -Path $watchFolder -Filter "$archiveBase.7z.*"
    foreach ($part in $parts) {
        Move-Item -Path $part.FullName -Destination (Join-Path $completedSubfolder $part.Name) -Force
        Log-Message "Moved $($part.Name) to $completedSubfolder" "DEBUG"
    }
} else {
    Log-Message "ERROR extracting $archiveBase" "ERROR"
    [System.Media.SystemSounds]::Beep.Play()
}
    

}

# === MAIN WATCHER LOOP WITH COLOR-CODED STATUS ===

while ($true) {
$archiveParts = Get-ChildItem -Path $watchFolder -Filter "*.7z.001"
if ($archiveParts.Count -eq 0) {
Log-Message "No archives found. Idle..." "SUCCESS"
} else {
Log-Message "Archives detected. Processing..." "PROCESS"
foreach ($file in $archiveParts) {
$baseName = ($file.Name -replace '.7z.001$', '')
ExtractArchive $baseName
}
}

    
Log-Message "Sleeping for $intervalMinutes minutes..." "DEBUG"
Start-Sleep -Seconds ($intervalMinutes * 60)
    

}
