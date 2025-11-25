# === CONFIGURE PATHS AND OPTIONS ===

$watchFolder      = "D:\Watchfolder\inbound"
$extractBase      = "D:\Watchfolder\Extracted"
$completedFolder  = "D:\Watchfolder\completed"
$logFolder        = "D:\Watchfolder\log"
$sevenZip         = "C:\Program Files\7-Zip\7z.exe"
$intervalMinutes  = 1  # check every 1 minute
$matrixMode       = $true  # Set to $true for Matrix-style background during idle
$matrixFadingTail = $true  # Set to $true to enable fading tail effect
$popupAlert       = $true  # Set to $true to show popup alert after extraction

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
if ($size -eq $prevSize) { $stableCount++ } else { $stableCount=0; $prevSize=$size }
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

    
if (-not (Test-Path $firstPart)) { Log-Message "File $firstPart does not exist, skipping..." "INFO"; return }
if (-not (Wait-ForFileComplete $firstPart)) { Log-Message "File $firstPart still writing, skipping..." "INFO"; return }

New-Item -ItemType Directory -Path $destination | Out-Null
Log-Message "Created folder $destination" "DEBUG"

Log-Message "Extracting $archiveBase into $destination ..." "PROCESS"
& "$sevenZip" x "$firstPart" -o"$destination" -y

if ($LASTEXITCODE -eq 0) {
    Log-Message "Extraction completed for $archiveBase" "SUCCESS"

    # Audio alert
    [System.Media.SystemSounds]::Exclamation.Play()

    # Popup alert only if enabled
    if ($popupAlert) {
        Add-Type -AssemblyName PresentationFramework
        [System.Windows.MessageBox]::Show("Extraction completed for $archiveBase","Extraction Complete","OK","Information")
    }

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

# === MATRIX-STYLE FALLING TEXT EFFECT WITH OPTIONAL FADING TAIL ===

function Show-RealMatrix($seconds, $statusMessage) {
$width = [Console]::WindowWidth
$height = [Console]::WindowHeight
$cols = $width
$rows = $height
$chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ@#$%&*"
$dropSpeed = 1..$cols | ForEach-Object { Get-Random -Minimum 1 -Maximum 5 }
$drops = 0..($cols-1) | ForEach-Object { Get-Random -Minimum 0 -Maximum $rows }
$trailLength = if ($matrixFadingTail) { 5 } else { 1 }

    
$endTime = (Get-Date).AddSeconds($seconds)
while ((Get-Date) -lt $endTime) {
    for ($x = 0; $x -lt $cols; $x++) {
        $y = $drops[$x]
        $char = $chars[(Get-Random -Minimum 0 -Maximum $chars.Length)]
        try { [Console]::SetCursorPosition($x, $y); Write-Host $char -ForegroundColor Green -NoNewline } catch {}

        # Fading tail
        for ($t=1; $t -lt $trailLength; $t++) {
            $trailY = ($y - $t) % $rows
            if ($trailY -lt 0) { $trailY += $rows }
            try { [Console]::SetCursorPosition($x, $trailY); Write-Host " " -NoNewline } catch {}
        }

        $drops[$x] = ($y + $dropSpeed[$x]) % $rows
    }

    # Overlay status message
    $statusX = 2
    $statusY = 1
    try {
        [Console]::SetCursorPosition($statusX, $statusY)
        Write-Host (" " * ($width-2)) -ForegroundColor Black -BackgroundColor Black -NoNewline
        [Console]::SetCursorPosition($statusX, $statusY)
        Write-Host $statusMessage -ForegroundColor Yellow -BackgroundColor DarkGreen
    } catch {}

    Start-Sleep -Milliseconds 80
}
    

}

# === MAIN WATCHER LOOP WITH MATRIX ===

while ($true) {
$archiveParts = Get-ChildItem -Path $watchFolder -Filter "*.7z.001"

    
if ($archiveParts.Count -eq 0) {
    $status = "No archives found. Idle..."
    Log-Message $status "SUCCESS"
    if ($matrixMode) { Show-RealMatrix ($intervalMinutes * 60) $status } 
    else { Start-Sleep -Seconds ($intervalMinutes * 60) }
} else {
    $status = "Archives detected. Processing..."
    Log-Message $status "PROCESS"
    foreach ($file in $archiveParts) {
        $baseName = ($file.Name -replace '\.7z\.001$', '')
        ExtractArchive $baseName
    }
}
    

}
