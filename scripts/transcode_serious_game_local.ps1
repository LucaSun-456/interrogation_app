# Local batch transcode for serious-game MP4 (Windows PowerShell)
# Requires ffmpeg in PATH: https://ffmpeg.org/download.html
#
# Usage (from project root):
#   .\scripts\transcode_serious_game_local.ps1
#
# Backs up originals to static\videos\serious-game-backup\ then overwrites with 720p H.264.

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$Dir = Join-Path $Root "static\videos\serious-game"
$Backup = Join-Path $Root "static\videos\serious-game-backup"

if (-not (Get-Command ffmpeg -ErrorAction SilentlyContinue)) {
    Write-Error "ffmpeg not found. Install ffmpeg and add to PATH."
}

if (-not (Test-Path $Dir)) {
    Write-Error "Directory not found: $Dir"
}

New-Item -ItemType Directory -Force -Path $Backup | Out-Null
$files = Get-ChildItem -Path $Dir -Filter "*.mp4"
if ($files.Count -eq 0) {
    Write-Error "No .mp4 files in $Dir"
}

Write-Host "Found $($files.Count) files. Backup -> $Backup"
foreach ($f in $files) {
    $dest = Join-Path $Backup $f.Name
    if (-not (Test-Path $dest)) {
        Copy-Item $f.FullName $dest
    }
}

$i = 0
foreach ($f in $files) {
    $i++
    $tmp = Join-Path $Dir ("tmp_" + $f.Name)
    Write-Host "[$i/$($files.Count)] $($f.Name)"
    & ffmpeg -y -i $f.FullName `
        -vf "scale=-2:720" `
        -c:v libx264 -preset medium -crf 23 `
        -c:a aac -b:a 128k `
        -movflags +faststart `
        $tmp
    if ($LASTEXITCODE -ne 0) { Write-Error "ffmpeg failed: $($f.Name)" }
    Move-Item -Force $tmp $f.FullName
}

Write-Host "Done. Upload static\videos\serious-game\ to server via WinSCP."
