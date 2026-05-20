# Install ffmpeg on Windows via winget (one-time)
# Run in PowerShell: .\scripts\install_ffmpeg_windows.ps1

$ErrorActionPreference = "Stop"

if (Get-Command ffmpeg -ErrorAction SilentlyContinue) {
    ffmpeg -version | Select-Object -First 1
    Write-Host "ffmpeg already installed."
    exit 0
}

if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
    Write-Error "winget not found. Install App Installer from Microsoft Store, or download ffmpeg from https://www.gyan.dev/ffmpeg/builds/"
}

Write-Host "Installing FFmpeg (Gyan.Essentials) via winget..."
winget install --id Gyan.FFmpeg.Essentials -e --accept-source-agreements --accept-package-agreements

$env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
            [System.Environment]::GetEnvironmentVariable("Path", "User")

if (Get-Command ffmpeg -ErrorAction SilentlyContinue) {
    ffmpeg -version | Select-Object -First 1
    Write-Host "OK. You can run: .\scripts\transcode_serious_game_local.ps1"
} else {
    Write-Warning "Installed but ffmpeg not in PATH yet. Close and reopen PowerShell, then retry."
}
