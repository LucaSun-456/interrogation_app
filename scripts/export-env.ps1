# Generate .env from Windows environment variables. Run from project root:
#   .\scripts\export-env.ps1
$ErrorActionPreference = "Stop"
$root = Resolve-Path (Join-Path $PSScriptRoot "..")

function Get-EnvOrEmpty($name) {
    $v = [Environment]::GetEnvironmentVariable($name, "User")
    if (-not $v) { $v = [Environment]::GetEnvironmentVariable($name, "Machine") }
    if (-not $v) { $v = [Environment]::GetEnvironmentVariable($name, "Process") }
    return $v
}

$secret = Get-EnvOrEmpty "SECRET_KEY"
if (-not $secret) {
    $secret = -join ((48..57) + (65..90) + (97..122) | Get-Random -Count 48 | ForEach-Object { [char]$_ })
}

$lines = @(
    "SECRET_KEY=$secret"
    "ADMIN_PASSWORD=$(Get-EnvOrEmpty 'ADMIN_PASSWORD')"
    ""
    "DEEPSEEK_API_KEY=$(Get-EnvOrEmpty 'DEEPSEEK_API_KEY')"
    "LIVEAVATAR_API_KEY=$(Get-EnvOrEmpty 'LIVEAVATAR_API_KEY')"
    "ELEVENLABS_API_KEY=$(Get-EnvOrEmpty 'ELEVENLABS_API_KEY')"
    ""
    "FLASK_DEBUG=false"
    "GUNICORN_WORKERS=1"
    "LOG_DIR=/app/logs"
)

$out = Join-Path $root ".env"
$lines | Set-Content -Path $out -Encoding utf8
Write-Host "Wrote $out"
$missing = @("DEEPSEEK_API_KEY", "LIVEAVATAR_API_KEY", "ELEVENLABS_API_KEY") | Where-Object { -not (Get-EnvOrEmpty $_) }
if ($missing.Count -gt 0) {
    Write-Host "WARNING: Missing env vars: $($missing -join ', ')"
}
Write-Host "Upload to server: scp .env root@SERVER:/opt/interrogation-app/.env"
