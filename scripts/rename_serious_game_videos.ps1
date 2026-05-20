# Rename serious-game MP4s to Arson_* / Theft_* convention (run once).
$ErrorActionPreference = "Stop"
$Dir = Join-Path (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)) "static\videos\serious-game"

$map = @{
    # Arson (纵火)
    "Guilty1.mp4" = "Arson_Guilty_1.mp4"
    "Guilty2-1.mp4" = "Arson_Guilty_2-1.mp4"
    "Guilty2-2.mp4" = "Arson_Guilty_2-2.mp4"
    "Guilty3.mp4" = "Arson_Guilty_3.mp4"
    "Guilty4-1.mp4" = "Arson_Guilty_4-1.mp4"
    "Guilty4-2.mp4" = "Arson_Guilty_4-2.mp4"
    "Guilty5.mp4" = "Arson_Guilty_5.mp4"
    "Guilty6-1.mp4" = "Arson_Guilty_6-1.mp4"
    "Guilty6-2.mp4" = "Arson_Guilty_6-2.mp4"
    "Guilty7.mp4" = "Arson_Guilty_7.mp4"
    "Innocent1.mp4" = "Arson_Innocent_1.mp4"
    "Innocent2-1.mp4" = "Arson_Innocent_2-1.mp4"
    "Innocent2-2.mp4" = "Arson_Innocent_2-2.mp4"
    "Innocent3.mp4" = "Arson_Innocent_3.mp4"
    "Innocent4-1.mp4" = "Arson_Innocent_4-1.mp4"
    "Innocent4-2.mp4" = "Arson_Innocent_4-2.mp4"
    "Innocent5.mp4" = "Arson_Innocent_5.mp4"
    "Innocent6-1.mp4" = "Arson_Innocent_6-1.mp4"
    "Innocent6-2.mp4" = "Arson_Innocent_6-2.mp4"
    "Innocent7.mp4" = "Arson_Innocent_7.mp4"
    # Theft (盗窃) — local *_Theft suffix
    "Guilty1_Theft.mp4" = "Theft_Guilty_1.mp4"
    "Guilty2-1_Theft.mp4" = "Theft_Guilty_2-1.mp4"
    "Guilty2-2_Theft.mp4" = "Theft_Guilty_2-2.mp4"
    "Guilty3_Theft.mp4" = "Theft_Guilty_3.mp4"
    "Guilty4-1_Theft.mp4" = "Theft_Guilty_4-1.mp4"
    "Guilty4-2_Theft.mp4" = "Theft_Guilty_4-2.mp4"
    "Guilty5_Theft.mp4" = "Theft_Guilty_5.mp4"
    "Guilty6-1_Theft.mp4" = "Theft_Guilty_6-1.mp4"
    "Guilty6-2_Theft.mp4" = "Theft_Guilty_6-2.mp4"
    "Guilty7_Theft.mp4" = "Theft_Guilty_7.mp4"
    "Innocent1_Theft.mp4" = "Theft_Innocent_1.mp4"
    "Innocent2-1_Theft.mp4" = "Theft_Innocent_2-1.mp4"
    "Innocent2-2_Theft.mp4" = "Theft_Innocent_2-2.mp4"
    "Innocent3_Theft.mp4" = "Theft_Innocent_3.mp4"
    "Innocent4-1_Theft.mp4" = "Theft_Innocent_4-1.mp4"
    "Innocent4-2_Theft.mp4" = "Theft_Innocent_4-2.mp4"
    "Innocent5_Theft.mp4" = "Theft_Innocent_5.mp4"
    "Innocent6-1_Theft.mp4" = "Theft_Innocent_6-1.mp4"
    "Innocent6-2_Theft.mp4" = "Theft_Innocent_6-2.mp4"
    "Innocent7_Theft.mp4" = "Theft_Innocent_7.mp4"
    # Legacy Theft_* names (if present)
    "Theft_Guilty1.mp4" = "Theft_Guilty_1.mp4"
    "Theft_Guilty2-1.mp4" = "Theft_Guilty_2-1.mp4"
    "Theft_Guilty2-2.mp4" = "Theft_Guilty_2-2.mp4"
    "Theft_Guilty3.mp4" = "Theft_Guilty_3.mp4"
    "Theft_Guilty4-1.mp4" = "Theft_Guilty_4-1.mp4"
    "Theft_Guilty4-2.mp4" = "Theft_Guilty_4-2.mp4"
    "Theft_Guilty5.mp4" = "Theft_Guilty_5.mp4"
    "Theft_Guilty6-1.mp4" = "Theft_Guilty_6-1.mp4"
    "Theft_Guilty6-2.mp4" = "Theft_Guilty_6-2.mp4"
    "Theft_Guilty7.mp4" = "Theft_Guilty_7.mp4"
    "Theft_Innocent1.mp4" = "Theft_Innocent_1.mp4"
    "Theft_Innocent2-1.mp4" = "Theft_Innocent_2-1.mp4"
    "Theft_Innocent2-2.mp4" = "Theft_Innocent_2-2.mp4"
    "Theft_Innocent3.mp4" = "Theft_Innocent_3.mp4"
    "Theft_Innocent4-1.mp4" = "Theft_Innocent_4-1.mp4"
    "Theft_Innocent4-2.mp4" = "Theft_Innocent_4-2.mp4"
    "Theft_Innocent5.mp4" = "Theft_Innocent_5.mp4"
    "Theft_Innocent6-1.mp4" = "Theft_Innocent_6-1.mp4"
    "Theft_Innocent6-2.mp4" = "Theft_Innocent_6-2.mp4"
    "Theft_Innocent7.mp4" = "Theft_Innocent_7.mp4"
}

foreach ($old in $map.Keys) {
    $src = Join-Path $Dir $old
    $dst = Join-Path $Dir $map[$old]
    if (-not (Test-Path $src)) { continue }
    if ((Test-Path $dst) -and ((Resolve-Path $src).Path -ne (Resolve-Path $dst).Path)) {
        Write-Warning "Skip $old -> $($map[$old]) (target exists)"
        continue
    }
    Rename-Item -LiteralPath $src -NewName $map[$old]
    Write-Host "OK: $old -> $($map[$old])"
}

$count = (Get-ChildItem $Dir -Filter "*.mp4").Count
Write-Host "Done. $count mp4 files in $Dir"
