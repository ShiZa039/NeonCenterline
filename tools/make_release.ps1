# Выпуск новой версии (запускать в PowerShell из корня репозитория):
#   .\tools\make_release.ps1 -Version 1.1.0 -NotesRu "Что нового" -NotesEn "What's new"
# 1) Перед этим: поднять NEON_VERSION в src\NeonUI.bas, собрать проект NeonBuild-ом, сохранить (Ctrl+S).
# 2) Скрипт копирует NeonCenterline.gms из папки GMS в release\, считает SHA-256 и пишет update.json.
# 3) Дальше: GitHub -> Releases -> новый релиз v<версия>, приложить release\NeonCenterline.gms;
#    закоммитить update.json в main. Плагины у пользователей увидят обновление.
param(
  [Parameter(Mandatory=$true)][string]$Version,
  [string]$NotesRu = "",
  [string]$NotesEn = "",
  [string]$Owner = "OWNER",
  [string]$Repo = "NeonCenterline"
)
$gms = Get-ChildItem "$env:APPDATA\Corel" -Recurse -Filter "NeonCenterline.gms" -ErrorAction SilentlyContinue |
       Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $gms) { Write-Error "NeonCenterline.gms not found under $env:APPDATA\Corel"; exit 1 }
New-Item -ItemType Directory -Force -Path "release" | Out-Null
Copy-Item $gms.FullName "release\NeonCenterline.gms" -Force
$hash = (Get-FileHash "release\NeonCenterline.gms" -Algorithm SHA256).Hash.ToLower()
$obj = [ordered]@{
  version  = $Version
  gms_url  = "https://github.com/$Owner/$Repo/releases/download/v$Version/NeonCenterline.gms"
  sha256   = $hash
  notes_ru = $NotesRu
  notes_en = $NotesEn
}
$json = $obj | ConvertTo-Json
[System.IO.File]::WriteAllText((Join-Path (Get-Location) "update.json"), $json, (New-Object System.Text.UTF8Encoding($false)))
# архив для пользователей: всё, что нужно для установки
$pkg = "release\NeonCenterline-$Version"
Remove-Item $pkg -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path "$pkg\addon" | Out-Null
Copy-Item "release\NeonCenterline.gms", "install.cmd", "install.ps1", "LICENSE.txt", "README.md", "CHANGELOG.md" $pkg
Copy-Item "addon\NeonCenterline" "$pkg\addon" -Recurse
Copy-Item "font" "$pkg" -Recurse
Compress-Archive -Path "$pkg\*" -DestinationPath "release\NeonCenterline-$Version.zip" -Force
Write-Host "release\NeonCenterline-$Version.zip  (for users)"
Write-Host "release\NeonCenterline.gms  sha256=$hash"
Write-Host "update.json written. Upload NeonCenterline.gms AND NeonCenterline-$Version.zip to release v$Version, then commit update.json."
