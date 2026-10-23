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
  [string]$Owner = "ShiZa039",
  [string]$Repo = "NeonCenterline"
)
# работаем из корня репозитория, откуда бы ни запустили скрипт
Set-Location (Split-Path -Parent $PSScriptRoot)
$gms = Get-ChildItem "$env:APPDATA\Corel" -Recurse -Filter "NeonCenterline.gms" -ErrorAction SilentlyContinue |
       Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $gms) { Write-Error "NeonCenterline.gms not found under $env:APPDATA\Corel"; exit 1 }
New-Item -ItemType Directory -Force -Path "release" | Out-Null
Copy-Item $gms.FullName "release\NeonCenterline.gms" -Force
$hash = (Get-FileHash "release\NeonCenterline.gms" -Algorithm SHA256).Hash.ToLower()
# копия .gms в репозитории (dist\) — запасной адрес через CDN jsDelivr по тегу версии
New-Item -ItemType Directory -Force -Path "dist" | Out-Null
Copy-Item "release\NeonCenterline.gms" "dist\NeonCenterline.gms" -Force
$obj = [ordered]@{
  version  = $Version
  gms_url  = "https://github.com/$Owner/$Repo/releases/download/$Version/NeonCenterline.gms"
  gms_url2 = "https://cdn.jsdelivr.net/gh/$Owner/$Repo@$Version/dist/NeonCenterline.gms"
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
Copy-Item "release\NeonCenterline.gms", "install.cmd", "install.ps1", "LICENSE.txt", "README.md", "README.txt", "CHANGELOG.md" $pkg
Copy-Item "addon\NeonCenterline" "$pkg\addon" -Recurse
Copy-Item "font" "$pkg" -Recurse
Compress-Archive -Path "$pkg\*" -DestinationPath "release\NeonCenterline-$Version.zip" -Force
Write-Host "release\NeonCenterline-$Version.zip  (for users)"
Write-Host "release\NeonCenterline.gms  sha256=$hash"
Write-Host "update.json written. Commit update.json and dist\NeonCenterline.gms, create tag/release $Version (tag without "v") and upload NeonCenterline.gms + NeonCenterline-$Version.zip."
