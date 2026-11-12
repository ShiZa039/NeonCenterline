# Выпуск новой версии (запускать в PowerShell из корня репозитория):
#   .\tools\make_release.ps1 -Version 1.1.2 -NotesRu 'Что нового' -NotesEn 'What is new'
# 1) Перед этим: поднять NEON_VERSION в src\NeonUI.bas, собрать проект NeonBuild-ом, сохранить (Ctrl+S).
# 2) Скрипт берёт свежий NeonCenterline.gms из папки GMS, снимает заголовок CorelDRAW 2026 (см. ниже),
#    кладёт в release\ и dist\, считает SHA-256, пишет update.json и собирает архив для пользователей.
# 3) Дальше — по BUILD.md: запушить dist\, создать релиз с файлами, потом запушить update.json.
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
# .gms из CorelDRAW 2026 начинается с заголовка "GMS" на 18 байт, дальше обычный файл OLE. Плагины
# 1.0.x-1.1.1 при обновлении проверяют подпись OLE в самом начале файла и отбраковывают такой .gms,
# поэтому в релиз и в dist\ идёт файл без заголовка: CorelDRAW 2026 его загружает (проверено 12.11).
$b = [System.IO.File]::ReadAllBytes($gms.FullName)
$ole = [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1)
$off = 0
if ($b.Length -gt 26 -and $b[0] -eq 0x47 -and $b[1] -eq 0x4D -and $b[2] -eq 0x53) { $off = 18 }
for ($i = 0; $i -lt 8; $i++) {
  if ($b[$off + $i] -ne $ole[$i]) { Write-Error "NeonCenterline.gms: no OLE signature at offset $off"; exit 1 }
}
$out = New-Object byte[] ($b.Length - $off)
[System.Array]::Copy($b, $off, $out, 0, $out.Length)
[System.IO.File]::WriteAllBytes((Join-Path (Get-Location) "release\NeonCenterline.gms"), $out)
if ($off -gt 0) { Write-Host "GMS header (18 bytes) removed for older plugin updaters" }
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
New-Item -ItemType Directory -Force -Path "$pkg\icons" | Out-Null
Copy-Item "icons\*.ico" "$pkg\icons"
Compress-Archive -Path "$pkg\*" -DestinationPath "release\NeonCenterline-$Version.zip" -Force
Write-Host "release\NeonCenterline-$Version.zip  (for users)"
Write-Host "release\NeonCenterline.gms  sha256=$hash"
Write-Host "update.json written. Next: 1) commit and push ONLY dist\NeonCenterline.gms; 2) create release $Version on main (tag without v), attach release\NeonCenterline.gms + release\NeonCenterline-$Version.zip; 3) then commit and push update.json."
