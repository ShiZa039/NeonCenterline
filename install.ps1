# NeonCenterline - installer (called from install.cmd).
# Finds the NEWEST installed CorelDRAW by the version of CorelDRW.exe (folder names can be misleading,
# old versions leave folders behind), then:
#   1) NeonCenterline.gms -> user GMS folder of that version (no admin rights);
#   2) addon\NeonCenterline -> Addons folder of that version (asks for administrator rights).
param([string]$AddonOnly = "")
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path

function Find-Corel {
  $roots = @("$env:ProgramFiles\Corel", "${env:ProgramFiles(x86)}\Corel") | Where-Object { $_ -and (Test-Path $_) }
  $exes = foreach ($r in $roots) { Get-ChildItem $r -Recurse -Filter "CorelDRW.exe" -ErrorAction SilentlyContinue }
  $best = $exes | Sort-Object { $_.VersionInfo.FileMajorPart }, { $_.LastWriteTime } -Descending | Select-Object -First 1
  return $best
}

if ($AddonOnly) {
  New-Item -ItemType Directory -Force -Path $AddonOnly | Out-Null
  Copy-Item "$here\addon\NeonCenterline\*" $AddonOnly -Force
  exit 0
}

$exe = Find-Corel
if (-not $exe) { Write-Host "CorelDRAW (CorelDRW.exe) not found."; exit 1 }
$ver = $exe.VersionInfo.FileMajorPart
$suite = $exe.Directory.Parent.FullName            # ...\CorelDRAW Graphics Suite XXXX
$addons = Join-Path $exe.Directory.FullName "Addons"
Write-Host "CorelDRAW version $ver : $suite"

# 1) GMS: папка пользователя с тем же именем, что и папка программы; иначе - самая свежая с Draw\GMS
$name = Split-Path $suite -Leaf
$gms = Join-Path "$env:APPDATA\Corel\$name" "Draw\GMS"
if (-not (Test-Path $gms)) {
  $cand = Get-ChildItem "$env:APPDATA\Corel" -Directory -ErrorAction SilentlyContinue |
          Where-Object { Test-Path (Join-Path $_.FullName "Draw\GMS") } |
          Sort-Object { (Get-Item (Join-Path $_.FullName "Draw")).LastWriteTime } -Descending | Select-Object -First 1
  if ($cand) { $gms = Join-Path $cand.FullName "Draw\GMS" }
}
if (Test-Path "$here\NeonCenterline.gms") {
  if (Test-Path $gms) {
    Copy-Item "$here\NeonCenterline.gms" $gms -Force
    Write-Host "GMS   -> $gms"
  } else { Write-Host "User GMS folder not found - copy NeonCenterline.gms manually (see README)." }
} else { Write-Host "NeonCenterline.gms is not in this folder - skipped (build it with NeonBuild)." }

# 2) Addons (права администратора)
$dst = Join-Path $addons "NeonCenterline"
Write-Host "Addon -> $dst  (administrator rights)"
Start-Process powershell -Verb RunAs -Wait -ArgumentList @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$($MyInvocation.MyCommand.Path)`"", "-AddonOnly", "`"$dst`"")
if (Test-Path (Join-Path $dst "AppUI.xslt")) { Write-Host "OK. Restart CorelDRAW: Window > Toolbars > NeonCenterline." }
else { Write-Host "Addon was not copied. Copy addon\NeonCenterline manually to: $addons" }
