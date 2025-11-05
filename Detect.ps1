# Detect.ps1 — Classic Teams leftovers (SYSTEM, 64-bit, system-level)
# Exits 1 if cleanup is needed, else 0

$ErrorActionPreference = 'SilentlyContinue'
$needs = $false

function Get-ProfileDirs {
  Get-ChildItem "$($env:SystemDrive)\Users" -Directory |
    Where-Object { $_.Name -notin @('Public','Default','Default User','All Users') }
}

# 1) Machine-Wide Installer product codes (x64 + legacy x86)
$MsiCodes = @(
  '{731F6BAA-A986-45A4-8936-7C3AAAAA760B}',
  '{39AF0813-FA7B-4860-ADBE-93B9B214B914}'
)
foreach ($code in $MsiCodes) {
  if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$code") { $needs = $true }
  if (Test-Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$code") { $needs = $true }
}

# 2) Run values that re-install Teams
$runRoots = @(
  'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run',
  'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run'
)
$runNames = @('TeamsMachineInstaller','TeamsMachineUninstallerLocalAppData','TeamsMachineUninstallerProgramData')
foreach ($r in $runRoots) {
  foreach ($n in $runNames) {
    if (Get-ItemProperty -Path $r -Name $n -ErrorAction SilentlyContinue) { $needs = $true }
  }
}

# 3) Installer directory
$installerDir = Join-Path ${Env:ProgramFiles(x86)} 'Teams Installer'
if (Test-Path $installerDir) { $needs = $true }

# 4) Per-user file system leftovers (all profiles)
$fsGlobs = @(
  'AppData\Local\Microsoft\Teams',
  'AppData\Roaming\Microsoft\Teams',
  'AppData\Local\Microsoft\SquirrelTemp',
  'AppData\Local\Microsoft\TeamsMeetingAddin',
  'AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Teams*.lnk'
)
foreach ($p in Get-ProfileDirs) {
  foreach ($g in $fsGlobs) {
    if (Get-ChildItem -Path (Join-Path $p.FullName $g) -Force -ErrorAction SilentlyContinue) { $needs = $true }
  }
}

# 5) Per-user registry leftovers (load offline hives)
function Test-LoadedHiveEntries {
  param([string]$MountKey)
  $found = $false

  $uninstallKey = "Registry::HKLM\$MountKey\Software\Microsoft\Windows\CurrentVersion\Uninstall\Teams"
  if (Test-Path $uninstallKey) { $found = $true }

  $assoc = "Registry::HKLM\$MountKey\SOFTWARE\Microsoft\Office\Teams\Capabilities\URLAssociations"
  if (Test-Path $assoc) {
    $val = Get-ItemProperty -Path $assoc -Name 'msteams' -ErrorAction SilentlyContinue
    if ($val) { $found = $true }
  }

  $runKey = "Registry::HKLM\$MountKey\Software\Microsoft\Windows\CurrentVersion\Run"
  if (Test-Path $runKey) {
    $val2 = Get-ItemProperty -Path $runKey -Name 'com.squirrel.Teams.Teams' -ErrorAction SilentlyContinue
    if ($val2) { $found = $true }
  }

  return $found
}

foreach ($p in Get-ProfileDirs) {
  $ntuser = Join-Path $p.FullName 'NTUSER.DAT'
  if (-not (Test-Path $ntuser)) { continue }
  $mount = "PR_DETECT_$($p.Name)"
  $loadCmd   = "REG LOAD HKLM\$mount `"$ntuser`""
  $unloadCmd = "REG UNLOAD HKLM\$mount"
  $loaded = $false
  try {
    $proc = Start-Process "$env:ComSpec" -ArgumentList "/c $loadCmd" -WindowStyle Hidden -Wait -PassThru
    if ($proc.ExitCode -eq 0) { $loaded = $true }
    if ($loaded -and (Test-LoadedHiveEntries -MountKey $mount)) { $needs = $true }
  } finally {
    if ($loaded) { Start-Process "$env:ComSpec" -ArgumentList "/c $unloadCmd" -WindowStyle Hidden -Wait | Out-Null }
  }
}

if ($needs) {  1 } else {  0 }

