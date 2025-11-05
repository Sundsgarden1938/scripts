# Remediate.ps1 — Classic Teams full cleanup (SYSTEM, 64-bit, system-level)
[CmdletBinding()]
param()

$ErrorActionPreference = 'SilentlyContinue'

# ---------- Logging ----------
$Global:LogFile = Join-Path $env:WINDIR "Temp\PR_ClassicTeams_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
New-Item -Path $Global:LogFile -ItemType File -Force | Out-Null
function Log { param([string]$m) Add-Content -Path $Global:LogFile -Value "$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss') $m" }

# ---------- Helpers ----------
function Get-ProfileDirs {
  Get-ChildItem "$($env:SystemDrive)\Users" -Directory |
    Where-Object { $_.Name -notin @('Public','Default','Default User','All Users') }
}
function Remove-Dir($Path) {
  if (Test-Path $Path) {
    try { Remove-Item -Path $Path -Recurse -Force -ErrorAction SilentlyContinue; Log "Removed: $Path" }
    catch { Log "Failed to remove $Path : $_" }
  }
}

# ---------- 0) Stop running Teams processes ----------
try {
  $procs = Get-Process -Name 'teams' -ErrorAction SilentlyContinue
  if ($procs) { Log "Stopping teams.exe"; $procs | Stop-Process -Force -ErrorAction SilentlyContinue }
} catch {}

# ---------- 1) Uninstall Machine-Wide Installer ----------
$MsiCodes = @(
  '{731F6BAA-A986-45A4-8936-7C3AAAAA760B}',  # x64
  '{39AF0813-FA7B-4860-ADBE-93B9B214B914}'   # legacy x86
)
foreach ($code in $MsiCodes) {
  foreach ($root in @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$code",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$code"
  )) {
    if (Test-Path $root) {
      Log "Uninstalling MSI $code"
      try {
        $p = Start-Process "msiexec.exe" -ArgumentList "/x $code /qn ALLUSERS=1" -WindowStyle Hidden -PassThru -Wait
        Log "msiexec exit $($p.ExitCode) for $code"
      } catch { Log "msiexec error for $code : $_" }
    }
  }
}

# ---------- 2) Remove machine Run values that re-install ----------
$runRoots = @(
  'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run',
  'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run'
)
$runNames = @('TeamsMachineInstaller','TeamsMachineUninstallerLocalAppData','TeamsMachineUninstallerProgramData')
foreach ($root in $runRoots) {
  foreach ($name in $runNames) {
    try {
      if (Get-ItemProperty -Path $root -Name $name -ErrorAction SilentlyContinue) {
        Remove-ItemProperty -Path $root -Name $name -Force -ErrorAction SilentlyContinue
        Log "Deleted Run value $name in $root"
      }
    } catch {}
  }
}

# ---------- 3) Remove installer directory ----------
$installerDir = Join-Path ${Env:ProgramFiles(x86)} 'Teams Installer'
Remove-Dir $installerDir

# ---------- 4) Clean per-user file system (all profiles) ----------
$fsTargets = @(
  'AppData\Local\Microsoft\Teams',
  'AppData\Roaming\Microsoft\Teams',
  'AppData\Local\Microsoft\SquirrelTemp',
  'AppData\Local\Microsoft\TeamsMeetingAddin'
)
foreach ($p in Get-ProfileDirs) {
  foreach ($rel in $fsTargets) { Remove-Dir (Join-Path $p.FullName $rel) }
  # Shortcuts
  Get-ChildItem -Path (Join-Path $p.FullName 'AppData\Roaming\Microsoft\Windows\Start Menu\Programs') `
    -Filter 'Microsoft Teams*.lnk' -ErrorAction SilentlyContinue |
    ForEach-Object {
      try { Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue; Log "Removed shortcut: $($_.FullName)" } catch {}
    }
}

# ---------- 5) Clean per-user registry (load offline NTUSER.DAT) ----------
function Clean-LoadedHive {
  param([string]$MountKey)
  try {
    $uninstallKey = "Registry::HKLM\$MountKey\Software\Microsoft\Windows\CurrentVersion\Uninstall\Teams"
    if (Test-Path $uninstallKey) {
      Remove-Item $uninstallKey -Recurse -Force -ErrorAction SilentlyContinue
      Log "Removed HKCU uninstall key (mounted $MountKey)"
    }

    $assoc = "Registry::HKLM\$MountKey\SOFTWARE\Microsoft\Office\Teams\Capabilities\URLAssociations"
    if (Test-Path $assoc) {
      if (Get-ItemProperty -Path $assoc -Name 'msteams' -ErrorAction SilentlyContinue) {
        Remove-ItemProperty -Path $assoc -Name 'msteams' -Force -ErrorAction SilentlyContinue
        Log "Removed URLAssociations\msteams (mounted $MountKey)"
      }
    }

    $runKey = "Registry::HKLM\$MountKey\Software\Microsoft\Windows\CurrentVersion\Run"
    if (Test-Path $runKey) {
      if (Get-ItemProperty -Path $runKey -Name 'com.squirrel.Teams.Teams' -ErrorAction SilentlyContinue) {
        Remove-ItemProperty -Path $runKey -Name 'com.squirrel.Teams.Teams' -Force -ErrorAction SilentlyContinue
        Log "Removed HKCU Run\com.squirrel.Teams.Teams (mounted $MountKey)"
      }
    }
  } catch {
    Log "Clean-LoadedHive error ($MountKey): $_"
  }
}

foreach ($p in Get-ProfileDirs) {
  $ntuser = Join-Path $p.FullName 'NTUSER.DAT'
  if (-not (Test-Path $ntuser)) { continue }
  $mount = "PR_CLEAN_$($p.Name)"
  $loadCmd   = "REG LOAD HKLM\$mount `"$ntuser`""
  $unloadCmd = "REG UNLOAD HKLM\$mount"
  $loaded = $false
  try {
    $proc = Start-Process "$env:ComSpec" -ArgumentList "/c $loadCmd" -WindowStyle Hidden -Wait -PassThru
    if ($proc.ExitCode -eq 0) {
      $loaded = $true
      Clean-LoadedHive -MountKey $mount
    } else {
      Log "REG LOAD failed ($($p.Name)) exit $($proc.ExitCode)"
    }
  } finally {
    if ($loaded) {
      Start-Process "$env:ComSpec" -ArgumentList "/c $unloadCmd" -WindowStyle Hidden -Wait | Out-Null
      Log "REG UNLOAD done ($($p.Name))"
    }
  }
}

# ---------- 6) Optional PatchCache cleanup ----------
foreach ($pc in @(
  'C:\Windows\Installer\$PatchCache$\Managed\AD7E5E9D92C699247949F5DDF5A4D661',
  'C:\Windows\Installer\$PatchCache$\Managed\3180FA93B7AF0684DAEB399B2B419B41'
)) { Remove-Dir $pc }

# ---------- Summary for IME logs ----------
[ordered]@{
  Timestamp = (Get-Date).ToString('s')
  RanAs = 'SYSTEM'
  Log = $Global:LogFile
} | ConvertTo-Json -Compress | Out-Null

exit 0
