<#
.SYNOPSIS
  Force-apply Swedish (sv-SE) language, locale, keyboard, formats and timezone. Manual, admin-run.
  Verbose output is ON by default. The script does NOT close immediately; it pauses before exit and writes a log file.

.DESCRIPTION
  - Installs the Swedish language (sv-SE). Falls back to DISM capabilities if Install-Language isn't available.
  - Sets current user UI language, culture, formats, and Swedish keyboard.
  - Sets system locale and default user (.DEFAULT) so the Welcome screen and new users are Swedish.
  - Enables Automatic Time Zone (requires Location service), otherwise sets "W. Europe Standard Time".
  - Exits with 3010 if a reboot is recommended for full effect.
  - Writes a transcript log to C:\ProgramData\LocaleFix\logs\locale_YYYYMMDD_HHMMSS.log
  - Keeps the window open by default (press Enter to exit). Use -NoPause to skip the pause.

.PARAMETER Lang
  Language tag to apply. Default: sv-SE

.PARAMETER NoPause
  Skip the final "Press Enter to exit" prompt (useful for automated runs).

.EXAMPLE
  .\locale_swe_manual_en_v3.ps1
  .\locale_swe_manual_en_v3.ps1 -Lang sv-SE -NoPause
#>

[CmdletBinding()]
param(
  [string]$Lang = 'sv-SE',
  [switch]$NoPause
)

# Make verbose the default experience
$VerbosePreference = 'Continue'

# Basic error handling: capture unexpected errors but continue sections
trap {
  Write-Error "Unexpected error: $($_.Exception.Message)"
  continue
}

# Prepare logging
$logRoot = 'C:\ProgramData\LocaleFix\logs'
try {
  if (-not (Test-Path $logRoot)) { New-Item -Path $logRoot -ItemType Directory -Force | Out-Null }
} catch {}
$logPath = Join-Path $logRoot ("locale_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
try {
  Start-Transcript -Path $logPath -Append -ErrorAction SilentlyContinue | Out-Null
  Write-Verbose "Transcript started: $logPath"
} catch {
  Write-Warning "Failed to start transcript: $($_.Exception.Message)"
}

# --- Helpers ---
function Test-IsAdmin {
  $wi = [Security.Principal.WindowsIdentity]::GetCurrent()
  $wp = New-Object Security.Principal.WindowsPrincipal($wi)
  return $wp.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Write-Info($msg){ Write-Host $msg }
function Write-Err($msg){ Write-Error $msg }

$global:RebootSuggested = $false

if (-not (Test-IsAdmin)) {
  Write-Err "Please start PowerShell as Administrator (Run as administrator)."
  if (-not $NoPause) { Write-Host ""; Read-Host "Press Enter to exit" | Out-Null }
  try { Stop-Transcript | Out-Null } catch {}
  exit 2
}

Write-Verbose "Starting at $(Get-Date) as $env:USERNAME on $env:COMPUTERNAME"

# Ensure 64-bit PowerShell (rare but helpful if launched from 32-bit host)
if ($env:PROCESSOR_ARCHITEW6432) {
  Write-Verbose "Detected 32-bit host on 64-bit OS. Relaunching in 64-bit PowerShell..."
  $args = @('-NoProfile','-ExecutionPolicy','Bypass','-File',"`"$PSCommandPath`"","-Lang","$Lang","-Verbose")
  if ($NoPause) { $args += '-NoPause' }
  Start-Process -FilePath "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -ArgumentList $args -Verb RunAs
  if (-not $NoPause) { Write-Host ""; Read-Host "Press Enter to exit (the 64-bit instance is running)" | Out-Null }
  try { Stop-Transcript | Out-Null } catch {}
  exit 0
}

# Ensure HKU: drive exists (needed to modify .DEFAULT hive)
try {
  if (-not (Get-PSDrive -Name 'HKU' -ErrorAction SilentlyContinue)) {
    New-PSDrive -Name 'HKU' -PSProvider Registry -Root 'HKEY_USERS' | Out-Null
    Write-Verbose "Created HKU: PSDrive."
  } else {
    Write-Verbose "HKU: PSDrive already present."
  }
} catch {
  Write-Warning "Failed to create HKU: PSDrive: $($_.Exception.Message)"
}

# --- 1) Install sv-SE language pack (best-effort) ---
function Ensure-LanguageInstalled {
  param([string]$Language = 'sv-SE')
  $installed = $null
  try {
    $installed = Get-InstalledLanguage -ErrorAction Stop | Where-Object Language -eq $Language
  } catch {
    Write-Verbose "Get-InstalledLanguage not available (older OS) or error: $($_.Exception.Message)"
  }

  if ($installed) {
    Write-Verbose "$Language already installed."
    return
  }

  $installedNow = $false
  try {
    if (Get-Command Install-Language -ErrorAction SilentlyContinue) {
      Write-Info "Installing language via Install-Language ($Language)..."
      Install-Language -Language $Language -ErrorAction Stop | Out-Null
      $installedNow = $true
    }
  } catch {
    Write-Warning "Install-Language failed: $($_.Exception.Message)"
  }

  if (-not $installedNow) {
    Write-Info "Attempting DISM capabilities path..."
    $caps = @(
      "Language.Basic~~~$Language~0.0.1.0",
      "Language.Handwriting~~~$Language~0.0.1.0",
      "Language.OCR~~~$Language~0.0.1.0",
      "Language.Speech~~~$Language~0.0.1.0",
      "Language.TextToSpeech~~~$Language~0.0.1.0"
    )
    foreach ($cap in $caps) {
      try {
        $capState = (Get-WindowsCapability -Online -Name $cap -ErrorAction Stop).State
        if ($capState -ne 'Installed') {
          Write-Verbose "Adding capability: $cap"
          Add-WindowsCapability -Online -Name $cap | Out-Null
        }
      } catch {
        Write-Warning "Capability $cap failed: $($_.Exception.Message)"
      }
    }
  }
}

Ensure-LanguageInstalled -Language $Lang

# --- 2) Current user: language list, UI override, culture, formats, keyboard ---
try {
  Write-Info "Applying $Lang to current user..."
  $list = New-WinUserLanguageList -Language $Lang
  # Keep English as fallback if present for the user
  try {
    $existing = Get-WinUserLanguageList
    $en = $existing | Where-Object { $_.LanguageTag -like 'en-*' }
    if ($en) { $list.Add($en[0]) }
  } catch {}
  Set-WinUserLanguageList -LanguageList $list -Force | Out-Null

  Set-WinUILanguageOverride -Language $Lang
  Set-Culture $Lang
  Set-WinHomeLocation -GeoId 221  # Sweden
  Write-Verbose "Current user language, UI override, culture and home location set."
} catch {
  Write-Warning "Failed to set current user language stack: $($_.Exception.Message)"
}

# Keyboard: ensure Swedish as primary
try {
  $kbdReg = 'HKCU:\Keyboard Layout\Preload'
  if (-not (Test-Path $kbdReg)) { New-Item -Path $kbdReg -Force | Out-Null }
  # Swedish layout ID
  New-ItemProperty -Path $kbdReg -Name '1' -Value '0000041D' -PropertyType String -Force | Out-Null
  Write-Verbose "Keyboard primary layout set to Swedish (0000041D)."
} catch {
  Write-Warning "Failed to set keyboard: $($_.Exception.Message)"
}

# --- 3) System locale & default user (.DEFAULT) ---
try {
  $curSys = (Get-WinSystemLocale).Name
  if ($curSys -ne $Lang) {
    Write-Info "Setting system locale to $Lang... (a reboot is typically required)"
    Set-WinSystemLocale -SystemLocale $Lang
    $global:RebootSuggested = $true
  } else {
    Write-Verbose "System locale already $Lang."
  }
} catch {
  Write-Warning "Set-WinSystemLocale failed: $($_.Exception.Message)"
}

# Default user hive
try {
  Write-Info "Setting default language for Welcome screen and new users..."
  $intlPath = 'HKU:\.DEFAULT\Control Panel\International'
  if (-not (Test-Path $intlPath)) { New-Item -Path $intlPath -Force | Out-Null }

  New-ItemProperty -Path $intlPath -Name 'LocaleName' -Value $Lang -PropertyType String -Force | Out-Null
  # Swedish formats
  New-ItemProperty -Path $intlPath -Name 'sShortDate' -Value 'yyyy-MM-dd' -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $intlPath -Name 'sLongDate'  -Value 'den d MMMM yyyy' -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $intlPath -Name 'sTimeFormat' -Value 'HH:mm:ss' -PropertyType String -Force | Out-Null

  $geoPath = 'HKU:\.DEFAULT\Control Panel\International\Geo'
  if (-not (Test-Path $geoPath)) { New-Item -Path $geoPath -Force | Out-Null }
  New-ItemProperty -Path $geoPath -Name 'Nation' -Value '221' -PropertyType String -Force | Out-Null

  $profPath = 'HKU:\.DEFAULT\Control Panel\International\User Profile'
  if (-not (Test-Path $profPath)) { New-Item -Path $profPath -Force | Out-Null }
  # REG_MULTI_SZ reset + set
  Remove-ItemProperty -Path $profPath -Name 'Languages' -ErrorAction SilentlyContinue
  New-ItemProperty -Path $profPath -Name 'Languages' -PropertyType MultiString -Value @($Lang) -Force | Out-Null

  # Default keyboard
  $preloadDef = 'HKU:\.DEFAULT\Keyboard Layout\Preload'
  if (-not (Test-Path $preloadDef)) { New-Item -Path $preloadDef -Force | Out-Null }
  New-ItemProperty -Path $preloadDef -Name '1' -Value '0000041D' -PropertyType String -Force | Out-Null
  Write-Verbose ".DEFAULT profile updated (formats, geo, keyboard)."
} catch {
  Write-Warning "Failed to set .DEFAULT profile: $($_.Exception.Message)"
}

# --- 4) Time zone: try automatic; otherwise set static ---
function Enable-ServiceSafe($name){
  try {
    Set-Service -Name $name -StartupType Automatic -ErrorAction Stop
    Start-Service -Name $name -ErrorAction Stop
  } catch {
    Write-Warning "Service $name failed: $($_.Exception.Message)"
  }
}

try {
  Write-Info "Enabling Automatic Time Zone (requires Location service)..."
  Enable-ServiceSafe -name 'lfsvc'         # Geolocation Service (Location)
  Enable-ServiceSafe -name 'tzautoupdate'  # Automatic Time Zone

  # Allow location access at device level (best-effort)
  $capHKLM = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location'
  if (-not (Test-Path $capHKLM)) { New-Item -Path $capHKLM -Force | Out-Null }
  New-ItemProperty -Path $capHKLM -Name 'Value' -Value 'Allow' -PropertyType String -Force | Out-Null

  $capHKCU = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location'
  if (-not (Test-Path $capHKCU)) { New-Item -Path $capHKCU -Force | Out-Null }
  New-ItemProperty -Path $capHKCU -Name 'Value' -Value 'Allow' -PropertyType String -Force | Out-Null
  Write-Verbose "Location consent set and services started."
} catch {
  Write-Warning "Automatic time zone enabling failed: $($_.Exception.Message)"
}

# Static fallback (always ensure W. Europe Standard Time)
try {
  tzutil /s "W. Europe Standard Time" | Out-Null
  Write-Verbose 'Time zone set to "W. Europe Standard Time".'
} catch {
  Write-Warning "tzutil failed: $($_.Exception.Message)"
}

# --- Done ---
if ($global:RebootSuggested) {
  Write-Info "Done. A restart is recommended for system-wide language changes."
  $exitCode = 3010
} else {
  Write-Info "Done. Some apps may require restart to switch language."
  $exitCode = 0
}

try { Stop-Transcript | Out-Null } catch {}

if (-not $NoPause) {
  Write-Host ""
  Read-Host "Press Enter to exit"
}

exit $exitCode
