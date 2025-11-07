<#PSScriptInfo
.VERSION 1.2.5
.GUID 75abbb52-e359-4945-81f6-3fdb711239a9
.AUTHOR asherto
.COMPANYNAME asheroto
.TAGS PowerShell, Microsoft Teams, remove, uninstall, delete, erase, uninstaller, widget, chat, enable, disable, change
.PROJECTURI https://github.com/asheroto/UninstallTeams
.RELEASENOTES
[Version 0.0.1] - Initial Release.
[Version 0.0.2] - Fixed typo and confirmed directory existence before removal.
[Version 0.0.3] - Added support for Uninstall registry key.
[Version 0.0.4] - Added to GitHub.
[Version 0.0.5] - Fixed signature.
[Version 0.0.6] - Fixed various bugs.
[Version 0.0.7] - Added removal AppxPackage.
[Version 0.0.8] - Added removal of startup entries.
[Version 1.0.0] - Added ability to optionally disable Chat widget (Win+C) which will reinstall Teams. Major refactor of code.
[Version 1.0.1] - Added URL to -CheckForUpdate function when script is out of date.
[Version 1.0.2] - Improve description.
[Version 1.0.3] - Fixed bug with -Version.
[Version 1.0.4] - Improved CheckForUpdate function by converting time to local time and switching to variables.
[Version 1.0.5] - Changed -CheckForUpdates to -CheckForUpdate.
[Version 1.1.0] - Various bug fixes. Added removal of Desktop and Start Menu shortcuts. Added method to prevent Office from installing Teams. Added folders and registry keys to detect.
[Version 1.1.1] - Improved Chat widget warning detection. Improved output into section headers.
[Version 1.1.2] - Improved DisableOfficeTeamsInstall by adding registry key if it doesn't exist.
[Version 1.1.3] - Added TeamsMachineInstaller registry key for deletion.
[Version 1.1.4] - Added Teams uninstall registry key for deletion.
[Version 1.2.0] - Improved functionality of uninstall key removal by detecting MsiExec product GUID to uninstall teams. Added additional startup registry keys.
[Version 1.2.1] - Added additional file and registry uninstall locations.
[Version 1.2.2] - Improved detection of registry uninstall keys. Improved error handling.
[Version 1.2.3] - Fixed bug when uninstalling Teams from the uninstall registry key and using MsiExec.exe.
[Version 1.2.4] - Added AutorunsDisabled registry keys for deletion.
[Version 1.2.5] - Improved path handling for Desktop and Programs folder paths by using special folders.
#>

<#
.SYNOPSIS
Uninstalls Microsoft Teams completely. Optional parameters to disable the Chat widget (Win+C) and prevent Office from installing Teams.

.DESCRIPTION
Uninstalls Microsoft Teams completely. Optional parameters to disable the Chat widget (Win+C) and prevent Office from installing Teams.

The script stops the Teams process, uninstalls Teams using the uninstall key, uninstalls Teams from the Program Files (x86) directory, uninstalls Teams from the AppData directory, removes the Teams AppxPackage, deletes the Microsoft Teams directory in AppData, deletes the Teams directory in AppData, removes the startup registry keys for Teams, and removes the Desktop and Start Menu icons for Teams.

.PARAMETER DisableChatWidget
Disables the Chat widget (Win+C) for Microsoft Teams.

.PARAMETER EnableChatWidget
Enables the Chat widget (Win+C) for Microsoft Teams.

.PARAMETER UnsetChatWidget
Removes the Chat widget registry value, effectively enabling it since that is the default.

.PARAMETER AllUsers
Applies the Chat widget setting to all user profiles on the machine.

.PARAMETER DisableOfficeTeamsInstall
Disable Office's ability to install Teams.

.PARAMETER EnableOfficeTeamsInstall
Enable Office's ability to install Teams.

.PARAMETER UnsetOfficeTeamsInstall
Removes the Office Teams registry value, effectively enabling it since that is the default.

.EXAMPLE
UninstallTeams -DisableChatWidget
Disables the Chat widget (Win+C) for Microsoft Teams.

.EXAMPLE
UninstallTeams -EnableChatWidget
Enables the Chat widget (Win+C) for Microsoft Teams.

.EXAMPLE
UninstallTeams -UnsetChatWidget
Removes the Chat widget value, effectively enabling it since that is the default.

.EXAMPLE
UninstallTeams -DisableChatWidget -AllUsers
Disables the Chat widget (Win+C) for Microsoft Teams for all user profiles on the machine.

.EXAMPLE
UninstallTeams -EnableChatWidget -AllUsers
Enables the Chat widget (Win+C) for Microsoft Teams for all user profiles on the machine.

.EXAMPLE
UninstallTeams -UnsetChatWidget -AllUsers
Removes the Chat widget value, effectively enabling it since that is the default, for all user profiles on the machine.

.EXAMPLE
UninstallTeams -DisableOfficeTeamsInstall
Disable Office's ability to install Teams.

.EXAMPLE
UninstallTeams -EnableOfficeTeamsInstall
Enable Office's ability to install Teams.

.EXAMPLE
UninstallTeams -UnsetOfficeTeamsInstall
Removes the Office Teams registry value, effectively enabling it since that is the default.

.NOTES
Version  : 1.2.5
Created by   : asheroto

.LINK
Project Site: https://github.com/asheroto/UninstallTeams

#>

#Requires -RunAsAdministrator

[CmdletBinding()]
param (
    [switch]$EnableChatWidget,
    [switch]$DisableChatWidget,
    [switch]$UnsetChatWidget,
    [switch]$EnableOfficeTeamsInstall,
    [switch]$DisableOfficeTeamsInstall,
    [switch]$UnsetOfficeTeamsInstall,
    [switch]$AllUsers,
    [switch]$Version,
    [switch]$Help,
    [switch]$CheckForUpdate
)

# Version
$CurrentVersion = '1.2.5'
$RepoOwner = 'asheroto'
$RepoName = 'UninstallTeams'
$PowerShellGalleryName = 'UninstallTeams'

# Versions
$ProgressPreference = 'SilentlyContinue' # Suppress progress bar (makes downloading super fast)
$ConfirmPreference = 'None' # Suppress confirmation prompts

# Display version if -Version is specified
if ($Version.IsPresent) {
    $CurrentVersion
    exit 0
}

# Display full help if -Help is specified
if ($Help) {
    Get-Help -Name $MyInvocation.MyCommand.Source -Full
    exit 0
}

# Display $PSVersionTable and Get-Host if -Verbose is specified
if ($PSBoundParameters.ContainsKey('Verbose') -and $PSBoundParameters['Verbose']) {
    $PSVersionTable
    Get-Host
}

function Get-GitHubRelease {
    <#
        .SYNOPSIS
        Fetches the latest release information of a GitHub repository.

        .DESCRIPTION
        This function uses the GitHub API to get information about the latest release of a specified repository, including its version and the date it was published.

        .PARAMETER Owner
        The GitHub username of the repository owner.

        .PARAMETER Repo
        The name of the repository.

        .EXAMPLE
        Get-GitHubRelease -Owner "asheroto" -Repo "winget-install"
        This command retrieves the latest release version and published datetime of the winget-install repository owned by asheroto.
    #>
    [CmdletBinding()]
    param (
        [string]$Owner,
        [string]$Repo
    )
    try {
        $url = "https://api.github.com/repos/$Owner/$Repo/releases/latest"
        $response = Invoke-RestMethod -Uri $url -ErrorAction Stop

        $latestVersion = $response.tag_name
        $publishedAt = $response.published_at

        # Convert UTC time string to local time
        $UtcDateTime = [DateTime]::Parse($publishedAt, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind)
        $PublishedLocalDateTime = $UtcDateTime.ToLocalTime()

        [PSCustomObject]@{
            LatestVersion     = $latestVersion
            PublishedDateTime = $PublishedLocalDateTime
        }
    } catch {
        Write-Error "Unable to check for updates.`nError: $_"
        exit 1
    }
}

function CheckForUpdate {
    param (
        [string]$RepoOwner,
        [string]$RepoName,
        [version]$CurrentVersion,
        [string]$PowerShellGalleryName
    )

    $Data = Get-GitHubRelease -Owner $RepoOwner -Repo $RepoName

    if ($Data.LatestVersion -gt $CurrentVersion) {
        Write-Output "`nA new version of $RepoName is available.`n"
        Write-Output "Current version: $CurrentVersion."
        Write-Output "Latest version: $($Data.LatestVersion)."
        Write-Output "Published at: $($Data.PublishedDateTime).`n"
        Write-Output "You can download the latest version from https://github.com/$RepoOwner/$RepoName/releases`n"
        if ($PowerShellGalleryName) {
            Write-Output "Or you can run the following command to update:"
            Write-Output "Install-Script $PowerShellGalleryName -Force`n"
        }
    } else {
        Write-Output "`n$RepoName is up to date.`n"
        Write-Output "Current version: $CurrentVersion."
        Write-Output "Latest version: $($Data.LatestVersion)."
        Write-Output "Published at: $($Data.PublishedDateTime)."
        Write-Output "`nRepository: https://github.com/$RepoOwner/$RepoName/releases`n"
    }
    exit 0
}

function Get-ChatWidgetStatus {
    param (
        [switch]$AllUsers
    )

    if ($AllUsers) {
        $RegistryPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Chat"
    } else {
        $RegistryPath = "HKCU:\SOFTWARE\Policies\Microsoft\Windows\Windows Chat"
    }

    if (Test-Path $RegistryPath) {
        $ChatIconValue = (Get-ItemProperty -Path $RegistryPath -Name "ChatIcon" -ErrorAction SilentlyContinue).ChatIcon
        if ($null -eq $ChatIconValue) {
            return "Unset (default is enabled)"
        } elseif ($ChatIconValue -eq 1) {
            return "Enabled"
        } elseif ($ChatIconValue -eq 2) {
            return "Hidden"
        } elseif ($ChatIconValue -eq 3) {
            return "Disabled"
        }
    }

    return "Unset (default is enabled)"
}

function Set-ChatWidgetStatus {
    param (
        [switch]$EnableChatWidget,
        [switch]$DisableChatWidget,
        [switch]$UnsetChatWidget,
        [switch]$AllUsers
    )

    if ($AllUsers) {
        $RegistryPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Chat"
    } else {
        $RegistryPath = "HKCU:\SOFTWARE\Policies\Microsoft\Windows\Windows Chat"
    }

    if ($EnableChatWidget) {
        $WhatChanged = "enabled"
        if (Test-Path $RegistryPath) {
            Set-ItemProperty -Path $RegistryPath -Name "ChatIcon" -Value 1 -Type DWord -Force
        } else {
            New-Item -Path $RegistryPath | Out-Null
            Set-ItemProperty -Path $RegistryPath -Name "ChatIcon" -Value 1 -Type DWord -Force
        }
    } elseif ($DisableChatWidget) {
        $WhatChanged = "disabled"
        if (Test-Path $RegistryPath) {
            Set-ItemProperty -Path $RegistryPath -Name "ChatIcon" -Value 3 -Type DWord -Force
        } else {
            New-Item -Path $RegistryPath | Out-Null
            Set-ItemProperty -Path $RegistryPath -Name "ChatIcon" -Value 3 -Type DWord -Force
        }
    } elseif ($UnsetChatWidget) {
        $WhatChanged = "unset"
        if (Test-Path $RegistryPath) {
            Remove-ItemProperty -Path $RegistryPath -Name "ChatIcon" -ErrorAction SilentlyContinue
        }
    }

    if ($AllUsers) {
        $AllUsersString = "all users"
    } else {
        $AllUsersString = "the current user"
    }
    Write-Output "Chat widget has been $WhatChanged for $AllUsersString."
}

function Get-OfficeTeamsInstallStatus {
    # According to Microsoft, HKLM is the only key that matters for this (no HKCU)
    $RegistryPath = "HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\OfficeUpdate"

    if (Test-Path $RegistryPath) {
        $OfficeTeamsInstallValue = (Get-ItemProperty -Path $RegistryPath).PreventTeamsInstall
        if ($null -eq $OfficeTeamsInstallValue) {
            return "Unset (default is enabled)"
        } elseif ($OfficeTeamsInstallValue -eq 0) {
            return "Enabled"
        } elseif ($OfficeTeamsInstallValue -eq 1) {
            return "Disabled"
        }
    }

    return "Unset (default is enabled)"
}

function Set-OfficeTeamsInstallStatus {
    param (
        [switch]$EnableOfficeTeamsInstall,
        [switch]$DisableOfficeTeamsInstall,
        [switch]$UnsetOfficeTeamsInstall
    )

    $RegistryPath = "HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\OfficeUpdate"

    if (-Not (Test-Path $RegistryPath)) {
        Write-Output "Creating registry path $RegistryPath."
        New-Item -Path $RegistryPath -Force | Out-Null
    }

    if ($EnableOfficeTeamsInstall) {
        $WhatChanged = "enabled"
        Set-ItemProperty -Path $RegistryPath -Name "PreventTeamsInstall" -Value 0 -Type DWord -Force
    } elseif ($DisableOfficeTeamsInstall) {
        $WhatChanged = "disabled"
        Set-ItemProperty -Path $RegistryPath -Name "PreventTeamsInstall" -Value 1 -Type DWord -Force
    } elseif ($UnsetOfficeTeamsInstall) {
        $WhatChanged = "unset (default is enabled)"
        Remove-ItemProperty -Path $RegistryPath -Name "PreventTeamsInstall" -ErrorAction SilentlyContinue
    }

    Write-Output "Office's ability to install Teams has been $WhatChanged."
}

function Check-GitHubRelease {
    param (
        [string]$Owner,
        [string]$Repo
    )
    try {
        $url = "https://api.github.com/repos/$Owner/$Repo/releases/latest"
        $response = Invoke-RestMethod -Uri $url -ErrorAction Stop

        $latestVersion = $response.tag_name
        $publishedAt = $response.published_at

        [PSCustomObject]@{
            LatestVersion = $latestVersion
            PublishedAt   = $publishedAt
        }
    } catch {
        Write-Error "Unable to check for updates. Error: $_"
        exit 1
    }
}

function Write-Section($text) {
    <#
        .SYNOPSIS
        Prints a text block surrounded by a section divider for enhanced output readability.

        .DESCRIPTION
        This function takes a string input and prints it to the console, surrounded by a section divider made of hash characters.
        It is designed to enhance the readability of console output.

        .PARAMETER text
        The text to be printed within the section divider.

        .EXAMPLE
        Write-Section "Downloading Files..."
        This command prints the text "Downloading Files..." surrounded by a section divider.
    #>
    Write-Output ""
    Write-Output ("#" * ($text.Length + 4))
    Write-Output "# $text #"
    Write-Output ("#" * ($text.Length + 4))
    Write-Output ""
}

# Get uninstall registry keys
function Get-UninstallRegistryKey {
    param (
        [string]$Match
    )

    $result = @()
    $uninstallKeys = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKCU:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )

    foreach ($key in $uninstallKeys) {
        if (Test-Path $key) {
            Get-Item $key | Get-ChildItem | Where-Object {
                $_.GetValue("DisplayName") -like "*${Match}*"
            } | ForEach-Object {
                $result += $_.PSPath
            }
        }
    }

    return $result
}

# Get uninstall string from registry key
function Get-UninstallString {
    param (
        [string]$Match
    )

    $result = @()
    $registryKeys = Get-UninstallRegistryKey -Match $Match

    foreach ($regKey in $registryKeys) {
        try {
            $displayName = (Get-ItemProperty -Path $regKey).DisplayName
            $uninstallString = (Get-ItemProperty -Path $regKey).UninstallString

            if ($displayName -and $uninstallString) {
                $obj = [PSCustomObject]@{
                    DisplayName     = $displayName
                    UninstallString = $uninstallString
                }
                $result += $obj
            }
        } catch {

        }
    }

    return $result
}

function Remove-Shortcut {
    param (
        [string]$ShortcutName,
        [string]$ShortcutPathName,
        [string]$UserPath,
        [string]$PublicPath
    )

    try {
        $userShortcutPath = [System.IO.Path]::Combine($UserPath, "$ShortcutName.lnk")
        $publicShortcutPath = [System.IO.Path]::Combine($PublicPath, "$ShortcutName.lnk")

        if (Test-Path -Path $userShortcutPath) {
            Write-Output "Deleting $ShortcutName from the user's $ShortcutPathName..."
            Remove-Item -Path $userShortcutPath
        }

        if (Test-Path -Path $publicShortcutPath) {
            Write-Output "Deleting $ShortcutName from the public $ShortcutPathName..."
            Remove-Item -Path $publicShortcutPath
        }
    } catch {
        Write-Output "An error occurred while attempting to delete the shortcut."
    }
}

function Remove-DesktopShortcuts {
    param (
        [string]$ShortcutName
    )

    $userDesktopPath = [System.Environment]::GetFolderPath('Desktop')
    $publicDesktopPath = [System.Environment]::GetFolderPath('CommonDesktopDirectory')

    Remove-Shortcut -ShortcutPathName "Desktop" -ShortcutName $ShortcutName -UserPath $userDesktopPath -PublicPath $publicDesktopPath
}

function Remove-StartMenuShortcuts {
    param (
        [string]$ShortcutName
    )

    $userStartMenuPath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath('StartMenu'), "Programs")
    $publicStartMenuPath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath('CommonStartMenu'), "Programs")

    Remove-Shortcut -ShortcutPathName "Start Menu" -ShortcutName $ShortcutName -UserPath $userStartMenuPath -PublicPath $publicStartMenuPath
}

# ============================================================================ #
# Initial checks
# ============================================================================ #

# Check for updates if -CheckForUpdate is specified
if ($CheckForUpdate) {
    CheckForUpdate -RepoOwner $RepoOwner -RepoName $RepoName -CurrentVersion $CurrentVersion -PowerShellGalleryName $PowerShellGalleryName
}

# Check if exactly one or none of -EnableChatWidget, -DisableChatWidget, or -UnsetChatWidget is specified
$chatWidgetCount = ($EnableChatWidget, $DisableChatWidget, $UnsetChatWidget).Where({ $_ }).Count

if ($chatWidgetCount -gt 1) {
    Write-Warning "Please choose only one of -EnableChatWidget, -DisableChatWidget, or -UnsetChatWidget."
    exit 1
}

# Check if -AllUsers is specified without one of -EnableChatWidget, -DisableChatWidget, or -UnsetChatWidget
if ($AllUsers -and $chatWidgetCount -eq 0) {
    Write-Error "The -AllUsers switch can only be used with -EnableChatWidget, -DisableChatWidget, or -UnsetChatWidget. UninstallTeams will always remove Teams for the local machine."
    exit 1
}

# Similar checks for -EnableOfficeTeamsInstall, -DisableOfficeTeamsInstall, or -UnsetOfficeTeamsInstall
$officeTeamsInstallCount = ($EnableOfficeTeamsInstall, $DisableOfficeTeamsInstall, $UnsetOfficeTeamsInstall).Where({ $_ }).Count

if ($officeTeamsInstallCount -gt 1) {
    Write-Warning "Please choose only one of -EnableOfficeTeamsInstall, -DisableOfficeTeamsInstall, or -UnsetOfficeTeamsInstall."
    exit 1
}

try {
    # Spacer
    Write-Output ""

    # Heading
    Write-Output "UninstallTeams $CurrentVersion"
    Write-Output "To check for updates, run UninstallTeams -CheckForUpdate"

    # Spacer
    Write-Output ""

    # Default
    $Uninstall = $true

    # Chat widget
    if ($EnableChatWidget) {
        Set-ChatWidgetStatus -EnableChatWidget -AllUsers:$AllUsers
        $Uninstall = $false
    } elseif ($DisableChatWidget) {
        Set-ChatWidgetStatus -DisableChatWidget -AllUsers:$AllUsers
        $Uninstall = $false
    } elseif ($UnsetChatWidget) {
        Set-ChatWidgetStatus -UnsetChatWidget -AllUsers:$AllUsers
        $Uninstall = $false
    }

    # Office Teams install
    if ($EnableOfficeTeamsInstall) {
        Set-OfficeTeamsInstallStatus -EnableOfficeTeamsInstall
        $Uninstall = $false
    } elseif ($DisableOfficeTeamsInstall) {
        Set-OfficeTeamsInstallStatus -DisableOfficeTeamsInstall
        $Uninstall = $false
    } elseif ($UnsetOfficeTeamsInstall) {
        Set-OfficeTeamsInstallStatus -UnsetOfficeTeamsInstall
        $Uninstall = $false
    }

    # Uninstall Teams
    if ($Uninstall -eq $true) {
        # Stopping Teams process
        Write-Output "Stopping Teams process..."
        Stop-Process -Name "*teams*" -Force -ErrorAction SilentlyContinue

        ###########################################################################
        # Start the process of uninstalling Teams
        Write-Output "Deleting Teams through uninstall registry key..."

        # Retrieve the uninstall information for Teams
        $uninstallInfo = Get-UninstallString -Match "Teams"

        foreach ($info in $uninstallInfo) {
            $uninstallString = $info.UninstallString

            if (-not [string]::IsNullOrWhiteSpace($uninstallString)) {
                Write-Debug "Found Teams uninstall string: $uninstallString"

                # Check if the uninstall string is an MSI command
                if ($uninstallString -match "msiexec.exe\s*/[XxIi]\{([^\}]+)\}") {
                    $productGUID = $matches[1]
                    Write-Debug "Found Teams product GUID: $productGUID"

                    # Construct the MSI uninstall command with the correct format for GUID
                    $filePath = "msiexec.exe"
                    $argList = "/x {${productGUID}} /qn"  # Correct format for GUID
                } else {
                    # For non-MSI packages, assume the uninstall string is a complete command
                    $filePath = $uninstallString.Split(" ")[0]
                    $argList = $uninstallString.Substring($filePath.Length).Trim()
                }

                # Execute the uninstall command
                if ($filePath -ieq "msiexec.exe" -or (Test-Path $filePath)) {
                    Write-Debug "Uninstalling Teams with command: $filePath $argList"
                    $proc = Start-Process -FilePath $filePath -ArgumentList $argList -PassThru
                    $proc.WaitForExit()
                } else {
                    Write-Warning "The path $filePath does not exist."
                }
            }
        }
        ###########################################################################

        # Uninstall from "Teams Installer"
        $TeamsPrgFiles = Join-Path ${env:ProgramFiles(x86)} "Teams Installer\Teams.exe"
        Write-Output "Checking Teams in `"$TeamsPrgFiles`..."
        if (Test-Path $TeamsPrgFiles) {
            Write-Output "Uninstalling Teams from `"$TeamsPrgFiles`..."
            $proc = Start-Process -FilePath $TeamsPrgFiles -ArgumentList "--uninstall" -PassThru
            $proc.WaitForExit()
        }

        # Uninstall from AppData\Microsoft\Teams
        $TeamsUpdateExePath = Join-Path $env:APPDATA "Microsoft\Teams\Update.exe"
        Write-Output "Checking Teams in `"$TeamsUpdateExePath`"..."
        if (Test-Path $TeamsUpdateExePath) {
            Write-Output "Uninstalling Teams from `"$TeamsUpdateExePath`"..."
            $proc = Start-Process -FilePath $TeamsUpdateExePath -ArgumentList "-uninstall -s" -PassThru
            $proc.WaitForExit()
        }

        # Uninstall from Program Files (x86)\Microsoft\Teams\current
        $TeamsUpdateExePathPrgX86 = Join-Path ${env:ProgramFiles(x86)} "Microsoft\Teams\current\Update.exe"
        Write-Output "Checking Teams in `"$TeamsUpdateExePathPrgX86`"..."
        if (Test-Path $TeamsUpdateExePathPrgX86) {
            Write-Output "Uninstalling Teams from `"$TeamsUpdateExePathPrgX86`"..."
            $proc = Start-Process -FilePath $TeamsUpdateExePathPrgX86 -ArgumentList "-uninstall -s" -PassThru
            $proc.WaitForExit()
        }

        # Remove via AppxPackage
        Write-Output "Removing Teams AppxPackage..."
        Get-AppxPackage "*Teams*" | Remove-AppxPackage -ErrorAction SilentlyContinue
        Get-AppxPackage "*Teams*" -AllUsers | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue

        # Delete Microsoft Teams directory
        $MicrosoftTeamsPath = Join-Path $env:LOCALAPPDATA "Microsoft Teams"
        Write-Output "Deleting `"$MicrosoftTeamsPath`"..."
        if (Test-Path $MicrosoftTeamsPath) {
            Remove-Item -Path $MicrosoftTeamsPath -Force -Recurse -ErrorAction SilentlyContinue
        }

        # Delete Teams directory
        $TeamsPath = Join-Path $env:LOCALAPPDATA "Microsoft\Teams"
        Write-Output "Deleting `"$TeamsPath`"..."
        if (Test-Path $TeamsPath) {
            Remove-Item -Path $TeamsPath -Force -Recurse -ErrorAction SilentlyContinue
        }

        # Remove from startup registry key
        Write-Output "Deleting Teams startup registry keys..."
        Remove-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run' -Name 'Teams', 'TeamsMachineUninstallerLocalAppData', 'TeamsMachineUninstallerProgramData', 'com.squirrel.Teams.Teams', 'TeamsMachineInstaller' -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        Remove-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run' -Name 'Teams', 'TeamsMachineUninstallerLocalAppData', 'TeamsMachineUninstallerProgramData', 'com.squirrel.Teams.Teams', 'TeamsMachineInstaller' -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        Remove-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run' -Name 'Teams', 'TeamsMachineUninstallerLocalAppData', 'TeamsMachineUninstallerProgramData', 'com.squirrel.Teams.Teams', 'TeamsMachineInstaller' -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        Remove-ItemProperty -Path 'HKCU:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run' -Name 'Teams', 'TeamsMachineUninstallerLocalAppData', 'TeamsMachineUninstallerProgramData', 'com.squirrel.Teams.Teams', 'TeamsMachineInstaller' -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

        # Remove from AutorunsDisabled registry key
        Write-Output "Deleting Teams startup registry keys from AutorunsDisabled..."
        Remove-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\AutorunsDisabled' -Name 'Teams', 'TeamsMachineUninstallerLocalAppData', 'TeamsMachineUninstallerProgramData', 'com.squirrel.Teams.Teams', 'TeamsMachineInstaller' -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        Remove-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run\AutorunsDisabled' -Name 'Teams', 'TeamsMachineUninstallerLocalAppData', 'TeamsMachineUninstallerProgramData', 'com.squirrel.Teams.Teams', 'TeamsMachineInstaller' -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        Remove-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\AutorunsDisabled' -Name 'Teams', 'TeamsMachineUninstallerLocalAppData', 'TeamsMachineUninstallerProgramData', 'com.squirrel.Teams.Teams', 'TeamsMachineInstaller' -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        Remove-ItemProperty -Path 'HKCU:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run\AutorunsDisabled' -Name 'Teams', 'TeamsMachineUninstallerLocalAppData', 'TeamsMachineUninstallerProgramData', 'com.squirrel.Teams.Teams', 'TeamsMachineInstaller' -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

        # Remove Teams uninstall registry keys
        Write-Output "Deleting Teams uninstall registry keys..."
        Remove-Item -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Teams" -Force -Recurse -ErrorAction SilentlyContinue
        Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Teams" -Force -Recurse -ErrorAction SilentlyContinue

        # Removing desktop shortcuts
        Write-Output "Deleting Teams desktop shortcuts..."
        Remove-DesktopShortcuts -ShortcutName "Microsoft Teams"

        # Removing start menu shortcuts
        Write-Output "Deleting Teams start menu shortcuts..."
        Remove-StartMenuShortcuts -ShortcutName "Microsoft Teams"
        Remove-StartMenuShortcuts -ShortcutName "Microsoft Teams classic (work or school)"

        # Removing Teams meeting addin
        Write-Output "Deleting Teams meeting addin..."
        $teamsMeetingAddin = "$env:LOCALAPPDATA\Microsoft\TeamsMeetingAddin"
        if (Test-Path $teamsMeetingAddin) {
            Remove-Item -Path $teamsMeetingAddin -Force -Recurse -ErrorAction SilentlyContinue
        }

        # Removing Teams meeting addin
        Write-Output "Deleting Teams presence addin..."
        $teamsPresenceAddin = "$env:LOCALAPPDATA\Microsoft\TeamsPresenceAddin"
        if (Test-Path $teamsPresenceAddin) {
            Remove-Item -Path $teamsPresenceAddin -Force -Recurse -ErrorAction SilentlyContinue
        }
    }
} catch {
    Write-Warning "An error occurred during the Teams uninstallation process: $_"
}

# Let user know nothing will change
if ($Uninstall -eq $true) {
    Write-Output ""
    Write-Output "Teams has been uninstalled, please restart your computer."
    Write-Output ""
    Write-Output "The information below is only information, the settings below will not change unless you use parameters to change them."
}

# Output the Chat widget status of both the current user and the local machine
$CurrentUserStatus = Get-ChatWidgetStatus
$LocalMachineStatus = Get-ChatWidgetStatus -AllUsers

# Determine the effective status
if ($CurrentUserStatus -ne "Unset (default is enabled)") {
    $effectiveStatus = $CurrentUserStatus
} elseif ($LocalMachineStatus -ne "Unset (default is enabled)") {
    $effectiveStatus = $LocalMachineStatus
} else {
    $effectiveStatus = "Enabled by default"
    Write-Output "Both Current User and Local Machine statuses are Unset (default is enabled). Enabled by default."
}

Write-Section("Chat widget")
Write-Output "Current User Status: $CurrentUserStatus"
Write-Output "Local Machine Status: $LocalMachineStatus"
Write-Output "Effective Status: $effectiveStatus"
Write-Output ""

# If Chat widget status is "Enabled" or "Enabled by default", show a warning
if ($effectiveStatus -eq "Enabled" -or $effectiveStatus -eq "Enabled by default") {
    Write-Warning "Teams Chat widget is enabled. Teams could be reinstalled if the user clicks 'Continue' after using Win+C or by clicking the Chat icon in the taskbar (if enabled). Use the '-DisableChatWidget' or '-DisableChatWidget -AllUsers' switch to disable it. Current user takes precedence unless unset. Use 'Get-Help UninstallTeams -Full' for more information."
}

# Output the Office Teams install status
$OfficeTeamsInstallStatus = Get-OfficeTeamsInstallStatus

# Chat widget status
Write-Section("Office's ability to install Teams")
Write-Output "Status: $OfficeTeamsInstallStatus"
Write-Output ""

# If Office Team install status is Enabled or unset, show a warning
if (($OfficeTeamsInstallStatus -eq "Enabled") -or ($OfficeTeamsInstallStatus -eq "Unset (default is enabled)")) {
    Write-Warning "Office is allowing Teams to install. Teams could be reinstalled if Office is installed or updated.`nUse the '-DisableOfficeTeamsInstall' switch to prevent Teams from installing with Office. Use 'Get-Help UninstallTeams -Full' for more information."
}

# Office note
Write-Section("Office Note")
Write-Output "If you just installed Microsoft Office, you may need to restart the computer once or`ntwice and then run UninstallTeams to prevent Teams from reinstalling."

# Spacer
Write-Output ""
# SIG # Begin signature block
# MIIhGQYJKoZIhvcNAQcCoIIhCjCCIQYCAQExDzANBglghkgBZQMEAgIFADCBiQYK
# KwYBBAGCNwIBBKB7MHkwNAYKKwYBBAGCNwIBHjAmAgMBAAAEEB/MO2BZSwhOtyTS
# xil+81ECAQACAQACAQACAQACAQAwQTANBglghkgBZQMEAgIFAAQwi10gIDOEquTL
# Dv15Ws3/yY0QUalGXSRG+gddxj7bE2AawYPzUcm72YxAQo9fbEAUoIIHZDCCA1kw
# ggLfoAMCAQICEA+4p0C5FY0DUUO8WdnwQCkwCgYIKoZIzj0EAwMwYTELMAkGA1UE
# BhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2lj
# ZXJ0LmNvbTEgMB4GA1UEAxMXRGlnaUNlcnQgR2xvYmFsIFJvb3QgRzMwHhcNMjEw
# NDI5MDAwMDAwWhcNMzYwNDI4MjM1OTU5WjBkMQswCQYDVQQGEwJVUzEXMBUGA1UE
# ChMORGlnaUNlcnQsIEluYy4xPDA6BgNVBAMTM0RpZ2lDZXJ0IEdsb2JhbCBHMyBD
# b2RlIFNpZ25pbmcgRUNDIFNIQTM4NCAyMDIxIENBMTB2MBAGByqGSM49AgEGBSuB
# BAAiA2IABLu0rCelSA2iU1+PLoE+L1N2uAiUopqqiouYtbHw/CoVu7mzpSIv/WrA
# veJVaGBrlzTBZlNxI/wa1cogDwJAoqNKWkajkVMrlfID6aum04d2L+dkn541UfzD
# YzV4duT4d6OCAVcwggFTMBIGA1UdEwEB/wQIMAYBAf8CAQAwHQYDVR0OBBYEFJtf
# sDa6nQauGSe9wKAiwIuLOHftMB8GA1UdIwQYMBaAFLPbSKT5ocXYrjZBzBFjaWIp
# vEvGMA4GA1UdDwEB/wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDAzB2BggrBgEF
# BQcBAQRqMGgwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBA
# BggrBgEFBQcwAoY0aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0
# R2xvYmFsUm9vdEczLmNydDBCBgNVHR8EOzA5MDegNaAzhjFodHRwOi8vY3JsMy5k
# aWdpY2VydC5jb20vRGlnaUNlcnRHbG9iYWxSb290RzMuY3JsMBwGA1UdIAQVMBMw
# BwYFZ4EMAQMwCAYGZ4EMAQQBMAoGCCqGSM49BAMDA2gAMGUCMHi9SZVlcQHQRldo
# ZQ5oqdw2CMHu/dSO20BlPw3/k6/CrmOGo37LtJFaeOwHA2cHfAIxAOefH/EHW6w0
# xji8taVQzubqOH4+eZDkpFurAg3oB/xWplqK3bNQst3y+mZ0ntAWYzCCBAMwggOJ
# oAMCAQICEAExw+sKUABDj0yZt5afTZQwCgYIKoZIzj0EAwMwZDELMAkGA1UEBhMC
# VVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMTwwOgYDVQQDEzNEaWdpQ2VydCBH
# bG9iYWwgRzMgQ29kZSBTaWduaW5nIEVDQyBTSEEzODQgMjAyMSBDQTEwHhcNMjQw
# MzA3MDAwMDAwWhcNMjUwMzA4MjM1OTU5WjBvMQswCQYDVQQGEwJVUzERMA8GA1UE
# CBMIT2tsYWhvbWExETAPBgNVBAcTCE11c2tvZ2VlMRwwGgYDVQQKExNBc2hlciBT
# b2x1dGlvbnMgSW5jMRwwGgYDVQQDExNBc2hlciBTb2x1dGlvbnMgSW5jMHYwEAYH
# KoZIzj0CAQYFK4EEACIDYgAExsP0nyCZ1QtY7aXin+tdZVcF0uPHJJjRpjVVgUmb
# 3iKJeKapvWBSAbroBouKIP9+Qoz197aNbZCSOBQsWX53SUyTu1Trvwku7ksL+eQh
# bJvnRJ20UqF566z5KbniyLrAo4IB8zCCAe8wHwYDVR0jBBgwFoAUm1+wNrqdBq4Z
# J73AoCLAi4s4d+0wHQYDVR0OBBYEFNdgDYHKEBunNDYgivfxKeS4YX0/MD4GA1Ud
# IAQ3MDUwMwYGZ4EMAQQBMCkwJwYIKwYBBQUHAgEWG2h0dHA6Ly93d3cuZGlnaWNl
# cnQuY29tL0NQUzAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMw
# gasGA1UdHwSBozCBoDBOoEygSoZIaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0Rp
# Z2lDZXJ0R2xvYmFsRzNDb2RlU2lnbmluZ0VDQ1NIQTM4NDIwMjFDQTEuY3JsME6g
# TKBKhkhodHRwOi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRHbG9iYWxHM0Nv
# ZGVTaWduaW5nRUNDU0hBMzg0MjAyMUNBMS5jcmwwgY4GCCsGAQUFBwEBBIGBMH8w
# JAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBXBggrBgEFBQcw
# AoZLaHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0R2xvYmFsRzND
# b2RlU2lnbmluZ0VDQ1NIQTM4NDIwMjFDQTEuY3J0MAkGA1UdEwQCMAAwCgYIKoZI
# zj0EAwMDaAAwZQIxAJHtFqbIBTSZ6AiYEyHsjjlZ7treTZfTSPiyyr8KAKBPKVXt
# B2859Jj8A3c9lEXrLgIwGTu2YV8DhFy9OqIDwkCZfoYH8oMo1LRtYhYZtVzkr3WF
# er8mkmAdOyNbW/DI0pZPMYIY+jCCGPYCAQEweDBkMQswCQYDVQQGEwJVUzEXMBUG
# A1UEChMORGlnaUNlcnQsIEluYy4xPDA6BgNVBAMTM0RpZ2lDZXJ0IEdsb2JhbCBH
# MyBDb2RlIFNpZ25pbmcgRUNDIFNIQTM4NCAyMDIxIENBMQIQATHD6wpQAEOPTJm3
# lp9NlDANBglghkgBZQMEAgIFAKCBjDAQBgorBgEEAYI3AgEMMQIwADAZBgkqhkiG
# 9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIB
# FTA/BgkqhkiG9w0BCQQxMgQwd9w8xu0Rix5tYLzcKTzF7kDRIlbzDk0cU81vzV0d
# gwrFxdQlXRQmCdJXhWhc3Yg8MAsGByqGSM49AgEFAARnMGUCMGRkibrNiwWZ7AuT
# hiAUzn/bdDuljghidZDFafipkTJ5F0k7uiE6KO7Cgc2dnz2uvgIxAK4Tq1w7QzSa
# 0Be6zvZ8Usc4q83u65CJJK+qTqxWYwKyNXLwZbmFADbz7AQV6YADO6GCF2Ewghdd
# BgorBgEEAYI3AwMBMYIXTTCCF0kGCSqGSIb3DQEHAqCCFzowghc2AgEDMQ8wDQYJ
# YIZIAWUDBAICBQAwgYgGCyqGSIb3DQEJEAEEoHkEdzB1AgEBBglghkgBhv1sBwEw
# QTANBglghkgBZQMEAgIFAAQwhJ6YKGPVftGbw/McHUTyuQNYJK/vTH78phkhOn/Y
# hsOkVSCxpSgdScCutaGW0ANBAhEA33/WtWKASA+jxKF8JD5dFBgPMjAyNDA4MDEw
# NzAwMzFaoIITCTCCBsIwggSqoAMCAQICEAVEr/OUnQg5pr/bP1/lYRYwDQYJKoZI
# hvcNAQELBQAwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMu
# MTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2IFRp
# bWVTdGFtcGluZyBDQTAeFw0yMzA3MTQwMDAwMDBaFw0zNDEwMTMyMzU5NTlaMEgx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjEgMB4GA1UEAxMX
# RGlnaUNlcnQgVGltZXN0YW1wIDIwMjMwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAw
# ggIKAoICAQCjU0WHHYOOW6w+VLMj4M+f1+XS512hDgncL0ijl3o7Kpxn3GIVWMGp
# kxGnzaqyat0QKYoeYmNp01icNXG/OpfrlFCPHCDqx5o7L5Zm42nnaf5bw9YrIBzB
# l5S0pVCB8s/LB6YwaMqDQtr8fwkklKSCGtpqutg7yl3eGRiF+0XqDWFsnf5xXsQG
# mjzwxS55DxtmUuPI1j5f2kPThPXQx/ZILV5FdZZ1/t0QoRuDwbjmUpW1R9d4KTlr
# 4HhZl+NEK0rVlc7vCBfqgmRN/yPjyobutKQhZHDr1eWg2mOzLukF7qr2JPUdvJsc
# srdf3/Dudn0xmWVHVZ1KJC+sK5e+n+T9e3M+Mu5SNPvUu+vUoCw0m+PebmQZBzcB
# kQ8ctVHNqkxmg4hoYru8QRt4GW3k2Q/gWEH72LEs4VGvtK0VBhTqYggT02kefGRN
# nQ/fztFejKqrUBXJs8q818Q7aESjpTtC/XN97t0K/3k0EH6mXApYTAA+hWl1x4Nk
# 1nXNjxJ2VqUk+tfEayG66B80mC866msBsPf7Kobse1I4qZgJoXGybHGvPrhvltXh
# EBP+YUcKjP7wtsfVx95sJPC/QoLKoHE9nJKTBLRpcCcNT7e1NtHJXwikcKPsCvER
# LmTgyyIryvEoEyFJUX4GZtM7vvrrkTjYUQfKlLfiUKHzOtOKg8tAewIDAQABo4IB
# izCCAYcwDgYDVR0PAQH/BAQDAgeAMAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/BAww
# CgYIKwYBBQUHAwgwIAYDVR0gBBkwFzAIBgZngQwBBAIwCwYJYIZIAYb9bAcBMB8G
# A1UdIwQYMBaAFLoW2W1NhS9zKXaaL3WMaiCPnshvMB0GA1UdDgQWBBSltu8T5+/N
# 0GSh1VapZTGj3tXjSTBaBgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsMy5kaWdp
# Y2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRSU0E0MDk2U0hBMjU2VGltZVN0YW1w
# aW5nQ0EuY3JsMIGQBggrBgEFBQcBAQSBgzCBgDAkBggrBgEFBQcwAYYYaHR0cDov
# L29jc3AuZGlnaWNlcnQuY29tMFgGCCsGAQUFBzAChkxodHRwOi8vY2FjZXJ0cy5k
# aWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRSU0E0MDk2U0hBMjU2VGltZVN0
# YW1waW5nQ0EuY3J0MA0GCSqGSIb3DQEBCwUAA4ICAQCBGtbeoKm1mBe8cI1Pijxo
# nNgl/8ss5M3qXSKS7IwiAqm4z4Co2efjxe0mgopxLxjdTrbebNfhYJwr7e09SI64
# a7p8Xb3CYTdoSXej65CqEtcnhfOOHpLawkA4n13IoC4leCWdKgV6hCmYtld5j9sm
# Viuw86e9NwzYmHZPVrlSwradOKmB521BXIxp0bkrxMZ7z5z6eOKTGnaiaXXTUORE
# Er4gDZ6pRND45Ul3CFohxbTPmJUaVLq5vMFpGbrPFvKDNzRusEEm3d5al08zjdSN
# d311RaGlWCZqA0Xe2VC1UIyvVr1MxeFGxSjTredDAHDezJieGYkD6tSRN+9NUvPJ
# YCHEVkft2hFLjDLDiOZY4rbbPvlfsELWj+MXkdGqwFXjhr+sJyxB0JozSqg21Lly
# ln6XeThIX8rC3D0y33XWNmdaifj2p8flTzU8AL2+nCpseQHc2kTmOt44OwdeOVj0
# fHMxVaCAEcsUDH6uvP6k63llqmjWIso765qCNVcoFstp8jKastLYOrixRoZruhf9
# xHdsFWyuq69zOuhJRrfVf8y2OMDY7Bz1tqG4QyzfTkx9HmhwwHcK1ALgXGC7KP84
# 5VJa1qwXIiNO9OzTF/tQa/8Hdx9xl0RBybhG02wyfFgvZ0dl5Rtztpn5aywGRu9B
# HvDwX+Db2a2QgESvgBBBijCCBq4wggSWoAMCAQICEAc2N7ckVHzYR6z9KGYqXlsw
# DQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0
# IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNl
# cnQgVHJ1c3RlZCBSb290IEc0MB4XDTIyMDMyMzAwMDAwMFoXDTM3MDMyMjIzNTk1
# OVowYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMTswOQYD
# VQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2IFRpbWVTdGFt
# cGluZyBDQTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAMaGNQZJs8E9
# cklRVcclA8TykTepl1Gh1tKD0Z5Mom2gsMyD+Vr2EaFEFUJfpIjzaPp985yJC3+d
# H54PMx9QEwsmc5Zt+FeoAn39Q7SE2hHxc7Gz7iuAhIoiGN/r2j3EF3+rGSs+Qtxn
# jupRPfDWVtTnKC3r07G1decfBmWNlCnT2exp39mQh0YAe9tEQYncfGpXevA3eZ9d
# rMvohGS0UvJ2R/dhgxndX7RUCyFobjchu0CsX7LeSn3O9TkSZ+8OpWNs5KbFHc02
# DVzV5huowWR0QKfAcsW6Th+xtVhNef7Xj3OTrCw54qVI1vCwMROpVymWJy71h6aP
# TnYVVSZwmCZ/oBpHIEPjQ2OAe3VuJyWQmDo4EbP29p7mO1vsgd4iFNmCKseSv6De
# 4z6ic/rnH1pslPJSlRErWHRAKKtzQ87fSqEcazjFKfPKqpZzQmiftkaznTqj1QPg
# v/CiPMpC3BhIfxQ0z9JMq++bPf4OuGQq+nUoJEHtQr8FnGZJUlD0UfM2SU2LINIs
# VzV5K6jzRWC8I41Y99xh3pP+OcD5sjClTNfpmEpYPtMDiP6zj9NeS3YSUZPJjAw7
# W4oiqMEmCPkUEBIDfV8ju2TjY+Cm4T72wnSyPx4JduyrXUZ14mCjWAkBKAAOhFTu
# zuldyF4wEr1GnrXTdrnSDmuZDNIztM2xAgMBAAGjggFdMIIBWTASBgNVHRMBAf8E
# CDAGAQH/AgEAMB0GA1UdDgQWBBS6FtltTYUvcyl2mi91jGogj57IbzAfBgNVHSME
# GDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0l
# BAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBpMCQGCCsGAQUFBzABhhhodHRw
# Oi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUHMAKGNWh0dHA6Ly9jYWNlcnRz
# LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3J0MEMGA1UdHwQ8
# MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0
# ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EMAQQCMAsGCWCGSAGG/WwHATAN
# BgkqhkiG9w0BAQsFAAOCAgEAfVmOwJO2b5ipRCIBfmbW2CFC4bAYLhBNE88wU86/
# GPvHUF3iSyn7cIoNqilp/GnBzx0H6T5gyNgL5Vxb122H+oQgJTQxZ822EpZvxFBM
# Yh0MCIKoFr2pVs8Vc40BIiXOlWk/R3f7cnQU1/+rT4osequFzUNf7WC2qk+RZp4s
# nuCKrOX9jLxkJodskr2dfNBwCnzvqLx1T7pa96kQsl3p/yhUifDVinF2ZdrM8HKj
# I/rAJ4JErpknG6skHibBt94q6/aesXmZgaNWhqsKRcnfxI2g55j7+6adcq/Ex8HB
# anHZxhOACcS2n82HhyS7T6NJuXdmkfFynOlLAlKnN36TU6w7HQhJD5TNOXrd/yVj
# mScsPT9rp/Fmw0HNT7ZAmyEhQNC3EyTN3B14OuSereU0cZLXJmvkOHOrpgFPvT87
# eK1MrfvElXvtCl8zOYdBeHo46Zzh3SP9HSjTx/no8Zhf+yvYfvJGnXUsHicsJttv
# FXseGYs2uJPU5vIXmVnKcPA3v5gA3yAWTyf7YGcWoWa63VXAOimGsJigK+2VQbc6
# 1RWYMbRiCQ8KvYHZE/6/pNHzV9m8BPqC3jLfBInwAM1dwvnQI38AC+R2AibZ8GV2
# QqYphwlHK+Z/GqSFD/yYlvZVVCsfgPrA8g4r5db7qS9EFUrnEw4d2zc4GqEr9u3W
# fPwwggWNMIIEdaADAgECAhAOmxiO+dAt5+/bUOIIQBhaMA0GCSqGSIb3DQEBDAUA
# MGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsT
# EHd3dy5kaWdpY2VydC5jb20xJDAiBgNVBAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQg
# Um9vdCBDQTAeFw0yMjA4MDEwMDAwMDBaFw0zMTExMDkyMzU5NTlaMGIxCzAJBgNV
# BAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdp
# Y2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRydXN0ZWQgUm9vdCBHNDCCAiIw
# DQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAL/mkHNo3rvkXUo8MCIwaTPswqcl
# LskhPfKK2FnC4SmnPVirdprNrnsbhA3EMB/zG6Q4FutWxpdtHauyefLKEdLkX9YF
# PFIPUh/GnhWlfr6fqVcWWVVyr2iTcMKyunWZanMylNEQRBAu34LzB4TmdDttceIt
# DBvuINXJIB1jKS3O7F5OyJP4IWGbNOsFxl7sWxq868nPzaw0QF+xembud8hIqGZX
# V59UWI4MK7dPpzDZVu7Ke13jrclPXuU15zHL2pNe3I6PgNq2kZhAkHnDeMe2scS1
# ahg4AxCN2NQ3pC4FfYj1gj4QkXCrVYJBMtfbBHMqbpEBfCFM1LyuGwN1XXhm2Tox
# RJozQL8I11pJpMLmqaBn3aQnvKFPObURWBf3JFxGj2T3wWmIdph2PVldQnaHiZdp
# ekjw4KISG2aadMreSx7nDmOu5tTvkpI6nj3cAORFJYm2mkQZK37AlLTSYW3rM9nF
# 30sEAMx9HJXDj/chsrIRt7t/8tWMcCxBYKqxYxhElRp2Yn72gLD76GSmM9GJB+G9
# t+ZDpBi4pncB4Q+UDCEdslQpJYls5Q5SUUd0viastkF13nqsX40/ybzTQRESW+UQ
# UOsxxcpyFiIJ33xMdT9j7CFfxCBRa2+xq4aLT8LWRV+dIPyhHsXAj6KxfgommfXk
# aS+YHS312amyHeUbAgMBAAGjggE6MIIBNjAPBgNVHRMBAf8EBTADAQH/MB0GA1Ud
# DgQWBBTs1+OC0nFdZEzfLmc/57qYrhwPTzAfBgNVHSMEGDAWgBRF66Kv9JLLgjEt
# UYunpyGd823IDzAOBgNVHQ8BAf8EBAMCAYYweQYIKwYBBQUHAQEEbTBrMCQGCCsG
# AQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQwYIKwYBBQUHMAKGN2h0
# dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RD
# QS5jcnQwRQYDVR0fBD4wPDA6oDigNoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29t
# L0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDARBgNVHSAECjAIMAYGBFUdIAAw
# DQYJKoZIhvcNAQEMBQADggEBAHCgv0NcVec4X6CjdBs9thbX979XB72arKGHLOyF
# XqkauyL4hxppVCLtpIh3bb0aFPQTSnovLbc47/T/gLn4offyct4kvFIDyE7QKt76
# LVbP+fT3rDB6mouyXtTP0UNEm0Mh65ZyoUi0mcudT6cGAxN3J0TU53/oWajwvy8L
# punyNDzs9wPHh6jSTEAZNUZqaVSwuKFWjuyk1T3osdz9HNj0d1pcVIxv76FQPfx2
# CWiEn2/K2yCNNWAcAgPLILCsWKAOQGPFmCLBsln1VWvPJ6tsds5vIy30fnFqI2si
# /xK4VC0nftg62fC2h5b9W9FcrBjDTZ9ztwGpn1eqXijiuZQxggOGMIIDggIBATB3
# MGMxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjE7MDkGA1UE
# AxMyRGlnaUNlcnQgVHJ1c3RlZCBHNCBSU0E0MDk2IFNIQTI1NiBUaW1lU3RhbXBp
# bmcgQ0ECEAVEr/OUnQg5pr/bP1/lYRYwDQYJYIZIAWUDBAICBQCggeEwGgYJKoZI
# hvcNAQkDMQ0GCyqGSIb3DQEJEAEEMBwGCSqGSIb3DQEJBTEPFw0yNDA4MDEwNzAw
# MzFaMCsGCyqGSIb3DQEJEAIMMRwwGjAYMBYEFGbwKzLCwskPgl3OqorJxk8ZnM9A
# MDcGCyqGSIb3DQEJEAIvMSgwJjAkMCIEINL25G3tdCLM0dRAV2hBNm+CitpVmq4z
# Fq9NGprUDHgoMD8GCSqGSIb3DQEJBDEyBDAwnfrhigwoeBnHeV0olbxN5cSK2FQ1
# 8yguLuyjK1lmfJybAkGj7r6aB4et0d4TSoUwDQYJKoZIhvcNAQEBBQAEggIAhuwb
# iULkZM9g6tTGUrHGLF7FWnqDsIwfyeLgZ42goUKEP6u66B/ySuj2l1MvORm2dIIl
# xC3D35NGwRuBg8twboumipoU9Rbpq5TVgcB3qmd36JxIwyioIhnN96/avwdtQQT6
# FuUVfXIP784NSkjMitftJUpLgdqHiuOpXg3jxN3vaS8ztkGVec3MRFEFlmKP21HG
# cRWHdizs4bxGeThXaJAzGCLbW6URbCLFlH1m9/MXnXaN+STgrb5FsL9JI/JRhbev
# zhqts6Qr3tU+H3FsI3DgAFQ41JfcaQw3pdr8juzgMyLEJh6+QZ2KRjZ3DvVlsMsE
# w6FTGVj9Fw+8bgqA6IiOFv2AlGIVR6NDVLMIYufcwLKWvwlKw3Djo6WvaRx1W6W5
# xfhd1jtqMI3iQ2opzS0HrWksKSGzn0jNMIirh26fjQcv7wlUzIwvsW41fKzpkXr4
# QWKvYe8QzIC4p58jqWHqCErxnLBvitc5Do71E5ow/SAZx4AEAx93fn+pq3nEtLP/
# W3hG1VUXEA7a0Ph+eykr4h2MgJm4RxTN1/xMvXzYiUbufC9/l61BaIrcPX7xQ3iB
# QYCmA19poY/WxGX+EzITv6of6ZwTh+RQl5bWyeOWWKYy3M75mCNVYyJ4+NbobbZx
# gP4MoSSGF3ecVFom33VqUmGPh/lfiIUq5SThIzg=
# SIG # End signature block
