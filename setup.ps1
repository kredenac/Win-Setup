#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Windows Setup Script - Automated configuration for new Windows installations
.DESCRIPTION
    Configures Windows settings, installs essential software, and sets up development environment
.NOTES
    Must be run as Administrator in PowerShell 7+
#>

# Color output functions
function Write-Success { param($Message) Write-Host "[SUCCESS] $Message" -ForegroundColor Green }
function Write-Info { param($Message) Write-Host "[INFO] $Message" -ForegroundColor Cyan }
function Write-Warning { param($Message) Write-Host "[WARNING] $Message" -ForegroundColor Yellow }
function Write-ErrorMsg { param($Message) Write-Host "[ERROR] $Message" -ForegroundColor Red }

# Track success/failure
$script:successCount = 0
$script:failureCount = 0
$script:warningCount = 0

function Invoke-Step {
    param(
        [string]$Name,
        [scriptblock]$Action
    )

    Write-Info "Running: $Name"
    try {
        & $Action
        Write-Success "$Name completed"
        $script:successCount++
    }
    catch {
        Write-ErrorMsg "$Name failed: $_"
        $script:failureCount++
    }
}

#region Pre-flight Checks
Write-Host "`n========================================" -ForegroundColor Magenta
Write-Host "  Windows Setup Script" -ForegroundColor Magenta
Write-Host "========================================`n" -ForegroundColor Magenta

# Load configuration
$configPath = Join-Path $PSScriptRoot "config.json"
if (Test-Path $configPath) {
    Write-Info "Loading configuration from config.json"
    $config = Get-Content $configPath | ConvertFrom-Json
} else {
    Write-Warning "config.json not found, using defaults"
    $config = @{
        git = @{
            username = "kredenac"
            email = "zacementirano@gmail.com"
        }
        gaming = $false
    }
    $script:warningCount++
}

Write-Info "Git Username: $($config.git.username)"
Write-Info "Git Email: $($config.git.email)"
Write-Info "Gaming Mode: $($config.gaming)"
Write-Host ""
#endregion

#region Registry Modifications & System Settings (FAST - Run First)
Write-Host "`n--- WINDOWS SETTINGS & REGISTRY ---`n" -ForegroundColor Yellow

Invoke-Step "Set taskbar alignment to left" {
    $registryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    Set-ItemProperty -Path $registryPath -Name "TaskbarAl" -Value 0 -Type DWord -Force
}

Invoke-Step "Enable dark mode" {
    # Apps dark mode
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -Value 0 -Type DWord -Force
    # System dark mode
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "SystemUsesLightTheme" -Value 0 -Type DWord -Force
}

Invoke-Step "Show hidden files in File Explorer" {
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "Hidden" -Value 1 -Type DWord -Force
}

Invoke-Step "Show file extensions in File Explorer" {
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "HideFileExt" -Value 0 -Type DWord -Force
}

Invoke-Step "Disable web search results in Start Menu" {
    $registryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search"
    if (-not (Test-Path $registryPath)) {
        New-Item -Path $registryPath -Force | Out-Null
    }
    Set-ItemProperty -Path $registryPath -Name "BingSearchEnabled" -Value 0 -Type DWord -Force
    Set-ItemProperty -Path $registryPath -Name "CortanaConsent" -Value 0 -Type DWord -Force
}

Invoke-Step "Remove search highlights from taskbar" {
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "SearchboxTaskbarMode" -Value 0 -Type DWord -Force
}

Invoke-Step "Remove widgets from taskbar" {
    $registryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    Set-ItemProperty -Path $registryPath -Name "TaskbarDa" -Value 0 -Type DWord -Force
}

Invoke-Step "Remove Meet Now from taskbar" {
    $registryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    if (-not (Test-Path $registryPath)) {
        New-Item -Path $registryPath -Force | Out-Null
    }
    Set-ItemProperty -Path $registryPath -Name "HideSCAMeetNow" -Value 1 -Type DWord -Force
}

Invoke-Step "Hide Music folder from File Explorer sidebar" {
    $registryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions\{a0c69a99-21c8-4671-8703-7934162fcf1d}\PropertyBag"
    if (-not (Test-Path $registryPath)) {
        New-Item -Path $registryPath -Force | Out-Null
    }
    Set-ItemProperty -Path $registryPath -Name "ThisPCPolicy" -Value "Hide" -Type String -Force
}

Invoke-Step "Show This PC icon on desktop" {
    $registryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"
    if (-not (Test-Path $registryPath)) {
        New-Item -Path $registryPath -Force | Out-Null
    }
    Set-ItemProperty -Path $registryPath -Name "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" -Value 0 -Type DWord -Force
}

Invoke-Step "Set short date format to d/M/yyyy" {
    Set-ItemProperty -Path "HKCU:\Control Panel\International" -Name "sShortDate" -Value "d/M/yyyy" -Type String -Force
}

Invoke-Step "Set first day of week to Monday" {
    Set-ItemProperty -Path "HKCU:\Control Panel\International" -Name "iFirstDayOfWeek" -Value "0" -Type String -Force
}

Invoke-Step "Set PowerShell execution policy to RemoteSigned" {
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
}

Invoke-Step "Disable UAC" {
    $registryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
    Set-ItemProperty -Path $registryPath -Name "EnableLUA" -Value 0 -Type DWord -Force
}

Invoke-Step "Enable Windows Developer Mode" {
    $registryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock"
    if (-not (Test-Path $registryPath)) {
        New-Item -Path $registryPath -Force | Out-Null
    }
    Set-ItemProperty -Path $registryPath -Name "AllowDevelopmentWithoutDevLicense" -Value 1 -Type DWord -Force
}

Invoke-Step "Disable hibernation" {
    powercfg /hibernate off
}

Invoke-Step "Set monitor refresh rate to maximum" {
    # Get all displays
    $displays = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorBasicDisplayParams
    foreach ($display in $displays) {
        try {
            # Get current resolution
            $currentMode = Get-DisplayResolution
            if ($currentMode) {
                # Get max refresh rate for current resolution
                $maxRefreshRate = (Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorListedSupportedSourceModes |
                    Where-Object { $_.HorizontalActivePixels -eq $currentMode.Width -and $_.VerticalActivePixels -eq $currentMode.Height } |
                    Measure-Object -Property RefreshRate -Maximum).Maximum

                if ($maxRefreshRate -and $maxRefreshRate -gt $currentMode.RefreshRate) {
                    Set-DisplayResolution -Width $currentMode.Width -Height $currentMode.Height -RefreshRate $maxRefreshRate
                    Write-Info "Set refresh rate to $maxRefreshRate Hz"
                }
            }
        }
        catch {
            Write-Warning "Could not automatically set refresh rate. Please set manually in Display Settings."
            $script:warningCount++
        }
    }
}

Invoke-Step "Add Serbian (Latin) keyboard layout" {
    $currentList = Get-WinUserLanguageList
    $serbianLatin = New-WinUserLanguageList -Language "sr-Latn-RS"

    # Check if Serbian Latin is already added
    $alreadyExists = $currentList | Where-Object { $_.LanguageTag -eq "sr-Latn-RS" }

    if (-not $alreadyExists) {
        $currentList += $serbianLatin
        Set-WinUserLanguageList -LanguageList $currentList -Force
    } else {
        Write-Info "Serbian (Latin) keyboard already installed"
    }
}

Invoke-Step "Unpin Copilot from taskbar" {
    # Disable Copilot button in taskbar
    $registryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    Set-ItemProperty -Path $registryPath -Name "ShowCopilotButton" -Value 0 -Type DWord -Force
}

Invoke-Step "Unpin Microsoft Store from taskbar" {
    $appName = "Microsoft Store"
    try {
        ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() |
            Where-Object { $_.Name -eq $appName }).Verbs() |
            Where-Object { $_.Name.replace('&', '') -match 'Unpin from taskbar' } |
            ForEach-Object { $_.DoIt() }
    }
    catch {
        Write-Warning "Could not unpin Microsoft Store automatically"
        $script:warningCount++
    }
}
#endregion

#region Software Installation (SLOW - Run in Parallel)
Write-Host "`n--- SOFTWARE INSTALLATION (Parallel) ---`n" -ForegroundColor Yellow

Write-Info "Starting parallel software installations..."

# Define software to install in parallel
$softwareJobs = @()

$parallelInstalls = @(
    @{ Name = "Google Chrome"; Id = "Google.Chrome" }
    @{ Name = "Visual Studio Code"; Id = "Microsoft.VisualStudioCode" }
    @{ Name = "VLC Media Player"; Id = "VideoLAN.VLC" }
    @{ Name = "Everything Search"; Id = "voidtools.Everything" }
    @{ Name = "7-Zip"; Id = "7zip.7zip" }
    @{ Name = "Python"; Id = "Python.Python.3.12" }
    @{ Name = "PowerToys"; Id = "Microsoft.PowerToys" }
)

# Add gaming software if enabled
if ($config.gaming -eq $true) {
    $parallelInstalls += @{ Name = "Steam"; Id = "Valve.Steam" }
    $parallelInstalls += @{ Name = "Discord"; Id = "Discord.Discord" }
}

# Start all installations as background jobs
foreach ($software in $parallelInstalls) {
    $job = Start-Job -ScriptBlock {
        param($Id, $Name)
        $result = winget install --id $Id --silent --accept-source-agreements --accept-package-agreements 2>&1
        return @{
            Name = $Name
            Success = $LASTEXITCODE -eq 0
            Output = $result
        }
    } -ArgumentList $software.Id, $software.Name

    $softwareJobs += @{
        Job = $job
        Name = $software.Name
    }
}

Write-Info "Waiting for $($softwareJobs.Count) parallel installations to complete..."

# Wait for all jobs and report results
foreach ($jobInfo in $softwareJobs) {
    $result = Receive-Job -Job $jobInfo.Job -Wait
    Remove-Job -Job $jobInfo.Job

    if ($result.Success -or $result.Output -like "*already installed*" -or $result.Output -like "*No available upgrade*") {
        Write-Success "$($jobInfo.Name) completed"
        $script:successCount++
    } else {
        Write-ErrorMsg "$($jobInfo.Name) failed"
        $script:failureCount++
    }
}

Write-Host ""

# Sequential installations (dependencies required)
Write-Host "--- SOFTWARE INSTALLATION (Sequential) ---`n" -ForegroundColor Yellow

Invoke-Step "Install Git" {
    winget install --id Git.Git --silent --accept-source-agreements --accept-package-agreements
}

Invoke-Step "Install nvm-windows" {
    winget install --id CoreyButler.NVMforWindows --silent --accept-source-agreements --accept-package-agreements
}

# Install Node.js LTS via nvm (needs to run after nvm installation)
Invoke-Step "Install Node.js LTS via nvm" {
    # Refresh environment variables to get nvm in PATH
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")

    # Check if nvm is available
    $nvmPath = "$env:APPDATA\nvm\nvm.exe"
    if (Test-Path $nvmPath) {
        & $nvmPath install lts
        & $nvmPath use lts
    } else {
        throw "nvm not found in expected location. May need to restart terminal."
    }
}
#endregion

#region Git Configuration
Write-Host "`n--- GIT CONFIGURATION ---`n" -ForegroundColor Yellow

# Refresh PATH to ensure git is available
$env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")

Invoke-Step "Configure Git username" {
    git config --global user.name "$($config.git.username)"
}

Invoke-Step "Configure Git email" {
    git config --global user.email "$($config.git.email)"
}

Invoke-Step "Set Git default branch to main" {
    git config --global init.defaultBranch main
}

Invoke-Step "Set Git default editor to VS Code" {
    git config --global core.editor "code --wait"
}

Invoke-Step "Configure Git credential manager" {
    git config --global credential.helper manager-core
}
#endregion

#region PowerShell Profile Configuration
Write-Host "`n--- POWERSHELL PROFILE ---`n" -ForegroundColor Yellow

Invoke-Step "Install posh-git module" {
    Install-Module -Name posh-git -Scope CurrentUser -Force -AllowClobber
}

Invoke-Step "Configure PowerShell profile with Git aliases" {
    $profileContent = @'
# Import posh-git for Git prompt integration
Import-Module posh-git

# Git shortcuts
function Get-GitCommit { & git add -A; git commit -m $args }
New-Alias -Name gac -Value Get-GitCommit

function Get-GitStatus { & git status }
New-Alias -Name gs -Value Get-GitStatus

function Get-GitMerge { & git fetch; git merge origin/main }
New-Alias -Name gfm -Value Get-GitMerge

function GitSquashUnpushed {
    param([string]$Message)

    $upstream = git rev-parse --abbrev-ref --symbolic-full-name "@{u}" 2>$null
    if (-not $upstream) {
        Write-Error "No upstream branch set. Push first or set upstream with 'git push --set-upstream origin <branch>'"
        return
    }

    $unpulled = [int](git rev-list --count "$upstream..HEAD")
    if ($unpulled -lt 2) {
        Write-Host "Nothing to squash ($unpulled unpushed commit)."
        return
    }

    git reset --soft "HEAD~$unpulled"
    git commit -m "$Message"
    Write-Host "Squashed $unpulled commits into one."
}

New-Alias -Name gsq -Value GitSquashUnpushed

# Custom prompt with posh-git integration and Windows Terminal support
function prompt
{
    $loc = Get-Location

    $prompt = & $GitPromptScriptBlock

    $prompt += "$([char]27)]9;12$([char]7)"
    if ($loc.Provider.Name -eq "FileSystem")
    {
        $prompt += "$([char]27)]9;9;`"$($loc.ProviderPath)`"$([char]27)\"
    }

    $prompt
}
'@

    # Get PowerShell profile path
    $profilePath = $PROFILE.CurrentUserAllHosts

    # Create profile directory if it doesn't exist
    $profileDir = Split-Path -Parent $profilePath
    if (-not (Test-Path $profileDir)) {
        New-Item -Path $profileDir -ItemType Directory -Force | Out-Null
    }

    # Write profile content
    Set-Content -Path $profilePath -Value $profileContent -Force
    Write-Info "PowerShell profile configured at: $profilePath"
}

Invoke-Step "Configure Windows Terminal keybindings" {
    $settingsPath = "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json"

    if (Test-Path $settingsPath) {
        # Read and parse JSON
        $settings = Get-Content $settingsPath -Raw | ConvertFrom-Json

        # Ensure keybindings array exists
        if (-not $settings.keybindings) {
            $settings | Add-Member -MemberType NoteProperty -Name "keybindings" -Value @()
        }

        # Remove existing bindings for these keys if they exist
        $settings.keybindings = @($settings.keybindings | Where-Object {
            $_.keys -ne "ctrl+shift+d" -and $_.keys -ne "ctrl+shift+s"
        })

        # Add new keybindings
        $settings.keybindings += @(
            @{
                id = "Terminal.DuplicatePaneRight"
                keys = "ctrl+shift+d"
            },
            @{
                id = "Terminal.DuplicatePaneDown"
                keys = "ctrl+shift+s"
            }
        )

        # Save settings back to file
        $settings | ConvertTo-Json -Depth 100 | Set-Content $settingsPath -Encoding UTF8
        Write-Success "Windows Terminal keybindings configured successfully!"
    } else {
        Write-Warning "Windows Terminal settings file not found. Install Windows Terminal first."
        $script:warningCount++
    }
}
#endregion

#region Summary
Write-Host "`n========================================" -ForegroundColor Magenta
Write-Host "  SETUP COMPLETE" -ForegroundColor Magenta
Write-Host "========================================`n" -ForegroundColor Magenta

Write-Host "Summary:" -ForegroundColor Cyan
Write-Success "$script:successCount tasks completed successfully"
if ($script:warningCount -gt 0) {
    Write-Warning "$script:warningCount warnings"
}
if ($script:failureCount -gt 0) {
    Write-ErrorMsg "$script:failureCount tasks failed"
}

Write-Host "`n" -NoNewline
Write-Host "IMPORTANT: " -ForegroundColor Yellow -NoNewline
Write-Host "Some changes require a restart to take effect."
Write-Host "This includes: UAC settings, Developer Mode, taskbar changes, date/time format`n"

$restart = Read-Host "Would you like to restart now? (Y/N)"
if ($restart -eq "Y" -or $restart -eq "y") {
    Write-Info "Restarting in 10 seconds... (Press Ctrl+C to cancel)"
    Start-Sleep -Seconds 10
    Restart-Computer -Force
} else {
    Write-Info "Please restart your computer when convenient to apply all changes."
}
#endregion
