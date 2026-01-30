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
# Handle both local execution and remote execution (irm | iex)
if ([string]::IsNullOrEmpty($PSScriptRoot)) {
    # Remote execution - no config file available
    Write-Warning "Running in remote mode (no config.json), using defaults"
    $config = @{
        git = @{
            username = "kredenac"
            email = "zacementirano@gmail.com"
        }
        gaming = $false
    }
    $script:warningCount++
} else {
    # Local execution - try to load config.json
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
}

Write-Info "Git Username: $($config.git.username)"
Write-Info "Git Email: $($config.git.email)"
Write-Info "Gaming Mode: $($config.gaming)"

# Check PowerShell version and install PowerShell 7 if needed
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Warning "Running on Windows PowerShell $($PSVersionTable.PSVersion.Major).$($PSVersionTable.PSVersion.Minor)"
    Write-Warning "This script requires PowerShell 7+. Installing PowerShell 7..."

    # Check if winget is available in Windows PowerShell
    $wingetAvailable = $null -ne (Get-Command winget -ErrorAction SilentlyContinue)

    if ($wingetAvailable) {
        Write-Info "Installing PowerShell 7 via winget..."
        try {
            winget install --id Microsoft.PowerShell --silent --accept-source-agreements --accept-package-agreements
            Write-Success "PowerShell 7 installed successfully!"
            Write-Host "`n" -NoNewline
            Write-Host "IMPORTANT: " -ForegroundColor Yellow -NoNewline
            Write-Host "Please close this Windows PowerShell window and re-run this script in PowerShell 7."
            Write-Host "You can find PowerShell 7 in the Start Menu as 'PowerShell 7' or run 'pwsh' from the command line.`n"
            return
        }
        catch {
            Write-ErrorMsg "Failed to install PowerShell 7 via winget: $_"
            Write-Warning "Falling back to direct download method..."
        }
    }

    # Fallback: Download and install PS7 directly (for Windows Sandbox or when winget fails)
    if (-not $wingetAvailable -or $LASTEXITCODE -ne 0) {
        Write-Info "Downloading PowerShell 7 from GitHub..."
        try {
            $ps7Version = "7.4.7"
            $url = "https://github.com/PowerShell/PowerShell/releases/download/v$ps7Version/PowerShell-$ps7Version-win-x64.msi"
            $output = "$env:TEMP\PowerShell-$ps7Version-win-x64.msi"

            # Download the MSI
            Invoke-WebRequest -Uri $url -OutFile $output -UseBasicParsing
            Write-Success "Downloaded PowerShell 7 installer"

            # Install silently
            Write-Info "Installing PowerShell 7..."
            Start-Process msiexec.exe -ArgumentList "/i `"$output`" /quiet /norestart" -Wait -NoNewWindow

            Write-Success "PowerShell 7 installed successfully!"
            Write-Host "`n" -NoNewline
            Write-Host "IMPORTANT: " -ForegroundColor Yellow -NoNewline
            Write-Host "Please close this Windows PowerShell window and re-run this script in PowerShell 7."
            Write-Host "You can find PowerShell 7 in the Start Menu as 'PowerShell 7' or run 'pwsh' from the command line.`n"
            return
        }
        catch {
            Write-ErrorMsg "Failed to install PowerShell 7: $_"
            Write-Host "Please install PowerShell 7 manually from: https://aka.ms/powershell"
            $global:LASTEXITCODE = 1
            return
        }
    }
}

Write-Info "Running on PowerShell $($PSVersionTable.PSVersion)"

# Check if winget is available
$wingetAvailable = $null -ne (Get-Command winget -ErrorAction SilentlyContinue)
if (-not $wingetAvailable) {
    Write-Warning "winget (Windows Package Manager) is not installed or not in PATH"
    Write-Info "Attempting to install winget..."

    try {
        # Download and install VCLibs dependency
        Write-Info "Downloading VCLibs dependency..."
        $vcLibsUrl = "https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx"
        $vcLibsPath = "$env:TEMP\Microsoft.VCLibs.x64.14.00.Desktop.appx"
        Invoke-WebRequest -Uri $vcLibsUrl -OutFile $vcLibsPath -UseBasicParsing
        Add-AppxPackage -Path $vcLibsPath
        Write-Success "VCLibs dependency installed"

        # Download and install UI.Xaml dependency (often needed)
        Write-Info "Downloading UI.Xaml dependency..."
        $uiXamlUrl = "https://github.com/microsoft/microsoft-ui-xaml/releases/download/v2.8.6/Microsoft.UI.Xaml.2.8.x64.appx"
        $uiXamlPath = "$env:TEMP\Microsoft.UI.Xaml.2.8.x64.appx"
        Invoke-WebRequest -Uri $uiXamlUrl -OutFile $uiXamlPath -UseBasicParsing
        Add-AppxPackage -Path $uiXamlPath
        Write-Success "UI.Xaml dependency installed"

        # Download and install Windows App Runtime dependency
        Write-Info "Downloading Windows App Runtime dependency..."
        # Using the 1.8.x runtime that matches the requirement
        $appRuntimeUrl = "https://github.com/microsoft/WindowsAppSDK/releases/download/v1.8.1/Microsoft.WindowsAppRuntime.1.8-x64.msix"
        $appRuntimePath = "$env:TEMP\Microsoft.WindowsAppRuntime.1.8-x64.msix"
        Invoke-WebRequest -Uri $appRuntimeUrl -OutFile $appRuntimePath -UseBasicParsing
        Add-AppxPackage -Path $appRuntimePath
        Write-Success "Windows App Runtime dependency installed"

        # Download and install App Installer (includes winget)
        Write-Info "Downloading App Installer (winget)..."
        $wingetUrl = "https://aka.ms/getwinget"
        $wingetPath = "$env:TEMP\Microsoft.DesktopAppInstaller.msixbundle"
        Invoke-WebRequest -Uri $wingetUrl -OutFile $wingetPath -UseBasicParsing
        Add-AppxPackage -Path $wingetPath
        Write-Success "App Installer installed"

        # Refresh PATH and check if winget is now available
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
        $wingetAvailable = $null -ne (Get-Command winget -ErrorAction SilentlyContinue)

        if (-not $wingetAvailable) {
            throw "winget was installed but is not available in PATH. Try restarting your terminal."
        }

        Write-Success "winget is now available"
    }
    catch {
        Write-ErrorMsg "Failed to install winget: $_"
        Write-ErrorMsg "winget is required for software installations. Cannot continue."
        Write-Host "Please install winget manually from: https://aka.ms/getwinget"
        $script:failureCount++
        return
    }
}

# Install Windows Terminal if not already installed
$terminalInstalled = Get-AppxPackage -Name Microsoft.WindowsTerminal -ErrorAction SilentlyContinue
if (-not $terminalInstalled) {
    Write-Info "Windows Terminal not found. Installing..."
    try {
        winget install --id Microsoft.WindowsTerminal --silent --accept-source-agreements --accept-package-agreements
        if ($LASTEXITCODE -eq 0) {
            Write-Success "Windows Terminal installed successfully"
        } else {
            Write-Warning "Windows Terminal installation may have failed (exit code: $LASTEXITCODE)"
            $script:warningCount++
        }
    }
    catch {
        Write-Warning "Failed to install Windows Terminal: $_"
        $script:warningCount++
    }
} else {
    Write-Info "Windows Terminal is already installed"
}

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
    try {
        Set-ItemProperty -Path $registryPath -Name "TaskbarDa" -Value 0 -Type DWord -Force -ErrorAction Stop
    }
    catch [System.UnauthorizedAccessException] {
        Write-Warning "Permission denied for TaskbarDa registry key (common in VMs)"
        $script:warningCount++
    }
}

Invoke-Step "Disable Task View from taskbar" {
    $registryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    Set-ItemProperty -Path $registryPath -Name "ShowTaskViewButton" -Value 0 -Type DWord -Force
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
    try {
        # Get all displays (may not be supported in VMs)
        $displays = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorBasicDisplayParams -ErrorAction Stop
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
    catch {
        # CIM operations not supported (common in VMs)
        Write-Warning "Monitor refresh rate detection not supported in this environment (VM or limited hardware access)"
        $script:warningCount++
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

Invoke-Step "Unpin Microsoft Edge from taskbar" {
    $appName = "Microsoft Edge"
    try {
        ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() |
            Where-Object { $_.Name -eq $appName }).Verbs() |
            Where-Object { $_.Name.replace('&', '') -match 'Unpin from taskbar' } |
            ForEach-Object { $_.DoIt() }
    }
    catch {
        Write-Warning "Could not unpin Microsoft Edge automatically"
        $script:warningCount++
    }
}
#endregion

#region Software Installation (SLOW - Run in Parallel)
Write-Host "`n--- SOFTWARE INSTALLATION (Parallel) ---`n" -ForegroundColor Yellow

# Skip software installation if winget is not available
if (-not $wingetAvailable) {
    Write-Warning "Skipping software installations (winget not available)"
    Write-Host ""
    # Skip to next section
} else {
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
}

# Sequential installations (dependencies required)
Write-Host "--- SOFTWARE INSTALLATION (Sequential) ---`n" -ForegroundColor Yellow

if (-not $wingetAvailable) {
    Write-Warning "Skipping sequential software installations (winget not available)"
    Write-Host ""
} else {
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
}
#endregion

#region Git Configuration
Write-Host "`n--- GIT CONFIGURATION ---`n" -ForegroundColor Yellow

# Refresh PATH to ensure git is available
$env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")

# Check if git is available
$gitAvailable = $null -ne (Get-Command git -ErrorAction SilentlyContinue)
if (-not $gitAvailable) {
    Write-Warning "Git is not installed or not in PATH. Skipping Git configuration."
    Write-Host ""
} else {
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
}
#endregion

#region Taskbar Pinning
Write-Host "`n--- TASKBAR PINNING ---`n" -ForegroundColor Yellow

# Helper function to pin app to taskbar
function Pin-ToTaskbar {
    param(
        [string]$AppName,
        [string]$AppPath
    )

    if (Test-Path $AppPath) {
        try {
            $shell = New-Object -ComObject Shell.Application
            $folder = $shell.Namespace((Split-Path $AppPath))
            $item = $folder.ParseName((Split-Path $AppPath -Leaf))
            $verb = $item.Verbs() | Where-Object { $_.Name -match 'Pin to taskbar' }
            if ($verb) {
                $verb.DoIt()
                Write-Info "Pinned $AppName to taskbar"
                return $true
            } else {
                Write-Warning "Could not find 'Pin to taskbar' option for $AppName"
                return $false
            }
        }
        catch {
            Write-Warning "Failed to pin $AppName to taskbar: $_"
            return $false
        }
    } else {
        Write-Warning "$AppName not found at $AppPath"
        return $false
    }
}

Invoke-Step "Pin Chrome to taskbar" {
    # Common Chrome installation paths
    $chromePaths = @(
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
        "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
    )

    $pinned = $false
    foreach ($path in $chromePaths) {
        if (Test-Path $path) {
            $pinned = Pin-ToTaskbar "Chrome" $path
            break
        }
    }

    if (-not $pinned) {
        throw "Chrome not found in expected locations"
    }
}

Invoke-Step "Pin Windows Terminal to taskbar" {
    # Windows Terminal is a UWP app, pinning works differently
    try {
        $shell = New-Object -ComObject Shell.Application
        $folder = $shell.Namespace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}')
        $item = $folder.Items() | Where-Object { $_.Name -eq "Terminal" -or $_.Name -eq "Windows Terminal" }
        if ($item) {
            $verb = $item.Verbs() | Where-Object { $_.Name -match 'Pin to taskbar' }
            if ($verb) {
                $verb.DoIt()
                Write-Info "Pinned Windows Terminal to taskbar"
            } else {
                throw "Could not find 'Pin to taskbar' option"
            }
        } else {
            throw "Windows Terminal not found"
        }
    }
    catch {
        Write-Warning "Failed to pin Windows Terminal: $_"
        $script:warningCount++
    }
}

Invoke-Step "Pin Visual Studio Code to taskbar" {
    # Common VS Code installation paths
    $vscodePaths = @(
        "$env:LOCALAPPDATA\Programs\Microsoft VS Code\Code.exe",
        "$env:ProgramFiles\Microsoft VS Code\Code.exe",
        "${env:ProgramFiles(x86)}\Microsoft VS Code\Code.exe"
    )

    $pinned = $false
    foreach ($path in $vscodePaths) {
        if (Test-Path $path) {
            $pinned = Pin-ToTaskbar "VS Code" $path
            break
        }
    }

    if (-not $pinned) {
        throw "VS Code not found in expected locations"
    }
}

Invoke-Step "Pin Everything to taskbar" {
    # Common Everything installation paths
    $everythingPaths = @(
        "$env:ProgramFiles\Everything\Everything.exe",
        "${env:ProgramFiles(x86)}\Everything\Everything.exe",
        "$env:LOCALAPPDATA\Programs\Everything\Everything.exe"
    )

    $pinned = $false
    foreach ($path in $everythingPaths) {
        if (Test-Path $path) {
            $pinned = Pin-ToTaskbar "Everything" $path
            break
        }
    }

    if (-not $pinned) {
        throw "Everything not found in expected locations"
    }
}
#endregion

#region PowerShell Profile Configuration
Write-Host "`n--- POWERSHELL PROFILE ---`n" -ForegroundColor Yellow

Invoke-Step "Install posh-git module" {
    # Install NuGet provider first without prompting
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction SilentlyContinue | Out-Null
    # Install posh-git module
    Install-Module -Name posh-git -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck
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
