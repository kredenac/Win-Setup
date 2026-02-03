#Requires -RunAsAdministrator

<# V 2
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

    # Download and install PS7 directly (faster than winget)
    Write-Info "Downloading PowerShell 7 from GitHub..."
    try {
        $ps7Version = "7.4.7"
        $url = "https://github.com/PowerShell/PowerShell/releases/download/v$ps7Version/PowerShell-$ps7Version-win-x64.msi"
        $output = "$env:TEMP\PowerShell-$ps7Version-win-x64.msi"

        # Download the MSI (disable progress bar for faster download)
        $ProgressPreference = 'SilentlyContinue'
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

Write-Info "Running on PowerShell $($PSVersionTable.PSVersion)"

# Check if winget is available
$wingetAvailable = $null -ne (Get-Command winget -ErrorAction SilentlyContinue)
if (-not $wingetAvailable) {
    Write-Warning "winget not available - will use direct downloads for software installation"
}

# Install Windows Terminal if not already installed (only if winget is available)
if ($wingetAvailable) {
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
} else {
    Write-Info "Skipping Windows Terminal installation (requires winget or Microsoft Store)"
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

Invoke-Step "Enable classic context menu" {
    if ([System.Environment]::OSVersion.Version.Build -ge 22000) {
        $registryPath = "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32"
        if (-not (Test-Path "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}")) {
            New-Item -Path "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}" -Force | Out-Null
        }
        if (-not (Test-Path $registryPath)) {
            New-Item -Path $registryPath -Force | Out-Null
        }
        Set-ItemProperty -Path $registryPath -Name "(Default)" -Type String -Value "" -Force
    } else {
        Write-Info "Classic context menu only applicable to Windows 11"
    }
}

Invoke-Step "Disable widgets in taskbar" {
    if ([System.Environment]::OSVersion.Version.Build -ge 22000) {
        $registryPath = "HKLM:\SOFTWARE\Policies\Microsoft\Dsh"
        if (-not (Test-Path $registryPath)) {
            New-Item -Path $registryPath -Force | Out-Null
        }
        Set-ItemProperty -Path $registryPath -Name "AllowNewsAndInterests" -Type DWord -Value 0 -Force
    }
}

Invoke-Step "Disable Cortana button in taskbar" {
    $registryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    Set-ItemProperty -Path $registryPath -Name "ShowCortanaButton" -Type DWord -Value 0 -Force
}

Invoke-Step "Disable People button in taskbar" {
    $registryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\People"
    if (-not (Test-Path $registryPath)) {
        New-Item -Path $registryPath -Force | Out-Null
    }
    Set-ItemProperty -Path $registryPath -Name "PeopleBand" -Type DWord -Value 0 -Force
}

Invoke-Step "Enable Remote Desktop" {
    Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -Value 0 -Type DWord -Force
    try {
        Enable-NetFirewallRule -DisplayGroup "Remote Desktop" -ErrorAction Stop
    }
    catch {
        Write-Warning "Could not configure firewall rules for Remote Desktop: $_"
        $script:warningCount++
    }
}

Invoke-Step "Enable Hyper-V" {
    try {
        $result = Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V -All -NoRestart -ErrorAction Stop
        if ($result.RestartNeeded) {
            Write-Info "Hyper-V enabled (restart required)"
        }
    }
    catch {
        Write-Warning "Could not enable Hyper-V. This feature may not be available on this edition of Windows or may require BIOS virtualization settings."
        $script:warningCount++
    }
}
#endregion

#region Software Installation (SLOW - Run in Parallel)
Write-Host "`n--- SOFTWARE INSTALLATION ---`n" -ForegroundColor Yellow

if ($wingetAvailable) {
    Write-Info "Using winget for software installation..."

    # Define software to install in parallel
    $softwareJobs = @()

    $parallelInstalls = @(
        @{ Name = "Google Chrome"; Id = "Google.Chrome" }
        @{ Name = "Visual Studio Code"; Id = "Microsoft.VisualStudioCode" }
        @{ Name = "VLC Media Player"; Id = "VideoLAN.VLC" }
        @{ Name = "Everything Search"; Id = "voidtools.Everything" }
        @{ Name = "7-Zip"; Id = "7zip.7zip" }
    )

    # Add gaming software if enabled
    if ($config.gaming -eq $true) {
        $parallelInstalls += @{ Name = "Steam"; Id = "Valve.Steam" }
        $parallelInstalls += @{ Name = "Discord"; Id = "Discord.Discord" }
    }

    # Start all installations as background jobs
    foreach ($software in $parallelInstalls) {
        $job = Start-Job -ScriptBlock {
            param($Id, $Name, $CustomArgs)
            if ($CustomArgs) {
                $result = winget install --id $Id --silent --override $CustomArgs --accept-source-agreements --accept-package-agreements 2>&1
            } else {
                $result = winget install --id $Id --silent --accept-source-agreements --accept-package-agreements 2>&1
            }
            return @{
                Name = $Name
                Success = $LASTEXITCODE -eq 0
                Output = $result
            }
        } -ArgumentList $software.Id, $software.Name, $software.Args

        $softwareJobs += @{
            Job = $job
            Name = $software.Name
        }
    }

    # Add Python as parallel job (direct download, faster than winget)
    $pythonJob = Start-Job -ScriptBlock {
        $pythonVersion = "3.12.8"
        $url = "https://www.python.org/ftp/python/$pythonVersion/python-$pythonVersion-amd64.exe"
        $output = "$env:TEMP\python-$pythonVersion-amd64.exe"

        try {
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri $url -OutFile $output -UseBasicParsing
            Start-Process -FilePath $output -ArgumentList '/quiet', 'InstallAllUsers=1', 'PrependPath=1', 'Include_pip=1', 'Include_test=0', 'Include_doc=0' -Wait -NoNewWindow

            $pythonInstallPath = "$env:ProgramFiles\Python312"
            if (Test-Path $pythonInstallPath) {
                $currentPath = [Environment]::GetEnvironmentVariable("Path", "Machine")
                if ($currentPath -notlike "*$pythonInstallPath*") {
                    [Environment]::SetEnvironmentVariable("Path", "$currentPath;$pythonInstallPath;$pythonInstallPath\Scripts", "Machine")
                }
            }

            return @{
                Name = "Python (direct download)"
                Success = $true
                Output = "Installed successfully"
            }
        }
        catch {
            return @{
                Name = "Python (direct download)"
                Success = $false
                Output = $_.Exception.Message
            }
        }
    }

    $softwareJobs += @{
        Job = $pythonJob
        Name = "Python (direct download)"
    }

    # Add Git as parallel job (direct download)
    $gitJob = Start-Job -ScriptBlock {
        try {
            $gitUrl = "https://github.com/git-for-windows/git/releases/download/v2.48.1.windows.1/Git-2.48.1-64-bit.exe"
            $gitPath = "$env:TEMP\git_installer.exe"
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri $gitUrl -OutFile $gitPath -UseBasicParsing
            Start-Process -FilePath $gitPath -ArgumentList "/VERYSILENT", "/NORESTART" -Wait -NoNewWindow

            return @{
                Name = "Git (direct download)"
                Success = $true
                Output = "Installed successfully"
            }
        }
        catch {
            return @{
                Name = "Git (direct download)"
                Success = $false
                Output = $_.Exception.Message
            }
        }
    }

    $softwareJobs += @{
        Job = $gitJob
        Name = "Git (direct download)"
    }

    # Add nvm-windows as parallel job (direct download)
    $nvmJob = Start-Job -ScriptBlock {
        try {
            $nvmVersion = "1.1.12"
            $nvmUrl = "https://github.com/coreybutler/nvm-windows/releases/download/$nvmVersion/nvm-setup.exe"
            $nvmPath = "$env:TEMP\nvm-setup.exe"
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri $nvmUrl -OutFile $nvmPath -UseBasicParsing
            Start-Process -FilePath $nvmPath -ArgumentList "/VERYSILENT" -Wait -NoNewWindow

            return @{
                Name = "nvm-windows (direct download)"
                Success = $true
                Output = "Installed successfully"
            }
        }
        catch {
            return @{
                Name = "nvm-windows (direct download)"
                Success = $false
                Output = $_.Exception.Message
            }
        }
    }

    $softwareJobs += @{
        Job = $nvmJob
        Name = "nvm-windows (direct download)"
    }

    # Add Claude CLI as parallel job
    $claudeJob = Start-Job -ScriptBlock {
        try {
            $result = Invoke-Expression "& { $(Invoke-RestMethod https://claude.ai/install.ps1) }" 2>&1
            return @{
                Name = "Claude CLI (Claude Code)"
                Success = $true
                Output = "Installed successfully"
            }
        }
        catch {
            return @{
                Name = "Claude CLI (Claude Code)"
                Success = $false
                Output = $_.Exception.Message
            }
        }
    }

    $softwareJobs += @{
        Job = $claudeJob
        Name = "Claude CLI (Claude Code)"
    }

    Write-Info "Waiting for $($softwareJobs.Count) parallel installations to complete..."

    # Wait for ALL jobs to complete first
    $allJobs = $softwareJobs | ForEach-Object { $_.Job }
    Wait-Job -Job $allJobs | Out-Null

    # Now collect all results
    foreach ($jobInfo in $softwareJobs) {
        $result = Receive-Job -Job $jobInfo.Job
        Remove-Job -Job $jobInfo.Job

        if ($result.Success -or $result.Output -like "*already installed*" -or $result.Output -like "*No available upgrade*") {
            Write-Success "$($jobInfo.Name) completed"
            $script:successCount++
        } else {
            Write-ErrorMsg "$($jobInfo.Name) failed"
            $script:failureCount++
        }
    }

    Invoke-Step "Set Chrome as default browser" {
        # Find Chrome installation
        $chromePaths = @(
            "${env:ProgramFiles}\Google\Chrome\Application\chrome.exe",
            "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
            "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
        )

        $chromePath = $chromePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

        if ($chromePath) {
            # Use Chrome's built-in flag to set as default browser
            Start-Process -FilePath $chromePath -ArgumentList "--make-default-browser" -NoNewWindow
            Write-Info "Set Chrome as default browser"
        } else {
            Write-Warning "Chrome not found, skipping default browser setup"
            $script:warningCount++
        }
    }

    # Install Node.js LTS via nvm (needs to run after nvm installation)
    Invoke-Step "Install Node.js LTS via nvm" {
        # Try multiple possible nvm installation locations
        $possibleNvmPaths = @(
            "$env:APPDATA\nvm\nvm.exe",
            "$env:ProgramFiles\nvm\nvm.exe",
            "${env:ProgramFiles(x86)}\nvm\nvm.exe"
        )

        $nvmPath = $null
        foreach ($path in $possibleNvmPaths) {
            if (Test-Path $path) {
                $nvmPath = $path
                Write-Info "Found nvm at: $path"
                break
            }
        }

        if ($nvmPath) {
            # Found nvm, install Node.js in a new PowerShell process with updated PATH
            $installScript = @"
& '$nvmPath' install lts
& '$nvmPath' use lts
"@
            $result = powershell -NoProfile -Command $installScript 2>&1
            Write-Info $result
        } else {
            Write-Warning "nvm not found. Please open a new terminal and run: nvm install lts && nvm use lts"
            $script:warningCount++
        }
    }

    Invoke-Step "Configure Claude CLI settings" {
        # Create .claude directory if it doesn't exist
        $claudeDir = "$env:USERPROFILE\.claude"
        if (-not (Test-Path $claudeDir)) {
            New-Item -Path $claudeDir -ItemType Directory -Force | Out-Null
        }

        # Download settings from dotfiles repo
        $claudeSettingsUrl = "https://raw.githubusercontent.com/kredenac/dotfiles/main/.claude/settings.json"
        $claudeSettingsPath = "$claudeDir\settings.json"
        try {
            $settingsContent = Invoke-WebRequest -Uri $claudeSettingsUrl -UseBasicParsing | Select-Object -ExpandProperty Content
            Set-Content -Path $claudeSettingsPath -Value $settingsContent -Force -Encoding UTF8
            Write-Info "Claude CLI settings downloaded and configured at: $claudeSettingsPath"
        }
        catch {
            Write-Warning "Failed to download Claude settings from GitHub: $_"
            Write-Info "You can manually download it from: $claudeSettingsUrl"
            throw
        }
    }
} else {
    Write-Info "Using direct downloads for software installation..."

    # Define software jobs for parallel installation
    $softwareJobs = @()

    # Chrome job
    $chromeJob = Start-Job -ScriptBlock {
        try {
            $chromeUrl = "https://dl.google.com/chrome/install/latest/chrome_installer.exe"
            $chromePath = "$env:TEMP\chrome_installer.exe"
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri $chromeUrl -OutFile $chromePath -UseBasicParsing
            Start-Process -FilePath $chromePath -ArgumentList "/silent", "/install" -Wait -NoNewWindow
            return @{ Name = "Google Chrome"; Success = $true; Output = "Installed successfully" }
        } catch {
            return @{ Name = "Google Chrome"; Success = $false; Output = $_.Exception.Message }
        }
    }
    $softwareJobs += @{ Job = $chromeJob; Name = "Google Chrome" }

    # VS Code job
    $vscodeJob = Start-Job -ScriptBlock {
        try {
            $vscodeUrl = "https://code.visualstudio.com/sha/download?build=stable&os=win32-x64"
            $vscodePath = "$env:TEMP\vscode_installer.exe"
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri $vscodeUrl -OutFile $vscodePath -UseBasicParsing
            Start-Process -FilePath $vscodePath -ArgumentList "/VERYSILENT", "/NORESTART", "/MERGETASKS=!runcode" -Wait -NoNewWindow
            return @{ Name = "Visual Studio Code"; Success = $true; Output = "Installed successfully" }
        } catch {
            return @{ Name = "Visual Studio Code"; Success = $false; Output = $_.Exception.Message }
        }
    }
    $softwareJobs += @{ Job = $vscodeJob; Name = "Visual Studio Code" }

    # Git job
    $gitJob = Start-Job -ScriptBlock {
        try {
            $gitUrl = "https://github.com/git-for-windows/git/releases/download/v2.48.1.windows.1/Git-2.48.1-64-bit.exe"
            $gitPath = "$env:TEMP\git_installer.exe"
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri $gitUrl -OutFile $gitPath -UseBasicParsing
            Start-Process -FilePath $gitPath -ArgumentList "/VERYSILENT", "/NORESTART" -Wait -NoNewWindow
            return @{ Name = "Git"; Success = $true; Output = "Installed successfully" }
        } catch {
            return @{ Name = "Git"; Success = $false; Output = $_.Exception.Message }
        }
    }
    $softwareJobs += @{ Job = $gitJob; Name = "Git" }

    # 7-Zip job
    $zipJob = Start-Job -ScriptBlock {
        try {
            $zipUrl = "https://www.7-zip.org/a/7z2408-x64.exe"
            $zipPath = "$env:TEMP\7zip_installer.exe"
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing
            Start-Process -FilePath $zipPath -ArgumentList "/S" -Wait -NoNewWindow
            return @{ Name = "7-Zip"; Success = $true; Output = "Installed successfully" }
        } catch {
            return @{ Name = "7-Zip"; Success = $false; Output = $_.Exception.Message }
        }
    }
    $softwareJobs += @{ Job = $zipJob; Name = "7-Zip" }

    # Python job
    $pythonJob = Start-Job -ScriptBlock {
        try {
            $pythonVersion = "3.12.8"
            $url = "https://www.python.org/ftp/python/$pythonVersion/python-$pythonVersion-amd64.exe"
            $output = "$env:TEMP\python-$pythonVersion-amd64.exe"
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri $url -OutFile $output -UseBasicParsing
            Start-Process -FilePath $output -ArgumentList '/quiet', 'InstallAllUsers=1', 'PrependPath=1', 'Include_pip=1', 'Include_test=0', 'Include_doc=0' -Wait -NoNewWindow

            $pythonInstallPath = "$env:ProgramFiles\Python312"
            if (Test-Path $pythonInstallPath) {
                $currentPath = [Environment]::GetEnvironmentVariable("Path", "Machine")
                if ($currentPath -notlike "*$pythonInstallPath*") {
                    [Environment]::SetEnvironmentVariable("Path", "$currentPath;$pythonInstallPath;$pythonInstallPath\Scripts", "Machine")
                }
            }
            return @{ Name = "Python"; Success = $true; Output = "Installed successfully" }
        } catch {
            return @{ Name = "Python"; Success = $false; Output = $_.Exception.Message }
        }
    }
    $softwareJobs += @{ Job = $pythonJob; Name = "Python" }

    # Everything Search job
    $everythingJob = Start-Job -ScriptBlock {
        try {
            $everythingUrl = "https://www.voidtools.com/Everything-1.4.1.1026.x64-Setup.exe"
            $everythingPath = "$env:TEMP\everything_installer.exe"
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri $everythingUrl -OutFile $everythingPath -UseBasicParsing
            Start-Process -FilePath $everythingPath -ArgumentList "/S" -Wait -NoNewWindow
            return @{ Name = "Everything Search"; Success = $true; Output = "Installed successfully" }
        } catch {
            return @{ Name = "Everything Search"; Success = $false; Output = $_.Exception.Message }
        }
    }
    $softwareJobs += @{ Job = $everythingJob; Name = "Everything Search" }

    Write-Info "Waiting for $($softwareJobs.Count) parallel installations to complete..."

    # Wait for ALL jobs to complete first
    $allJobs = $softwareJobs | ForEach-Object { $_.Job }
    Wait-Job -Job $allJobs | Out-Null

    # Now collect all results
    foreach ($jobInfo in $softwareJobs) {
        $result = Receive-Job -Job $jobInfo.Job
        Remove-Job -Job $jobInfo.Job

        if ($result.Success) {
            Write-Success "$($jobInfo.Name) completed"
            $script:successCount++
        } else {
            Write-ErrorMsg "$($jobInfo.Name) failed: $($result.Output)"
            $script:failureCount++
        }
    }

    Write-Info "Skipping some software (no direct download URLs available): VLC, PowerToys, nvm, Node.js"
    $script:warningCount++
}

# Wait for all installers to fully complete and register
Write-Info "Waiting for installations to finalize..."
Start-Sleep -Seconds 5
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
    # Get PowerShell profile path
    $profilePath = $PROFILE.CurrentUserAllHosts

    # Create profile directory if it doesn't exist
    $profileDir = Split-Path -Parent $profilePath
    if (-not (Test-Path $profileDir)) {
        New-Item -Path $profileDir -ItemType Directory -Force | Out-Null
    }

    # Download profile from dotfiles repo
    $profileUrl = "https://raw.githubusercontent.com/kredenac/dotfiles/main/PowerShell-Profile.ps1"
    try {
        $profileContent = Invoke-WebRequest -Uri $profileUrl -UseBasicParsing | Select-Object -ExpandProperty Content
        Set-Content -Path $profilePath -Value $profileContent -Force
        Write-Info "PowerShell profile downloaded and configured at: $profilePath"
    }
    catch {
        Write-Warning "Failed to download PowerShell profile from GitHub: $_"
        Write-Info "You can manually download it from: $profileUrl"
        throw
    }
}

Invoke-Step "Configure Windows Terminal settings" {
    $settingsPath = "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json"

    if (Test-Path $settingsPath) {
        # Download settings from dotfiles repo
        $settingsUrl = "https://raw.githubusercontent.com/kredenac/dotfiles/main/WindowsTerminal-settings.json"
        try {
            $settingsContent = Invoke-WebRequest -Uri $settingsUrl -UseBasicParsing | Select-Object -ExpandProperty Content
            Set-Content -Path $settingsPath -Value $settingsContent -Force -Encoding UTF8
            Write-Info "Windows Terminal settings downloaded and configured from dotfiles"
        }
        catch {
            Write-Warning "Failed to download Windows Terminal settings from GitHub: $_"
            Write-Info "You can manually download it from: $settingsUrl"
            throw
        }
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

Write-Host "REMINDER: " -ForegroundColor Cyan -NoNewline
Write-Host "You might want to pin these apps to the taskbar:"
Write-Host "  - Google Chrome"
Write-Host "  - Windows Terminal"
Write-Host "  - Visual Studio Code"
Write-Host "  - Everything Search`n"

$restart = Read-Host "Would you like to restart now? (Y/N)"
if ($restart -eq "Y" -or $restart -eq "y") {
    Write-Info "Restarting in 10 seconds... (Press Ctrl+C to cancel)"
    Start-Sleep -Seconds 10
    Restart-Computer -Force
} else {
    Write-Info "Please restart your computer when convenient to apply all changes."
}
#endregion
