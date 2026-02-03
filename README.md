# Windows Setup Script

Automated Windows configuration script for new machine setup. Designed for development and general use.

## How to Run

This script requires **PowerShell 7+** (not Windows PowerShell 5.1). If you run it in Windows PowerShell by mistake, the script will automatically install PowerShell 7 and prompt you to re-run it.

### Option 1: One-liner (Remote Execution)

Run directly from an elevated **PowerShell 7+** session:

```powershell
irm https://raw.githubusercontent.com/kredenac/win-setup/main/setup.ps1 | iex
```

**Note:** This method uses default configuration. For custom settings, use Option 2.

**Refresh dotfiles only** (PowerShell profile, Windows Terminal settings, Claude CLI settings):

```powershell
$DotfilesOnly=$true; irm https://raw.githubusercontent.com/kredenac/win-setup/main/setup.ps1 | iex
```

### Option 2: Local Execution (Custom Configuration)

1. Download or clone this repository to your machine
2. Edit `config.json` with your Git settings
3. Right-click **Windows Terminal** and select **Run as Administrator** (this opens PowerShell 7 by default)
4. Navigate to the script directory:
   ```powershell
   cd C:\path\to\win-setup
   ```
5. Run the script:
   ```powershell
   # Full setup
   .\setup.ps1

   # Or refresh dotfiles only
   .\setup.ps1 -DotfilesOnly
   ```

## Dotfiles-Only Mode

Use the `-DotfilesOnly` flag to refresh your configuration files without running the full setup. This mode:

- **Skips** all Windows settings, registry modifications, and software installation
- **Refreshes** only your dotfiles:
  - PowerShell profile (aliases, prompt customization)
  - Windows Terminal settings (theme, keybindings, profiles)
  - Claude CLI settings
- **No restart required** - just restart your terminal sessions

This is useful when you've updated your dotfiles repository and want to pull the latest changes without re-running the entire setup.

## What This Script Does

### Windows UI Customization
- Sets taskbar alignment to left (not centered)
- Enables dark mode for Windows and apps
- Shows hidden files and file extensions in File Explorer
- Hides Music folder from File Explorer sidebar
- Shows "This PC" icon on desktop
- Removes weather/search highlights from taskbar
- Removes widgets, Meet Now, and Task View button from taskbar
- Unpins Microsoft Store, Copilot, and Microsoft Edge from taskbar
- Pins Chrome, Windows Terminal, Visual Studio Code, and Everything to taskbar (in that order, after File Explorer)
- Disables web search results in Start Menu (keeps local search)

### System Settings
- Sets monitor refresh rate to maximum available
- Disables UAC (User Account Control)
- Enables Windows Developer Mode
- Sets PowerShell execution policy to RemoteSigned
- Disables hibernation (frees up disk space)
- Sets date format to d/M/yyyy
- Sets first day of week to Monday
- Adds Serbian (Latin) keyboard layout

### Software Installation
**Always Installed:**
- Google Chrome
- Visual Studio Code
- Git
- VLC Media Player
- Everything Search
- 7-Zip
- Python 3.12
- PowerToys
- nvm-windows (Node Version Manager)
- Node.js LTS (via nvm)

**Gaming Mode (when `gaming: true` in config):**
- Steam
- Discord

### Git Configuration
- Sets username and email (from config.json)
- Sets default branch name to "main"
- Sets default editor to VS Code
- Configures credential manager

## Configuration

Edit `config.json` to customize settings:

```json
{
  "git": {
    "username": "your-username",
    "email": "your-email@example.com"
  },
  "gaming": false
}
```

- **git.username**: Your Git username
- **git.email**: Your Git email address
- **gaming**: Set to `true` to install Steam and Discord

## After Running

1. The script will display a summary of completed tasks
2. You'll be prompted to restart your computer
3. Some changes (UAC, Developer Mode, taskbar settings) require a restart to take effect

## Troubleshooting

### "Running on Windows PowerShell" warning
If you see a warning about Windows PowerShell and the script installs PowerShell 7:
1. The script will automatically install PowerShell 7 for you
2. Close the current window after installation
3. Open a new **PowerShell 7** window (not Windows PowerShell) as Administrator
4. Re-run the script

### "Script cannot be loaded because running scripts is disabled"
Run this command in an elevated PowerShell session:
```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
```

## Safety

- Script is idempotent - safe to run multiple times
- Each task is wrapped in error handling - failures won't stop the entire script
- Color-coded output shows success/warning/error status for each step

## License

MIT License - Feel free to modify and use as needed.
