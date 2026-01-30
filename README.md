# Windows Setup Script

Automated Windows configuration script for new machine setup. Designed for development and general use.

## How to Run

### Option 1: One-liner (Remote Execution)

Run directly from an elevated PowerShell session:

```powershell
irm https://raw.githubusercontent.com/kredenac/win-setup/main/setup.ps1 | iex
```

**Note:** This method uses default configuration. For custom settings, use Option 2.

### Option 2: Simple Execution (Custom Configuration)

1. Download or clone this repository to your machine
2. Edit `config.json` with your Git settings
3. Right-click **Windows Terminal** and select **Run as Administrator**
4. Navigate to the script directory:
   ```powershell
   cd C:\path\to\win-setup
   ```
5. Run the script:
   ```powershell
   .\setup.ps1
   ```

## What This Script Does

### Windows UI Customization
- Sets taskbar alignment to left (not centered)
- Enables dark mode for Windows and apps
- Shows hidden files and file extensions in File Explorer
- Hides Music folder from File Explorer sidebar
- Shows "This PC" icon on desktop
- Removes weather/search highlights from taskbar
- Removes widgets, Meet Now from taskbar
- Unpins Microsoft Store and Copilot from taskbar
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

## Requirements

- **Windows 10 (20H2+) or Windows 11**
- **PowerShell 7+** (not Windows PowerShell 5.1)
- **Windows Terminal** (pre-installed on Windows 11)
- **Administrator privileges**
- **Internet connection** for software downloads

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

## Post-Setup Manual Steps

Some items may need manual attention:

- **Monitor refresh rate**: If automatic detection fails, set manually in Display Settings
- **Taskbar unpinning**: If Microsoft Store wasn't unpinned, right-click and unpin manually
- **nvm**: If Node.js installation via nvm fails, open a new terminal and run:
  ```powershell
  nvm install lts
  nvm use lts
  ```

## Troubleshooting

### "Script cannot be loaded because running scripts is disabled"
Run this command in an elevated PowerShell session:
```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
```

### "winget not found"
Ensure you're on Windows 10 20H2+ or Windows 11. Update Windows if necessary.

### Software installation fails
Check your internet connection. You can re-run the script safely - it will skip already-installed software.

### Changes not taking effect
Restart your computer. Many Windows registry changes require a restart or logoff/logon.

## Safety

- Script is idempotent - safe to run multiple times
- Each task is wrapped in error handling - failures won't stop the entire script
- Color-coded output shows success/warning/error status for each step
- No restore point is created (changes are intentional and documented)

## License

MIT License - Feel free to modify and use as needed.
