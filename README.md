# GraphModuleStatus

A PowerShell module that checks the status of Microsoft Graph PowerShell modules and helps keep them up to date.

## Features

- **Automatic Status Check** - Shows installed vs available versions of Microsoft.Graph and Microsoft.Graph.Beta on PowerShell startup
- **Clean Update Process** - Performs a complete uninstall and reinstall to resolve version conflicts and assembly loading errors
- **Profile Integration** - Easily add/remove the status check from your PowerShell profile

## Installation

### From GitHub (One-Liner)

```powershell
Invoke-RestMethod "https://raw.githubusercontent.com/markorr321/GraphModuleStatus/main/GraphModuleStatus.psm1" -OutFile "$HOME\Documents\PowerShell\Modules\GraphModuleStatus\GraphModuleStatus.psm1"; Invoke-RestMethod "https://raw.githubusercontent.com/markorr321/GraphModuleStatus/main/GraphModuleStatus.psd1" -OutFile "$HOME\Documents\PowerShell\Modules\GraphModuleStatus\GraphModuleStatus.psd1"; Invoke-RestMethod "https://raw.githubusercontent.com/markorr321/GraphModuleStatus/main/Update-MicrosoftGraph.ps1" -OutFile "$HOME\Documents\PowerShell\Modules\GraphModuleStatus\Update-MicrosoftGraph.ps1"
```

Or use this simpler version:

```powershell
$m="$HOME\Documents\PowerShell\Modules\GraphModuleStatus";New-Item $m -ItemType Directory -Force|Out-Null;@("psm1","psd1")|%{Invoke-RestMethod "https://raw.githubusercontent.com/markorr321/GraphModuleStatus/main/GraphModuleStatus.$_" -OutFile "$m\GraphModuleStatus.$_"};Invoke-RestMethod "https://raw.githubusercontent.com/markorr321/GraphModuleStatus/main/Update-MicrosoftGraph.ps1" -OutFile "$m\Update-MicrosoftGraph.ps1"
```

### From PowerShell Gallery (Coming Soon)

```powershell
Install-Module -Name GraphModuleStatus -Scope CurrentUser
```

### Manual Installation

Clone the repo and copy to your modules directory:

```powershell
Copy-Item -Path ".\GraphModuleStatus" -Destination "$HOME\Documents\PowerShell\Modules\GraphModuleStatus" -Recurse
```

### After Installation

Import the module:

```powershell
Import-Module GraphModuleStatus
```

To load it automatically on every PowerShell session, add to your profile:

```powershell
Add-GraphModuleStatusToProfile
```

## Usage

### Check Module Status

```powershell
Get-GraphModuleStatus
```

Output (when updates available):
```
  [Microsoft.Graph]      v2.25.0 → v2.26.0
  [Microsoft.Graph.Beta] v2.25.0 ● Current

  Update available! Run Update-GraphModule to update.
```

Output (when all current):
```
  [Microsoft.Graph]      v2.25.0 ● Current
  [Microsoft.Graph.Beta] v2.25.0 ● Current
```

Output (when not installed):
```
  [Microsoft.Graph]      ○ Not installed
  [Microsoft.Graph.Beta] ○ Not installed
```

### Update Modules

Run a clean uninstall and reinstall of Microsoft Graph modules:

```powershell
Update-GraphModule
```

This is useful when:
- Updates are available
- You're experiencing assembly loading errors
- Commands aren't being recognized after updates
- You want a fresh, clean installation

### Add to PowerShell Profile

Automatically check module status every time you open PowerShell:

```powershell
Add-GraphModuleStatusToProfile
```

### Remove from Profile

```powershell
Remove-GraphModuleStatusFromProfile
```

## Functions

| Function | Description |
|----------|-------------|
| `Get-GraphModuleStatus` | Displays installed vs available versions of Graph modules |
| `Update-GraphModule` | Runs the clean reinstall script for Graph modules |
| `Add-GraphModuleStatusToProfile` | Adds the status check to your PowerShell profile |
| `Remove-GraphModuleStatusFromProfile` | Removes the status check from your profile |

## Parameters

### Get-GraphModuleStatus

| Parameter | Type | Description |
|-----------|------|-------------|
| `-Silent` | Switch | Suppresses output and returns objects instead |
| `-NoPrompt` | Switch | Does not show the update prompt when updates are available |

## Requirements

- PowerShell 5.1 or later (PowerShell 7+ recommended)
- Microsoft.PowerShell.PSResourceGet module
- Internet connectivity to check PSGallery for updates

## What the Update Script Does

The `Update-GraphModule` function runs a comprehensive update process:

1. **Session Cleanup** - Removes loaded Graph modules from the current session
2. **Module Uninstallation** - Iteratively uninstalls all Graph module versions
3. **Folder Cleanup** - Removes any leftover module folders
4. **Fresh Installation** - Installs the latest versions from PowerShell Gallery
5. **Validation** - Verifies the installation was successful

This resolves common issues like:
- "Assembly with same name is already loaded"
- "Could not load file or assembly 'Microsoft.Graph.Authentication'"
- Commands not being recognized after module updates
- Multiple authentication prompts

## Author

Mark Orr

## License

MIT License - See [LICENSE](LICENSE) file for details.
