# GraphModuleStatus

A PowerShell module that checks the status of Microsoft Graph PowerShell modules and helps keep them up to date.

## Features

- **Automatic Status Check** - Shows installed vs available versions of Microsoft.Graph and Microsoft.Graph.Beta on PowerShell startup
- **Fast Version Checking** - Checks PSGallery via URL redirect with no download and a 5-second timeout
- **Interactive Install** - Prompts to install Graph modules when none are detected, with module and scope selection
- **Clean Update Process** - Performs a complete uninstall and reinstall to resolve version conflicts and assembly loading errors
- **Dual Scope Detection** - Checks both CurrentUser and AllUsers installation scopes
- **Auto-Elevation** - Automatically launches an elevated session when All Users scope requires Administrator rights
- **Profile Integration** - Easily add/remove the status check from your PowerShell profile

## Installation

**PSResourceGet (recommended):**
```powershell
Install-PSResource -Name GraphModuleStatus -Repository PSGallery -TrustRepository
```

**PowerShellGet:**
```powershell
Install-Module -Name GraphModuleStatus -Scope CurrentUser
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

  Update available. Run Update-GraphModule now? (Y/N):
```

Output (when all current):
```
  [Microsoft.Graph]      v2.26.0 ● Current
  [Microsoft.Graph.Beta] v2.26.0 ● Current
```

Output (when not installed):
```
  [Microsoft.Graph]      ○ Not installed  (v2.26.0 available on PSGallery)
  [Microsoft.Graph.Beta] ○ Not installed  (v2.26.0 available on PSGallery)

  No Microsoft Graph modules are installed.
  The following versions are available from PSGallery:

    [Microsoft.Graph]      v2.26.0
    [Microsoft.Graph.Beta] v2.26.0

  Would you like to install them now? (Y/N):
```

When installing, you choose which modules to install and the scope:
```
  Which modules would you like to install?

    [1] Microsoft.Graph  v2.26.0
    [2] Microsoft.Graph.Beta  v2.26.0
    [A] All  (default)

  Install scope:
    [1] All Users  (Recommended)
    [2] Current User Only
```

If All Users is selected and the session is not elevated, an elevated PowerShell window is launched automatically to complete the installation.

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

## Available Cmdlets

### Get-GraphModuleStatus
Displays the status of Microsoft Graph modules (installed vs available versions).

**Syntax:**
```powershell
Get-GraphModuleStatus [-Silent] [-NoPrompt]
```

**Parameters:**
- `-Silent` (Switch) - Suppresses console output and returns objects instead
- `-NoPrompt` (Switch) - Does not prompt to run the update script when updates are available

**Examples:**
```powershell
# Check module status with console output
Get-GraphModuleStatus

# Get status as objects for scripting
$status = Get-GraphModuleStatus -Silent

# Check status without update prompt
Get-GraphModuleStatus -NoPrompt
```

### Update-GraphModule
Runs the Microsoft Graph module update script to perform a clean uninstall and reinstall.

**Syntax:**
```powershell
Update-GraphModule
```

**Parameters:** None

**Examples:**
```powershell
# Run the update process
Update-GraphModule
```

### Add-GraphModuleStatusToProfile
Adds the GraphModuleStatus module import and status check to your PowerShell profile.

**Syntax:**
```powershell
Add-GraphModuleStatusToProfile
```

**Examples:**
```powershell
# Add to profile
Add-GraphModuleStatusToProfile
```

### Remove-GraphModuleStatusFromProfile
Removes the GraphModuleStatus module import and status check from your PowerShell profile.

**Syntax:**
```powershell
Remove-GraphModuleStatusFromProfile
```

**Examples:**
```powershell
# Remove from profile
Remove-GraphModuleStatusFromProfile
```

## Requirements

- PowerShell 5.1 or later (PowerShell 7+ recommended)
- Microsoft.PowerShell.PSResourceGet module
- Internet connectivity to check PSGallery for updates

## What the Update Script Does

The `Update-GraphModule` function runs `Update-MicrosoftGraph.ps1`, a comprehensive 6-step process:

1. **Session Cleanup** - Removes loaded Graph modules from the current session to prevent file locking
2. **Module Uninstallation** - Iteratively uninstalls all selected module versions using both `Uninstall-PSResource` (PSResourceGet) and `Uninstall-Module` (PowerShellGet), with garbage collection between passes
3. **Folder Cleanup** - Removes any leftover module subfolders in all known PowerShell module paths
4. **Fresh Installation** - Installs your choice of Microsoft.Graph and/or Microsoft.Graph.Beta from PowerShell Gallery, with scope selection (All Users or Current User)
5. **Module Import** - Imports `Microsoft.Graph.Authentication` immediately; all other modules load on demand
6. **Validation** - Verifies installed versions and checks for version mismatches between Graph and Graph.Beta

### Interactive Prompts

The update script prompts you to:
- **Choose which modules to uninstall** (Graph stable, Graph.Beta, or both)
- **Choose which modules to reinstall** (same options, plus skip)
- **Choose the installation scope** (All Users or Current User)

A real-time elapsed timer is shown in the window title bar during the process.

This resolves common issues like:
- "Assembly with same name is already loaded"
- "Could not load file or assembly 'Microsoft.Graph.Authentication'"
- Commands not being recognized after module updates
- Multiple authentication prompts

## Author

Mark Orr

## License

MIT License - See [LICENSE](LICENSE) file for details.
