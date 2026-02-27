##############################################################################
# GraphModuleStatus
# Shows the status of Microsoft Graph PowerShell modules on profile load
##############################################################################

# Path to the update script (bundled with module)
$script:UpdateScriptPath = Join-Path $PSScriptRoot "Update-MicrosoftGraph.ps1"

##############################################################################
# Get-LatestPSGalleryVersion
# Fast version check via URL redirect — no download, 5-second timeout
# Adapted from Entra-PIM module (github.com/markdomansky/Entra-PIM)
##############################################################################
function Get-LatestPSGalleryVersion {
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )
    try {
        $Url = "https://www.powershellgallery.com/packages/$Name"
        try {
            $null = Invoke-WebRequest -Uri $Url -UseBasicParsing -MaximumRedirection 0 -TimeoutSec 5 -ErrorAction Stop
        }
        catch {
            if ($_.Exception.Response -and $_.Exception.Response.Headers) {
                try {
                    $Location = $_.Exception.Response.Headers.GetValues('Location') | Select-Object -First 1
                    if ($Location) {
                        return [version](Split-Path -Path $Location -Leaf)
                    }
                }
                catch { }
            }
        }
    }
    catch { }
    return $null
}

##############################################################################
# Get-GraphModuleStatus
# Shows installed vs available versions of Graph modules
##############################################################################
Function Get-GraphModuleStatus {
    <#
    .SYNOPSIS
    Shows the status of Microsoft Graph modules (installed vs available versions)

    .DESCRIPTION
    Checks Microsoft.Graph and Microsoft.Graph.Beta modules and displays their
    current installed version compared to the latest available version in PSGallery.
    When updates are available, prompts to run the update script.

    .PARAMETER Silent
    If specified, suppresses output and returns objects instead.

    .PARAMETER NoPrompt
    If specified, does not prompt to run the update script when updates are available.

    .EXAMPLE
    Get-GraphModuleStatus

    .EXAMPLE
    Get-GraphModuleStatus -Silent

    .EXAMPLE
    Get-GraphModuleStatus -NoPrompt

    .LINK
    https://github.com/yourusername/GraphModuleStatus
    #>

    [CmdletBinding()]
    param(
        [switch]$Silent,
        [switch]$NoPrompt
    )

    $Modules = @(
        @{ Name = "Microsoft.Graph"; Display = "Microsoft.Graph" },
        @{ Name = "Microsoft.Graph.Beta"; Display = "Microsoft.Graph.Beta" }
    )

    $Results = @()

    if (-not $Silent) {
        Write-Host ""
    }

    foreach ($Module in $Modules) {
        # Check both CurrentUser and AllUsers scopes
        $Installed = @(
            Get-InstalledPSResource -Name $Module.Name -Scope CurrentUser -ErrorAction SilentlyContinue
            Get-InstalledPSResource -Name $Module.Name -Scope AllUsers -ErrorAction SilentlyContinue
        ) | Sort-Object Version -Descending | Select-Object -First 1

        $Status = [PSCustomObject]@{
            Name             = $Module.Name
            InstalledVersion = $null
            AvailableVersion = $null
            UpdateAvailable  = $false
            Installed        = $false
        }

        if ($Installed) {
            $Status.Installed = $true
            $Status.InstalledVersion = $Installed.Version.ToString()

            # Check for updates via fast URL redirect (no download needed)
            $LatestVersion = Get-LatestPSGalleryVersion -Name $Module.Name
            if ($LatestVersion) {
                $Status.AvailableVersion = $LatestVersion.ToString()
                $Status.UpdateAvailable = ($LatestVersion -gt [version]$Status.InstalledVersion)
            }

            if (-not $Silent) {
                if ($Status.UpdateAvailable) {
                    # Update available
                    Write-Host "  [$($Module.Display)]" -ForegroundColor White -NoNewline
                    Write-Host " v$($Status.InstalledVersion) " -ForegroundColor Yellow -NoNewline
                    Write-Host "→" -ForegroundColor DarkGray -NoNewline
                    Write-Host " v$($Status.AvailableVersion)" -ForegroundColor Green
                } else {
                    # Up to date
                    Write-Host "  [$($Module.Display)]" -ForegroundColor Cyan -NoNewline
                    Write-Host " v$($Status.InstalledVersion) " -NoNewline
                    Write-Host "●" -ForegroundColor Green -NoNewline
                    Write-Host " Current" -ForegroundColor DarkGray
                }
            }
        } else {
            # Not installed - check what's available on PSGallery via fast URL redirect
            $LatestVersion = Get-LatestPSGalleryVersion -Name $Module.Name
            if ($LatestVersion) {
                $Status.AvailableVersion = $LatestVersion.ToString()
            }

            if (-not $Silent) {
                Write-Host "  [$($Module.Display)]" -ForegroundColor DarkGray -NoNewline
                Write-Host " ○ Not installed" -ForegroundColor Red -NoNewline
                if ($Status.AvailableVersion) {
                    Write-Host "  (v$($Status.AvailableVersion) available on PSGallery)" -ForegroundColor DarkGray
                } else {
                    Write-Host ""
                }
            }
        }

        $Results += $Status
    }

    if (-not $Silent) {
        Write-Host ""
    }

    # Check if any updates are available and prompt user
    $UpdatesAvailable = $Results | Where-Object { $_.UpdateAvailable -eq $true }
    
    if ($UpdatesAvailable -and -not $Silent -and -not $NoPrompt) {
        if (Test-Path -Path $script:UpdateScriptPath) {
            $Response = Read-Host "  Update available. Run Update-GraphModule now? (Y/N)"
            if ($Response -eq 'Y' -or $Response -eq 'y') {
                Write-Host ""
                Update-GraphModule
            } else {
                Write-Host ""
            }
        }
    }

    # If none are installed, offer to install them now
    $NoneInstalled = -not ($Results | Where-Object { $_.Installed -eq $true })
    $NotInstalledModules = @($Results | Where-Object { -not $_.Installed })

    if ($NoneInstalled -and $NotInstalledModules.Count -gt 0 -and -not $Silent -and -not $NoPrompt) {
        Write-Host "  No Microsoft Graph modules are installed." -ForegroundColor Yellow
        $HasVersionInfo = @($NotInstalledModules | Where-Object { $null -ne $_.AvailableVersion })
        if ($HasVersionInfo.Count -gt 0) {
            Write-Host "  The following versions are available from PSGallery:" -ForegroundColor Yellow
        }
        Write-Host ""
        foreach ($Mod in $NotInstalledModules) {
            Write-Host "    [$($Mod.Name)]" -ForegroundColor White -NoNewline
            if ($null -ne $Mod.AvailableVersion) {
                Write-Host " v$($Mod.AvailableVersion)" -ForegroundColor Cyan
            } else {
                Write-Host " (latest)" -ForegroundColor DarkGray
            }
        }
        Write-Host ""

        $InstallPrompt = Read-Host "  Would you like to install them now? (Y/N)"

        if ($InstallPrompt -match '^[Yy]') {
            Write-Host ""

            # If more than one module available, let the user choose which to install
            $ModulesToInstall = $NotInstalledModules
            if ($NotInstalledModules.Count -gt 1) {
                Write-Host "  Which modules would you like to install?" -ForegroundColor Cyan
                Write-Host ""
                $i = 1
                foreach ($Mod in $NotInstalledModules) {
                    Write-Host "    [$i] $($Mod.Name)$(if ($Mod.AvailableVersion) { "  v$($Mod.AvailableVersion)" })" -ForegroundColor White
                    $i++
                }
                Write-Host "    [A] All  (default)" -ForegroundColor White
                Write-Host ""
                $ModChoice = Read-Host "  Enter your choice"

                if ($ModChoice -match '^\d+$') {
                    $Index = [int]$ModChoice - 1
                    if ($Index -ge 0 -and $Index -lt $NotInstalledModules.Count) {
                        $ModulesToInstall = @($NotInstalledModules[$Index])
                    } else {
                        Write-Host "  Invalid choice. Installing all modules." -ForegroundColor DarkGray
                    }
                }
                # Any other input (A, Enter, etc.) installs all — no action needed
                Write-Host ""
            }

            Write-Host "  Install scope:" -ForegroundColor Cyan
            Write-Host "    [1] All Users  (Recommended)" -ForegroundColor White
            Write-Host "    [2] Current User Only" -ForegroundColor White
            Write-Host ""
            $ScopeChoice = Read-Host "  Enter your choice (1-2) [default: 1]"

            $InstallScope = if ($ScopeChoice -eq "2") { "CurrentUser" } else { "AllUsers" }

            # All Users requires elevation — auto-launch an elevated session if needed
            if ($InstallScope -eq "AllUsers") {
                $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
                if (-not $IsAdmin) {
                    Write-Host ""
                    Write-Host "  All Users requires Administrator rights." -ForegroundColor Yellow
                    Write-Host "  Launching elevated session to complete the installation..." -ForegroundColor Yellow
                    Write-Host ""

                    # Write install commands to a temp script — avoids -Command quoting issues
                    $TempScript = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.ps1'
                    $ScriptLines = @(
                        "Write-Host ''",
                        "Write-Host '  Installing Microsoft Graph modules for All Users...' -ForegroundColor Cyan",
                        "Write-Host ''"
                    )
                    foreach ($Mod in $ModulesToInstall) {
                        $ScriptLines += "Write-Host '  Installing $($Mod.Name)...' -ForegroundColor Yellow"
                        $ScriptLines += "Write-Host '  (This installs all sub-modules and may take several minutes)' -ForegroundColor Gray"
                        $ScriptLines += "Write-Host ''"
                        $ScriptLines += "Install-PSResource -Name '$($Mod.Name)' -Scope AllUsers -TrustRepository -AcceptLicense"
                        $ScriptLines += "Write-Host ''"
                        $ScriptLines += "Write-Host '  $($Mod.Name) installed successfully.' -ForegroundColor Green"
                        $ScriptLines += "Write-Host ''"
                    }
                    $ScriptLines += "Write-Host '  Done. Open a new PowerShell window before using Graph commands.' -ForegroundColor Green"
                    $ScriptLines += "Write-Host ''"
                    $ScriptLines += "Read-Host '  Press Enter to close'"
                    $ScriptLines += "Remove-Item -Path '$TempScript' -Force -ErrorAction SilentlyContinue"
                    $ScriptLines | Out-File -FilePath $TempScript -Encoding utf8

                    # Use the current process executable so elevation uses the same PS version
                    $PwshExe = (Get-Process -Id $PID).MainModule.FileName
                    try {
                        Start-Process $PwshExe -Verb RunAs -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $TempScript
                        Write-Host "  Elevated installer launched — check the new window for progress." -ForegroundColor Cyan
                        Write-Host ""
                    }
                    catch {
                        Remove-Item -Path $TempScript -Force -ErrorAction SilentlyContinue
                        Write-Host "  ERROR: Could not launch elevated session - $_" -ForegroundColor Red
                        Write-Host "  Run PowerShell as Administrator and call Get-GraphModuleStatus again." -ForegroundColor Yellow
                        Write-Host ""
                    }
                    return
                }
            }

            $ScopeDisplay = if ($InstallScope -eq "AllUsers") { "All Users" } else { "Current User Only" }
            Write-Host ""
            Write-Host "  Modules will be installed for: $ScopeDisplay" -ForegroundColor Gray
            Write-Host ""

            $InstallSuccess = 0
            $InstallFailed = 0

            foreach ($Mod in $ModulesToInstall) {
                Write-Host "  Installing $($Mod.Name)$(if ($Mod.AvailableVersion) { " v$($Mod.AvailableVersion)" })..." -ForegroundColor Yellow
                Write-Host "  (This installs all sub-modules and may take several minutes)" -ForegroundColor Gray
                Write-Host ""
                try {
                    Install-PSResource -Name $Mod.Name -Scope $InstallScope -TrustRepository -AcceptLicense -ErrorAction Stop
                    Write-Host "  $($Mod.Name) installed successfully." -ForegroundColor Green
                    $InstallSuccess++
                }
                catch {
                    Write-Host "  ERROR: Failed to install $($Mod.Name) - $_" -ForegroundColor Red
                    $InstallFailed++
                }
                Write-Host ""
            }

            Write-Host "  Install complete: $InstallSuccess succeeded, $InstallFailed failed." -ForegroundColor $(if ($InstallFailed -gt 0) { "Yellow" } else { "Green" })
            Write-Host ""
            if ($InstallSuccess -gt 0) {
                Write-Host "  Open a new PowerShell window before running any Graph commands." -ForegroundColor Yellow
                Write-Host ""
            }
        }
    }

    if ($Silent) {
        return $Results
    }
}

##############################################################################
# Update-GraphModule
# Runs the update script to clean install/reinstall Graph modules
##############################################################################
Function Update-GraphModule {
    <#
    .SYNOPSIS
    Runs the Microsoft Graph module update script

    .DESCRIPTION
    Invokes the Update-MicrosoftGraph.ps1 script which performs a clean
    uninstall and reinstall of Microsoft Graph modules.

    .EXAMPLE
    Update-GraphModule

    .LINK
    https://github.com/yourusername/GraphModuleStatus
    #>

    [CmdletBinding()]
    param()

    if (Test-Path -Path $script:UpdateScriptPath) {
        Write-Host "Running Microsoft Graph update script..." -ForegroundColor Cyan
        Write-Host ""
        & $script:UpdateScriptPath
    } else {
        Write-Host "Update script not found at: $script:UpdateScriptPath" -ForegroundColor Red
        Write-Host "Please ensure the script exists or update the path in the module." -ForegroundColor Yellow
    }
}

##############################################################################
# Add-GraphModuleStatusToProfile
# Adds GraphModuleStatus check to PowerShell profile
##############################################################################
Function Add-GraphModuleStatusToProfile {
    <#
    .SYNOPSIS
    Add GraphModuleStatus check to PowerShell Profile

    .DESCRIPTION
    Adds the GraphModuleStatus module import and status check to your PowerShell profile
    so it runs automatically when you start PowerShell.

    Needs to be executed separately for PowerShell v5 and v7.

    .EXAMPLE
    Add-GraphModuleStatusToProfile

    .LINK
    https://github.com/yourusername/GraphModuleStatus
    #>

    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param()

    $ProfileContent = @"

# GraphModuleStatus: Check Microsoft Graph module status on startup
Import-Module -Name GraphModuleStatus -ErrorAction SilentlyContinue
Get-GraphModuleStatus
"@

    if (-not (Test-Path -Path $Profile)) {
        # No Profile found
        Write-Host "No PowerShell Profile exists. Creating new Profile with GraphModuleStatus setup." -ForegroundColor Yellow
        $ProfileContent | Out-File -FilePath $Profile -Encoding utf8 -Force
        Write-Host "Profile created at: $Profile" -ForegroundColor Green
    } else {
        # Profile found
        $ExistingContent = Get-Content -Path $Profile -Encoding utf8 -Raw
        $Match = $ExistingContent | Where-Object { $_ -match "GraphModuleStatus" }

        if ($Match) {
            # GraphModuleStatus already in Profile
            Write-Host "GraphModuleStatus is already in your PowerShell Profile." -ForegroundColor Yellow
        } else {
            # GraphModuleStatus not in Profile
            Write-Host "Adding GraphModuleStatus to existing PowerShell Profile..." -ForegroundColor Yellow
            Add-Content -Path $Profile -Value $ProfileContent -Encoding utf8
            Write-Host "GraphModuleStatus added to: $Profile" -ForegroundColor Green
        }
    }
}

##############################################################################
# Remove-GraphModuleStatusFromProfile
# Removes GraphModuleStatus from PowerShell profile
##############################################################################
Function Remove-GraphModuleStatusFromProfile {
    <#
    .SYNOPSIS
    Remove GraphModuleStatus from PowerShell Profile

    .DESCRIPTION
    Removes the GraphModuleStatus module import and status check from your PowerShell profile.

    .EXAMPLE
    Remove-GraphModuleStatusFromProfile

    .LINK
    https://github.com/yourusername/GraphModuleStatus
    #>

    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param()

    if (-not (Test-Path -Path $Profile)) {
        Write-Host "No PowerShell Profile exists." -ForegroundColor Yellow
        return
    }

    $Content = Get-Content -Path $Profile -Encoding utf8
    $NewContent = $Content | Where-Object { 
        $_ -notmatch "GraphModuleStatus" -and 
        $_ -notmatch "# GraphModuleStatus:" 
    }

    if ($Content.Count -ne $NewContent.Count) {
        $NewContent | Out-File -FilePath $Profile -Encoding utf8 -Force
        Write-Host "GraphModuleStatus removed from PowerShell Profile." -ForegroundColor Green
    } else {
        Write-Host "GraphModuleStatus was not found in your PowerShell Profile." -ForegroundColor Yellow
    }
}

##############################################################################
# Module Load Message
##############################################################################
if (-not (Test-Path -Path $Profile)) {
    Write-Host "Tip: Run Add-GraphModuleStatusToProfile to check Graph module status on startup." -ForegroundColor DarkGray
} else {
    $Content = Get-Content -Path $Profile -Encoding utf8 -Raw -ErrorAction SilentlyContinue
    if ($Content -notmatch "GraphModuleStatus") {
        Write-Host "Tip: Run Add-GraphModuleStatusToProfile to check Graph module status on startup." -ForegroundColor DarkGray
    }
}
