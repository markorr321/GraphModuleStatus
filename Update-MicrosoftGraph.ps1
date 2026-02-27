<#
.SYNOPSIS
    Performs a clean installation of Microsoft Graph PowerShell modules.

.DESCRIPTION
    This script provides a comprehensive solution for resolving Microsoft Graph
    PowerShell module version conflicts and ensuring a clean, consistent installation. It is
    designed to address common issues that occur when multiple versions of modules are installed,
    which can cause assembly loading errors and command failures.

    The script performs the following operations:

    1. SESSION CLEANUP
       Removes all currently loaded Microsoft Graph modules from the PowerShell session
       to prevent file locking issues during uninstallation.

    2. MODULE UNINSTALLATION (Iterative)
       Systematically uninstalls all installed modules using an iterative approach with
       garbage collection to handle dependencies and file locks. Uses both Get-InstalledModule
       and Get-Module -ListAvailable for comprehensive detection.

    3. FOLDER CLEANUP
       Scans common PowerShell module directories for any leftover module folders
       that may have been orphaned, and removes them to ensure a clean slate.

    4. FRESH INSTALLATION
       Installs your choice of Microsoft.Graph and/or Microsoft.Graph.Beta
       modules from the PowerShell Gallery.

    5. VALIDATION
       Verifies the installation was successful and confirms that all module versions
       are aligned to prevent future version mismatch issues.

    This script is particularly useful when encountering errors such as:
    - "Assembly with same name is already loaded"
    - "Could not load file or assembly 'Microsoft.Graph.Authentication'"
    - Commands not being recognized after module updates
    - Multiple authentication prompts when using Graph/Entra cmdlets

.INPUTS
    None. This script does not accept pipeline input.

.OUTPUTS
    Console output showing the progress and results of each step. Upon completion,
    displays a summary of installed module versions.

.EXAMPLE
    .\Update-MicrosoftGraph.ps1

    Runs the script to perform a complete reset and fresh installation of
    Microsoft Graph modules.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Update-MicrosoftGraph.ps1

    Runs the script with execution policy bypass if scripts are restricted.

.NOTES
    File Name      : Update-MicrosoftGraph.ps1
    Author         : Mark Orr
    Prerequisite   : PowerShell 5.1 or later
                     Internet connectivity to PowerShell Gallery
                     Administrator rights may be needed for system-wide module locations
    Version        : 1.1

.LINK
    https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation

.LINK
    https://www.powershellgallery.com/packages/Microsoft.Graph
#>

# ============================================================
# Self-elevation to Administrator if not already elevated
# ============================================================
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $IsAdmin) {
    Write-Host ""
    Write-Host "This script requires Administrator privileges." -ForegroundColor Yellow
    Write-Host "Elevating to Administrator..." -ForegroundColor Yellow
    Write-Host ""

    # Build the argument list to re-run this script
    $ScriptPath = $MyInvocation.MyCommand.Definition

    try {
        # Start new elevated PowerShell process
        $Process = Start-Process -FilePath "pwsh.exe" -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$ScriptPath`"" -Verb RunAs -Wait -PassThru
        exit $Process.ExitCode
    }
    catch {
        Write-Host "ERROR: Failed to elevate to Administrator." -ForegroundColor Red
        Write-Host "Please run this script as Administrator manually." -ForegroundColor Red
        Write-Host ""
        Read-Host "Press Enter to exit"
        exit 1
    }
}

Write-Host ""
Write-Host "Running as Administrator" -ForegroundColor Green
Write-Host ""

# Save original window title to restore later
$OriginalTitle = $Host.UI.RawUI.WindowTitle

# Create synchronized hashtable to share state with background runspace
$script:TitleState = [hashtable]::Synchronized(@{
    StartTime = $null
    CurrentStep = ""
    IsRunning = $false
    OriginalTitle = $OriginalTitle
    RawUI = $Host.UI.RawUI
})

# Create background runspace for real-time clock updates
$script:TitleRunspace = [runspacefactory]::CreateRunspace()
$script:TitleRunspace.ApartmentState = "STA"
$script:TitleRunspace.ThreadOptions = "ReuseThread"
$script:TitleRunspace.Open()
$script:TitleRunspace.SessionStateProxy.SetVariable("TitleState", $script:TitleState)

$script:TitlePipeline = $script:TitleRunspace.CreatePipeline()
$script:TitlePipeline.Commands.AddScript({
    $LastTitle = ""
    while ($TitleState.IsRunning) {
        if ($TitleState.StartTime) {
            $Elapsed = [DateTime]::Now - $TitleState.StartTime
            $TimeStr = "{0:D2}:{1:D2}:{2:D2}" -f [int]$Elapsed.TotalHours, $Elapsed.Minutes, $Elapsed.Seconds
            if ($TitleState.CurrentStep) {
                $NewTitle = "Microsoft Graph Update - Elapsed: $TimeStr - $($TitleState.CurrentStep)"
            }
            else {
                $NewTitle = "Microsoft Graph Update - Elapsed: $TimeStr"
            }
            # Only update if title changed to reduce flicker
            if ($NewTitle -ne $LastTitle) {
                $TitleState.RawUI.WindowTitle = $NewTitle
                $LastTitle = $NewTitle
            }
        }
        Start-Sleep -Seconds 1
    }
    $TitleState.RawUI.WindowTitle = $TitleState.OriginalTitle
})

# Function to update the current step (background runspace handles the clock)
function Update-WindowTitle {
    param (
        [string]$CurrentStep = ""
    )
    $script:TitleState.CurrentStep = $CurrentStep
}

# Step header function
function Write-Progress-Step {
    param (
        [int]$Step,
        [int]$TotalSteps,
        [string]$StepName,
        [string]$Status = "In Progress"
    )

    Update-WindowTitle -CurrentStep "Step $Step of $TotalSteps : $StepName"

    Write-Host ""
    Write-Host "  ── Step $Step of $TotalSteps : $StepName" -ForegroundColor Cyan
    Write-Host ""
}

$TotalSteps = 6

Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "      Microsoft Graph Module Updater            " -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  This script will:" -ForegroundColor White
Write-Host "    1. Clear loaded modules from session" -ForegroundColor Gray
Write-Host "    2. Uninstall existing modules (iterative)" -ForegroundColor Gray
Write-Host "    3. Clean up leftover module folders" -ForegroundColor Gray
Write-Host "    4. Install packages (you choose which)" -ForegroundColor Gray
Write-Host "    5. Import the new modules" -ForegroundColor Gray
Write-Host "    6. Validate the installation" -ForegroundColor Gray
Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# Discover installed modules
Write-Host "  Scanning for installed modules..." -ForegroundColor Gray

# Check for any Microsoft.Graph modules (stable - not Beta)
# Checks both old PowerShellGet (Get-InstalledModule) and new PSResourceGet (Get-InstalledPSResource)
$GraphModules = @()
$GraphModules += Get-InstalledModule Microsoft.Graph* -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike "Microsoft.Graph.Beta*" }
$GraphModules += Get-InstalledPSResource -Name "Microsoft.Graph*" -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike "Microsoft.Graph.Beta*" }
$GraphPathModules = Get-Module -ListAvailable Microsoft.Graph* -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike "Microsoft.Graph.Beta*" }
$HasGraph = ($GraphModules.Count -gt 0) -or ($GraphPathModules.Count -gt 0)
$GraphVer = if ($GraphModules.Count -gt 0) { "v$($GraphModules[0].Version)" } elseif ($GraphPathModules.Count -gt 0) { "v$($GraphPathModules[0].Version) (path only)" } else { "" }

# Check for any Microsoft.Graph.Beta modules
$BetaModules = @()
$BetaModules += Get-InstalledModule Microsoft.Graph.Beta* -ErrorAction SilentlyContinue
$BetaModules += Get-InstalledPSResource -Name "Microsoft.Graph.Beta*" -ErrorAction SilentlyContinue
$BetaPathModules = Get-Module -ListAvailable Microsoft.Graph.Beta* -ErrorAction SilentlyContinue
$HasBeta = ($BetaModules.Count -gt 0) -or ($BetaPathModules.Count -gt 0)
$BetaVer = if ($BetaModules.Count -gt 0) { "v$($BetaModules[0].Version)" } elseif ($BetaPathModules.Count -gt 0) { "v$($BetaPathModules[0].Version) (path only)" } else { "" }

Write-Host ""
Write-Host "  Discovered modules:" -ForegroundColor White
if ($HasGraph) {
    Write-Host "    - Microsoft.Graph (stable)    $GraphVer" -ForegroundColor Green
}
if ($HasBeta) {
    Write-Host "    - Microsoft.Graph.Beta        $BetaVer" -ForegroundColor Green
}
if (-not $HasGraph -and -not $HasBeta) {
    Write-Host "    (none found)" -ForegroundColor Yellow
}
Write-Host ""

# Prompt user for which modules to manage
Write-Host "  Which modules would you like to uninstall?" -ForegroundColor Cyan
Write-Host ""
Write-Host "    [1] Microsoft.Graph (stable) only" -ForegroundColor White
Write-Host "    [2] Microsoft.Graph.Beta only" -ForegroundColor White
Write-Host "    [3] Both Microsoft.Graph and Microsoft.Graph.Beta" -ForegroundColor White
Write-Host ""
$ModuleChoice = Read-Host "  Enter your choice (1-3) [default: 3]"

if ([string]::IsNullOrWhiteSpace($ModuleChoice)) {
    $ModuleChoice = "3"
}

# Set flags based on choice
$script:IncludeGraph = $false
$script:IncludeBeta = $false

switch ($ModuleChoice) {
    "1" { $script:IncludeGraph = $true }
    "2" { $script:IncludeBeta = $true }
    "3" { $script:IncludeGraph = $true; $script:IncludeBeta = $true }
    default {
        Write-Host "  Invalid choice. Defaulting to Microsoft.Graph and Microsoft.Graph.Beta." -ForegroundColor Yellow
        $script:IncludeGraph = $true
        $script:IncludeBeta = $true
    }
}

# Build module filter pattern based on selection
$script:ModulePatterns = @()
if ($script:IncludeGraph -or $script:IncludeBeta) {
    $script:ModulePatterns += "Microsoft.Graph*"
}

Write-Host ""
Write-Host "  Selected for uninstall:" -ForegroundColor Gray
if ($script:IncludeGraph) { Write-Host "    - Microsoft.Graph (stable)" -ForegroundColor Yellow }
if ($script:IncludeBeta) { Write-Host "    - Microsoft.Graph.Beta" -ForegroundColor Yellow }
Write-Host ""
Start-Sleep -Seconds 2

# Start the stopwatch for timing (script scope so Update-WindowTitle can access it)
$script:Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# Start the background title updater
$script:TitleState.StartTime = [DateTime]::Now
$script:TitleState.IsRunning = $true
$script:TitlePipeline.InvokeAsync()

# ============================================================
# Step 1: Remove loaded modules from current session (0% -> 20%)
# ============================================================
Write-Progress-Step -Step 1 -TotalSteps $TotalSteps -StepName "Clearing loaded modules from session"

# Get loaded modules based on user selection
$LoadedModules = @()
foreach ($Pattern in $script:ModulePatterns) {
    $LoadedModules += Get-Module $Pattern | Select-Object -ExpandProperty Name
}
$LoadedModules = $LoadedModules | Select-Object -Unique

# Apply selection filter so only the chosen module family is affected
if ($script:IncludeGraph -and -not $script:IncludeBeta) {
    $LoadedModules = @($LoadedModules | Where-Object { $_ -notlike "Microsoft.Graph.Beta*" })
} elseif (-not $script:IncludeGraph -and $script:IncludeBeta) {
    $LoadedModules = @($LoadedModules | Where-Object { $_ -like "Microsoft.Graph.Beta*" })
}

if ($LoadedModules) {
    $LoadedTotal = @($LoadedModules).Count
    $LoadedCounter = 0
    $LoadedSuccess = 0
    $LoadedFailed = 0

    foreach ($Module in $LoadedModules) {
        $LoadedCounter++
        try {
            Remove-Module -Name $Module -Force -ErrorAction Stop
            $LoadedSuccess++
        }
        catch {
            $LoadedFailed++
        }
    }

    if ($LoadedFailed -eq 0) {
        Write-Host "  Cleared $LoadedTotal loaded module(s) from session." -ForegroundColor Green
    } else {
        Write-Host "  Cleared $LoadedSuccess of $LoadedTotal loaded module(s). $LoadedFailed could not be removed." -ForegroundColor Yellow
    }
}
else {
    Write-Host "  No matching modules loaded in session." -ForegroundColor Green
}

# ============================================================
# Step 2: Uninstall modules (iterative with garbage collection)
# ============================================================
Write-Progress-Step -Step 2 -TotalSteps $TotalSteps -StepName "Uninstalling existing modules"

$Iteration = 1
$MaxIterations = 10

do {
    # Get installed modules via old PowerShellGet (Get-InstalledModule)
    $InstalledModulesOld = @()
    foreach ($Pattern in $script:ModulePatterns) {
        $InstalledModulesOld += Get-InstalledModule $Pattern -ErrorAction SilentlyContinue
    }

    # Get installed modules via new PSResourceGet (Get-InstalledPSResource)
    $InstalledModulesNew = @()
    foreach ($Pattern in $script:ModulePatterns) {
        $InstalledModulesNew += Get-InstalledPSResource -Name $Pattern -ErrorAction SilentlyContinue
    }

    # Merge both lists, normalising to Name/Version/ModuleBase
    $InstalledModules = @()
    $InstalledModules += $InstalledModulesOld | Select-Object -Property Name, Version, @{N='ModuleBase';E={$_.InstalledLocation}}, @{N='Source';E={'Old'}}
    foreach ($PSRMod in $InstalledModulesNew) {
        $AlreadyListed = $InstalledModules | Where-Object { $_.Name -eq $PSRMod.Name -and $_.Version -eq $PSRMod.Version.ToString() }
        if (-not $AlreadyListed) {
            $InstalledModules += [PSCustomObject]@{
                Name      = $PSRMod.Name
                Version   = $PSRMod.Version.ToString()
                ModuleBase = $PSRMod.InstalledLocation
                Source    = 'New'
            }
        }
    }

    # Get available modules using Get-Module -ListAvailable (catches modules not in gallery)
    $AvailableModules = @()
    foreach ($Pattern in $script:ModulePatterns) {
        $AvailableModules += Get-Module -ListAvailable $Pattern -ErrorAction SilentlyContinue
    }
    $AvailableModules = $AvailableModules | Select-Object -Unique -Property Name, Version, ModuleBase

    # Apply selection filter so only the chosen module family is uninstalled
    if ($script:IncludeGraph -and -not $script:IncludeBeta) {
        $InstalledModules = @($InstalledModules | Where-Object { $_.Name -notlike "Microsoft.Graph.Beta*" })
        $AvailableModules = @($AvailableModules | Where-Object { $_.Name -notlike "Microsoft.Graph.Beta*" })
    } elseif (-not $script:IncludeGraph -and $script:IncludeBeta) {
        $InstalledModules = @($InstalledModules | Where-Object { $_.Name -like "Microsoft.Graph.Beta*" })
        $AvailableModules = @($AvailableModules | Where-Object { $_.Name -like "Microsoft.Graph.Beta*" })
    }

    $TotalFound = @($InstalledModules).Count + @($AvailableModules).Count

    if ($TotalFound -eq 0) {
        Write-Host "  No modules found. Cleanup complete!" -ForegroundColor Green
        break
    }

    Write-Host "  Get-InstalledModule found:      $(@($InstalledModulesOld).Count) module(s)" -ForegroundColor Gray
    Write-Host "  Get-InstalledPSResource found:  $(@($InstalledModulesNew).Count) module(s)" -ForegroundColor Gray
    Write-Host "  Get-Module -ListAvailable found: $(@($AvailableModules).Count) module(s)" -ForegroundColor Gray
    Write-Host ""

    # Uninstall gallery-installed modules — try PSResourceGet first, fall back to old PowerShellGet
    if ($InstalledModules) {
        # Group sub-modules by their root package
        $InstallGroups = @{}
        foreach ($Module in $InstalledModules) {
            $Root = if ($Module.Name -like "Microsoft.Graph.Beta*") { "Microsoft.Graph.Beta" }
                    elseif ($Module.Name -like "Microsoft.Graph*") { "Microsoft.Graph" }
                    else { $Module.Name }
            if (-not $InstallGroups.ContainsKey($Root)) { $InstallGroups[$Root] = [System.Collections.Generic.List[object]]::new() }
            $InstallGroups[$Root].Add($Module)
        }

        foreach ($Root in ($InstallGroups.Keys | Sort-Object)) {
            Write-Host "  Uninstalling $Root..." -ForegroundColor Yellow
            Write-Host "  (This removes all sub-modules)" -ForegroundColor Gray
            Write-Host ""

            $PendingCount = 0
            $FailMessages = [System.Collections.Generic.List[string]]::new()
            foreach ($Module in $InstallGroups[$Root]) {
                $Uninstalled = $false

                # Try Uninstall-PSResource (exact version)
                try {
                    Uninstall-PSResource -Name $Module.Name -Version $Module.Version -ErrorAction Stop
                    $Uninstalled = $true
                }
                catch { $LastError = $_.Exception.Message }

                # Try Uninstall-PSResource (no version — catches all)
                if (-not $Uninstalled) {
                    try {
                        Uninstall-PSResource -Name $Module.Name -ErrorAction Stop
                        $Uninstalled = $true
                    }
                    catch { $LastError = $_.Exception.Message }
                }

                # Fall back to old Uninstall-Module
                if (-not $Uninstalled) {
                    try {
                        Uninstall-Module -Name $Module.Name -AllVersions -Force -ErrorAction Stop
                        $Uninstalled = $true
                    }
                    catch { $LastError = $_.Exception.Message }
                }

                if (-not $Uninstalled) {
                    $PendingCount++
                    $FailMessages.Add("    $($Module.Name) v$($Module.Version): $LastError")
                }
            }

            if ($PendingCount -eq 0) {
                Write-Host "  $Root uninstalled successfully." -ForegroundColor Green
            } else {
                Write-Host "  ${Root}: $PendingCount sub-module(s) could not be uninstalled:" -ForegroundColor Yellow
                foreach ($Msg in $FailMessages) {
                    Write-Host $Msg -ForegroundColor DarkGray
                }
            }
            Write-Host ""
        }
    }

    # Handle modules found via Get-Module -ListAvailable that aren't tracked by either package manager.
    # Try proper uninstall first — only fall back to folder deletion if that fails.
    if ($AvailableModules) {
        # Filter to only modules not already handled above
        $OrphanModules = $AvailableModules | Where-Object {
            $ModName = $_.Name
            -not ($InstalledModules | Where-Object { $_.Name -eq $ModName })
        }

        if ($OrphanModules) {
            # Group by root module name
            $OrphanGroups = @{}
            foreach ($Module in $OrphanModules) {
                $Root = if ($Module.Name -like "Microsoft.Graph.Beta*") { "Microsoft.Graph.Beta" }
                        elseif ($Module.Name -like "Microsoft.Graph*") { "Microsoft.Graph" }
                        elseif ($Module.Name -like "Microsoft.Entra*") { "Microsoft.Entra" }
                        else { $Module.Name }
                if (-not $OrphanGroups.ContainsKey($Root)) { $OrphanGroups[$Root] = [System.Collections.Generic.List[object]]::new() }
                $OrphanGroups[$Root].Add($Module)
            }

            foreach ($Root in ($OrphanGroups.Keys | Sort-Object)) {
                Write-Host "  Uninstalling $Root (untracked)..." -ForegroundColor Yellow
                Write-Host "  (This removes all sub-modules)" -ForegroundColor Gray
                Write-Host ""

                $FailCount = 0
                foreach ($Module in $OrphanGroups[$Root]) {
                    $Uninstalled = $false

                    # Try Uninstall-PSResource first
                    try {
                        Uninstall-PSResource -Name $Module.Name -ErrorAction Stop
                        $Uninstalled = $true
                    }
                    catch { }

                    # Fall back to Uninstall-Module
                    if (-not $Uninstalled) {
                        try {
                            Uninstall-Module -Name $Module.Name -AllVersions -Force -ErrorAction Stop
                            $Uninstalled = $true
                        }
                        catch { }
                    }

                    # Last resort: folder deletion
                    if (-not $Uninstalled) {
                        try {
                            if (Test-Path $Module.ModuleBase) {
                                Remove-Item -Path $Module.ModuleBase -Recurse -Force -ErrorAction Stop
                                $Uninstalled = $true
                            }
                        }
                        catch {
                            Write-Host "    Failed to remove: $($Module.ModuleBase) - $_" -ForegroundColor Red
                            $FailCount++
                        }
                    }
                }

                if ($FailCount -eq 0) {
                    Write-Host "  $Root removed successfully." -ForegroundColor Green
                } else {
                    Write-Host "  ${Root}: $FailCount sub-module(s) could not be deleted." -ForegroundColor Yellow
                }
                Write-Host ""
            }
        }
    }

    # Force garbage collection to release any file locks
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 2

    $Iteration++

    if ($Iteration -gt $MaxIterations) {
        Write-Host "  Reached maximum iterations ($MaxIterations). Some modules may need manual removal." -ForegroundColor Yellow
        break
    }

} while ($true)

Write-Host ""
Write-Host "  Uninstall phase complete." -ForegroundColor Green

# ============================================================
# Step 3: Clean up leftover module folders (40% -> 60%)
# ============================================================
Write-Progress-Step -Step 3 -TotalSteps $TotalSteps -StepName "Cleaning up leftover module folders"

# All possible module locations to check (PowerShell 7 only)
$ModulePaths = @(
    "$env:USERPROFILE\Documents\PowerShell\Modules",
    "$env:USERPROFILE\OneDrive\Documents\PowerShell\Modules",
    "$env:ProgramFiles\PowerShell\Modules",
    "$env:ProgramFiles\PowerShell\7\Modules",
    "$env:LOCALAPPDATA\PowerShell\Modules"
)

# Also add paths from PSModulePath environment variable (exclude WindowsPowerShell paths)
$PSModulePaths = $env:PSModulePath -split ';' | Where-Object { $_ -ne '' -and $_ -notlike '*WindowsPowerShell*' }
foreach ($PSPath in $PSModulePaths) {
    if ($PSPath -and ($ModulePaths -notcontains $PSPath)) {
        $ModulePaths += $PSPath
    }
}

# Remove duplicates and non-existent paths
$ModulePaths = $ModulePaths | Select-Object -Unique | Where-Object { Test-Path $_ -ErrorAction SilentlyContinue }

# Build folder filter patterns based on user selection
$FolderPatterns = @()
if ($script:IncludeGraph -or $script:IncludeBeta) {
    $FolderPatterns += "Microsoft.Graph*"
}

# Collect all items to delete across all module paths
$ItemsToDelete = [System.Collections.Generic.List[string]]::new()

foreach ($Path in $ModulePaths) {
    $TargetFolders = @()
    foreach ($Pattern in $FolderPatterns) {
        $TargetFolders += Get-ChildItem -Path $Path -Directory -Filter $Pattern -ErrorAction SilentlyContinue
    }

    # Apply selection filter so only the chosen module family is cleaned up
    if ($script:IncludeGraph -and -not $script:IncludeBeta) {
        $TargetFolders = @($TargetFolders | Where-Object { $_.Name -notlike "Microsoft.Graph.Beta*" })
    } elseif (-not $script:IncludeGraph -and $script:IncludeBeta) {
        $TargetFolders = @($TargetFolders | Where-Object { $_.Name -like "Microsoft.Graph.Beta*" })
    }

    foreach ($Folder in $TargetFolders) {
        $ItemsToDelete.Add($Folder.FullName)
    }
}

if ($ItemsToDelete.Count -eq 0) {
    Write-Host "  No leftover folders found." -ForegroundColor Green
}
else {
    Write-Host "  Found $($ItemsToDelete.Count) item(s) to remove." -ForegroundColor Gray
    Write-Host ""

    $MaxRetries = 3
    $RetryCount = 0
    $Pending = $ItemsToDelete.ToArray()
    $TotalRemoved = 0

    while ($Pending.Count -gt 0 -and $RetryCount -lt $MaxRetries) {
        if ($RetryCount -gt 0) {
            Write-Host "  Retrying $($Pending.Count) locked item(s) (attempt $($RetryCount + 1) of $MaxRetries)..." -ForegroundColor Yellow
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()
            Start-Sleep -Seconds 3
        }

        $StillPending = [System.Collections.Generic.List[string]]::new()
        foreach ($ItemPath in $Pending) {
            if (-not (Test-Path $ItemPath)) {
                $TotalRemoved++
                continue
            }
            try {
                Remove-Item -Path $ItemPath -Recurse -Force -ErrorAction Stop
                $TotalRemoved++
            }
            catch {
                $StillPending.Add($ItemPath)
            }
        }
        $Pending = $StillPending.ToArray()
        $RetryCount++
    }

    if ($Pending.Count -eq 0) {
        Write-Host "  Folder cleanup complete. Removed $TotalRemoved item(s)." -ForegroundColor Green
    }
    else {
        Write-Host "  $($Pending.Count) item(s) could not be deleted (locked by session DLLs):" -ForegroundColor Yellow
        foreach ($Remaining in $Pending) {
            Write-Host "    - $Remaining" -ForegroundColor Yellow
        }
        Write-Host ""
        Write-Host "  These are most likely DLL files still locked by this PowerShell session." -ForegroundColor Yellow
        Write-Host "  To complete the cleanup:" -ForegroundColor Yellow
        Write-Host "    1. Close this PowerShell window" -ForegroundColor White
        Write-Host "    2. Open a new elevated PowerShell window" -ForegroundColor White
        Write-Host "    3. Re-run this script" -ForegroundColor White
        Write-Host ""
        Write-Host "  The fresh install in Step 4 will still proceed." -ForegroundColor DarkGray
    }
}

# ============================================================
# Step 4: Install modules (60% -> 80%)
# ============================================================
Write-Progress-Step -Step 4 -TotalSteps $TotalSteps -StepName "Installing selected modules"

Write-Host ""
Write-Host "  Which modules would you like to install?" -ForegroundColor Cyan
Write-Host ""
Write-Host "    [1] Microsoft.Graph (stable) only" -ForegroundColor White
Write-Host "    [2] Microsoft.Graph.Beta only" -ForegroundColor White
Write-Host "    [3] Both Microsoft.Graph and Microsoft.Graph.Beta" -ForegroundColor White
Write-Host "    [0] Skip installation" -ForegroundColor White
Write-Host ""
$InstallChoice = Read-Host "  Enter your choice (0-3) [default: 3]"

if ([string]::IsNullOrWhiteSpace($InstallChoice)) {
    $InstallChoice = "3"
}

# Set install flags based on choice (script scope for Steps 5 & 6)
$script:InstallGraph = $false
$script:InstallBeta = $false

switch ($InstallChoice) {
    "0" {
        Write-Host "  Skipping installation." -ForegroundColor Yellow
    }
    "1" { $script:InstallGraph = $true }
    "2" { $script:InstallBeta = $true }
    "3" { $script:InstallGraph = $true; $script:InstallBeta = $true }
    default {
        Write-Host "  Invalid choice. Defaulting to Microsoft.Graph and Microsoft.Graph.Beta." -ForegroundColor Yellow
        $script:InstallGraph = $true
        $script:InstallBeta = $true
    }
}

# Ask for installation scope if user chose to install modules
$script:InstallScope = "AllUsers"
if ($InstallChoice -ne "0") {
    Write-Host ""
    Write-Host "  Where would you like to install the modules?" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    [1] All Users" -ForegroundColor White
    Write-Host "    [2] Current User Only" -ForegroundColor White
    Write-Host ""
    Write-Host "        Recommended: All Users" -ForegroundColor Gray
    Write-Host ""
    $ScopeChoice = Read-Host "  Enter your choice (1-2) [default: 1]"

    if ([string]::IsNullOrWhiteSpace($ScopeChoice)) {
        $ScopeChoice = "1"
    }

    switch ($ScopeChoice) {
        "1" { $script:InstallScope = "AllUsers" }
        "2" { $script:InstallScope = "CurrentUser" }
        default {
            Write-Host "  Invalid choice. Defaulting to All Users." -ForegroundColor Yellow
            $script:InstallScope = "AllUsers"
        }
    }
    $ScopeDisplay = if ($script:InstallScope -eq "AllUsers") { "All Users" } else { "Current User Only" }
    Write-Host ""
    Write-Host "  Modules will be installed: $ScopeDisplay" -ForegroundColor Gray
}

if ($InstallChoice -ne "0") {
    Write-Host ""

    $InstallSuccess = 0
    $InstallFailed = 0

    # Suppress native Install-Module progress output
    $PrevProgressPreference = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'

    if ($script:InstallGraph) {
        Write-Host "  Installing Microsoft.Graph..." -ForegroundColor Yellow
        Write-Host "  (This will install all sub-modules - may take several minutes)" -ForegroundColor Gray
        Write-Host ""
        try {
            Install-Module Microsoft.Graph -Scope $script:InstallScope -Force -AllowClobber -ErrorAction Stop
            Write-Host "  Microsoft.Graph installed successfully." -ForegroundColor Green
            $InstallSuccess++
        }
        catch {
            Write-Host "  ERROR: Failed to install Microsoft.Graph - $_" -ForegroundColor Red
            Write-Host "  Check your internet connection and try again." -ForegroundColor DarkGray
            $InstallFailed++
        }
        Write-Host ""
    }

    if ($script:InstallBeta) {
        Write-Host "  Installing Microsoft.Graph.Beta..." -ForegroundColor Yellow
        Write-Host "  (This will install all sub-modules - may take several minutes)" -ForegroundColor Gray
        Write-Host ""
        try {
            Install-Module Microsoft.Graph.Beta -Scope $script:InstallScope -Force -AllowClobber -ErrorAction Stop
            Write-Host "  Microsoft.Graph.Beta installed successfully." -ForegroundColor Green
            $InstallSuccess++
        }
        catch {
            Write-Host "  ERROR: Failed to install Microsoft.Graph.Beta - $_" -ForegroundColor Red
            Write-Host "  Check your internet connection and try again." -ForegroundColor DarkGray
            $InstallFailed++
        }
        Write-Host ""
    }

    $ProgressPreference = $PrevProgressPreference

    # Summary message
    Write-Host "  Install complete: $InstallSuccess succeeded, $InstallFailed failed." -ForegroundColor Green

    if ($InstallFailed -gt 0) {
        Write-Host "  WARNING: Some packages failed to install." -ForegroundColor Yellow
    }
}

# ============================================================
# Step 5: Import the new modules (80% -> 90%)
# ============================================================
Write-Progress-Step -Step 5 -TotalSteps $TotalSteps -StepName "Importing installed modules"

Write-Host "  Skipping bulk import - PowerShell auto-loads modules on demand." -ForegroundColor Gray
Write-Host "  Importing only Microsoft.Graph.Authentication for immediate use..." -ForegroundColor Gray
Write-Host ""

try {
    Import-Module Microsoft.Graph.Authentication -Force -ErrorAction Stop
    Write-Host "    Imported: Microsoft.Graph.Authentication" -ForegroundColor Green
}
catch {
    Write-Host "    Failed to import Microsoft.Graph.Authentication: $_" -ForegroundColor Yellow
}

Write-Host ""

# ============================================================
# Step 6: Validation (90% -> 100%)
# ============================================================
Write-Progress-Step -Step 6 -TotalSteps $TotalSteps -StepName "Validating installation"

Write-Host "  Checking installed modules..." -ForegroundColor Gray

# Get module info based on what was installed
$GraphModule = $null
$GraphBetaModule = $null
$AuthModule = $null

if ($script:InstallGraph) {
    $GraphModule = Get-InstalledModule Microsoft.Graph -ErrorAction SilentlyContinue
    if (-not $GraphModule) { $GraphModule = Get-InstalledPSResource -Name "Microsoft.Graph" -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1 }
    $AuthModule = Get-InstalledModule Microsoft.Graph.Authentication -ErrorAction SilentlyContinue
    if (-not $AuthModule) { $AuthModule = Get-InstalledPSResource -Name "Microsoft.Graph.Authentication" -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1 }
}
if ($script:InstallBeta) {
    $GraphBetaModule = Get-InstalledModule Microsoft.Graph.Beta -ErrorAction SilentlyContinue
    if (-not $GraphBetaModule) { $GraphBetaModule = Get-InstalledPSResource -Name "Microsoft.Graph.Beta" -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1 }
}

# Stop the stopwatch
$script:Stopwatch.Stop()
$ElapsedTime = $script:Stopwatch.Elapsed
$TimeFormatted = "{0:D2}:{1:D2}:{2:D2}" -f $ElapsedTime.Hours, $ElapsedTime.Minutes, $ElapsedTime.Seconds

Write-Host ""
Write-Host "  INSTALLED MODULES:" -ForegroundColor White
Write-Host "  -------------------------------------------" -ForegroundColor Gray

# Display Graph modules if installed
if ($script:InstallGraph) {
    if ($GraphModule) {
        Write-Host "  Microsoft.Graph:                v$($GraphModule.Version)" -ForegroundColor Green
    }
    else {
        Write-Host "  Microsoft.Graph:                NOT INSTALLED" -ForegroundColor Red
    }
    
    if ($AuthModule) {
        Write-Host "  Microsoft.Graph.Authentication: v$($AuthModule.Version)" -ForegroundColor Green
    }
    else {
        Write-Host "  Microsoft.Graph.Authentication: NOT INSTALLED" -ForegroundColor Red
    }
}

# Display Beta modules if installed
if ($script:InstallBeta) {
    if ($GraphBetaModule) {
        Write-Host "  Microsoft.Graph.Beta:           v$($GraphBetaModule.Version)" -ForegroundColor Green
    }
    else {
        Write-Host "  Microsoft.Graph.Beta:           NOT INSTALLED" -ForegroundColor Red
    }
}

Write-Host "  -------------------------------------------" -ForegroundColor Gray

# Final status check
$AllSuccess = $true
if ($script:InstallGraph -and (-not $GraphModule)) { $AllSuccess = $false }
if ($script:InstallBeta -and (-not $GraphBetaModule)) { $AllSuccess = $false }

# Version match check for Graph modules only (if both installed)
$VersionMismatch = $false
if ($script:InstallGraph -and $script:InstallBeta -and $GraphModule -and $GraphBetaModule) {
    if ($GraphModule.Version -ne $GraphBetaModule.Version) {
        $VersionMismatch = $true
    }
}

if ($AllSuccess -and -not $VersionMismatch) {
    Write-Host ""
    Write-Host "  STATUS: SUCCESS" -ForegroundColor Green
    Write-Host "  All selected modules installed successfully!" -ForegroundColor Green
}
elseif ($AllSuccess -and $VersionMismatch) {
    Write-Host ""
    Write-Host "  STATUS: WARNING" -ForegroundColor Yellow
    Write-Host "  Modules installed but Graph versions do not match." -ForegroundColor Yellow
}
else {
    Write-Host ""
    Write-Host "  STATUS: FAILED" -ForegroundColor Red
    Write-Host "  One or more modules failed to install." -ForegroundColor Red
}

Write-Host ""
Write-Host "  IMPORTANT: Open a NEW PowerShell window before running any Graph commands." -ForegroundColor Yellow
Write-Host ""
Write-Host "  Total time: $TimeFormatted" -ForegroundColor Cyan
Write-Host ""

# Stop the background title updater and restore original title
$script:TitleState.IsRunning = $false
Start-Sleep -Milliseconds 600  # Give runspace time to restore title
$script:TitleRunspace.Close()
$script:TitleRunspace.Dispose()

Read-Host "  Press Enter to exit"
