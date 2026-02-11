<#
.SYNOPSIS
    Performs a clean installation of Microsoft Graph and/or Microsoft Entra PowerShell modules.

.DESCRIPTION
    This script provides a comprehensive solution for resolving Microsoft Graph and Entra
    PowerShell module version conflicts and ensuring a clean, consistent installation. It is
    designed to address common issues that occur when multiple versions of modules are installed,
    which can cause assembly loading errors and command failures.

    The script performs the following operations:

    1. SESSION CLEANUP
       Removes all currently loaded Microsoft Graph/Entra modules from the PowerShell session
       to prevent file locking issues during uninstallation.

    2. MODULE UNINSTALLATION (Iterative)
       Systematically uninstalls all installed modules using an iterative approach with
       garbage collection to handle dependencies and file locks. Uses both Get-InstalledModule
       and Get-Module -ListAvailable for comprehensive detection.

    3. FOLDER CLEANUP
       Scans common PowerShell module directories for any leftover module folders
       that may have been orphaned, and removes them to ensure a clean slate.

    4. FRESH INSTALLATION
       Installs your choice of Microsoft.Graph, Microsoft.Graph.Beta, and/or Microsoft.Entra
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

# Progress tracking function
function Write-Progress-Step {
    param (
        [int]$Step,
        [int]$TotalSteps,
        [string]$StepName,
        [string]$Status = "In Progress"
    )

    # Update window title with current step
    Update-WindowTitle -CurrentStep "Step $Step of $TotalSteps : $StepName"

    $PercentComplete = [math]::Round(($Step / $TotalSteps) * 100)
    $ProgressBar = ""
    $BarLength = 30
    $FilledLength = [math]::Round(($PercentComplete / 100) * $BarLength)

    for ($i = 0; $i -lt $BarLength; $i++) {
        if ($i -lt $FilledLength) {
            $ProgressBar += [char]0x2588  # Full block
        }
        else {
            $ProgressBar += [char]0x2591  # Light shade
        }
    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  $ProgressBar $PercentComplete%" -ForegroundColor Cyan
    Write-Host "  Step $Step of $TotalSteps : $StepName" -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
}

# Sub-progress function for operations within a step
function Write-SubProgress {
    param (
        [int]$Current,
        [int]$Total,
        [string]$ItemName,
        [string]$Status = "Processing"
    )

    # Note: Title updates removed to prevent bouncing - step-level only shows in title bar

    $PercentComplete = [math]::Round(($Current / $Total) * 100)
    $ProgressBar = ""
    $BarLength = 20
    $FilledLength = [math]::Round(($PercentComplete / 100) * $BarLength)

    for ($i = 0; $i -lt $BarLength; $i++) {
        if ($i -lt $FilledLength) {
            $ProgressBar += [char]0x2588
        }
        else {
            $ProgressBar += [char]0x2591
        }
    }

    Write-Host "  [$ProgressBar] $PercentComplete% - $Status : $ItemName" -ForegroundColor Gray
}

$TotalSteps = 6

Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "   Microsoft Graph & Entra Module Updater       " -ForegroundColor Cyan
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
$GraphModules = @()
$GraphModules += Get-InstalledModule Microsoft.Graph* -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike "Microsoft.Graph.Beta*" }
$GraphPathModules = Get-Module -ListAvailable Microsoft.Graph* -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike "Microsoft.Graph.Beta*" }
$HasGraph = ($GraphModules.Count -gt 0) -or ($GraphPathModules.Count -gt 0)
$GraphVer = if ($GraphModules.Count -gt 0) { "v$($GraphModules[0].Version)" } elseif ($GraphPathModules.Count -gt 0) { "v$($GraphPathModules[0].Version) (path only)" } else { "" }

# Check for any Microsoft.Graph.Beta modules
$BetaModules = @()
$BetaModules += Get-InstalledModule Microsoft.Graph.Beta* -ErrorAction SilentlyContinue
$BetaPathModules = Get-Module -ListAvailable Microsoft.Graph.Beta* -ErrorAction SilentlyContinue
$HasBeta = ($BetaModules.Count -gt 0) -or ($BetaPathModules.Count -gt 0)
$BetaVer = if ($BetaModules.Count -gt 0) { "v$($BetaModules[0].Version)" } elseif ($BetaPathModules.Count -gt 0) { "v$($BetaPathModules[0].Version) (path only)" } else { "" }

# Check for any Microsoft.Entra modules
$EntraModules = @()
$EntraModules += Get-InstalledModule Microsoft.Entra* -ErrorAction SilentlyContinue
$EntraPathModules = Get-Module -ListAvailable Microsoft.Entra* -ErrorAction SilentlyContinue
$HasEntra = ($EntraModules.Count -gt 0) -or ($EntraPathModules.Count -gt 0)
$EntraVer = if ($EntraModules.Count -gt 0) { "v$($EntraModules[0].Version)" } elseif ($EntraPathModules.Count -gt 0) { "v$($EntraPathModules[0].Version) (path only)" } else { "" }

Write-Host ""
Write-Host "  Discovered modules:" -ForegroundColor White
if ($HasGraph) {
    Write-Host "    - Microsoft.Graph (stable)    $GraphVer" -ForegroundColor Green
}
if ($HasBeta) {
    Write-Host "    - Microsoft.Graph.Beta        $BetaVer" -ForegroundColor Green
}
if ($HasEntra) {
    Write-Host "    - Microsoft.Entra             $EntraVer" -ForegroundColor Green
}
if (-not $HasGraph -and -not $HasBeta -and -not $HasEntra) {
    Write-Host "    (none found)" -ForegroundColor Yellow
}
Write-Host ""

# Prompt user for which modules to manage
Write-Host "  Which modules would you like to uninstall?" -ForegroundColor Cyan
Write-Host ""
Write-Host "    [1] Microsoft.Graph (stable) only" -ForegroundColor White
Write-Host "    [2] Microsoft.Graph.Beta only" -ForegroundColor White
Write-Host "    [3] Both Microsoft.Graph and Microsoft.Graph.Beta" -ForegroundColor White
Write-Host "    [4] Microsoft.Entra only" -ForegroundColor White
Write-Host "    [5] All (Graph, Graph.Beta, and Entra)" -ForegroundColor White
Write-Host ""
$ModuleChoice = Read-Host "  Enter your choice (1-5) [default: 3]"

if ([string]::IsNullOrWhiteSpace($ModuleChoice)) {
    $ModuleChoice = "3"
}

# Set flags based on choice
$script:IncludeGraph = $false
$script:IncludeBeta = $false
$script:IncludeEntra = $false

switch ($ModuleChoice) {
    "1" { $script:IncludeGraph = $true }
    "2" { $script:IncludeBeta = $true }
    "3" { $script:IncludeGraph = $true; $script:IncludeBeta = $true }
    "4" { $script:IncludeEntra = $true }
    "5" { $script:IncludeGraph = $true; $script:IncludeBeta = $true; $script:IncludeEntra = $true }
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
if ($script:IncludeEntra) {
    $script:ModulePatterns += "Microsoft.Entra*"
}

Write-Host ""
Write-Host "  Selected for uninstall:" -ForegroundColor Gray
if ($script:IncludeGraph) { Write-Host "    - Microsoft.Graph (stable)" -ForegroundColor Yellow }
if ($script:IncludeBeta) { Write-Host "    - Microsoft.Graph.Beta" -ForegroundColor Yellow }
if ($script:IncludeEntra) { Write-Host "    - Microsoft.Entra" -ForegroundColor Yellow }
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

if ($LoadedModules) {
    $LoadedTotal = @($LoadedModules).Count
    $LoadedCounter = 0
    $LoadedSuccess = 0
    $LoadedFailed = 0

    Write-Host "  Found $LoadedTotal modules loaded in session..." -ForegroundColor Gray
    Write-Host ""

    foreach ($Module in $LoadedModules) {
        $LoadedCounter++
        Write-SubProgress -Current $LoadedCounter -Total $LoadedTotal -ItemName $Module -Status "Removing"
        try {
            Remove-Module -Name $Module -Force -ErrorAction Stop
            Write-Host "    Removed: $Module" -ForegroundColor Green
            $LoadedSuccess++
        }
        catch {
            Write-Host "    Failed: $Module" -ForegroundColor Yellow
            $LoadedFailed++
        }
    }
    Write-Host ""
    Write-Host "  Clear complete: $LoadedSuccess succeeded, $LoadedFailed failed." -ForegroundColor Green
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
    Write-Host "  === Iteration $Iteration ===" -ForegroundColor Blue
    Write-Host ""

    # Get installed modules using Get-InstalledModule (gallery-installed)
    $InstalledModules = @()
    foreach ($Pattern in $script:ModulePatterns) {
        $InstalledModules += Get-InstalledModule $Pattern -ErrorAction SilentlyContinue
    }
    $InstalledModules = $InstalledModules | Select-Object -Unique -Property Name, Version, @{N='ModuleBase';E={$_.InstalledLocation}}

    # Get available modules using Get-Module -ListAvailable (catches modules not in gallery)
    $AvailableModules = @()
    foreach ($Pattern in $script:ModulePatterns) {
        $AvailableModules += Get-Module -ListAvailable $Pattern -ErrorAction SilentlyContinue
    }
    $AvailableModules = $AvailableModules | Select-Object -Unique -Property Name, Version, ModuleBase

    $TotalFound = @($InstalledModules).Count + @($AvailableModules).Count

    if ($TotalFound -eq 0) {
        Write-Host "  No modules found. Cleanup complete!" -ForegroundColor Green
        break
    }

    Write-Host "  Found $(@($InstalledModules).Count) installed modules (PowerShell Gallery)" -ForegroundColor White
    Write-Host "  Found $(@($AvailableModules).Count) modules in module paths" -ForegroundColor White
    Write-Host ""

    # First, uninstall using Uninstall-Module for gallery-installed modules
    if ($InstalledModules) {
        Write-Host "  Uninstalling gallery-installed modules..." -ForegroundColor Yellow
        $Counter = 0
        $TotalModules = @($InstalledModules).Count

        foreach ($Module in $InstalledModules) {
            $Counter++
            Write-SubProgress -Current $Counter -Total $TotalModules -ItemName "$($Module.Name) v$($Module.Version)" -Status "Uninstalling"
            try {
                Uninstall-Module -Name $Module.Name -RequiredVersion $Module.Version -Force -ErrorAction Stop
                Write-Host "    Uninstalled: $($Module.Name) v$($Module.Version)" -ForegroundColor Green
            }
            catch {
                # Try uninstalling all versions
                try {
                    Uninstall-Module -Name $Module.Name -AllVersions -Force -ErrorAction Stop
                    Write-Host "    Uninstalled: $($Module.Name) (all versions)" -ForegroundColor Green
                }
                catch {
                    Write-Host "    Pending cleanup: $($Module.Name)" -ForegroundColor Yellow
                }
            }
        }
        Write-Host ""
    }

    # Handle modules found via Get-Module -ListAvailable that aren't in the gallery
    # These need to be deleted directly by removing their folders
    if ($AvailableModules) {
        # Filter to only modules not already handled by Uninstall-Module
        $OrphanModules = $AvailableModules | Where-Object {
            $ModName = $_.Name
            -not ($InstalledModules | Where-Object { $_.Name -eq $ModName })
        }

        if ($OrphanModules) {
            Write-Host "  Removing orphaned module folders..." -ForegroundColor Yellow
            $Counter = 0
            $TotalOrphans = @($OrphanModules).Count

            foreach ($Module in $OrphanModules) {
                $Counter++
                Write-SubProgress -Current $Counter -Total $TotalOrphans -ItemName "$($Module.Name) v$($Module.Version)" -Status "Deleting"
                try {
                    if (Test-Path $Module.ModuleBase) {
                        Remove-Item -Path $Module.ModuleBase -Recurse -Force -ErrorAction Stop
                        Write-Host "    Deleted: $($Module.ModuleBase)" -ForegroundColor Green
                    }
                }
                catch {
                    Write-Host "    Failed to delete: $($Module.ModuleBase) - $_" -ForegroundColor Yellow
                }
            }
            Write-Host ""
        }
    }

    # Force garbage collection to release any file locks
    Write-Host "  Releasing file locks..." -ForegroundColor Gray
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
if ($script:IncludeEntra) {
    $FolderPatterns += "Microsoft.Entra*"
}

$TotalPaths = $ModulePaths.Count
$PathCounter = 0
$LeftoverCount = 0
$RemovedCount = 0

Write-Host "  Scanning $TotalPaths module locations..." -ForegroundColor Gray
Write-Host ""

foreach ($Path in $ModulePaths) {
    $PathCounter++
    Write-SubProgress -Current $PathCounter -Total $TotalPaths -ItemName (Split-Path $Path -Leaf) -Status "Scanning"

    # Find all matching folders based on user selection
    $TargetFolders = @()
    foreach ($Pattern in $FolderPatterns) {
        $TargetFolders += Get-ChildItem -Path $Path -Directory -Filter $Pattern -ErrorAction SilentlyContinue
    }

    foreach ($Folder in $TargetFolders) {
        Write-Host "    Found: $($Folder.FullName)" -ForegroundColor Yellow

        # Delete everything inside the folder
        $FolderContents = Get-ChildItem -Path $Folder.FullName -Force -ErrorAction SilentlyContinue
        foreach ($Item in $FolderContents) {
            try {
                Remove-Item -Path $Item.FullName -Recurse -Force -ErrorAction Stop
                Write-Host "      Deleted: $($Item.Name)" -ForegroundColor Green
                $RemovedCount++
            }
            catch {
                Write-Host "      Could not delete: $($Item.Name) (may need admin rights or file is locked)" -ForegroundColor Red
                $LeftoverCount++
            }
        }
    }
}

# Second pass: Verify cleanup was complete
Write-Host ""
Write-Host "  Performing verification scan..." -ForegroundColor Gray

$RemainingItems = @()
foreach ($Path in $ModulePaths) {
    foreach ($Pattern in $FolderPatterns) {
        $TargetFolders = Get-ChildItem -Path $Path -Directory -Filter $Pattern -ErrorAction SilentlyContinue
        foreach ($Folder in $TargetFolders) {
            $Contents = Get-ChildItem -Path $Folder.FullName -Force -ErrorAction SilentlyContinue
            if ($Contents) {
                foreach ($Item in $Contents) {
                    $RemainingItems += $Item.FullName
                }
            }
        }
    }
}

Write-Host ""
if ($RemainingItems.Count -eq 0) {
    Write-Host "  Folder cleanup complete. Deleted $RemovedCount items." -ForegroundColor Green
    Write-Host "  Verification: All target module folders are now empty." -ForegroundColor Green
}
else {
    Write-Host "  Warning: $($RemainingItems.Count) items could not be deleted:" -ForegroundColor Yellow
    foreach ($Remaining in $RemainingItems) {
        Write-Host "    - $Remaining" -ForegroundColor Yellow
    }
    Write-Host ""
    Write-Host "  Try running PowerShell as Administrator, or manually delete these items." -ForegroundColor Yellow
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
Write-Host "    [4] Microsoft.Entra only" -ForegroundColor White
Write-Host "    [5] All (Graph, Graph.Beta, and Entra)" -ForegroundColor White
Write-Host "    [0] Skip installation" -ForegroundColor White
Write-Host ""
$InstallChoice = Read-Host "  Enter your choice (0-5) [default: 3]"

if ([string]::IsNullOrWhiteSpace($InstallChoice)) {
    $InstallChoice = "3"
}

# Set install flags based on choice (script scope for Steps 5 & 6)
$script:InstallGraph = $false
$script:InstallBeta = $false
$script:InstallEntra = $false

switch ($InstallChoice) {
    "0" { 
        Write-Host "  Skipping installation." -ForegroundColor Yellow
    }
    "1" { $script:InstallGraph = $true }
    "2" { $script:InstallBeta = $true }
    "3" { $script:InstallGraph = $true; $script:InstallBeta = $true }
    "4" { $script:InstallEntra = $true }
    "5" { $script:InstallGraph = $true; $script:InstallBeta = $true; $script:InstallEntra = $true }
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
    Write-Host "  Installing selected packages..." -ForegroundColor Gray

    $InstallSuccess = 0
    $InstallFailed = 0

    if ($script:InstallGraph) {
        Write-Host "  Installing Microsoft.Graph..." -ForegroundColor Yellow
        Write-Host "  (This will install all sub-modules - may take several minutes)" -ForegroundColor Gray
        Write-Host ""
        try {
            Install-Module Microsoft.Graph -Scope $script:InstallScope -Force -AllowClobber -ErrorAction Stop
            Write-Host ""
            Write-Host "  Microsoft.Graph installed successfully." -ForegroundColor Green
            $InstallSuccess++
        }
        catch {
            Write-Host "  ERROR: Failed to install Microsoft.Graph - $_" -ForegroundColor Red
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
            Write-Host ""
            Write-Host "  Microsoft.Graph.Beta installed successfully." -ForegroundColor Green
            $InstallSuccess++
        }
        catch {
            Write-Host "  ERROR: Failed to install Microsoft.Graph.Beta - $_" -ForegroundColor Red
            $InstallFailed++
        }
        Write-Host ""
    }

    if ($script:InstallEntra) {
        Write-Host "  Installing Microsoft.Entra..." -ForegroundColor Yellow
        Write-Host "  (This will install all sub-modules - may take several minutes)" -ForegroundColor Gray
        Write-Host ""
        try {
            Install-Module Microsoft.Entra -Scope $script:InstallScope -Force -AllowClobber -ErrorAction Stop
            Write-Host ""
            Write-Host "  Microsoft.Entra installed successfully." -ForegroundColor Green
            $InstallSuccess++
        }
        catch {
            Write-Host "  ERROR: Failed to install Microsoft.Entra - $_" -ForegroundColor Red
            $InstallFailed++
        }
        Write-Host ""
    }

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
$EntraModule = $null

if ($script:InstallGraph) {
    $GraphModule = Get-InstalledModule Microsoft.Graph -ErrorAction SilentlyContinue
    $AuthModule = Get-InstalledModule Microsoft.Graph.Authentication -ErrorAction SilentlyContinue
}
if ($script:InstallBeta) {
    $GraphBetaModule = Get-InstalledModule Microsoft.Graph.Beta -ErrorAction SilentlyContinue
}
if ($script:InstallEntra) {
    $EntraModule = Get-InstalledModule Microsoft.Entra -ErrorAction SilentlyContinue
}

# Stop the stopwatch
$script:Stopwatch.Stop()
$ElapsedTime = $script:Stopwatch.Elapsed
$TimeFormatted = "{0:D2}:{1:D2}:{2:D2}" -f $ElapsedTime.Hours, $ElapsedTime.Minutes, $ElapsedTime.Seconds

Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "            INSTALLATION COMPLETE               " -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# Final progress bar at 100%
$FinalBar = [string]::new([char]0x2588, 30)
Write-Host "  $FinalBar 100%" -ForegroundColor Green
Write-Host ""
Write-Host "  TOTAL TIME: $TimeFormatted" -ForegroundColor Cyan
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

# Display Entra modules if installed
if ($script:InstallEntra) {
    if ($EntraModule) {
        Write-Host "  Microsoft.Entra:                v$($EntraModule.Version)" -ForegroundColor Green
    }
    else {
        Write-Host "  Microsoft.Entra:                NOT INSTALLED" -ForegroundColor Red
    }
}

Write-Host "  -------------------------------------------" -ForegroundColor Gray

# Final status check
$AllSuccess = $true
if ($script:InstallGraph -and (-not $GraphModule)) { $AllSuccess = $false }
if ($script:InstallBeta -and (-not $GraphBetaModule)) { $AllSuccess = $false }
if ($script:InstallEntra -and (-not $EntraModule)) { $AllSuccess = $false }

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
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  IMPORTANT: Open a NEW PowerShell window      " -ForegroundColor Yellow
Write-Host "  before running any Graph commands.           " -ForegroundColor Yellow
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# Stop the background title updater and restore original title
$script:TitleState.IsRunning = $false
Start-Sleep -Milliseconds 600  # Give runspace time to restore title
$script:TitleRunspace.Close()
$script:TitleRunspace.Dispose()
