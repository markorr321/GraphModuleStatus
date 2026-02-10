##############################################################################
# GraphModuleStatus
# Shows the status of Microsoft Graph PowerShell modules on profile load
##############################################################################

# Path to the update script (bundled with module)
$script:UpdateScriptPath = Join-Path $PSScriptRoot "Update-MicrosoftGraph.ps1"

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
        $Installed = Get-InstalledPSResource -Name $Module.Name -ErrorAction SilentlyContinue | 
                     Sort-Object Version -Descending | Select-Object -First 1

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

            # Check for updates
            $Available = Find-PSResource -Name $Module.Name -Repository PSGallery -ErrorAction SilentlyContinue

            if ($Available) {
                $Status.AvailableVersion = $Available.Version.ToString()
                $Status.UpdateAvailable = ($Available.Version -gt $Installed.Version)
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
            # Not installed
            if (-not $Silent) {
                Write-Host "  [$($Module.Display)]" -ForegroundColor DarkGray -NoNewline
                Write-Host " ○ Not installed" -ForegroundColor Red
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
            Write-Host "  Update available! Run " -NoNewline -ForegroundColor Yellow
            Write-Host "Update-GraphModule" -NoNewline -ForegroundColor Cyan
            Write-Host " to update." -ForegroundColor Yellow
            Write-Host ""
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
