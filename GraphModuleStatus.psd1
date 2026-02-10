# Module manifest for module 'GraphModuleStatus'
# Generated on: 2/10/2026

@{

# Script module or binary module file associated with this manifest.
RootModule = 'GraphModuleStatus.psm1'

# Version number of this module.
ModuleVersion = '1.0.0'

# Supported PSEditions
CompatiblePSEditions = @('Core', 'Desktop')

# ID used to uniquely identify this module
GUID = '03dafafc-8e8b-407d-bf2e-1861df94c5a5'

# Author of this module
Author = 'Mark Orr'

# Company or vendor of this module
CompanyName = ''

# Copyright statement for this module
Copyright = '(c) 2026. All rights reserved.'

# Description of the functionality provided by this module
Description = 'Check the status of Microsoft Graph PowerShell modules and update them when needed'

# Minimum version of the PowerShell engine required by this module
PowerShellVersion = '5.1'

# Modules that must be imported into the global environment prior to importing this module
RequiredModules = @(
    @{ModuleName = 'Microsoft.PowerShell.PSResourceGet'; GUID = 'e4e0bda1-0703-44a5-b70d-8fe704cd0643'; ModuleVersion = '1.0.0'}
)

# Functions to export from this module
FunctionsToExport = @('Get-GraphModuleStatus', 'Update-GraphModule', 'Add-GraphModuleStatusToProfile', 'Remove-GraphModuleStatusFromProfile')

# Cmdlets to export from this module
CmdletsToExport = @()

# Variables to export from this module
VariablesToExport = @()

# Aliases to export from this module
AliasesToExport = @()

# List of all files packaged with this module
FileList = @('GraphModuleStatus.psm1', 'GraphModuleStatus.psd1', 'Update-MicrosoftGraph.ps1', 'README.md', 'LICENSE')

# Private data to pass to the module specified in RootModule/ModuleToProcess
PrivateData = @{
    PSData = @{
        Tags = @('Microsoft365', 'Graph', 'MicrosoftGraph', 'Admin', 'Module', 'Update')
        LicenseUri = ''
        ProjectUri = ''
        ReleaseNotes = 'Initial release - Get-GraphModuleStatus and Update-GraphModule functions'
    }
}

}
