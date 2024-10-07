@{
    RootModule = 'AD-AzureSyncPrepTool.psm1'
    ModuleVersion = '1.1.0'
    GUID = '595f1f01-b7da-4c26-bf84-3c9ff909649c'  # Replace with an actual GUID, e.g., (New-Guid).Guid
    Author = 'Pablo Clemons'
    CompanyName = 'WhyCantI'
    Description = 'Module for synchronizing and managing AD and Azure AD user attributes, designed to run prior to federating with Azure AD Connect to preserve cloud attributes.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @(
        'Start-ADAzureSyncPrep'
    )
    CmdletsToExport   = @()  # Specify cmdlets if any
    VariablesToExport = '*'  # Export all variables or specify as needed
    AliasesToExport   = @()  # Specify aliases if any
    FileList = @(
        'AD-AzureSyncPrepTool.psm1',
        'Scripts\AD-AzureSyncPrepTool.ps1',
        'Functions\Common.ps1',
        'Functions\RunAllActiveFunctions.ps1',
        'Functions\Update-City.ps1',
        'Functions\Update-Country.ps1',
        'Functions\Update-Department.ps1',
        'Functions\Update-DisplayName.ps1',
        'Functions\Update-Email.ps1',
        'Functions\Update-GivenName.ps1',
        'Functions\Update-JobTitle.ps1',
        'Functions\Update-Manager.ps1',
        'Functions\Update-MobileNumber.ps1',
        'Functions\Update-Office.ps1',
        'Functions\Update-PhoneNumber.ps1',
        'Functions\Update-PostalCode.ps1',
        'Functions\Update-ProxyAddresses.ps1',
        'Functions\Update-State.ps1',
        'Functions\Update-StreetAddress.ps1',
        'Functions\Update-Surname.ps1',
        'Functions\Update-UserLogonName.ps1',
        'GUI\Initialize-GUI.ps1',
        'GUI\Setup-Controls.ps1',
        'GUI\Update-GUI.ps1',
        'GUI\syncStatus.txt'
    )
    RequiredModules = @(
        'ActiveDirectory',
        'AzureAD',
        'ExchangeOnlineManagement'
    )
    PrivateData = @{
        PSData = @{
            ReleaseNotes = @'
Version 1.1.0:
- Added ability to list users alphabetically and display Azure licenses in the GUI.
- Intelligent filtering includes showing disabled and unlicensed users.
- Sync status color codes: orange for cloud and on-prem, green for federated, black for single environment.
- Bulk updates via CSV with import/export features for mass attribute migration.
- Buttons disable when attributes match and turn green or gray depending on their state.
- "Run Sync" button turns orange when updates occur, indicating a new sync is required.
- Future enhancements: timeout after Azure updates and ability to open previous CSVs for viewing/restoration.
- CSVs are renamed to maintain historical logs.
'@
        }
    }
}