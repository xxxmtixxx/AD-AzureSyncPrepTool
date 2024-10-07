![image](https://github.com/user-attachments/assets/8170141c-eaa2-4946-b0b5-6b74fb63111c)

# AD-AzureSyncPrepTool

## Overview

AD-AzureSyncPrepTool is a comprehensive PowerShell-based tool designed for managing and synchronizing user attributes between on-premises Active Directory (AD) and Azure AD environments. This tool is particularly valuable when preparing to federate with Azure AD Connect, ensuring that cloud attributes are not overwritten or wiped out by on-prem attributes during synchronization. The tool provides a flexible and intuitive interface for administrators to compare, update, and synchronize user data across both environments, ensuring consistency and compliance.

### Key Features

1. **User Management and Comparison:**
   - **Alphabetical User Listing:** Lists users in alphabetical order for easy navigation and management.
   - **Licensing Information:** Pulls licensing data from Azure AD and displays it directly in the GUI.
   - **Sync Status Visualization:**
     - **Orange:** Indicates accounts exist in both environments but are not federated.
     - **Green:** Indicates federated accounts.
     - **Black:** Indicates accounts exist only in one environment.
   - **Attribute Matching Visualization:**
     - Both sides of the attributes turn **green** if matched, **red** if they do not match, and **black** if N/A.
   - **Intelligent Filtering Options:**
     - **Show Disabled AD Users:** Displays users disabled in Active Directory.
     - **Show Disabled Azure AD Users:** Displays users disabled in Azure AD.
     - **Show Unlicensed Users:** Displays users without licenses.
     - **Show Only Matched Users:** Displays only accounts that exist in both AD and Azure environments.

2. **Attribute Management and Synchronization:**
   - **Selective Synchronization:** Allows migration of individual attributes per user, environment-specific synchronization, or bulk updates via CSV.
   - **Disable Matched Buttons:** Buttons are automatically disabled if the attributes are already matched, turning **green** when enabled and **gray** when not, preventing unnecessary actions.
   - **Run All Buttons:** Blue "Run All" buttons only activate if individual attribute buttons are enabled. They execute updates only for enabled buttons.
   - **Run Sync Status:** When an attribute is updated, the "Run Sync" button turns **orange**, indicating that a new sync is required to populate updated data from AD and Azure.

3. **CSV Integration for Bulk Updates:**
   - **Export CSV:** Export a complete template CSV with all currently populated attributes from both environments.
   - **Import CSV:** Re-import the modified CSV, and the tool intelligently checks for changes and applies updates to both AD and Azure environments.
   - **Seamless Bulk Operations:** Whether migrating a single attribute, all attributes for a user between environments, or performing bulk updates via CSV, the tool handles it all smoothly.

4. **Logging and Error Handling:**
   - **Detailed Logging:** Logs every action and validation step to a log file, aiding in troubleshooting and record-keeping. Logs are stored in `C:\Temp\AD-AzureSyncPrepTool`.
   - **Historical CSV Renaming:** CSVs are renamed with a timestamp for historical purposes, ensuring past data is preserved and easily retrievable.
   - **Error Handling:** Plans to add additional catch blocks and message blocks to further enhance the robustness of the tool.

5. **GUI Enhancements:**
   - **Modern Interface:** Clean and user-friendly interface with clear indicators for attribute matching and synchronization status.
   - **Message Blocks and Feedback:** Planned improvements include adding more feedback messages and improving the user experience.

### Installation and Setup

1. **Run the Tool Manually:**
   - **Download and Extract:**
     - Download the zip file from the [GitHub repository](https://github.com/xxxmtixxx/AD-AzureSyncPrepTool/archive/refs/heads/main.zip) and extract it to your desired location.
   - **Start the GUI:**
     - Navigate to the `GUI` folder, right-click on `Initialize-GUI.ps1`, and select **Run with PowerShell**. The script will automatically ensure it runs with administrative privileges.

2. **Use the One-Liner for Quick Installation:**
   - Run the following command in PowerShell to download and set up the module automatically:

     ```powershell
     $documentsPath=[Environment]::GetFolderPath('MyDocuments');$url='https://github.com/xxxmtixxx/AD-AzureSyncPrepTool/archive/refs/heads/main.zip';$moduleName='AD-AzureSyncPrepTool';$modulePath=Join-Path $documentsPath 'WindowsPowerShell\Modules';$tempPath=Join-Path $env:TEMP ($moduleName+'.zip');Invoke-WebRequest -Uri $url -OutFile $tempPath;$tempDir='.'+$moduleName+'_temp';$extractPath=Join-Path $HOME $tempDir;Expand-Archive -Path $tempPath -DestinationPath $extractPath -Force;$sourceFolder=Join-Path $extractPath 'AD-AzureSyncPrepTool-main';$destinationFolder=Join-Path $modulePath $moduleName;if (!(Test-Path $destinationFolder)) {New-Item -Path $destinationFolder -ItemType Directory | Out-Null};Copy-Item -Path "$sourceFolder\*" -Destination $destinationFolder -Recurse -Force;Remove-Item $tempPath, $extractPath -Recurse -Force
     ```

   - **Import the Module:**
     ```powershell
     Import-Module AD-AzureSyncPrepTool
     ```
   - **Start the GUI:**
     ```powershell
     Start-ADAzureSyncPrep
     ```

### File Structure

- **`Functions` Folder:** Contains individual PowerShell scripts for updating specific attributes like `Update-GivenName.ps1`, `Update-Manager.ps1`, etc.
- **`GUI` Folder:** Contains GUI-related scripts like `Initialize-GUI.ps1`, `Setup-Controls.ps1`, and `Update-GUI.ps1` that define the layout and functionality of the dashboard.
- **`Scripts` Folder:** Houses the main script, `AD-AzureSyncPrepTool.ps1`, which coordinates the overall synchronization process.
- **`AD-AzureSyncPrepTool.psd1` and `AD-AzureSyncPrepTool.psm1`:** Module files that define the tool’s configuration and the primary module script.

### Modules Used

- `ActiveDirectory`
- `AzureAD`
- `ExchangeOnline`

### Future Enhancements

- Additional error handling and feedback improvements.
- Reorganizing code for consistency.
- The ability to open previously written CSVs to view and restore attributes.
- Implementing a timeout after updating an Azure attribute before you can run "Run Sync" again, to account for Azure propagation delays and ensure accurate data retrieval.

### Conclusion

The AD-AzureSyncPrepTool is a powerful solution for IT administrators needing precise control over user synchronization between AD and Azure AD. Whether managing individual users, performing bulk updates via CSV, or ensuring that cloud attributes are preserved before federating with Azure AD Connect, the tool's intuitive interface and robust feature set make it an invaluable asset for maintaining data integrity across environments.
