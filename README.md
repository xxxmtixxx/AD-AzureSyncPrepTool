# AD-AzureSyncPrepTool: Azure AD Connect Preparation Scripts

**AD-AzureSyncPrepTool** is a robust suite of PowerShell scripts, meticulously crafted to synchronize user details between on-premises Active Directory and Azure Active Directory for Office 365. This toolkit is designed to streamline the process of user data synchronization, ensuring consistency across your infrastructure.

## Key Features

- **User Data Export:** The scripts export user details from both directories into CSV files, providing a structured format for data analysis and processing.
- **User Matching:** The toolkit matches users based on their account names, ensuring accurate synchronization between the two directories.
- **Proxy Addresses Handling:** The scripts are equipped to handle proxy addresses, enabling the synchronization of user aliases.
- **Country Name to ISO Code Mapping:** The tool maps full country names to their respective ISO codes for matched users, providing standardized country information.
- **Unmatched Users Identification:** The toolkit identifies users that exist in one directory but not the other, generating separate CSV files for unmatched users in each directory.

## Scripts

### AD-AzureSyncPrepTool-Export.ps1

This script performs the following tasks:

- Exports user details from on-premises Active Directory into a CSV file (`ad_users.csv`).
- Exports user details from Azure Active Directory for Office 365 into another CSV file (`O365_users.csv`).
- Matches users from both directories based on their account names.
- Maps full country names to their ISO codes for matched users.
- Identifies users that exist in one directory but not the other.
- Exports the details of matched users into a CSV file (`matched_users.csv`).
- Creates separate CSV files for users found only in on-premises Active Directory (`ad_users_not_matched.csv`) and users found only in Azure Active Directory (`o365_users_not_matched.csv`).
- Retrieves the list of domains configured in the Office 365 tenant and exports it to a CSV file (`O365_domains.csv`).

### AD-AzureSyncPrepTool-Import.ps1

This script performs the following tasks:

- Defines the file path to the CSV containing matched user details.
- Imports matched users from the CSV file.
- Iterates through each user in the imported CSV.
- Retrieves the corresponding AD user based on the SamAccountName.
- Updates the AD user's properties if they are not empty in the CSV.
- Skips updating any properties that are empty in the CSV.
- Updates the AD user's proxy addresses if the Aliases field is not empty in the CSV.
- Defines the file path to the CSV containing Office 365 domains.
- Imports domains from the CSV file.
- Prompts for and adds UPN suffixes based on the imported domains.
- Catches and logs any errors that occur during the update process.
- Outputs a message upon successful update of users in Active Directory and addition of UPN suffixes.

## CSV Files

The toolkit generates several CSV files to facilitate the process of user data synchronization:

- **ad_users.csv:** Contains the details of users found in the on-premises Active Directory.
- **O365_users.csv:** Contains the details of users found in Azure Active Directory for Office 365.
- **matched_users.csv:** Contains the details of users found in both directories. This CSV file is used to update the user details in Active Directory, including the proxy addresses.
- **ad_users_not_matched.csv:** Lists the users found only in the on-premises Active Directory.
- **o365_users_not_matched.csv:** Lists the users found only in Azure Active Directory.

Each CSV file is structured with each row representing a user and columns representing user attributes such as SamAccountName, DisplayName, GivenName, Surname, EmailAddress, StreetAddress, City, PostalCode, State, Country, TelephoneNumber, Title, Department, and Aliases.

With **AD-AzureSyncPrepTool**, you can ensure seamless synchronization of user details between your on-premises Active Directory and Azure Active Directory for Office 365, making it an invaluable part of any organization's Azure AD Connect setup.
