# Import the Active Directory module for on-prem AD users
Import-Module ActiveDirectory

# Define the path to the CSV file containing the matched users
$matchedCsvFilePath = "C:\temp\matched_users.csv"

# Import the matched users from the CSV file
$matchedUsers = Import-Csv -Path $matchedCsvFilePath

# Loop through each user in the CSV file
foreach ($user in $matchedUsers) {
    $adUser = Get-ADUser -Filter "SamAccountName -eq '$($user.SamAccountName)'"

    if ($adUser) {
        try {
            # Create a hashtable to hold the parameters for Set-ADUser
            $params = @{
                Identity = $adUser.DistinguishedName
                ErrorAction = 'Stop'
            }

            # Add parameters conditionally based on whether the field has a value
            if (![string]::IsNullOrWhiteSpace($user.DisplayName)) {
                $params['DisplayName'] = $user.DisplayName
            }
            if (![string]::IsNullOrWhiteSpace($user.GivenName)) {
                $params['GivenName'] = $user.GivenName
            }
            if (![string]::IsNullOrWhiteSpace($user.Surname)) {
                $params['Surname'] = $user.Surname
            }
            if (![string]::IsNullOrWhiteSpace($user.EmailAddress)) {
                $params['EmailAddress'] = $user.EmailAddress
            }
            if (![string]::IsNullOrWhiteSpace($user.StreetAddress)) {
                $params['StreetAddress'] = $user.StreetAddress
            }
            if (![string]::IsNullOrWhiteSpace($user.City)) {
                $params['City'] = $user.City
            }
            if (![string]::IsNullOrWhiteSpace($user.PostalCode)) {
                $params['PostalCode'] = $user.PostalCode
            }
            if (![string]::IsNullOrWhiteSpace($user.State)) {
                $params['State'] = $user.State
            }
            if (![string]::IsNullOrWhiteSpace($user.Country)) {
                $params['Country'] = $user.Country
            }
            if (![string]::IsNullOrWhiteSpace($user.TelephoneNumber)) {
                $params['OfficePhone'] = $user.TelephoneNumber
            }
            if (![string]::IsNullOrWhiteSpace($user.Title)) {
                $params['Title'] = $user.Title
            }
            if (![string]::IsNullOrWhiteSpace($user.Department)) {
                $params['Department'] = $user.Department
            }
            if (![string]::IsNullOrWhiteSpace($user.Aliases)) {
                # Split the Aliases string back into an array
                $proxyAddresses = $user.Aliases -split ';'
                # Use the Replace parameter to update the ProxyAddresses attribute
                $params['Replace'] = @{ProxyAddresses = $proxyAddresses}
            }

            # Update the user in Active Directory with the specified parameters
            Set-ADUser @params
        } catch {
            Write-Warning "Failed to update user $($user.SamAccountName): $_"
        }
    } else {
        Write-Warning "User $($user.SamAccountName) not found in AD."
    }
}

Write-Host "Active Directory users have been updated with the matched details."
Write-Host ""

# Define the path to the CSV file containing the Office 365 domains
$domainsCsvFilePath = "C:\temp\O365_domains.csv"

# Import the domains from the CSV file
$domains = Import-Csv -Path $domainsCsvFilePath

# Output the list of domains and ask which ones to add as UPN suffixes
Write-Host "The following domains are available to add as UPN suffixes:"
$domains.Name

$domainsToAdd = Read-Host -Prompt "Enter the domains to add as UPN suffixes (separate multiple domains with commas)"

# Split the input into an array of domains
$domainsToAdd = $domainsToAdd.Split(',')

# Get the current forest
$forest = Get-ADForest

# Get the current list of UPN suffixes
$currentUPNSuffixes = $forest.UPNSuffixes

foreach ($domain in $domainsToAdd) {
    # Trim whitespace
    $domain = $domain.Trim()

    if ($domains.Name -contains $domain -and -not $currentUPNSuffixes -contains $domain) {
        # Add the domain as a UPN suffix
        try {
            Set-ADForest -Identity $forest -UPNSuffixes @{Add=$domain}
            Write-Host "Successfully added UPN suffix: $domain"
        } catch {
            Write-Error "An error occurred while adding UPN suffix: $domain. Error: $_"
        }
    } elseif ($currentUPNSuffixes -contains $domain) {
        Write-Host "UPN suffix $domain is already present."
    } else {
        Write-Warning "Domain $domain is not recognized. Skipping..."
    }
}

Write-Host "UPN suffixes have been updated."