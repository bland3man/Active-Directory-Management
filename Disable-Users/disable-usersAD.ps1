<# Created by Bland D. Wallace III
   bdw3@live.com
   This should disable users in Active Directory.  It will remove groups, properties, disable the user or users, move and modify ACL on the user's home drive (if applicable),
   and finally move the user to the OU you need them to be moved to.  Some administrators will keep these disabled users in a separate OU in case a user comes back.  This
   script will record the user's data in a recorded-disabledUsers.xlsx file.  This will be in the case that you want to restore the user at a later time.
#>

# Import the ImportExcel module
Import-Module ImportExcel

# Import the Active Directory module
Import-Module ActiveDirectory

# Set the path to the Excel file (.xlsx)
$excelFilePath = "\\path\to\disableUsersList.xlsm (or .xlsx which ever you decide to use)"

# Define the path to the new Excel file to write data to
$newExcelFilePath = "\\path\to\recordedDisableUsers.xlsx"

# Specify the root OU to search within
$RootOU = "OU=<parentOU>,DC=<DC>,DC=<DC>,DC=<DC>"

# Specify the target OU to move the user to
$targetOU = "OU=<locationOfUsersOU>,OU=<subOU>,OU=<parentOU>,DC=<DC>,DC=<DC>,DC=<DC>"

# Initialize an empty array to store user properties
$userPropertiesArray = @()

# Initialize an empty array to store user group memberships
$userGroupMembershipsArray = @()

# Function to remove user from home drive folder and move it to another folder
function Remove-UserFromHomeDrive {
    param (
        [string]$UserName
    )

    $homeDriveRoot = "\\path\to\currentLocationHomeDrive"
    $folderName = $UserName -replace '[^\p{L}\d]', '.'
    $folderPath = Join-Path -Path $homeDriveRoot -ChildPath $folderName

    # Check if the folder exists
    if (Test-Path $folderPath) {
        # Remove the user from the folder's security tab
        $user = Get-ADUser -Filter { SamAccountName -eq $UserName }
        if ($user) {
            $userSID = $user.SID.Value
            $acl = Get-Acl -Path $folderPath
            $accessRules = $acl.GetAccessRules($true, $false, [System.Security.Principal.NTAccount])
            $modifiedAcl = $false

            foreach ($rule in $accessRules) {
                if ($rule.IdentityReference.Value -eq $userSID) {
                    $acl.RemoveAccessRule($rule)
                    $modifiedAcl = $true
                    Write-Host "Removed user '$UserName' from the security tab of the folder."
                }
            }

            if ($modifiedAcl) {
                Set-Acl -Path $folderPath -AclObject $acl
            }
        }

        # Move the user's folder to a different location
        $destinationPath = "\\path\to\destination"
        Move-Item -Path $folderPath -Destination $destinationPath -Force
        Write-Host "Moved folder for user '$UserName' to '$destinationPath'."
    } else {
        Write-Host "Home drive folder not found for user '$UserName'."
    }
}

# Read data from Excel file
try {
    # Read data from the Excel file, specifying column headers
    $excelData = Import-Excel -Path $excelFilePath -ErrorAction Stop
    Write-Host "Successfully read data from the Excel file."
}
catch {
    Write-Host "Error reading data from the Excel file: $_"
    return
}

# Display the retrieved data
if ($excelData.Count -eq 0) {
    Write-Host "No data found in the Excel file after the header."
}
else {
    foreach ($row in $excelData) {
        $firstName = $row.'FirstName'
        $lastName = $row.'LastName'
        $middleInitial = $row.'MiddleInitial'

        # Construct the user's distinguished name
        $userDN = "CN=$lastName $firstName$($middleInitial -replace '\s',''),$RootOU"

        # Search for the user in Active Directory
        $user = Get-ADUser -Filter { GivenName -eq $firstName -and Surname -eq $lastName } -SearchBase $RootOU -ErrorAction SilentlyContinue

        if (-not $user) {
            # If no user found, switch to FirstNameMiddleInitial.LastName format
            $userDN = "CN=$firstName$middleInitial.$lastName,$RootOU"
            $user = Get-ADUser -Filter { GivenName -eq $firstName -and Surname -eq "$middleInitial.$lastName" } -SearchBase $RootOU -ErrorAction SilentlyContinue
        }

        # Additional search for hyphenated name
        if (-not $user -and $lastName -like "*-*") {
            $hyphenatedLastName = $lastName -replace '-', '.'
            $userDN = "CN=$lastName $firstName$($middleInitial -replace '\s',''),$RootOU"
            $user = Get-ADUser -Filter { GivenName -eq $firstName -and Surname -eq $hyphenatedLastName } -SearchBase $RootOU -ErrorAction SilentlyContinue
        }

        # Additional search for apostrophe in name
        if (-not $user -and $lastName -like "*'*") {
            $apostropheLastName = $lastName -replace "'", '.'
            $userDN = "CN=$lastName $firstName$($middleInitial -replace '\s',''),$RootOU"
            $user = Get-ADUser -Filter { GivenName -eq $firstName -and Surname -eq $apostropheLastName } -SearchBase $RootOU -ErrorAction SilentlyContinue
        }

        if ($user) {
			# Retrieve the termination date from the Excel data
            $terminationDate = $row.'Termed/Resigned/Retired Date'
			
			# Update the user's description in Active Directory
            if ($terminationDate) {
                Set-ADUser -Identity $user -Description $terminationDate
                Write-Host "Updated description for $($user.SamAccountName) with termination date: $terminationDate"
            } else {
                Write-Host "Termination date not found for $($user.SamAccountName). Description not updated."
			}
			
            # Retrieve the user's distinguished name
            $userDN = $user.DistinguishedName
			
			# Retrieve the user's logon scriptPath
			$oldLogonScript = (Get-ADUser -Identity $user -Properties ScriptPath).ScriptPath
			
            
            # Retrieve user's group memberships excluding '<FQDN>/Users/Domain Users'
            $userGroups = Get-ADUser $user | Get-ADPrincipalGroupMembership | Where-Object { $_.DistinguishedName -ne 'CN=Domain Users,CN=<CN>,DC=<DC>,DC=<DC>,DC=<DC>' }
            $groupNames = $userGroups.Name

            # Create a hashtable to store user group memberships
            $userGroupMemberships = @{
                'GroupMemberships' = $groupNames
            }

            # Add user group memberships hashtable to the array
            $userGroupMembershipsArray += $userGroupMemberships
            
            # Retrieve all properties of the user Object
            $user = Get-ADUser -Identity $user -Properties *

            # Create a hashtable to store user properties
            $userProperties = @{
                'SamAccountName' = $user.samAccountName
                'DistinguishedName' = $userDN
                'Description' = $user.Description
                'TelephoneNumber' = $user.telephoneNumber
                'Office' = $user.physicalDeliveryOfficeName
                'PostalCode' = $user.postalCode
                'City' = $user.l
                'State' = $user.st
                'StreetAddress' = $user.streetAddress
                'Title' = $user.title
                'LogonScript' = $oldLogonScript
            }

            # Add user properties hashtable to the array
            $userPropertiesArray += $userProperties

            # Rearrange keys in the hashtable to ensure desired order
            $userPropertiesArray = $userPropertiesArray | ForEach-Object {
                [ordered]@{
                    'SamAccountName' = $_.SamAccountName
                    'DistinguishedName' = $_.DistinguishedName
                    'Description' = $_.Description
                    'Office' = $_.Office
                    'TelephoneNumber' = $_.TelephoneNumber
                    'StreetAddress' = $_.StreetAddress
                    'City' = $_.City
                    'State' = $_.State
                    'PostalCode' = $_.PostalCode
                    'Title' = $_.Title
                    'LogonScript' = $_.LogonScript
                }
            }

            # Output the user properties array
			$userPropertiesArray

			# Output the user group memberships array
			$userGroupMembershipsArray

			# Assemble user properties into a hashtable
			$userProperties = @{
				'Name' = $userDN.Split(',')[0] -replace 'CN=', ''  # Extract user's CN from DN
				'SamAccountName' = $user.samAccountName
				'DN' = $userDN
				'Properties' = @(
					"Description: $($user.Description)",
					"TelephoneNumber: $($user.telephoneNumber)",
					"Office: $($user.physicalDeliveryOfficeName)",
					"PostalCode: $($user.postalCode)",
					"City: $($user.l)",
					"State: $($user.st)",
					"StreetAddress: $($user.streetAddress)",
					"Title: $($user.title)",
					"LogonScript: $oldLogonScript"
				) -join "`n"  # Join properties with new line
				'Groups' = $groupNames -join "`n"  # Join groups with new line
				'Date Disabled' = (Get-Date).ToString('MM/dd/yyyy')  # Current date
			}

			# Export the user properties hashtable to the new Excel file without headers and append to 'Disabled Users' sheet
			New-Object PSObject -Property $userProperties | Export-Excel -Path $newExcelFilePath -WorksheetName 'Disabled Users' -NoHeader -Append
			
			# Remove empty rows after row 1 in the Excel file
			$excelPackage = Open-ExcelPackage -Path $newExcelFilePath
			$worksheet = $excelPackage.Workbook.Worksheets['Disabled Users']

			# Find the last row with data in the worksheet
			$lastRow = $worksheet.Dimension.End.Row

			# Iterate through rows starting from row 2 to the last row
			for ($row = $lastRow; $row -gt 1; $row--) {
				$isEmpty = $true
				for ($col = 1; $col -le $worksheet.Dimension.Columns; $col++) {
					if ($worksheet.Cells[$row, $col].Text -ne "") {
						$isEmpty = $false
						break
					}
				}
				if ($isEmpty) {
					$worksheet.DeleteRow($row)
				}
			}

			# Save the modified Excel file
			$excelPackage.Save()

            # Remove user from all groups except 'Domain Users'(You can add more exclusions here)
			$userGroups = Get-ADUser $user | Get-ADPrincipalGroupMembership | Where-Object { $_.DistinguishedName -ne 'CN=Domain Users,CN=Users,DC=<DC>,DC=<DC>,DC=<DC>' }

			foreach ($group in $userGroups) {
				Remove-ADGroupMember -Identity $group -Members $user -Confirm:$false
				Write-Host "Removed $($user.SamAccountName) from group $($group.Name)"
			}
			
			# Remove specified properties from the user object
			$propertiesToRemove = @('TelephoneNumber', 'physicalDeliveryOfficeName', 'PostalCode', 'l', 'st', 'StreetAddress', 'Title', 'ScriptPath')

			foreach ($property in $propertiesToRemove) {
				$propertyValue = $user.$property
				if ($propertyValue -ne $null) {
					Set-ADUser -Identity $user -Clear $property
					Write-Host "Removed property '$property' from $($user.SamAccountName)"
				} else {
					Write-Host "Property '$property' is already null for $($user.SamAccountName). Skipping removal."
				}
			}
			
			# Disable the user
			Set-ADUser -Identity $user -Enabled $false
			Write-Host "Disabled user $($user.SamAccountName)"

            # Move the user to the target OU
            Move-ADObject -Identity $user.DistinguishedName -TargetPath $targetOU
			Write-Host "Moved user $($user.SamAccountName) to $targetOU"

            # Call the function to remove the user from the home drive folder
            Remove-UserFromHomeDrive -UserName "$($user.SamAccountName)"
			
			# Clear the arrays for next iteration
			$userPropertiesArray = @()
			Write-Host "Cleared the Properties Array..."
			$userGroupMembershipsArray = @()
			Write-Host "Cleared the Groups Array..."
        }
    }
}

# Clear user data in the Excel file
$excelPackage = Open-ExcelPackage -Path $excelFilePath
$worksheet = $excelPackage.Workbook.Worksheets['Sheet1']

# Clear data starting from row 2
for ($row = 2; $row -le $worksheet.Dimension.Rows; $row++) {
    for ($col = 1; $col -le $worksheet.Dimension.Columns; $col++) {
        $worksheet.Cells[$row, $col].Value = $null
    }
}

# Save the modified Excel file without changing formatting
$excelPackage.Save()
