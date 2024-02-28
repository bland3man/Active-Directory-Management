<# Created by Bland D. Wallace III
   bdw3@live.com
   This should move users in Active Directory.  It will first remove any existing properties like office, location, etc. but not the other important properties that the
   user will need when moved.  It removes the current group memberships as well.  After it does this for the user, it will grab the template user of the OU where the
   user is to be moved to, copy its properties and group memberships to the user, and finally move the user to the target OU.  This was designed to use in conjunction with
   a *.xlsm where I have provided an example *.xlsm file to go by.  Of course, associate your own OU to worksheet mappings to your liking.  This was just an example so that
   you can get an idea of how you can structure this script to read from the *.xlsm file.  I left my own mappings so you can see you can do this for every department
   that is part of its own OU.
   The script is designed to clear the excel sheet after user's have been moved.  This way, it will be a fresh sheet each time you use it.
#>

# Import the ImportExcel module
Import-Module ImportExcel

# Import the Active Directory module
Import-Module ActiveDirectory

# Set the path to the Excel file (.xlsx or .xlsm - your choice)
$excelFilePath = "C:\path\to\fileContainingData"

# Specify the root OU to search within
$RootOU = "OU=<parentOU>,DC=<DC>,DC=<DC>,DC=<DC>"

# Specify an array of worksheet names
$worksheetNames = @("Administration", 
                    "BWCI", 
                    "Central VOP", 
                    "CNTI", 
                    "DHR", 
                    "Dover Probation", 
                    "DRCGtown", 
                    "EDCTrain", 
                    "Education", 
                    "Georgetown Probation", 
                    "HRYCI", 
                    "JTVCC", 
                    "MCCC", 
                    "Medical", 
                    "MIS", 
                    "NCCCourt", 
                    "New Castle Probation", 
                    "Plummer", 
                    "SCI", 
                    "Seaford Probation", 
                    "SWRC", 
                    "Wilmington Probation", 
                    "WTC")

# Define the mapping of worksheet names to OUs and template users
$ouMapping = @{
    "Administration" = @{
        OU = "OU=Users,OU=Administration,$RootOU"
        TemplateUser = "CN=!ATemplateMgmt,OU=Users,OU=Administration,$RootOU"
    };
	"BWCI" = @{
        OU = "OU=Users,OU=BWCI,$RootOU"
        TemplateUser = "CN=!ATemplateBWCI,OU=Users,OU=BWCI,$RootOU"
    };
	"Central VOP" = @{
        OU = "OU=Users,OU=Central VOP,$RootOU"
        TemplateUser = "CN=!ATemplateCentralVOP,OU=Users,OU=Central VOP,$RootOU"
    };
	"CNTI" = @{
        OU = "OU=Users,OU=CNTI,$RootOU"
        TemplateUser = "CN=!ATemplateCNTI,OU=Users,OU=CNTI,$RootOU"
    };
	"DHR" = @{
        OU = "OU=Users,OU=DHR,$RootOU"
        TemplateUser = "CN=!ATemplateDHR,OU=Users,OU=DHR,$RootOU"
    };
	"Dover Probation" = @{
        OU = "OU=Users,OU=Dover Probation,$RootOU"
        TemplateUser = "CN=!ATemplateDPP,OU=Users,OU=Dover Probation,$RootOU"
    };
	"DRCGtown" = @{
        OU = "OU=Users,OU=DRCGtown,$RootOU"
        TemplateUser = "CN=!ATemplateDRC,OU=Users,OU=DRCGtown,$RootOU"
    };
	"EDCTrain" = @{
        OU = "OU=Users,OU=EDCTrain,$RootOU"
        TemplateUser = "CN=!ATemplateEDC,OU=Users,OU=EDCTrain,$RootOU"
    };
	"Education" = @{
        OU = "OU=Users,OU=Education,$RootOU"
        TemplateUser = "CN=!ATemplateEDU,OU=Users,OU=Education,$RootOU"
    };
	"Georgetown Probation" = @{
        OU = "OU=Users,OU=Georgetown Probation,$RootOU"
        TemplateUser = "CN=!ATemplateGtown,OU=Users,OU=Georgetown Probation,$RootOU"
    };
	"HRYCI" = @{
        OU = "OU=Users,OU=HRYCI,$RootOU"
        TemplateUser = "CN=!ATemplateGander,OU=Users,OU=HRYCI,$RootOU"
    };
	"JTVCC" = @{
        OU = "OU=Users,OU=JTVCC,$RootOU"
        TemplateUser = "CN=!ATemplateDCC,OU=Users,OU=JTVCC,$RootOU"
    };
	"MCCC" = @{
        OU = "OU=Users,OU=MCCC,$RootOU"
        TemplateUser = "CN=!ATemplateMCCC,OU=Users,OU=MCCC,$RootOU"
    };
	"Medical" = @{
        OU = "OU=Users,OU=Medical,$RootOU"
        TemplateUser = "CN=!ATemplateMed,OU=Users,OU=Medical,$RootOU"
    };
	"MIS" = @{
        OU = "OU=Users,OU=MIS,$RootOU"
        TemplateUser = "CN=!ATemplateMIS,OU=Users,OU=MIS,$RootOU"
    };
	"NCCCourt" = @{
        OU = "OU=Users,OU=NCCCourt,$RootOU"
        TemplateUser = "CN=!ATemplateNCCCourt,OU=Users,OU=NCCCourt,$RootOU"
    };
	"New Castle Probation" = @{
        OU = "OU=Users,OU=New Castle Probation,$RootOU"
        TemplateUser = "CN=!ATemplateHares,OU=Users,OU=New Castle Probation,$RootOU"
    };
	"Plummer" = @{
        OU = "OU=Users,OU=Plummer,$RootOU"
        TemplateUser = "CN=!ATemplatePCCC,OU=Users,OU=Plummer,$RootOU"
    };
	"SCI" = @{
        OU = "OU=Users,OU=SCI,$RootOU"
        TemplateUser = "CN=!A_SCITemplate,OU=Users,OU=SCI,$RootOU"
    };
	"Seaford Probation" = @{
        OU = "OU=Users,OU=Seaford Probation,$RootOU"
        TemplateUser = "CN=!ATemplateSeaford,OU=Users,OU=Seaford Probation,$RootOU"
    };
	"SWRC" = @{
        OU = "OU=Users,OU=SWRC,$RootOU"
        TemplateUser = "CN=!ATemplateSWRC,OU=Users,OU=SWRC,$RootOU"
    };
	"Wilmington Probation" = @{
        OU = "OU=Users,OU=Wilmington Probation,$RootOU"
        TemplateUser = "CN=!ATemplateWilm,OU=Users,OU=Wilmington Probation,$RootOU"
    };
	"WTC" = @{
        OU = "OU=Users,OU=WTC,$RootOU"
        TemplateUser = "CN=!ATemplateWTC,OU=Users,OU=WTC,$RootOU"
    }
    # Add mappings for other worksheets as needed
}

# Function to remove properties and groups from an existing user
function Remove-Properties {
    param (
        [PSCustomObject]$user,
        [hashtable]$ouMapping,
        [string]$RootOU
    )

    try {
        # 1. Get the OU from the DN
        $ou = $user.DistinguishedName -replace '^CN=[^,]+,(.*)$', '$1'

        # 2. Check if the OU is under the RootOU
        if ($ou -like "*$RootOU") {
            # Construct the comparison OU for ouMapping
            $ouToSearch = $ou -replace 'OU=Users,', ''  # Strip the initial "OU=Users,"

            # 3. Search for worksheetName string in the DN
            $worksheetNameMatch = $worksheetNames | Where-Object { $ouToSearch -match $_ }

            if ($worksheetNameMatch) {
                # 4. Retrieve ouMapping based on the matched worksheetName
                $ouMappingForUser = $ouMapping[$worksheetNameMatch]

                # Output the ouMapping for the user
                Write-Host "ouMapping for the user: $($ouMappingForUser | Out-String)"

                # Store ouMapping as an array for the user
                $ouMappingArrayForUser = @($ouMappingForUser)

                # 5. Retrieve user object with all properties
                $fullUserObject = Get-AdUser -Identity $user.SamAccountName -Properties *

                # Output values of specified properties before removal
                $propertyMappings = @{
                    'l' = 'City'
                    'St' = 'State/province'
                    'physicalDeliveryOfficeName' = 'Office'
                    'postalCode' = 'Postal Code'
                    'streetAddress' = 'Street Address'
                    'telephoneNumber' = 'Telephone'
                }

                foreach ($propertyMap in $propertyMappings.GetEnumerator()) {
                    $ldapProperty = $propertyMap.Key
                    $friendlyName = $propertyMap.Value
                    $existingValue = $fullUserObject.$ldapProperty

                    if ($existingValue -ne $null) {
                        Set-AdUser -Identity $user.SamAccountName -Clear $ldapProperty -ErrorAction SilentlyContinue
                        Write-Host "Property $($friendlyName) removed from $($user.SamAccountName). Existing value: $($existingValue)"
                    } else {
                        Write-Host "Property $($friendlyName) not found for $($user.SamAccountName)"
                    }
                }

                # 6. Retrieve groups from the template user
                $templateUserObject = Get-AdUser -Identity $ouMappingForUser.TemplateUser -Properties MemberOf
                $templateGroups = $templateUserObject.MemberOf

                # Output the template user's groups
                Write-Host "Groups from template user: $($templateGroups -join ', ')"

                # 7. Remove groups from the user object
                $userGroups = Get-AdUser -Identity $user.SamAccountName -Properties MemberOf | Select-Object -ExpandProperty MemberOf

                foreach ($group in $templateGroups) {
					try {
						if ($userGroups -contains $group) {
							Remove-AdGroupMember -Identity $group -Members $user.SamAccountName -Confirm:$false -ErrorAction Stop
							Write-Host "User removed from group: $group"
						} else {
							Write-Host "User is not a member of group: $group"
						}
					} catch {
						Write-Host "Error removing user from group: $group - $_"
						continue  # Skip to the next group if an error occurs
					}
				}

            } else {
                Write-Host "No matching sheet found for the user DN: $($user.DistinguishedName)"
            }
        } else {
            Write-Host "Skipping OU: $ou. No action taken."
        }
    } catch {
        Write-Host "Error removing properties for $($user.SamAccountName): $_"
    }

}

# Function to move a user and add properties and groups
Function Move-UserAddProperties {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        $user,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        $ouMapping,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        $worksheetName
    )

    try {
        # Retrieve the ouMapping for the specified worksheetName
        $ouMappingForWorksheet = $ouMapping[$worksheetName]

        if (-not $ouMappingForWorksheet) {
            Write-Host "ouMapping not found for $($worksheetName). Skipping."
            return
        }

        # Output the ouMapping details
        Write-Host "ouMapping details for $($worksheetName): $($ouMappingForWorksheet | Out-String)"

        # Output the worksheet name
        Write-Host "Worksheet name for $($user.SamAccountName): $($worksheetName)"

        # Retrieve target template user CN from ouMapping
		$targetTemplateUserCN = ($ouMappingForWorksheet.TemplateUser -split ',', 2)[0] -replace 'CN=', ''

        # Check if TargetOU is specified in the ouMapping
        if (-not $ouMappingForWorksheet.OU) {
            Write-Host "TargetOU not specified in ouMapping. Skipping."
            return
        }

        # Search for targetTemplateUser in Active Directory in the specified OU
		$templateUserObject = Get-AdUser -Filter { CN -eq $targetTemplateUserCN } -Properties *

        if (-not $templateUserObject) {
            Write-Host "Template user not found in Active Directory: $($targetTemplateUserCN). Skipping."
            return
        }

        # Create a hashtable for properties to set
        $propertiesToSet = @{
            'TelephoneNumber' = 'Telephone'
            'StreetAddress' = 'Street Address'
            'PostalCode' = 'Postal Code'
            'l' = 'City'
            'St' = 'State/province'
            'physicalDeliveryOfficeName' = 'Office'
        }

        # Create a hashtable to store user properties from the template
        $userProperties = @{}

        # Display the retrieved properties
        Write-Host "Properties from template user ($targetTemplateUserCN):"
        foreach ($property in $propertiesToSet.Keys) {
            $ldapProperty = $property
            $friendlyName = $propertiesToSet[$property]
            $propertyValue = $templateUserObject.$ldapProperty

            if ($propertyValue -ne $null -and $propertyValue -ne "") {
                Write-Host "$($friendlyName): $($propertyValue)"

                # Store property and value in the hashtable
                $userProperties[$ldapProperty] = $propertyValue
            }
            else {
                Write-Host "$($friendlyName): Not found for $($targetTemplateUserCN)"
            }
        }

        # Set properties on the user object
		foreach ($property in $userProperties.Keys) {
			$propertyValue = $userProperties[$property]
			Set-AdUser -Identity $user.SamAccountName -Replace @{$property = $propertyValue}
		}
		
		# Retrieve the updated user object
		$updatedUser = Get-AdUser -Identity $user.SamAccountName

		# Display the properties (either retrieved from template user or updated)
		Write-Host "Properties for $($user.SamAccountName):"
		foreach ($property in $propertiesToSet.Keys) {
			$ldapProperty = $property
			$friendlyName = $propertiesToSet[$property]
			$propertyValue = (Get-AdUser -Filter { DistinguishedName -eq $user.DistinguishedName } -Properties $ldapProperty).$ldapProperty

			if ($propertyValue -ne $null -and $propertyValue -ne "") {
				Write-Host "$($friendlyName): $($propertyValue)"
			}
			else {
				Write-Host "$($friendlyName): Not found for $($user.SamAccountName)"
			}
		}

        # Retrieve groups from the template user
        $templateGroups = Get-AdUser -Identity $templateUserObject.SamAccountName -Properties MemberOf | Select-Object -ExpandProperty MemberOf

        # Output the template user's groups
        Write-Host "Groups from template user: $($templateGroups -join ', ')"

        # Add groups to the user object
        foreach ($group in $templateGroups) {
            Add-AdGroupMember -Identity $group -Members $user.SamAccountName -ErrorAction SilentlyContinue
            Write-Host "User added to group: $group"
        }

        # Move the user to the targetOU
        Move-AdObject -Identity $user.DistinguishedName -TargetPath $ouMappingForWorksheet.OU

        # Display the move information
        Write-Host "$($user.SamAccountName) has been successfully moved from $($user.DistinguishedName) to $($ouMappingForWorksheet.OU)"

    }
    catch {
        Write-Host "Error moving user and adding properties for $($user.SamAccountName): $_"
    }
}

# Main Script
foreach ($worksheetName in $worksheetNames) {
    try {
        # Read data from the current worksheet, specifying column headers
        $worksheetData = Import-Excel -Path $excelFilePath -WorksheetName $worksheetName -ErrorAction Stop
        Write-Host "Successfully read data from worksheet: $worksheetName"

        # Display the retrieved data
        if ($worksheetData.Count -eq 0) {
            Write-Host "No data found in the Excel file after the header for worksheet: $worksheetName"
        }
        else {
            foreach ($row in $worksheetData) {
                $firstName = $row.'FirstName'
                $lastName = $row.'LastName'
                $middleInitial = $row.'MiddleInitial'

                # Debugging statement to output the entire $row
                Write-Host "Row data: $($row | Format-List | Out-String)"

                # Debugging statement to output the 'Job Title' value
                $jobTitleFromSheet = $row.'Job Title'  # Use 'Job Title' instead of 'JobTitle'
                Write-Host "Job Title from sheet: $($jobTitleFromSheet)"

                # Prompt for sheet description if data exists
                $sheetDescription = Read-Host "Enter a description for the sheet $worksheetName (press Enter to skip):"

                # Construct samAccountName in both formats
                $samAccountNameFormat1 = "$firstName.$lastName"
                $samAccountNameFormat2 = "$firstName$middleInitial.$lastName"

                # Construct the user's distinguished name
                $userDN = "CN=$lastName $firstName$($middleInitial -replace '\s',''),$RootOU"

                # Query Active Directory for user information
                $user = Get-AdUser -Filter { (SamAccountName -eq $samAccountNameFormat1) -or (SamAccountName -eq $samAccountNameFormat2) } -Properties DistinguishedName, SamAccountName, Description, Title

                if ($user) {
                    Write-Host "User found in Active Directory for worksheet $($worksheetName):"
                    Write-Host "  Distinguished Name: $($user.DistinguishedName)"
                    Write-Host "  SamAccountName: $($user.SamAccountName)"
                    Write-Host "  FirstName: $($firstName)"
                    Write-Host "  LastName: $($lastName)"
                    Write-Host "  MiddleInitial: $($middleInitial)"

                    # Output current user description
                    Write-Host "Current Description: $($user.Description)"

                    # Check if there is data in the "Job Title" column
                    if ($jobTitleFromSheet) {
                        # Update with Job Title from the sheet
                        Set-AdUser -Identity $user.SamAccountName -Description $jobTitleFromSheet -Title $jobTitleFromSheet
                        Write-Host "Description and job title updated with Job Title from the sheet."
                    }
                    elseif ($sheetDescription) {
                        # Update with sheet description
                        Set-AdUser -Identity $user.SamAccountName -Description $sheetDescription -Title $sheetDescription
                        Write-Host "Description and job title updated with sheet description."
                    }
                    else {
                        Write-Host "No data in 'Job Title' column. Keeping existing description."
                    }
					
					# Call the Remove-Properties function
					Remove-Properties -user $user -ouMapping $ouMapping -RootOU $RootOU

                    # Call the Move-UserAddProperties function
                    Move-UserAddProperties -user $user -ouMapping $ouMapping -worksheetName $worksheetName

                    # ... Additional user information can be displayed here ...
                } else {
                    Write-Host "User not found in Active Directory for worksheet $worksheetName, FirstName: $($firstName), LastName: $($lastName), MiddleInitial: $($middleInitial)"
                }
            }
        }
    }
    catch {
        Write-Host "Error reading data from worksheet $($worksheetName): $_"
    }
}

# Clear user data in all worksheets of the Excel file
foreach ($worksheetName in $worksheetNames) {
    try {
        $excelPackage = Open-ExcelPackage -Path $excelFilePath
        $worksheet = $excelPackage.Workbook.Worksheets[$worksheetName]

        if ($worksheet -eq $null) {
            Write-Host "Worksheet not found: $worksheetName"
            continue
        }

        # Clear data starting from row 2
        for ($row = 2; $row -le $worksheet.Dimension.Rows; $row++) {
            for ($col = 1; $col -le $worksheet.Dimension.Columns; $col++) {
                $worksheet.Cells[$row, $col].Value = $null
            }
        }

        # Save the modified Excel file without changing formatting
        $excelPackage.Save()

        # Output message after clearing Excel file for each sheet
        Write-Host "User data in the Excel file ($worksheetName) has been successfully cleared."
    }
    catch {
        Write-Host "Error clearing data from worksheet $($worksheetName): $_"
    }
}
