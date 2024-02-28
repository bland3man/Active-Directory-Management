<# Created by Bland D. Wallace III
   bdw3@live.com
   This should create bulk users in AD.  Each sheet in the excel file represents its own department and is mapped to an OU.  This will create users with the template user
   of their respective OU to be created in according to the excel sheet their data is found.  Of course, go ahead and change it to your liking, but I have the code in this
   script that created those users for my purpose at work.    you will have to adjust the formatting of the samAccountName and such according to your specifications.  I 
   have kept the formatting as reading FirstName.LastName or FirstNameMiddleInitial.LastName.  If your AD lists users lastName firstName you have pretty much nothing to change
   other than what is specific to your standards.  This is pretty standard for most AD's I have worked on though.
#>

# Import the ImportExcel module
Import-Module ImportExcel

# Import the Active Directory module
Import-Module ActiveDirectory

# Set the path to the Excel file (.xlsx)
$excelFilePath = "C:\path\to\userDataExcelFile"

# Specify the root OU to search within
$RootOU = "OU=<parentOU>,DC=<DC>,DC=<DC>,DC=<DC>"

# Specify the domain to search for user's domain wide
$Domain = "DC=<DC>,DC=<DC>,DC=<DC>"

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

# Function to search Active Directory for a user
function Search-ADUser {
    param (
        [string]$FirstName,
        [string]$LastName,
        [string]$MiddleInitial
    )

    # Build the filter based on provided information
    $filter = "(&(givenName=$FirstName)"
    if ($LastName) {
        $filter += "(sn=$LastName)"
    }
    if ($MiddleInitial) {
        $filter += "(initials=$MiddleInitial)"
    }
    $filter += ")"

    # Search Active Directory
    $users = Get-ADUser -LDAPFilter $filter

    return $users
}

# Function to create a new user in Active Directory
function Create-NewUser {
    param (
        [string]$FirstName,
        [string]$LastName,
        [string]$MiddleInitial,
        [string]$OU,
        [string]$SheetDescription,
        [string]$JobTitle,
        [string]$TargetTemplateUser,
		[string]$HomeDriveStatus
    )

    # Define properties to copy
    $propertiesToCopy = @(
        "physicalDeliveryOfficeName",
        "telephoneNumber",
        "streetAddress",
        "l",
        "st",
        "postalCode",
        "scriptPath",
        "department",
        "Company"
    )

    # Construct the user's CN
    $userCN = "$LastName $FirstName"
    if ($MiddleInitial) {
        $userCN += " $MiddleInitial"
    }

    # Construct possible samAccountNames and UPNs
    $possibleConstructedNames = @()

    # Original user data
    $originalName = "$FirstName$MiddleInitial$LastName"
    $possibleConstructedNames += $originalName

    # Remove hyphens and apostrophes
    $cleanedName = $originalName -replace "['-]"

    # Construct possible names based on cleaned data
    $possibleNames = @()
    if ($cleanedName -ne $originalName) {
        $possibleNames += $cleanedName
        $parts = $cleanedName -split '\.'
        if ($parts.Count -eq 2) {
            $possibleNames += "$($parts[0]).$($parts[1].Replace('.', ''))"
            $possibleNames += "$($parts[1].Replace('.', '')).$($parts[0])"
        }
    }

    # Add variations to the hashtable
    $possibleConstructedNames += $possibleNames

    # Now, $possibleConstructedNames contains variations based on both original and cleaned data

    # Construct samAccountName based on FirstName.LastName format
    $samAccountName = "$FirstName.$LastName"

    # Remove apostrophes and hyphens from samAccountName
    $samAccountName = $samAccountName -replace "['-]"

    # Construct the User Principal Name (UPN)
    $userPrincipalName = "$FirstName.$LastName@delaware.gov"

    # Remove apostrophes and hyphens from UPN
    $userPrincipalName = $userPrincipalName -replace "['-]"

    # Check for an existing user with the same samAccountName
    $existingUser = Search-ADUser -FirstName $FirstName -LastName $LastName -MiddleInitial $MiddleInitial

    # If there's a conflict, handle it
    if ($existingUser) {
        $samAccountName = "$FirstName$MiddleInitial$LastName"
        $samAccountName = $samAccountName -replace "['-]"
        $includeMiddleInitial = Read-Host "User with samAccountName '$samAccountName' already exists. Include Middle Initial in samAccountName? (1 = Yes, 2 = No)"
        if ($includeMiddleInitial -eq '1') {
            $samAccountName = "$FirstName$MiddleInitial.$LastName"
        } else {
            Write-Host "Skipping user creation for $userCN, proceeding to process the next user in the excel file"
            Add-Content -Path "\\path\to\skippedUser.txt" -Value $userCN
            return
        }
        $userPrincipalName = "$samAccountName@<domain.gov or something.com>"
    }

    # Determine the value for 'description' and 'title' properties
    if ($JobTitle -ne "") {
        $description = $JobTitle
        $title = $JobTitle
    } else {
        $description = $SheetDescription
        $title = $SheetDescription
    }

    # Create a new user
    $newUser = New-ADUser -Name $userCN -SamAccountName $samAccountName -GivenName $FirstName -Surname $LastName -Initials $MiddleInitial -Path $OU -Enabled $true -UserPrincipalName $userPrincipalName -DisplayName "$LastName, $FirstName $MiddleInitial. (whatever you need to be displayed here after the user's name)" -Title $title -Description $description -AccountPassword (ConvertTo-SecureString "<temporaryPassword>" -AsPlainText -Force) -ChangePasswordAtLogon $true

    Write-Host "User created successfully."
    Write-Host "Distinguished Name (DN): CN=$userCN,$OU"

    # Wait for up to 40 seconds, checking every 10 seconds
    for ($i = 0; $i -lt 4; $i++) {
        # Check if the new user exists in the domain
        $userExists = Get-ADUser -Filter {SamAccountName -eq $samAccountName}

        if ($userExists) {
            Write-Host "User '$samAccountName' found."

            # Wait for an additional 10 seconds to allow properties to replicate
            Write-Host "Waiting for an additional 10 seconds for Active Directory to replicate..."
            Start-Sleep -Seconds 10

            # Retrieve the new user again
            $newUser = Get-ADUser -Filter {SamAccountName -eq $samAccountName}

            Write-Host "User '$samAccountName' found. Distinguished Name (DN): $($newUser.DistinguishedName)"
            break
        }

        Write-Host "User '$samAccountName' not found. Waiting for 10 seconds..."
        Start-Sleep -Seconds 10
    }

    # Call Set-UserProperties function
    Set-UserProperties -User $newUser -OU $OU -TargetTemplateUser $TargetTemplateUser -HomeDriveStatus $HomeDriveStatus
}

# Function to set additional properties for a user in Active Directory
function Set-UserProperties {
    param (
        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADUser]$User,
        
        [Parameter(Mandatory=$true)]
        [string]$OU,

        [Parameter(Mandatory=$true)]
        [string]$TargetTemplateUser,
		
		[string]$HomeDriveStatus
    )

    try {
        # Output the retrieved values
        Write-Host "Retrieved targetOU: $OU"
        Write-Host "Retrieved targetTemplateUser: $TargetTemplateUser"

        # Retrieve the target template user from Active Directory
        $templateUserObject = Get-AdUser -Filter { DistinguishedName -eq $TargetTemplateUser } -Properties *

        if (-not $templateUserObject) {
            Write-Host "Template user not found in Active Directory: $($TargetTemplateUser). Skipping."
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
			'scriptPath' = 'Login script'
			'department' = 'Department'
			'Company' = 'Company'
            # Add more properties as needed
        }

        # Create a hashtable to store user properties from the template
        $userProperties = @{}

        # Display the retrieved properties
        Write-Host "Properties from template user ($TargetTemplateUser):"
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
                Write-Host "$($friendlyName): Not found for $($TargetTemplateUser)"
            }
        }

        <## Set properties on the user object
        foreach ($property in $userProperties.Keys) {
            $propertyValue = $userProperties[$property]
            Set-AdUser -Identity $User.DistinguishedName -Replace @{$property = $propertyValue}
        }#> # Use this way if you rather use the DN, but it won't work because you will need to change other coding to accept this parameter
		
		# Set properties on the user object
		foreach ($property in $userProperties.Keys) {
			$propertyValue = $userProperties[$property]
			Set-AdUser -Identity $User.SamAccountName -Replace @{$property = $propertyValue}
		}

        # Retrieve groups from the template user
        $templateGroups = Get-AdUser -Filter { DistinguishedName -eq $TargetTemplateUser } -Properties MemberOf | Select-Object -ExpandProperty MemberOf

        # Output the template user's groups
        Write-Host "Groups from template user: $($templateGroups -join ', ')"

        # Add groups to the user object
        foreach ($group in $templateGroups) {
            Add-AdGroupMember -Identity $group -Members $User.SamAccountName -ErrorAction SilentlyContinue
            Write-Host "User added to group: $group"
        }

        Write-Host "User properties and groups set successfully."
		
		# After setting user properties, call Create-UserHomeDrive
		Create-UserHomeDrive -User $User -HomeDriveStatus $HomeDriveStatus

    } catch {
        Write-Warning "Error setting user properties and groups: $_"
    }
}

# Function to create user's home drive folder and set permissions
function Create-UserHomeDrive {
    param (
        [Microsoft.ActiveDirectory.Management.ADUser]$User,
        [string]$HomeDriveStatus
    )

    $homeDriveRoot = "\\path\to\userHomeDrives"

    if ($User) {
        $firstName = $User.GivenName
        $lastName = $User.Surname
        $folderName = "$firstName.$lastName"
        $folderPath = Join-Path -Path $homeDriveRoot -ChildPath $folderName

        # Check if the folder should be created based on Home Drive status
        if ($HomeDriveStatus -eq "Yes") {
            # Create the user's home drive folder if it doesn't exist
            if (-not (Test-Path $folderPath)) {
                New-Item -Path $folderPath -ItemType Directory -Force
                Write-Host "Home drive folder created for user '$firstName $lastName'."

                # Retry mechanism to wait for user identity replication
                $maxRetries = 12
                $retryInterval = 5  # seconds
                $retryCount = 0
                $userSID = $null

                while ($retryCount -lt $maxRetries -and -not $userSID) {
                    $retryCount++
                    Write-Host "Waiting for user identity replication... Retry $retryCount of $maxRetries"
                    Start-Sleep -Seconds $retryInterval

                    $user = Get-ADUser -Filter { SamAccountName -eq $User.SamAccountName }
                    if ($user) {
                        $userSID = $user.SID.Value
                    }
                }

                if ($userSID) {
                    # Add the user to the folder's security tab with modify permissions
                    $acl = Get-Acl -Path $folderPath
                    $identityReference = New-Object System.Security.Principal.SecurityIdentifier($userSID)
                    $aclRule = New-Object System.Security.AccessControl.FileSystemAccessRule($identityReference, "Modify", "ContainerInherit,ObjectInherit", "None", "Allow")
                    $acl.AddAccessRule($aclRule)
                    Set-Acl -Path $folderPath -AclObject $acl

                    Write-Host "User '$firstName $lastName' added to the security tab of the folder with modify permissions."

                    # Return a custom object with relevant information
                    return [PSCustomObject]@{
                        UserName = $User.SamAccountName
                        UserDisplayName = "$firstName $lastName"
                        HomeDriveFolder = $folderPath
                        SecurityPermissions = "Modify"
                    }
                } else {
                    Write-Host "User identity could not be resolved after $maxRetries retries."
                }
            } else {
                Write-Host "Home drive folder already exists for user '$firstName $lastName'."
            }
        } else {
            Write-Host "Home drive folder creation skipped for user '$firstName $lastName'."
        }
    } else {
        Write-Host "User object is null. Cannot create home drive folder."
    }
}

# Main Script
foreach ($worksheetName in $worksheetNames) {
    try {
        # Import data from the Excel worksheet
        $userData = Import-Excel -Path $excelFilePath -WorksheetName $worksheetName -ErrorAction Stop

        # Check if there is user data on the sheet
        if ($userData) {
            # Display the worksheet name
            Write-Host "Worksheet: $worksheetName"

            # Retrieve the ouMapping based on the sheet name
            $mapping = $ouMapping[$worksheetName]

            # Check if ouMapping exists for the sheet
            if ($mapping) {
                # Extract OU and TemplateUser from the mapping
                $targetOU = $mapping.OU
                $targetTemplateUser = $mapping.TemplateUser

                # Output the retrieved values
                Write-Host "Retrieved OU: $targetOU"
                Write-Host "Retrieved TemplateUser: $targetTemplateUser"

                # Prompt for sheet description
                $sheetDescription = Read-Host "Enter description for the sheet '$worksheetName':"

                # Loop through each row of user data
                foreach ($user in $userData) {
                    # Display user data for debugging
                    Write-Host "User Data:"
                    $user | Format-Table
                    Write-Host "------------------------"
					
					# Retrieve 'Home Drive' status from user data
					$homeDriveStatus = $user.'Home Drive'
					
					# Retrieve 'Job Title' from user data
					$JobTitle = $user.'Job Title'
					
					# Display 'Job Title' for debugging
					Write-Host "Job Title retrieved: $JobTitle"
										
					# Set 'No' as default value if 'Home Drive' status is null
					if ($null -eq $homeDriveStatus) {
					$homeDriveStatus = 'No'
					}
					
					# Retrieve the value of the 'Home Drive'
					Write-Host "Home Drive Status retrieved: $homeDriveStatus"

                    # Display the user being searched for
                    Write-Host "Searching for user: $($user.FirstName) $($user.MiddleInitial) $($user.LastName)"

                    # Search Active Directory for the user
                    $adUsers = Search-ADUser -FirstName $user.FirstName -LastName $user.LastName -MiddleInitial $user.MiddleInitial

                    if ($adUsers.Count -eq 1) {
                        # Only one user found, proceed with that user
                        $adUser = $adUsers[0]
                        Write-Host "User found in Active Directory."
                        Write-Host "Distinguished Name (DN): $($adUser.DistinguishedName)"
                        # Add your logic here for handling existing users
                    } elseif ($adUsers.Count -gt 1) {
                        # Multiple users found, additional logic may be needed
                        Write-Host "Multiple users found in Active Directory."
                        # Add your logic here for handling multiple users
                    } else {
                        Write-Host "User not found in Active Directory."
                        # Add your logic here for handling non-existing users
                    }
                    Write-Host "------------------------"
					
                    # Call Create-NewUser function if user is not found
					if (-not $adUser) {
						# Call Create-NewUser function
						$newUser = Create-NewUser -FirstName $user.FirstName -LastName $user.LastName -MiddleInitial $user.MiddleInitial -OU $targetOU -JobTitle $user.'Job Title' -SheetDescription $sheetDescription -TargetTemplateUser $targetTemplateUser -HomeDriveStatus $HomeDriveStatus

						# Check if the user creation was successful
						if ($newUser) {
							# Display user properties
							Write-Host "User properties:"
							$newUser | Format-List
						}
					}
				}
            } else {
                Write-Warning "ouMapping not found for worksheet '$worksheetName'"
            }
        }
    } catch {
        # Handle any errors, and continue to the next iteration
        Write-Warning $_.Exception.Message
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