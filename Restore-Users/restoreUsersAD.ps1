<# Created by Bland D. Wallace III
   bdw3@live.com
   This should restore user's that were disabled and still exist in AD.  You must already have a disabled user's list with the proper information.  Please use
   the example text sheet I have provided (it is a blank, but this is where you would list your users according to how they are listed in the disable-users.xlsx.
   So, for example, if in the *.xlsx 'Name' column a person is listed as FirstName LastName, you would populate in the text file FirstName LastName.  This is only if
   you choose to restore multiple users.  There is an option to restore a single user, which you would still need to input the user's name as what it appears in the
   *.xlsx file.
   I have created a disable-users.ps1 which records all of this information in a separate excel file.  When that is populated,
   and you have accidently deleted a user's account in AD, this script will take the recorded data of that user and populate it back to the user object.  It will
   proceed to also move the user back to the original OU.  All the necessary information in that file would have been populated in the recorded-users.xlsx file
   so that if a user was to be disabled accidently or mistakenly then that user will be restored.
#>

# Import the ImportExcel module
Import-Module ImportExcel

# Import the Active Directory module
Import-Module ActiveDirectory

# Set the path to the Excel file (.xlsx) containing disabled user records
$newExcelFilePath = "\\path\to\recordedList"

# Specify the target OU to move disabled users back from
$targetOU = "OU=<Users>,OU=<targetOU>,OU=<taretOUParent>,DC=<FQDN>,DC=<FQDN>,DC=<FQDN>" # Make sure this is completed with the DN of the OU you want to move users from

# Mapping of Excel properties to corresponding Active Directory properties
$propertyMapping = @{
    'Office' = 'physicalDeliveryOfficeName'
    'City' = 'l'
    'State' = 'st'
    'LogonScript' = 'scriptPath'
}

# Function to restore user's home drive folder and set permissions
function Restore-UserHomeDrive {
    param (
        [Microsoft.ActiveDirectory.Management.ADUser]$User
    )

    $separatedHomeDriveRoot = "\\path\to\currentLocation\homeDrive"
    $homeDriveRoot = "\\path\to\moveToLocation"

    if ($User) {
        $firstName = $User.GivenName
        $lastName = $User.Surname
        $folderName = "$firstName.$lastName"
        $separatedFolderPath = Join-Path -Path $separatedHomeDriveRoot -ChildPath $folderName
        $folderPath = Join-Path -Path $homeDriveRoot -ChildPath $folderName

        # Check if the separated folder exists
        if (Test-Path $separatedFolderPath) {
            # Add the user back to the folder's security tab with modify permissions
            $acl = Get-Acl -Path $separatedFolderPath
            $identityReference = New-Object System.Security.Principal.NTAccount("$($User.SamAccountName)")
            $aclRule = New-Object System.Security.AccessControl.FileSystemAccessRule($identityReference, "Modify", "ContainerInherit,ObjectInherit", "None", "Allow")
            $acl.AddAccessRule($aclRule)
            Set-Acl -Path $separatedFolderPath -AclObject $acl

            Write-Host "User '$firstName $lastName' added back to the security tab of the separated folder with modify permissions."

            # Move the folder from separated directory to users directory
            Move-Item -Path $separatedFolderPath -Destination $folderPath -Force
            Write-Host "User's home drive folder moved to '$folderPath'."
            
            # Return a custom object with relevant information
            return [PSCustomObject]@{
                UserName = $User.SamAccountName
                UserDisplayName = "$firstName $lastName"
                HomeDriveFolder = $folderPath
                SecurityPermissions = "Modify"
            }
        } else {
            Write-Host "Separated home drive folder not found for user '$firstName $lastName'."
        }
    } else {
        Write-Host "User object is null. Cannot restore home drive folder."
    }
}

# Function to restore a user based on the provided username
function Restore-User {
    param(
        [string]$restoreUsername
    )

    # Transform the username format from "LastName FirstName" to "FirstName.LastName"
    $lastName, $firstName = $restoreUsername -split " "
    $samAccountName = "$firstName.$lastName"

    # Search for the user in the disabled users data from Excel
    $userRecordToRestore = $disabledUsersData | Where-Object { $_.'Name' -eq $restoreUsername }

    # Check if user was found in the Excel data
    if ($userRecordToRestore) {
        Write-Host "User found in Disabled Users record. Information for $($userRecordToRestore.Name):"
        Write-Host "DN: $($userRecordToRestore.DN)"
        
        # Create empty hashtables to store properties and groups
        $userProperties = @{}
        $userGroups = @{}
        
        Write-Host "Properties:"
        # Split the 'Properties' column by newlines and process each line
        $userRecordToRestore.Properties -split "`n" | ForEach-Object {
            # Split each line by colon into property name and value
            $propertyLine = $_ -split ":", 2
            $propertyName = $propertyLine[0].Trim()
            $propertyValue = $propertyLine[1].Trim()
            
            # Check if the property needs to be mapped
            if ($propertyMapping.ContainsKey($propertyName)) {
                $mappedPropertyName = $propertyMapping[$propertyName]
                # Store the property in the hashtable with the mapped name
                $userProperties[$mappedPropertyName] = $propertyValue
                # Output property name and value
                Write-Host "  ${mappedPropertyName}: $propertyValue"
            }
            else {
                # Store the property in the hashtable with the original name
                $userProperties[$propertyName] = $propertyValue
                # Output property name and value
                Write-Host "  ${propertyName}: $propertyValue"
            }
        }

        Write-Host "Groups:"
        # Split the 'Groups' column by newlines and process each group
        $userRecordToRestore.Groups -split "`n" | ForEach-Object {
            $groupName = $_.Trim()
            # Store the group in the hashtable
            $userGroups[$groupName] = $true  # You can set any value for the group, true is just an example
            Write-Host "  $_"
        }

        # Search for the user in Active Directory
        $user = Get-ADUser -Filter { SamAccountName -eq $samAccountName } -SearchBase $targetOU -ErrorAction SilentlyContinue

        # Check if user was found in Active Directory
        if ($user) {
            Write-Host "User found in Active Directory in the specified OU."
            # Update the user in Active Directory with the properties
            foreach ($key in $userProperties.Keys) {
                $value = $userProperties[$key]
                Set-ADUser -Identity $user.SamAccountName -Replace @{$key = $value}
            }

            # Update the user's group membership in Active Directory
            foreach ($groupName in $userGroups.Keys) {
                Add-ADGroupMember -Identity $groupName -Members $user.SamAccountName
            }
            Write-Host "User updated in Active Directory with properties and group memberships."
			
			# Restore the user's home drive
            Restore-UserHomeDrive $user
			
        } else {
            Write-Host "User not found in Active Directory in the specified OU."
        }
        
        # Set the password to "<temporaryPassword>"
        $password = ConvertTo-SecureString -String "Welcome@doc1" -AsPlainText -Force

        # Enable the user account in Active Directory and set the password to change at first login
        Enable-ADAccount -Identity $user.SamAccountName
        Set-ADAccountPassword -Identity $user.SamAccountName -NewPassword $password
        Set-ADUser -Identity $user.SamAccountName -ChangePasswordAtLogon $true

        Write-Host "User account enabled in Active Directory. Password set to 'Welcome@doc1' and set to change at first login."
        
        # Extract the original OU from the DN column in the Excel file
        $originalOU = $userRecordToRestore.DN -replace '^CN=[^,]+,(?<OU>OU=.+),$','$1'
        
        # Extract the parent container or Organizational Unit (OU) from the DN
        $parentContainer = ($userRecordToRestore.DN -split ",", 2)[1]

        # Move the user object to the original OU
        Move-ADObject -Identity $user.DistinguishedName -TargetPath $parentContainer

        Write-Host "User moved to the original OU: $parentContainer"
        
    } else {
        Write-Host "User $restoreUsername not found in the Disabled Users record."
    }
}

# Read data from the disabled users record Excel file
try {
    # Read data from the Excel file, specifying column headers
    $disabledUsersData = Import-Excel -Path $newExcelFilePath -WorksheetName 'Disabled Users' -ErrorAction Stop
    Write-Host "Successfully read data from the Disabled Users record Excel file."
}
catch {
    Write-Host "Error reading data from the Disabled Users record Excel file: $_"
    return
}

# Prompt for the choice of user restoration method
Write-Host "How would you like to select users for restoration?"
Write-Host "1. Enter the username manually"
Write-Host "2. Choose from a list in a text file"

$choice = Read-Host "Enter your choice (1 or 2)"

if ($choice -eq '1') {
    # Prompt for the username to restore
    Write-Host "Enter the username of the user you wish to restore (Format: LastName FirstName)"
    $restoreUsername = Read-Host "Username (Format: LastName FirstName)"

    # Restore the user based on the entered username
    Restore-User $restoreUsername
}
elseif ($choice -eq '2') {
    # Read the list of users from the text file
    $userListFilePath = "\\path\to\listOfUsers"  # Specify the path to your text file
    if (Test-Path $userListFilePath) {
        $userList = Get-Content $userListFilePath
        foreach ($username in $userList) {
            # Restore the user based on the username from the list
            Restore-User $username
        }
    } else {
        Write-Host "User list file not found at: $userListFilePath"
    }
}
else {
    Write-Host "Invalid choice. Please enter 1 or 2."
}
