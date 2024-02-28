Created by Bland D. Wallace III
For Use by the DOC - Administering Active Directory

*** Please read carefully ***
*** You must run a powershell window in administration mode ***
*** Please make sure you are in the correct path when running the script, otherwise it will not work ***
*** Also, ensure the correct path to the file is corrected when running this.  I know I have it working from my local computer, but you should be able to use the network
path for the file and update the script accordingly. ***

This script is designed to move users who exist from one OU to another.  It will:
Remove existing groups and add the appropriate groups according to the sheet the user data was on
Rename the description and job title fields
Automate the moving from one OU to another and copying over properties like address and such from the template user of the target OU the user is to be moved to.
