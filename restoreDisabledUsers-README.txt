Created by Bland D. Wallace III
For Use by the DOC - Administering Active Directory

*** Please read carefully ***
*** You must run a powershell window in administration mode ***
*** Please make sure you are in the correct path when running the script, otherwise it will not work ***
*** Also, ensure the correct path to the file is corrected when running this.  I know I have it working from my local computer, but you should be able to use the network
path for the file and update the script accordingly. ***

This script will prompt you for a user name to restore if a user had been accidently deleted.
It will give you an option for a single user to restore or a list of users to restore.

You will need to formulate the list like this in the text file:

LastName FirstName

So each name will be on a new line and in that format because the disabled-users Record excel file has those names listed in that fashion.

*** After restoring, you should delete the row of data in the disabled-users-record excel file only for the user or users that were restored.