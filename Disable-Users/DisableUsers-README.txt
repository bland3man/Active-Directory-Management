Created by Bland D. Wallace III
For Use by the DOC - Administering Active Directory

*** Please read carefully ***
*** You must run a powershell window in administration mode ***
*** Please make sure you are in the correct path when running the script, otherwise it will not work ***

The scripts to disable users have the same functionality, but they were split into two different versions.

The medical one will automatically disable the users in a file that Todd Kramer has located in the medical share drive:
\\DOCFPADMIN02\medical Folder$\01 Vital Core Folder\Terminations\VitalCore terminations - disable-ADUsers.xlsm
This will append names to a file of users who have been terminated.

The normal version can be run any time and does the same thing as the medical version.  I believe it will not create a file of users that were disabled though.
One could copy the code from the medical to the regular in order to keep a record of those users that are disabled.

I have included a disabled user's list for you to fill out with the appropriate information.  Once you have that filled out, close the file so that the script won't
error out.  This is because it will automatically clear the file of the user data so it will stay a new blank file each time.