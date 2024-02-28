Created by Bland D. Wallace III
For Use by the DOC - Administering Active Directory

*** Please read carefully ***
*** You must run a powershell window in administration mode ***
*** Please make sure you are in the correct path when running the script, otherwise it will not work ***
*** Also, ensure the correct path to the file is corrected when running this.  I know I have it working from my local computer, but you should be able to use the network
path for the file and update the script accordingly. ***

This script is to be run to create new users in active directory.  It will automatically do everything we are tasked to do for each user like:
Format the user's name
Copy the template user properties to the newly created user
Create a home drive and add that user to the folder with modify permissions
If the user doesn't have a description, you are able to get a description set for each sheet when prompted, which will be applied to the user's Description and Job Title fields.

I have included the spreadsheet representing each department or "OU".  This is mapped in the script for each user to template user assignments.

This script is designed to use multiple ways to construct a user's name from the data.  It will go through multiple ways to get the user's account created and unique domain wide.
