Each folder has its own powershell script and README along with a sample excel or text file.

Create-User does just that.  It will create a single user or multiple users according to the data in the excel file.  Each worksheet in the excel file represents a department and is mapped to a specific OU and a specific template user in that OU.

Disable-User does just that.  It will disable a single user or multiple users according to the data in the excel file.  The included excel file for this should work as well to your liking.  It also includes code to record the users that were disabled along with their data just in case a user was not supposed to be on the list got disabled and needs restored.  It will change the password for first logon as well, I think I have a sample password in this script you will need to modify according to your environment.

Move-User does just that.  It will move a single user or multiple users according to the data in the excel file.  Each worksheet in the excel file is the same as the create-user excel file, but for moving the users instead.

Restore-User does just that.  It will restore a single user or multiple users according to the data in a text file.  I have included a example text file.  This is the trick for this script though:
  1.  You can select single user which you will then input the name of the user.  In my case, I had the data saved to the disableUser.xlsx from the disable powershell script.  This recorded my user data as the CN of my user (so it was lastName firstName).  Make sure the format you use when prompted is the same as in the disableUsers.xlsx file.
  2.  If you select multiple users from a text file, you will populate the text file according to how the user's name looks in the excel file.
  3.  This will ensure proper execution.

Each script does what is intended and nothing more.  If you wish to modify it according to your environment then do so.

By creating these scripts it made Help Desk jobs more efficient.  I tried to make an all inclusive script using all of this logic combined, but I couldn't quite get that working the way these work individually.  Gets confusing when you try to tie in everything because you need to ensure the variables being passed and processes for those functions.

I know by using these scripts it will make your job a lot easier to do.  We all have mundane tasks, and sometimes get caught up in other projects, so when you have tools like these at your disposal it makes life better.
