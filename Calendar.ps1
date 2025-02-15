﻿#Connect to exchange powershell. Not MSonline, exchange. 
$O365Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $O365Cred -Authentication Basic -AllowRedirection
Import-PSSession $o365Session

#Put conference room titles in an array for adding permissions
$HTXConference = "CCR@example.com", "FCR@example.com", "LRC@example.com", "SCR@example.com"
$ATXConference = "ATXCONF@example.com", "ATXCONF2@example.com"
$AllConference = $HTXConference + $ATXConference

#Prompt the user what they would like to do:
do {
$StartInput = Read-host "Enter what you would like to do:
1. Add Calendar Permissions
2. Remove Calendar Permissions
3. Check Calendar Permissions"
#Begin first switch statement
Switch ($StartInput) {
1 {
#Prompt the user to enter the email of the user being given calendar permissions
#Note that after every switch statement, the connection from exchange is removed via Remove-PSSession
#Also note that this is a nested DO loop inside of a switch. Do Loop inside of a switch inside of a do loop. 
do {
$userinput = Read-Host "Enter the email of the user that you are adding to the calendar: "
$ChoiceInput = Read-Host "What calendar would you like to change permissions for?
1. For HTX Conference Rooms
2. For ATX Conference Rooms
3. For every Conference room
4. For a specific user"

#Note that the way each switch works is that first, users are added as delegates to the resource mailbox. 
#Next, the calendar value is added to each value in the array
#Next, the users are given editing rights to all the calendars
Switch ($Choiceinput) {
1 {
   foreach ($HTXConf in $HTXConference) {
   Set-CalendarProcessing $HTXConf -ResourceDelegates $userinput
   $HTXConf += ":\calendar"
   Add-MailboxFolderPermission $HTXConf -user $userinput -accessrights Editor}
   }
2 {
   foreach ($ATXConf in $ATXConference) {
   Set-CalendarProcessing $ATXConf -ResourceDelegates $userinput
   $ATXConf += ":\calendar"
   Add-MailboxFolderPermission $ATXConf -user $userinput -accessrights Editor}
   }
3 {
   foreach ($Allconf in $Allconference) {
   Set-CalendarProcessing $AllConf -ResourceDelegates $userinput
   $AllConf += ":\calendar"
   Add-MailboxFolderPermission $AllConf -user $userinput -accessrights Editor}
   }
4 {
#This section gives access rights to specific user calendars. Helpful for anyone that needs access to an Excec calendar or other confused people trying to share their calendar
$userchoice = Read-host "Enter the email of the user you would like to give" $userinput "access to."
$userchoice = $userchoice + ":\calendar"
$userrights = Read-Host "Enter the access you would like the user to have: (Owner/Editor/Reviewer/AvailabilityOnly/LimitedDetails)"
Add-MailboxFolderpermission $userchoice -user $userinput -accessrights $userrights
Remove-Pssession $O365Session
}
Default {continue}
}
}
While ($input -ne 1 -and $input -ne 2 -and $input -ne 3 -and $input -ne 4)
}
#Second choice within do loop, this section will be expanded for removing calendar permissions instead of adding
2 {Write-Host "Input 2"}
#Begin third choice, this is for viewing permissions. Helpful for viewing permissions of any random user as well as the permissions on a specific calendar. 
#This calendar also shows the different permissions on resource calendars. Permission on the calendar mailbox itself != a resource delegate.
3 {
do {
$ChoiceInput = Read-Host "What Calendar would you like to view permissions for?
1. For HTX Conference Rooms
2. For ATX Conference Rooms
3. For every Conference room
4. For a specific user"
Switch ($Choiceinput) {
1 {
   foreach ($HTXConf in $HTXConference) {
   Get-CalendarProcessing $HTXConf | Select Identity,ResourceDelegates
   $HTXConf += ":\calendar"}
   foreach ($HTXCONF in $HTXConference) {
   $HTXConf += ":\calendar"
   Get-MailboxFolderPermission $HTXConf | Select identity,User,AccessRights | Format-Table}
   }
2 {
   foreach ($ATXConf in $ATXConference) {
   Get-CalendarProcessing $ATXConf | Select Identity,ResourceDelegates
   $ATXConf += ":\calendar"}
   foreach ($ATXCONF in $ATXConference) {
   $ATXConf += ":\calendar"
   Get-MailboxFolderPermission $ATXConf | Select identity,User,AccessRights | Format-Table}
   }
3 {
   foreach ($Allconf in $Allconference) {
   Get-CalendarProcessing $AllConf | Select Identity,ResourceDelegates
   $AllConf += ":\calendar"}
   foreach ($AllCONF in $AllConference) {
   $HTXConf += ":\calendar"
   Get-MailboxFolderPermission $AllConf | Select identity,User,AccessRights | Format-Table}
   }
4 {
$userchoice = Read-host "Enter the email of the user you would like to check permissions for:"
$userchoice = $userchoice + ":\calendar"
Get-MailboxFolderPermission $userchoice
Remove-Pssession $O365Session
}
Default {continue}
}
}
While ($input -ne 1 -and $input -ne 2 -and $input -ne 3 -and $input -ne 4)
}
Default {continue}
}
}
while ($input -ne 1 -and $input -ne 2 -and $input -ne 3)
