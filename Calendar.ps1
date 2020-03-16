#Connect to exchange powershell. Not MSonline, exchange. 
$O365Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $O365Cred -Authentication Basic -AllowRedirection
Import-PSSession $o365Session

#Put conference room titles in an array
$HTXConference = "CCR@trioltd.com:\calendar", "FCR@trioltd.com:\calendar", "LRC@trioltd.com:\calendar", "SCR@trioltd.com:\calendar"
$ATXConference = "ATXCONF@trioltd.com:\calendar", "ATXCONF2@trioltd.com:\calendar"
$AllConference = $HTXConference + $ATXConference

#Prompt the user what they would like to do:
do {
$StartInput = Read-host "Enter what you would like to do:
1. Add Calendar Permissions
2. Remove Calendar Permissions
3. Check Calendar Permissions"
Switch ($StartInput) {
1 {
#Prompt the user to enter the email of the user being given calendar permissions
#Note that after every switch statement, the connection from exchange is removed via Remove-PSSession
do {
$userinput = Read-Host "Enter the email of the user that you are adding to the calendar: "
$ChoiceInput = Read-Host "What calendar would you like to change permissions for?
1. For HTX Conference Rooms
2. For ATX Conference Rooms
3. For every Conference room
4. For a specific user"
Switch ($Choiceinput) {
1 {
   foreach ($HTXConf in $HTXConference) {
   Add-MailboxFolderPermission $HTXConf -user $userinput -accessrights Editor
   Set-CalendarProcessing $HTXConf -ResourceDelegates $userinput}
   Remove-Pssession $O365Session
   }
2 {
   foreach ($ATXConf in $ATXConference) {
   Add-MailboxFolderPermission $ATXConf -user $userinput -accessrights Editor
   Set-CalendarProcessing $ATXConf -ResourceDelegates $userinput}
   Remove-Pssession $O365Session
   }
3 {
   foreach ($Allconf in $Allconference) {
   Add-MailboxFolderPermission $AllConf -user $userinput -accessrights Editor
   Set-CalendarProcessing $AllConf -ResourceDelegates $userinput}
   Remove-Pssession $O365Session
   }
4 {
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
2 {Write-Host "Input 2"}
3 {Write-Host "Input 3"}
Default {continue}
}
}
while ($input -ne 1 -and $input -ne 2 -and $input -ne 3)
