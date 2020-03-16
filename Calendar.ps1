#Connect to exchange powershell. Not MSonline, exchange. 
$O365Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $O365Cred -Authentication Basic -AllowRedirection
Import-PSSession $o365Session
$HTXConference = "CCR@trioltd.com:\calendar", "FCR@trioltd.com:\calendar", "LRC@trioltd.com:\calendar", "SCR@trioltd.com:\calendar"
$ATXConference = "ATXCONF@trioltd.com:\calendar", "ATXCONF2@trioltd.com:\calendar"

#Prompt the user to enter the email of the user being given calendar permissions
#Note that after every switch statement, the connection from exchange is removed via Remove-PSSession
$userinput = Read-Host "Enter the email of the user that you are adding to the calendar: "
do {
$ChoiceInput = Read-Host "What calendar would you like to change permissions for?
1. For HTX Conference Rooms
2. For ATX Conference Rooms
3. For every Conference room
4. For a specific user"
Switch ($Choiceinput) {
1 {Add-MailboxFolderPermission CCR@trioltd.com:\calendar -user $userinput -accessrights Editor
   Add-MailboxFolderPermission FCR@trioltd.com:\calendar -user $userinput -accessrights Editor
   Add-MailboxFolderPermission LRC@trioltd.com:\calendar -user $userinput -accessrights Editor
   Add-MailboxFolderPermission SCR@trioltd.com:\calendar -user $userinput -accessrights Editor
   Remove-Pssession $O365Session
   }
2 {Add-MailboxFolderPermission ATXCONF@trioltd.com:\calendar -user $userinput -accessrights Editor
   Add-MailboxFolderPermission ATXCONF2@trioltd.com:\calendar -user $userinput -accessrights Editor
   Remove-Pssession $O365Session
   }
3 {Add-MailboxFolderPermission CCR@trioltd.com:\calendar -user $userinput -accessrights Editor
   Add-MailboxFolderPermission FCR@trioltd.com:\calendar -user $userinput -accessrights Editor
   Add-MailboxFolderPermission LRC@trioltd.com:\calendar -user $userinput -accessrights Editor
   Add-MailboxFolderPermission SCR@trioltd.com:\calendar -user $userinput -accessrights Editor
   Add-MailboxFolderPermission ATXCONF@trioltd.com:\calendar -user $userinput -accessrights Editor
   Add-MailboxFolderPermission ATXCONF2@trioltd.com:\calendar -user $userinput -accessrights Editor
   Remove-Pssession $O365Session
   }
4 {$userchoice = "Other"}
Default {continue}
}
}
While ($input -ne 1 -and $input -ne 2 -and $input -ne 3 -and $input -ne 4)

#Prompts the user to enter the email of who they would like to give $userinput access to
#Also removes the exchange connection at the end
if ($userchoice -eq "Other") {
$userchoice = Read-host "Enter the email of the user you would like to give" $userinput "access to."
$userchoice = $userchoice + ":\calendar"
$userrights = Read-Host "Enter the access you would like the user to have: (Owner/Editor/Reviewer/AvailabilityOnly/LimitedDetails)"
Add-MailboxFolderpermission $userchoice -user $userinput -accessrights $userrights
Remove-Pssession $O365Session
}
