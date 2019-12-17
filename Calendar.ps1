$O365Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $O365Cred -Authentication Basic -AllowRedirection
Import-PSSession $o365Session

$userinput = Read-Host "Enter the email of the user that you are adding to the calendar: "
Get-Mailbox $userinput
do {
$ChoiceInput = Read-Host "What calendar would you like to change permissions for?
1. For HTX Conference Rooms
2. For ATX Conference Rooms
3. For every Conference room
4. For a specific user"
Switch ($Choiceinput) {
1 {Add-MailboxFolderPermission CCR@trioltd.com:\calendar -user $userinput -accessrights Editor
   Add-MailboxFolderPermission FCRR@trioltd.com:\calendar -user $userinput -accessrights Editor
   Add-MailboxFolderPermission LRC@trioltd.com:\calendar -user $userinput -accessrights Editor
   Add-MailboxFolderPermission SCR@trioltd.com:\calendar -user $userinput -accessrights Editor
   }
2 {Add-MailboxFolderPermission ATXCONF@trioltd.com:\calendar -user $userinput -accessrights Editor
   Add-MailboxFolderPermission ATXCONF2@trioltd.com:\calendar -user $userinput -accessrights Editor
   }
3 {Add-MailboxFolderPermission CCR@trioltd.com:\calendar -user $userinput -accessrights Editor
   Add-MailboxFolderPermission FCRR@trioltd.com:\calendar -user $userinput -accessrights Editor
   Add-MailboxFolderPermission LRC@trioltd.com:\calendar -user $userinput -accessrights Editor
   Add-MailboxFolderPermission SCR@trioltd.com:\calendar -user $userinput -accessrights Editor
   Add-MailboxFolderPermission ATXCONF@trioltd.com:\calendar -user $userinput -accessrights Editor
   Add-MailboxFolderPermission ATXCONF2@trioltd.com:\calendar -user $userinput -accessrights Editor
   }
4 {}
Default {continue}
}
}
While ($input -ne 1 -and $input -ne 2 -and $input -ne 3 -and $input -ne 4)
