<# $O365Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $O365Cred -Authentication Basic -AllowRedirection
Import-PSSession $o365Session #>

Write-Host "What calendar would you like to change permissions for?
1. For HTX Conference Rooms
2. For ATX Conference Rooms
3. For every Conference room
4. For a specific user
"