Create-PingCastleReport.ps1
===================

Create-PingCastleReport.ps1 script launches the PingCastle analysis with the default parameters and sends the report by mail.

***To use this script you must configure these variables :***
- **$domains** line 32 with domain name "contoso.com"
- **$PSScriptRoot** line 34 with script path "c:\scripts\
- **$SmtpServer** line 166 with ip address or SMTP name server "10.2.3.1" or "smtp.contoso.com"
- **$EmailFrom** line 167 with sender address "report@contoso.com"
- **$EmailTo** line 168 with receiver address "receiver@contoso.com"

![alt tag](Images/Create-PingCastleReport-Mail.png)

Fix-OwnerComputer.ps1
===================

Fix-OwnerComputer.ps1 can list and fix the computers objects where owner is not Domain Admins.

For each computer where owner is not Domain Admins group, It will then retrieve the computer name.


##Examples
Using the script to list computers with owner is not Domain Admins:
```PowerShell
Fix-OwnerComputer.ps1 -List
```
![alt tag](Images/Fix-OwnerComputer-List.png)

Using the script to List and Fix computers with owner is not Domain Admins:
```PowerShell
Fix-OwnerComputer.ps1 -Fix
```
![alt tag](Images/Fix-OwnerComputer-Fix.png)


Get-AzureADGroupEmpty.ps1
===================

Get-AzureADGroupEmpty.ps1 can list empty sync and cloud groups in Microsoft Entra.

For each group without members, It will then retrieve the name group.


##Examples
Using the script to list synced groups without members:
```PowerShell
Get-AzureADGroupEmpty.ps1 -ListSync
```

Using the script to list Cloud groups without members:
```PowerShell
Get-AzureADGroupEmpty.ps1 -ListCloud
```

Using the script to Remove synced groups without members:
```PowerShell
Get-AzureADGroupEmpty.ps1 -Remove
```

