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


