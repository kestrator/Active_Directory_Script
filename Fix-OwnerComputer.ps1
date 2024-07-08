<#
.SYNOPSIS
    The Fix-OwnerComputer.ps1 script allows you to get information about Active Directory Computers owner.

.DESCRIPTION
    The Fix-OwnerComputer.ps1 script allows you to get information about Active Directory Computers owner.
    You can specify: List to list computers objects where owner is not Domain Admins or Fix to replace owner by Domain Admins

.PARAMETER List
	To list computers objects where owner is not Domain Admins

.PARAMETER Fix
    To replace owner by Domain Admins


.EXAMPLE
    Fix-OwnerComputer.ps1 -List

    This will show all the computers in the current domain where owner is not Domain Admins

.EXAMPLE
    Fix-OwnerComputer.ps1 -Fix

    This will show all the computers in the current domain where owner is not Domain Admins and replace owner by Domain Admins

.NOTES
    NAME:    Fix-OwnerComputer.ps1
    AUTHOR:  Jonathan BAUZONE JBE
    DATE:    2024/07/05
    WWW:     N/A
    TWITTER: @BznJnthn

    VERSION HISTORY:
    1.0 2024.07.05
        Initial Version
#>

Param(
[Parameter(Mandatory=$false,HelpMessage="list computer")]
[Switch]$list,
[Parameter(Mandatory=$false,HelpMessage="Fix Owner")]
[Switch]$fix
)

$domain = Get-ADDomain
$Name = $domain.NetBIOSName
$SID = $domain.DomainSID
$DA = Get-ADGroup -Filter "SID -eq '$SID-512'"
$DA = $DA.Name
$group = $Name+"\"+$DA

$user = New-Object System.Security.Principal.NTAccount("$DA")
$computers = Get-ADComputer -Filter *
if ($list){
    foreach ($computer in $computers)
    {
        $ACL = Get-Acl -Path "AD:$($computer.DistinguishedName)"
        If($ACL.Owner -ne $group)
        {
            $computer.Name
        }
    }
}

if ($fix){
    foreach ($computer in $computers)
    {
        $ACL = Get-Acl -Path "AD:$($computer.DistinguishedName)"
        If($ACL.Owner -ne $group)
        {
            $computer.Name
            $usertoremove = $ACL.Owner
            $ACL.SetOwner($user)
            Foreach($ACE in $ACL.Access)
            {
                If($ACE.IdentityReference.Value -eq $usertoremove)
                {
                    $ACL.RemoveAccessRule($ACE)
                }
            }
            Set-Acl -Path "AD:$($Computer.DistinguishedName)" $ACL
        }
    }
}
