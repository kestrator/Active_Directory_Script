<#
.SYNOPSIS
    The Get-AzureADGroupEmpty.ps1 script allows you to get information about Entra ID Groups Empty.

.DESCRIPTION
    The Get-AzureADGroupEmpty.ps1 script allows you to get information about Entra ID Groups Empty.
    You can specify: List to list Entra ID Groups Empty or Remove to Remove Entra ID Groups Empty.
    WARNING : The security groups cannot be restored

.PARAMETER ListSync
	To list Entra ID Groups Sync Empty

.PARAMETER ListCloud
	To list Entra ID Groups Cloud Empty

.PARAMETER Remove
    To remove Entra ID Groups Empty

.EXAMPLE
    Get-AzureADGroupEmpty.ps1 -ListSync

    This will show all Entra ID Groups Sync Empty

.EXAMPLE
    Get-AzureADGroupEmpty.ps1 -ListCloud

    This will show all Entra ID Groups Cloud Empty

.EXAMPLE
    Get-AzureADGroupEmpty.ps1 -Remove

    This will show all Entra ID Groups Cloud Empty and remove them

.NOTES
    NAME:    Get-AzureADGroupEmpty.ps1
    AUTHOR:  Jonathan BAUZONE JBE
    DATE:    2024/07/08
    WWW:     N/A
    TWITTER: @BznJnthn

    VERSION HISTORY:
    1.0 2024.07.08
        Initial Version
#>

Param(
[Parameter(Mandatory=$false,HelpMessage="list Empty Sync groups")]
[Switch]$ListSync,
[Parameter(Mandatory=$false,HelpMessage="list Empty Cloud groups ")]
[Switch]$ListCloud,
[Parameter(Mandatory=$false,HelpMessage="Remove Empty groups")]
[Switch]$Remove
)

if(-not (Get-Module AzureAD -ListAvailable)){
    Install-Module AzureAD -Scope CurrentUser -Force
    }
Connect-AzureAD
$result = @()

if($ListSync){
    $groups = Get-AzureADGroup -All:$true | Where-Object DirSyncEnabled -eq $true
    if($groups){
        foreach ($group in $groups)
        {
            If((Get-AzureADGroupMember -ObjectId $group.ObjectId).count -eq 0)
            {
                $result += $($group.displayname + ";" + $group.OnPremisesSecurityIdentifier)
            }
        }
    }
    if($result){
        $result
    }
    Else{
        Write-Warning "No empty Sync groups"
    }
    $result = @()
}

if($ListCloud){
    $groups = Get-AzureADGroup -All:$true | Where-Object DirSyncEnabled -eq $null
    foreach ($group in $groups)
    {
        If((Get-AzureADGroupMember -ObjectId $group.ObjectId).count -eq 0)
        {
            $result += $group.displayname
        }
    }
    if($result){
        $result
    }
    Else{
        Write-Warning "No empty Cloud groups"
    }
    $result = @()
}

if($Remove){
    $groups = Get-AzureADGroup -All:$true | Where-Object DirSyncEnabled -eq $null
    if($groups){
        Write-Warning "The security groups cannot be restored, are you sure to continue ?"
        $Continue = Read-Host "Press Y to continue or any key to cancel"
        if($Continue -eq "Y"){
            foreach ($group in $groups)
            {
                If((Get-AzureADGroupMember -ObjectId $group.ObjectId).count -eq 0)
                {
                    $group.displayname
                    Remove-AzureADGroup -ObjectId $group.ObjectId
                    Write-Host "The group $($group.displayname) has been deleted" -ForegroundColor Green
                }
            }
        }
    }
    Else {
        Write-Warning "No empty Cloud groups"
    }
}


