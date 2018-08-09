<#
.SYNOPSIS
    A script to manage mailbox permissions when using security groups.
.DESCRIPTION
    This script manages mailbox delegation for group mailboxes by granting the group
    "Send As"and"Send on Behalf" rights to the mailbox.  It then takes the members of
    the group and add them individually to the mailbox so the "AutoMapping" property
    automatically assigns the mailbox to the user in Outlook.
.PARAMETER Group
    The name of the security group.
.PARAMETER Mailbox
    The name of the mailbox in the john.doe@example.org format.
.EXAMPLE
    C:\PS> .\MailboxPermissions -Group "Mailbox Security Group" -Mailbox "john.doe@example.org"
.LINK
    https://github.com/DoctorKisow/ExchangeGroupMailbox 
.NOTES
    Author: Matthew R. Kisow, D.Sc.
    Date:   August 9, 2018
#>

Param(
    [Parameter(Mandatory=$True, Position=0)]
    [string]$Group,
    [Parameter(Mandatory=$True, Position=1)]
    [string]$Mailbox
)

Function ErrorChecking
{
    $IsGroup = Get-DistributionGroup -Id $Group -ErrorAction 'SilentlyContinue'
    if (-not $IsGroup)
    {
         Write-Host "The group $Group does not exist or is not a group."
         exit 1
    }
    
    $IsUser = Get-Recipient -ANR $Mailbox -ErrorAction 'SilentlyContinue'
    if (-not $IsUser)
    {
         Write-Host "The mailbox $Mailbox does not exist or is not a mailbox."
         exit 1
    }
}

Function VerifyUpdates
{
    # The de-facto ARE YOU SURE function used to review and proceed with script execution.
    Write-Host "Are you sure that you want to update the security on this group mailbox."
    $CONTINUE = Read-Host "[Y]es or [N]o"

    while("Y","N" -notcontains $CONTINUE)
    {
        Write-Host "Incorrect response, please try again."
	    $CONTINUE = Read-Host "[Y]es or [N]o"
    }

    IF ($CONTINUE -eq 'N')
    {
        Write-Host "The mailbox security updates have been aborted at the operators request."
        exit 1
    }
}

Function UpdateMailboxSecurity
{
    # Get the members from the security group.
    $DISTRIBUTION_LIST = Get-DistributionGroupMember $Group | Select-Object -ExpandProperty Name

    foreach ($USER in $DISTRIBUTION_LIST)
    {
        # Add the "FullAccess" permissions for each security group member to the group mailbox.
    	Add-MailboxPermission -Identity $Mailbox -User $USER -AccessRights 'FullAccess' -InheritanceType 'all' -AutoMapping '$true'
    }

    # Set the "SendAs" permission to the security group.
    Get-User -identity $Mailbox| Add-ADPermission -User $Group -ExtendedRights Send-As

    # Remove and reset the "Send on Behalf" permission to the security group.
    Get-Mailbox -identity $Mailbox | Set-Mailbox -GrantSendOnBehalfTo $null
    Get-Mailbox -identity $Mailbox | Set-Mailbox -GrantSendOnBehalfTo $Group
}

### Main Script

# If not already loaded, load the Exchange server PowerShell modules.
if ((Get-PSSnapin -Name *Exchange* -ErrorAction SilentlyContinue) -eq $null )
{
    Add-PSSnapin *Exchange*
}

ErrorChecking
VerifyUpdates
UpdateMailboxSecurity

# Exit Script
exit 0
