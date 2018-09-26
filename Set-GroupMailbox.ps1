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
.PARAMETER Convert
    Boolean value represented as $True or $False; used to indicate the conversion of
    the mailbox to a shared mailbox.  Default value is $False.
.PARAMETER UpdateSecurity
    Boolean value represented as $True or $False; used to update mailbox security only.
    Default value $False.
.EXAMPLE
    C:\PS> .\MailboxPermissions -Group "Mailbox Security Group" -Mailbox "john.doe@example.org"
    C:\PS> .\MailboxPermissions -Group "Mailbox Security Group" -Mailbox "john.doe@example.org" -Convert $True
.LINK
    https://github.com/DoctorKisow/ExchangeGroupMailbox 
.NOTES
    Author: Matthew R. Kisow, D.Sc.
    Date:   September 26, 2018
#>

Param(
 [Parameter(Mandatory=$True, Position=0)]
 [string]$Group,
 [Parameter(Mandatory=$True, Position=1)]
 [string]$Mailbox,
 [Parameter(Mandatory=$False)]
 [bool]$Convert,
 [Parameter(Mandatory=$False)]
 [bool]$UpdateSecurity
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

    # This is set to ensure that if the Convert parameter is not passed it will be set to $False by default.
    if ($Convert -eq $Null -or $Convert -eq '')
    {
         $Convert = 'False'
    }

    # This is set to ensure that if the UpdateSecurity parameter is not passed it will be set to $False by default.
    if ($UpdateSecurity -eq $Null -or $UpdateSecurity -eq '')
    {
         $UpdateSecurity = 'False'
    }
}

Function VerifyUpdates
{
    # The de-facto "CONFIRM" and "ARE YOU SURE" function used to review and proceed with script execution.
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
    Write-Host "Updating mailbox permissions."

    # Get the members from the security group.
    $DISTRIBUTION_LIST = Get-DistributionGroupMember $Group | Select-Object -ExpandProperty Name

    # Get the current members from the assigned to the mailbox.
    $MAILBOX_PERMISSION_LIST = Get-MailboxPermission -Identity $Mailbox | ?{$_.User.ToString() -ne "NT AUTHORITY\SELF" -and $_.IsInherited -eq $false}

    # Remove old mailbox permissions.
    Write-Host "  Removing outdated 'Full Access' permissions."
    ForEach ($ObjFound in $MAILBOX_PERMISSION_LIST)
    {
        Remove-MailboxPermission -Identity $Mailbox -User $ObjFound.User -AccessRights FullAccess -Confirm:$false
    }

    Write-Host "  Adding updated 'Full Access' permissions."
    foreach ($USER in $DISTRIBUTION_LIST)
    {
        $ACE = Get-MailboxPermission -Identity $Mailbox -User $USER | Select-Object -ExpandProperty AccessRights
        IF ($ACE -ne 'FullAccess')
        {
             # Add the "FullAccess" permissions for each security group member to the group mailbox.
    	     Add-MailboxPermission -Identity $Mailbox -User $USER -AccessRights FullAccess -InheritanceType all -AutoMapping $true
        }
    }

    # Set the "Send As" permission to the security group.
    $ACE = Get-RecipientPermission -Identity $Mailbox -Trustee $Group | Select-Object -ExpandProperty AccessRights
    if($ACE -ne 'SendAs')
    {
        Write-Host "  Updating the 'Send-As' permissions."
        Add-RecipientPermission -Identity $Mailbox -AccessRights SendAs -Trustee $Group -Confirm:$false
    }

    $GMB = [bool](Get-Mailbox $Mailbox -RecipientTypeDetails UserMailbox, LegacyMailbox, LinkedMailbox, GroupMailbox, RoomMailbox, EquipmentMailbox -ErrorAction 'SilentlyContinue')
    IF ($GMB -eq 'True')
    {
        # Remove and reset the "Send on Behalf" permission to the security group.
        Write-Host "  Updating 'Send on Behalf' permissions."
        Get-Mailbox -Identity $Mailbox | Set-Mailbox -GrantSendOnBehalfTo $null -ErrorAction 'SilentlyContinue'
        Get-Mailbox -Identity $Mailbox | Set-Mailbox -GrantSendOnBehalfTo $Group -ErrorAction 'SilentlyContinue'
    }
}

Function ConvertShared
{
    Write-Host "Converting the group mailbox to a shared mailbox."

    # Verify the mailbox type and convert if it is not a "Shared" mailbox.
    Write-Host "  Verifying the mailbox type."
    $SMB = [bool](Get-Mailbox $Mailbox -RecipientTypeDetails SharedMailbox -ErrorAction 'SilentlyContinue')
    IF ($SMB -eq $False)
    {
        Write-Host "  Converting the mailbox to a shared mailbox."
        Set-Mailbox -Identity $Mailbox -Type Share
    }
    ELSE
    {
        Write-Host "  Mailbox is already a shared mailbox, not converting."
    }
}

### Main Script

# If not already loaded, load the Exchange server PowerShell modules.
if ( (Get-PSSnapin -Name *Exchange* -ErrorAction SilentlyContinue) -eq $null )
{
    Add-PSSnapin *Exchange*
}


Write-Host -ForegroundColor White "Set-GroupMailbox"
Write-Host ""

ErrorChecking
VerifyUpdates
IF ($Convert -eq 'True')
{
    ConvertShared
}
UpdateMailboxSecurity

# Exit Script
exit 0
