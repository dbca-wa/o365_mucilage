Import-Module -Force 'C:\cron\creds.psm1'
$ErrorActionPreference = "Stop"

Function Log {
   Param ([string]$logstring)
   Add-content "C:\cron\group_wrangler.log" -value $("{0} ({1} - {2}): {3}" -f $(Get-Date), $(GCI $MyInvocation.PSCommandPath | Select -Expand Name), $pid, $logstring)
}

# download distribution group list from Exchange Online
$dgrps = Invoke-command -session $session -Command { Get-DistributionGroup -ResultSize unlimited }

Log $("Processing {0} groups" -f $dgrps.Length)

try {
    # filter the Exchange Online groups list for groups that are "In cloud", and
    # don't include the above org-derived groups (i.e. Alias doesn't have db- prefix)
    $msolgroups = $dgrps | where isdirsynced -eq $false | where Alias -notlike db-* | select @{name="members";expression={Get-MsolGroupMember -All -GroupObjectId $_.ExternalDirectoryObjectId}}, *
    Log $("Syncing {0} MailSecurity groups" -f $msolgroups.length)
    
    # fetch all the shadow groups from the local AD
    $localgroups = Get-DistributionGroup -OrganizationalUnit $mail_security_ou -ResultSize Unlimited
    # delete any shadow groups where the original is no longer online
    $localgroups | select @{name='exists';expression={$msolgroups | where ExternalDirectoryObjectId -like $_.CustomAttribute1}}, Identity | where exists -eq $null | Remove-DistributionGroup -BypassSecurityGroupManagerCheck -Confirm:$false
    # create/update the rest of the shadow groups to match Exchange Online
    ForEach ($msolgroup in $msolgroups) {
        $group = $localgroups | where CustomAttribute1 -like $msolgroup.ExternalDirectoryObjectId
        $name = $msolgroup.Alias.replace("`r", "").replace("`n", " ").TrimEnd();
        if (-not $group) { 
            $group = New-DistributionGroup -OrganizationalUnit $mail_security_ou -PrimarySmtpAddress $msolgroup.PrimarySmtpAddress -Name $name -Type Security 
        }
        Set-DistributionGroup $group -Name $msolgroup.Alias -CustomAttribute1 $msolgroup.ExternalDirectoryObjectId -Alias $msolgroup.Alias -DisplayName $msolgroup.DisplayName -PrimarySmtpAddress $msolgroup.PrimarySmtpAddress
        Update-DistributionGroupMember $group -Members $msolgroup.members.EmailAddress  -BypassSecurityGroupManagerCheck -Confirm:$false
        
        # we need to ensure an admin group is added as an owner the O365 group object. 
        # why? because without this, even God Mode Exchange admins can't use ECP to manage the group owners! 
        # instead they have to open PowerShell and do the exact same thing... which will fail until 
        # you add -BypassSecurityGroupManagerCheck, which all of a sudden makes it totally okay! 
        #$gsmtp = $msolgroup.PrimarySmtpAddress;
        #Invoke-command -session $session -ScriptBlock $([ScriptBlock]::Create("Set-DistributionGroup -Identity $gsmtp -ManagedBy @{Add=`"$admin_msolgroup`"} -BypassSecurityGroupManagerCheck"))
        
    }
} catch [System.Exception] {
    Log "ERROR: Exception caught, skipping rest of MailSecurity"
    Log $($_ | convertto-json)
}

Log "Finished"

# cleanup
Get-PSSession | Remove-PSSession
