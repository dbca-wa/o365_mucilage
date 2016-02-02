Import-Module -Force 'C:\cron\creds.psm1'
$ErrorActionPreference = "Stop"

Function Log {
   Param ([string]$logstring)
   Add-content "C:\cron\group_wrangler.log" -value $("{0} ({1} - {2}): {3}" -f $(Get-Date), $(GCI $MyInvocation.PSCommandPath | Select -Expand Name), $pid, $logstring)
}

Function smash-groups {
    param([Object[]]$grps, [Object[]]$localgroups, [String]$ou)
    foreach ($grp in $grps) {
        $group, $ogroup, $diff = $null, $null, @()
        $group = $localgroups | where Alias -like $grp.id
        $ogroup = $dgrps | where Alias -like $grp.id
        $name = $grp.name.Substring(0,[System.Math]::Min(64, $grp.name.Length)).replace("`r", "").replace("`n", " ").TrimEnd();
        $id, $email, $dname, $owner = $grp.id, $grp.email, $grp.name, $grp.owner;
        
        if (-not $group) { 
            $group = New-DistributionGroup -OrganizationalUnit $ou -Alias $id -PrimarySmtpAddress $email -Name $name -Type Security 
        }
        if (-not $ogroup) { 
            $ogroup = Invoke-command -session $session -ScriptBlock $([ScriptBlock]::Create("New-DistributionGroup -Alias $id -PrimarySmtpAddress $email -Name `"$name`" -Type Security"))     
        }
        if (-not ($group -and $ogroup)) { 
            continue 
        }

        $mtip = "Please contact the Office for Information Management (OIM) to correct membership information for this group."
        Set-DistributionGroup $group -Name $name -DisplayName $dname -PrimarySmtpAddress $email -ManagedBy $owner -MailTip $mtip  -BypassSecurityGroupManagerCheck -Confirm:$false
        Invoke-command -session $session -ScriptBlock $([ScriptBlock]::Create("Set-DistributionGroup `"$ogroup`" -Name `"$name`" -DisplayName `"$dname`" -PrimarySmtpAddress `"$email`" -ManagedBy `"$owner`" -MailTip `"$mtip`" -BypassSecurityGroupManagerCheck -Confirm:`$false"))
        try { 
            $diff = compare $((Get-DistributionGroupMember $group -ResultSize Unlimited).primarysmtpaddress | foreach { $([string]$_).toLower() }) $($grp.members | foreach { $_.toLower() }) -PassThru
            if ($diff.Length -eq 0) { 
                continue 
            } 
        } catch [System.Exception] { 
            $diff = $grp.members 
        }
        Log $("Updating {3}/{2} in {0} managed by {1}" -f $group, $owner, $($grp.members.length), $($diff.Length))
        if ($diff.Length -lt 5) { 
            Log $($diff | convertto-json) 
        }
        Invoke-command -session $session -ScriptBlock $([ScriptBlock]::Create("Update-DistributionGroupMember `"$ogroup`" -Members `"$($grp.members -join '","')`" -BypassSecurityGroupManagerCheck -Confirm:`$false"))
        Update-DistributionGroupMember $group -Members $grp.members  -BypassSecurityGroupManagerCheck -Confirm:$false
    }
}

# download org structure as JSON from OIM CMS
$org_structure = Invoke-RestMethod ("{0}?org_structure" -f $user_api)
# download distribution group list from Exchange Online
$dgrps = Invoke-command -session $session -Command { Get-DistributionGroup -ResultSize unlimited }

Log $("Processing {0} groups" -f $dgrps.Length)

try {
    Log "Loading OrgUnit groups..."
    $orgunits = $org_structure.objects | where id -like "db-org_*" | where email -like "*@*"
    $localgroups = Get-DistributionGroup -OrganizationalUnit 'corporateict.domain/Groups/OrgUnit' -ResultSize Unlimited
    smash-groups -grps $orgunits -localgroups $localgroups -ou 'corporateict.domain/Groups/OrgUnit'
} catch [System.Exception] {
    Log "ERROR: Exception caught, skipping rest of OrgUnit"
    Log $($_ | convertto-json)
}

try {
    Log "Loading CostCentre groups..."
    $costcentres = $org_structure.objects | where id -like "db-cc_*" | where email -like "*@*"
    $localgroups = Get-DistributionGroup -OrganizationalUnit 'corporateict.domain/Groups/CostCentre' -ResultSize Unlimited
    smash-groups -grps $costcentres -localgroups $localgroups -ou 'corporateict.domain/Groups/CostCentre'
} catch [System.Exception] {
    Log "ERROR: Exception caught, skipping rest of CostCentre"
    Log $($_ | convertto-json)
}

try {
    Log "Loading Location groups"
    $locations = $org_structure.objects | where id -like "db-loc*_*" | where email -like "*@*"
    $localgroups = Get-DistributionGroup -OrganizationalUnit 'corporateict.domain/Groups/Location' -ResultSize Unlimited
    smash-groups -grps $locations -localgroups $localgroups -ou 'corporateict.domain/Groups/Location'
} catch [System.Exception] {
    Log "ERROR: Exception caught, skipping rest of Location"
    Log $($_ | convertto-json)
}

# cache org structure
$org_structure | convertto-json > C:\cron\org_structure.json

try {
    Log "Syncing MailSecurity groups"
    # filter the Exchange Online groups list for groups that are "In cloud", and
    # don't include the above org-derived groups (i.e. Alias doesn't have db- prefix)
    $msolgroups = $dgrps | where isdirsynced -eq $false | where Alias -notlike db-* | select @{name="members";expression={Get-MsolGroupMember -All -GroupObjectId $_.ExternalDirectoryObjectId}}, *
    $localgroups = Get-DistributionGroup -OrganizationalUnit 'corporateict.domain/Groups/MailSecurity' -ResultSize Unlimited
    # delete groups not online anymore
    $localgroups | select @{name='exists';expression={$msolgroups | where ExternalDirectoryObjectId -like $_.CustomAttribute1}}, Identity | where exists -eq $null | Remove-DistributionGroup -BypassSecurityGroupManagerCheck -Confirm:$false
    # create/update rest
    ForEach ($msolgroup in $msolgroups) {
        $group = $localgroups | where CustomAttribute1 -like $msolgroup.ExternalDirectoryObjectId
        $name = $msolgroup.Alias.replace("`r", "").replace("`n", " ").TrimEnd();
        if (-not $group) { 
            $group = New-DistributionGroup -OrganizationalUnit 'corporateict.domain/Groups/MailSecurity' -PrimarySmtpAddress $msolgroup.PrimarySmtpAddress -Name $name -Type Security 
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
