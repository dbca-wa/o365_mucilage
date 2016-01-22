Import-Module -Force 'C:\cron\creds.psm1'
$ErrorActionPreference = "Stop"

Function Log {
   Param ([string]$logstring)
   Add-content "C:\cron\group_wrangler.log" -value $("{0} ({1} - {2}): {3}" -f $(Get-Date), $(GCI $MyInvocation.PSCommandPath | Select -Expand Name), $pid, $logstring)
}

try {
    $org_structure = Invoke-RestMethod ("{0}?org_structure" -f $user_api)
    $dgrps = Invoke-command -session $session -Command { Get-DistributionGroup -ResultSize unlimited }
    Log $("Processing {0} groups" -f $dgrps.Length)

    function smash-groups {
        param([Object[]]$grps, [Object[]]$localgroups, [String]$ou)
        foreach ($grp in $grps) {
            $group, $ogroup, $diff = $null, $null, @()
            $group = $localgroups | where Alias -like $grp.id
            $ogroup = $dgrps | where Alias -like $grp.id
            $name, $id, $email, $dname, $owner = $grp.name.Substring(0,[System.Math]::Min(64, $grp.name.Length)), $grp.id, $grp.email, $grp.name, $grp.owner;
            $name = $name.replace("`r", "").replace("`n", " ").trim();
            if (-not $group) { $group = New-DistributionGroup -OrganizationalUnit $ou -Alias $id -PrimarySmtpAddress $email -Name $name -Type Security }
            if (-not $ogroup) { $ogroup = Invoke-command -session $session -ScriptBlock $([ScriptBlock]::Create("New-DistributionGroup -Alias $id -PrimarySmtpAddress $email -Name `"$name`" -Type Security")) }
            if (-not ($group -and $ogroup)) { continue }
            $mtip = "Please contact the Office for Information Management (OIM) to correct membership information for this group."
            Set-DistributionGroup $group -Name $name -DisplayName $dname -PrimarySmtpAddress $email -ManagedBy $owner -MailTip $mtip  -BypassSecurityGroupManagerCheck -Confirm:$false
            Invoke-command -session $session -ScriptBlock $([ScriptBlock]::Create("Set-DistributionGroup `"$ogroup`" -Name `"$name`" -DisplayName `"$dname`" -PrimarySmtpAddress `"$email`" -ManagedBy `"$owner`" -MailTip `"$mtip`" -BypassSecurityGroupManagerCheck -Confirm:`$false"))
            try { $diff = compare $((Get-DistributionGroupMember $group -ResultSize Unlimited).primarysmtpaddress | foreach { $([string]$_).toLower() }) $($grp.members | foreach { $_.toLower() }) -PassThru
                if ($diff.Length -eq 0) { continue } } catch [System.Exception] { $diff = $grp.members }
            Log $("Updating {3}/{2} in {0} managed by {1}" -f $group, $owner, $($grp.members.length), $($diff.Length))
            if ($diff.Length -lt 5) { Log $($diff | convertto-json) }
            Invoke-command -session $session -ScriptBlock $([ScriptBlock]::Create("Update-DistributionGroupMember `"$ogroup`" -Members `"$($grp.members -join '","')`" -BypassSecurityGroupManagerCheck -Confirm:`$false"))
            Update-DistributionGroupMember $group -Members $grp.members  -BypassSecurityGroupManagerCheck -Confirm:$false
        }
    }

    $orgunits = $org_structure.objects | where id -like "db-org_*" | where email -like "*@*"
    $localgroups = Get-DistributionGroup -OrganizationalUnit 'corporateict.domain/Groups/OrgUnit' -ResultSize Unlimited
    smash-groups -grps $orgunits -localgroups $localgroups -ou 'corporateict.domain/Groups/OrgUnit'
    $costcentres = $org_structure.objects | where id -like "db-cc_*" | where email -like "*@*"
    $localgroups = Get-DistributionGroup -OrganizationalUnit 'corporateict.domain/Groups/CostCentre' -ResultSize Unlimited
    smash-groups -grps $costcentres -localgroups $localgroups -ou 'corporateict.domain/Groups/CostCentre'
    $locations = $org_structure.objects | where id -like "db-loc*_*" | where email -like "*@*"
    $localgroups = Get-DistributionGroup -OrganizationalUnit 'corporateict.domain/Groups/Location' -ResultSize Unlimited
    smash-groups -grps $locations -localgroups $localgroups -ou 'corporateict.domain/Groups/Location'
    $org_structure | convertto-json > C:\cron\org_structure.json

    $msolgroups = $dgrps | where isdirsynced -eq $false | where Alias -notlike db-* | select @{name="members";expression={Get-MsolGroupMember -All -GroupObjectId $_.ExternalDirectoryObjectId}}, *
    $localgroups = Get-DistributionGroup -OrganizationalUnit 'corporateict.domain/Groups/MailSecurity' -ResultSize Unlimited
    # delete groups not online anymore
    $localgroups | select @{name='exists';expression={$msolgroups | where ExternalDirectoryObjectId -like $_.CustomAttribute1}}, Identity | where exists -eq $null | Remove-DistributionGroup -BypassSecurityGroupManagerCheck -Confirm:$false
    # create/update rest
    ForEach ($msolgroup in $msolgroups) {
        $group = $localgroups | where CustomAttribute1 -like $msolgroup.ExternalDirectoryObjectId
        $name = $msolgroup.Alias.replace("`r", "").replace("`n", " ");
        if (-not $group) { $group = New-DistributionGroup -OrganizationalUnit 'corporateict.domain/Groups/MailSecurity' -PrimarySmtpAddress $msolgroup.PrimarySmtpAddress -Name $name -Type Security }
        Set-DistributionGroup $group -Name $msolgroup.Alias -CustomAttribute1 $msolgroup.ExternalDirectoryObjectId -Alias $msolgroup.Alias -DisplayName $msolgroup.DisplayName -PrimarySmtpAddress $msolgroup.PrimarySmtpAddress
        Update-DistributionGroupMember $group -Members $msolgroup.members.EmailAddress  -BypassSecurityGroupManagerCheck -Confirm:$false
    }
    Log "Finished"
} catch [System.Exception] {
    Log "ERROR: Exception caught, dying =("
    Log $($_ | convertto-json)
}

# cleanup
Get-PSSession | Remove-PSSession
