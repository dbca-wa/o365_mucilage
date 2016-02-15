Import-Module -Force 'C:\cron\creds.psm1'
$ErrorActionPreference = "Stop"

Function Log {
   Param ([string]$logstring)
   Add-content "C:\cron\org_wrangler.log" -value $("{0} ({1} - {2}): {3}" -f $(Get-Date), $(GCI $MyInvocation.PSCommandPath | Select -Expand Name), $pid, $logstring)
}

# download distribution group list from Exchange Online
$dgrps = Invoke-command -session $session -Command { Get-DistributionGroup -ResultSize unlimited }

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


# update org unit groups
try {
    $orgunits = $org_structure.objects | where id -like "db-org_*" | where email -like "*@*"
    Log $("Loading {0} OrgUnit groups..." -f $orgunits.length)
    $localgroups = Get-DistributionGroup -OrganizationalUnit $org_unit_ou -ResultSize Unlimited
    smash-groups -grps $orgunits -localgroups $localgroups -ou $org_unit_ou
} catch [System.Exception] {
    Log "ERROR: Exception caught, skipping rest of OrgUnit"
    Log $($_ | convertto-json)
}

# update cost centre groups
try {
    $costcentres = $org_structure.objects | where id -like "db-cc_*" | where email -like "*@*"
    Log $("Loading {0} CostCentre groups..." -f $costcentres.length)
    $localgroups = Get-DistributionGroup -OrganizationalUnit $cost_centre_ou -ResultSize Unlimited
    smash-groups -grps $costcentres -localgroups $localgroups -ou $cost_centre_ou
} catch [System.Exception] {
    Log "ERROR: Exception caught, skipping rest of CostCentre"
    Log $($_ | convertto-json)
}

# update location groups
try {
    $locations = $org_structure.objects | where id -like "db-loc*_*" | where email -like "*@*"
    Log $("Loading {0} Location groups" -f $locations.length)
    $localgroups = Get-DistributionGroup -OrganizationalUnit $location_ou -ResultSize Unlimited
    smash-groups -grps $locations -localgroups $localgroups -ou $location_ou
} catch [System.Exception] {
    Log "ERROR: Exception caught, skipping rest of Location"
    Log $($_ | convertto-json)
}

# cache org structure
$org_structure | convertto-json > C:\cron\org_structure.json

Log "Finished"

# cleanup
Get-PSSession | Remove-PSSession