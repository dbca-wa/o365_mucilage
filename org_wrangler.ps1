Import-Module -Force 'C:\cron\creds.psm1';
$ErrorActionPreference = "Stop";

Function Log {
   Param ([string]$logstring)
   Add-content "C:\cron\org_wrangler.log" -value $("{0} ({1} - {2}): {3}" -f $(Get-Date), $(GCI $MyInvocation.PSCommandPath | Select -Expand Name), $pid, $logstring);
}

# download distribution group list from Exchange Online
$dgrps = Invoke-command -session $session -Command { Get-DistributionGroup -ResultSize unlimited };

Function smash-groups {
    param([Object[]]$grps, [Object[]]$localgroups, [String]$ou)
    # loop through all the groups taken from the OIM CMS org structure
    foreach ($grp in $grps) {
        $group, $ogroup, $diff = $null, $null, @();

        # find the equivalent group in AD
        $group = $localgroups | where Alias -like $grp.id;

        # find the equivalent group in Exchange Online
        $ogroup = $dgrps | where Alias -like $grp.id;

        # get the group's name, clean it up a bit for AD
        $name = $grp.name.Substring(0,[System.Math]::Min(64, $grp.name.Length)).replace("`r", "").replace("`n", " ").TrimEnd();
        $id, $email, $dname, $owner = $grp.id, $grp.email, $grp.name, $grp.owner;
        
        # create Outlook Online group, if it doesn't exist
        if (-not $ogroup) { 
            try {
                $ogroup = Invoke-command -session $session -ScriptBlock $([ScriptBlock]::Create("New-DistributionGroup -Alias $id -PrimarySmtpAddress $email -Name `"$name`" -Type Security"));
            } catch [System.Exception] {
                # bump email and try again
                $email = $email.split("@")[0]+"-orgunit@"+$email.split("@")[1];
                $ogroup = Invoke-command -session $session -ScriptBlock $([ScriptBlock]::Create("New-DistributionGroup -Alias $id -PrimarySmtpAddress $email -Name `"$name`" -Type Security"));
            }
        }

        # create AD group, if it doesn't exist
        if (-not $group) { 
            $group = New-DistributionGroup -OrganizationalUnit $ou -Alias $id -PrimarySmtpAddress $email -Name $name -Type Security;
        }

        # bail out if either of those operations return $null
        if (-not ($group -and $ogroup)) { 
            continue;
        }

        # set most of the attributes of the AD and Outlook Online groups to match OIM CMS
        $mtip = "Please contact the Office for Information Management (OIM) to correct membership information for this group.";
        try {
            Invoke-command -session $session -ScriptBlock $([ScriptBlock]::Create("Set-DistributionGroup `"$id`" -Name `"$name`" -DisplayName `"$dname`" -PrimarySmtpAddress `"$email`" -ManagedBy `"$owner`" -MailTip `"$mtip`" -BypassSecurityGroupManagerCheck -Confirm:`$false"));
        } catch [System.Exception] {
            # bump email and try again
            $email = $email.split("@")[0]+"-orgunit@"+$email.split("@")[1];
            Invoke-command -session $session -ScriptBlock $([ScriptBlock]::Create("Set-DistributionGroup `"$id`" -Name `"$name`" -DisplayName `"$dname`" -PrimarySmtpAddress `"$email`" -ManagedBy `"$owner`" -MailTip `"$mtip`" -BypassSecurityGroupManagerCheck -Confirm:`$false"));
        }
        Set-DistributionGroup $id -Name $name -DisplayName $dname -PrimarySmtpAddress $email -ManagedBy $owner -MailTip $mtip  -BypassSecurityGroupManagerCheck -Confirm:$false;
        
        # check for a change in group membership, before doing the expensive update group members operation
        try { 
            $diff = compare $((Get-DistributionGroupMember $group -ResultSize Unlimited).primarysmtpaddress | foreach { $([string]$_).toLower() }) $($grp.members | foreach { $_.toLower() }) -PassThru;
            if ($diff.Length -eq 0) { 
                continue;
            } 
        } catch [System.Exception] { 
            $diff = $grp.members;
        }
        Log $("Updating {3}/{2} in {0} managed by {1}" -f $group, $owner, $($grp.members.length), $($diff.Length));
        if ($diff.Length -lt 5) { 
            Log $($diff | convertto-json);
        }

        # update members of the Outlook Online and AD groups to match OIM CMS
        Invoke-command -session $session -ScriptBlock $([ScriptBlock]::Create("Update-DistributionGroupMember `"$id`" -Members `"$($grp.members -join '","')`" -BypassSecurityGroupManagerCheck -Confirm:`$false"));
        Update-DistributionGroupMember $id -Members $grp.members  -BypassSecurityGroupManagerCheck -Confirm:$false;
    }
}

# Start with a result of success
$result = "Success";
$result | Out-File "C:\cron\org_wrangler_result.txt";

# download org structure as JSON from OIM CMS
$org_structure = Invoke-RestMethod ("{0}?org_structure=true&sync_o365=true" -f $user_api) -WebSession $oimsession;


# update org unit groups
try {
    $orgunits = $org_structure.objects | where id -like "db-org_*" | where email -like "*@*";
    Log $("Loading {0} OrgUnit groups..." -f $orgunits.length);
    $localgroups = Get-DistributionGroup -OrganizationalUnit $org_unit_ou -ResultSize Unlimited;
    smash-groups -grps $orgunits -localgroups $localgroups -ou $org_unit_ou;
} catch [System.Exception] {
    Log "ERROR: Exception caught, skipping rest of OrgUnit";
    Log $($_ | convertto-json);
    $result = "Failure"
    $result | Out-File "C:\cron\org_wrangler_result.txt"
}

# update cost centre groups
try {
    $costcentres = $org_structure.objects | where id -like "db-cc_*" | where email -like "*@*";
    Log $("Loading {0} CostCentre groups..." -f $costcentres.length);
    $localgroups = Get-DistributionGroup -OrganizationalUnit $cost_centre_ou -ResultSize Unlimited;
    smash-groups -grps $costcentres -localgroups $localgroups -ou $cost_centre_ou;
} catch [System.Exception] {
    Log "ERROR: Exception caught, skipping rest of CostCentre";
    Log $($_ | convertto-json);
    $result = "Failure"
    $result | Out-File "C:\cron\org_wrangler_result.txt"
}

# update location groups
try {
    $locations = $org_structure.objects | where id -like "db-loc*_*" | where email -like "*@*";
    Log $("Loading {0} Location groups" -f $locations.length);
    $localgroups = Get-DistributionGroup -OrganizationalUnit $location_ou -ResultSize Unlimited;
    smash-groups -grps $locations -localgroups $localgroups -ou $location_ou;
} catch [System.Exception] {
    Log "ERROR: Exception caught, skipping rest of Location";
    Log $($_ | convertto-json);
    $result = "Failure"
    $result | Out-File "C:\cron\org_wrangler_result.txt"
}

# cache org structure
$org_structure | convertto-json > C:\cron\org_structure.json;

Log "Finished";
$result = "Hello world"
$result | Out-File "C:\cron\org_wrangler_result.txt"

# cleanup
Get-PSSession | Remove-PSSession;