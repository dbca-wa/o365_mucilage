Import-Module -Force 'C:\cron\creds.psm1';
$ErrorActionPreference = "Stop";

Function Log {
   Param ([string]$logstring)
   $output = $("{0} ({1} - {2}): {3}" -f $(Get-Date), $(GCI $MyInvocation.PSCommandPath | Select -Expand Name), $pid, $logstring);
   Write-Host $output;
   Add-content "C:\cron\org_wrangler.log" -value $output;
}

# download distribution group list from Exchange Online
$dgrps = Invoke-command -session $session -Command { Get-DistributionGroup -ResultSize unlimited -Filter "(Alias -like 'db-*')" };

# download user list from Office 365
$users = Get-MsolUser -All;

Function smash-groups {
    param([Object[]]$grps)
    # loop through all the groups taken from the OIM CMS org structure
    foreach ($grp in $grps) {
        $ogroup, $diff = $null, $null, @();

        # find the equivalent group in Exchange Online
        $ogroup = $dgrps | where Alias -like $grp.id;

        # get the group's name, clean it up a bit for AD
        $name = $grp.name.Substring(0,[System.Math]::Min(64, $grp.name.Length)).replace("`r", "").replace("`n", " ").TrimEnd();
        $id, $email, $dname, $owner = $grp.id, $grp.email, $grp.name, $grp.owner;
        $email_sec = $email.split("@")[0]+"-orgunit@"+$email.split("@")[1];

        If ((-not $owner) -or ($owner -eq 'support@dbca.wa.gov.au')) {
            $owner = "`"admin@dbca.wa.gov.au`"";
        } else {
            $owner = "`"$owner`""
        }

        # hack to avoid updating groups for domains in transit
        $domain = $email.Split('@', 2)[1];
        If ($domain -in $domain_skip) {
            Log $('Skipping group {0} due to domain {1}' -f $dname,$domain)
            Continue;
        }

        # check if the owner is in the O365 directory
        # FIXME: support@dpaw.wa.gov.au isn't a UPN? which then breaks this
        #If ($owner -notin $users.UserPrincipalName) {
        #    Log $("Skipping {0}, couldn't find ManagedBy in O365:  {1}" -f $email,$owner);
        #    continue;
        #}

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

        # bail out if either of those operations return $null
        if (-not $ogroup) { 
            continue;
        }

        # set most of the attributes of the AD and Outlook Online groups to match OIM CMS
        $mtip = "Please contact the Office for Information Management (OIM) to correct membership information for this group.";

        If (($ogroup.alias -ne $id) -or -not (($ogroup.primarysmtpaddress -eq $email) -or ($ogroup.primarysmtpaddress -eq $email_sec)) -or ($ogroup.displayname -ne $dname)) {

            try {
                Log "Set-DistributionGroup `"$id`" -DisplayName `"$dname`" -PrimarySmtpAddress `"$email`" -ManagedBy $owner -MailTip `"$mtip`" -BypassSecurityGroupManagerCheck -Confirm:`$false";
                Invoke-command -session $session -ScriptBlock $([ScriptBlock]::Create("Set-DistributionGroup `"$id`" -DisplayName `"$dname`" -PrimarySmtpAddress `"$email`" -ManagedBy $owner -MailTip `"$mtip`" -BypassSecurityGroupManagerCheck -Confirm:`$false"));
            } catch [System.Exception] {
                # bump email and try again
                Log "Set-DistributionGroup `"$id`" -DisplayName `"$dname`" -PrimarySmtpAddress `"$email_sec`" -ManagedBy $owner -MailTip `"$mtip`" -BypassSecurityGroupManagerCheck -Confirm:`$false";
                try {
                    Invoke-command -session $session -ScriptBlock $([ScriptBlock]::Create( "Set-DistributionGroup `"$id`" -DisplayName `"$dname`" -PrimarySmtpAddress `"$email_sec`" -ManagedBy $owner -MailTip `"$mtip`" -BypassSecurityGroupManagerCheck -Confirm:`$false"));
                } catch [System.Exception] {
                    Log "Failed to update distribution group, need to manually check!";
                }
            }
        }
        
        $missing = $grp.members | Where {$_ -notin $users.UserPrincipalName};
        If ($missing) {
             Log $("Couldn't find following DepartmentUsers in O365 for group {0}:  {1}" -f $email,$($missing -join ', '));
        }
        #Log $("Writing DepartmentUsers in O365 for group {0}:  {1}" -f $email,$($members -join ', '));
        $members = $grp.members | Where {$_ -in $users.UserPrincipalName};

        # update members of Outlook Online to match OIM CMS
        Invoke-command -session $session -ScriptBlock $([ScriptBlock]::Create("Update-DistributionGroupMember `"$id`" -Members `"$($members -join '","')`" -BypassSecurityGroupManagerCheck -Confirm:`$false"));
    }
}

# Start with a result of success
$result = "Success";
$result | Out-File "C:\cron\org_wrangler_result.txt";

# download org structure as JSON from OIM CMS
$org_structure = Invoke-RestMethod ("{0}?org_structure=true&sync_o365=true&populate_groups=true" -f $user_api) -WebSession $oimsession;


# update org unit groups
try {
    $orgunits = $org_structure.objects | where id -like "db-org_*" | where email -like "*@*";
    Log $("Loading {0} OrgUnit groups..." -f $orgunits.length);
    smash-groups -grps $orgunits;
} catch [System.Exception] {
    Log "ERROR: Exception caught, skipping rest of OrgUnit";
    Log $($_ | convertto-json);
    $result = "Failure"
    $result | Out-File "C:\cron\org_wrangler_result.txt"
}

# update location groups
try {
    $locations = $org_structure.objects | where id -like "db-loc*_*" | where email -like "*@*";
    Log $("Loading {0} Location groups" -f $locations.length);
    smash-groups -grps $locations;
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