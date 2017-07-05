Import-Module -Force 'C:\cron\creds.psm1';
$ErrorActionPreference = "Stop";

Function Log {
   Param ([string]$logstring)
   $output = $("{0} ({1} - {2}): {3}" -f $(Get-Date), $(GCI $MyInvocation.PSCommandPath | Select -Expand Name), $pid, $logstring);
   Write-Host $output;
   Add-content "C:\cron\cloud_wrangler.log" -value $output;
}

# get all of the mailboxes from 365
$mailboxes = Invoke-command -session $session -Command { Get-Mailbox -ResultSize unlimited };

# get all of the users from 365
$o365_users = Get-MsolUser -All;
$o365_map = @{};
ForEach ($user in $o365_users) {
    $o365_map[$user.userprincipalname] = $user;
}
$o365_updated = $false

# Rig the UPN for each user account so that it matches the primary SMTP address.
ForEach ($mb in ($mailboxes | Where {$_.userprincipalname -ne $_.primarysmtpaddress})) {
    Log $("Changing UPN from {0} to {1}" -f $mb.UserPrincipalName,$mb.PrimarySmtpAddress);
    Set-MsolUserPrincipalName -UserPrincipalName $mb.UserPrincipalName -NewUserPrincipalName $mb.PrimarySmtpAddress -Verbose;
    $o365_updated = $true;
}


if ($o365_updated) {
    $mailboxes = Invoke-command -session $session -Command { Get-Mailbox -ResultSize unlimited };
    $o365_users = Get-MsolUser -All;
}



# Read the full user DB from OIM CMS (all DepartmentUser objects) via the OIM CMS API.
# NOTE: $user_api is set in C:\cron\creds.psm1
$users = Invoke-RestMethod ("{0}?all" -f $user_api) -WebSession $oimsession;
# Deserialise response into JSON (bypass the MaxJsonLength property of 2 MB).
if (-not $users.objects) {
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions");
    $json = New-Object -TypeName System.Web.Script.Serialization.JavaScriptSerializer;
    $json.MaxJsonLength = 104857600;
    $users = $json.Deserialize($users, [System.Object]);
}
$cmsusers_updated = $false;




# set archiving for all cloud mailboxes which don't have it
$non_archive = $mailboxes | where {$_.ArchiveState -eq 'None'} | where {$_.userprincipalname -notin $o365_users.UserPrincipalName};
ForEach ($mb in $non_archive) {
    $email = $mb.PrimarySmtpAddress;
    Log $("Adding archive mailbox for {0}" -f $email);
    Invoke-command -session $session -ScriptBlock $([ScriptBlock]::Create("Enable-Mailbox -Identity $email -Archive"));
}

# set auditing for all mailboxes which don't have it
$non_audit = $mailboxes | where {-not $_.AuditEnabled};
ForEach ($mb in $non_audit) {
    $email = $mb.PrimarySmtpAddress;
    Log $("Adding access audit rules for {0}" -f $email);
    Invoke-command -session $session -ScriptBlock $([ScriptBlock]::Create("Set-Mailbox -Identity $email -AuditEnabled $true -AuditAdmin 'SendAs' -AuditDelegate 'SendAs' -AuditOwner 'MailboxLogin' "));
}


$untracked_users = $o365_users | where {$_.userprincipalname -notin $users.objects.email};

ForEach ($user in $untracked_users) {
    $simpleuser =  $user |   select @{name='ObjectGUID'; expression={$null}},
                                    @{name='EmailAddress'; expression={$_.userprincipalname}}, 
                                    @{name='DistinguishedName'; expression={$null}}, 
                                    @{name='SamAccountName'; expression={$_.userprincipalname.split('@')[0].replace('.','').replace('#', '').replace(',', '')}}, 
                                    @{name='AccountExpirationDate'; expression={$null}}, 
                                    @{name='Enabled'; expression={$false}},
                                    @{name='DisplayName'; expression={$_.DisplayName}}, 
                                    @{name='Title'; expression={$_.Title}}, 
                                    @{name='GivenName'; expression={$_.firstname}}, 
                                    @{name='Surname'; expression={$_.lastname}}, 
                                    @{name='Modified'; expression={$null}}
                                    ;

    # For every push to the API, we need to explicitly convert to UTF8 bytes
    # to avoid the stupid moon-man encoding Windows uses for strings.
    # Without this, e.g. users with Unicode names will fail as the server expects UTF8.
    $userjson = [System.Text.Encoding]::UTF8.GetBytes($($simpleuser | ConvertTo-Json));
    # Here we POST to the API endpoint to create a new DepartmentUser in the CMS.
    try {
        Log $("Creating a new OIM CMS object for {0}" -f $simpleuser.EmailAddress);
        Write-Host $("Creating a new OIM CMS object for {0} ({1})" -f $simpleuser.EmailAddress, $simpleuser.SamAccountName);
        $response = Invoke-RestMethod $user_api -Body $userjson -Method Post -ContentType "application/json" -Verbose -WebSession $oimsession;
        # Note that a change has occurred.
        $cmsusers_updated = $true;
    } catch [System.Exception] {
        # Log any failures to sync AD data into the OIM CMS, for reference.
        Log $("ERROR: creating new OIM CMS user failed for {0}" -f $simpleuser.EmailAddress);
        Log $_.Exception.ToString();
        Log $($userjson);
    }
}
