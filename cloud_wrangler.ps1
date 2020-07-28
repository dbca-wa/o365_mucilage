﻿Import-Module -Force 'C:\cron\creds.psm1';
$ErrorActionPreference = "Stop";

Function Log {
   Param ([string]$logstring)
   $output = $("{0} ({1} - {2}): {3}" -f $(Get-Date), $(GCI $MyInvocation.PSCommandPath | Select -Expand Name), $pid, $logstring);
   Write-Host $output;
   Add-content "C:\cron\cloud_wrangler.log" -value $output;
}

Log "Loading cloud information...";

# get all of the mailboxes from 365
$mailboxes = Invoke-command -session $session -Command { Get-Mailbox -ResultSize unlimited };

# get all of the users from 365
$o365_users = Get-MsolUser -All | where {[string]$_.licenses.accountskuid -like "*ENTERPRISEPREMIUM*"};
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

# Enforce MFA for all synced users
$mfa_exclude = Get-MsolGroupMember -GroupObjectId $mfa_exclude;

$mfaauth = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
$mfaauth.RelyingParty = "*"
$mfaauth.State = "Enforced"

$users_mfa_off = $o365_users | where { ($_.StrongAuthenticationRequirements.State -eq "Enforced") -and ($mfa_exclude.objectid -contains $_.objectid) };
$users_mfa_on = $o365_users | where { ($_.StrongAuthenticationRequirements.State -ne "Enforced") -and ($mfa_exclude.objectid -notcontains $_.objectid)};

ForEach ($user in $users_mfa_on) {
    try {
        Log $("Enforcing MFA for {0}" -f $user.UserPrincipalName)
        Set-MsolUser -UserPrincipalName $user.UserPrincipalName -StrongAuthenticationRequirements $mfaauth;
    } catch [System.Exception] {
        Log "ERROR: Couldn't run Set-MsolUser";
        $except = $_;
        Log $($except);#| convertto-json);
    }
}
ForEach ($user in $users_mfa_off) {
    try {
        Log $("Disabling MFA for {0}" -f $user.UserPrincipalName)
        Set-MsolUser -UserPrincipalName $user.UserPrincipalName -StrongAuthenticationRequirements @();
    } catch [System.Exception] {
        Log "ERROR: Couldn't run Set-MsolUser";
        $except = $_;
        Log $($except);#| convertto-json);
    }
}


$cloud_only = $o365_users | where {-not $_.LastDirSyncTime};


# Read the full user DB from OIM CMS (all DepartmentUser objects) via the OIM CMS API.
# NOTE: $user_api is set in C:\cron\creds.psm1
$users = Invoke-RestMethod ("{0}?all" -f $user_api) -WebSession $oimsession -TimeoutSec 300;
# Deserialise response into JSON (bypass the MaxJsonLength property of 2 MB).
if (-not $users.objects) {
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions");
    $json = New-Object -TypeName System.Web.Script.Serialization.JavaScriptSerializer;
    $json.MaxJsonLength = 104857600;
    $users = $json.Deserialize($users, [System.Object]);
}
$cmsusers_updated = $false;




# set archiving for all cloud mailboxes which don't have it
$non_archive = $mailboxes | where {$_.ArchiveStatus -eq 'None'} | where {$_.userprincipalname -in $cloud_only.UserPrincipalName} | where {-not $_.managedfoldermailboxpolicy};
ForEach ($mb in $non_archive) {
    try {
        $email = $mb.PrimarySmtpAddress;
        Log $("Adding archive mailbox for {0}" -f $email);
        Invoke-command -session $session -ScriptBlock $([ScriptBlock]::Create("Enable-Mailbox -Identity `"$email`" -Archive"));
    } catch [System.Exception] {
        Log "ERROR: Couldn't run Enable-Mailbox";
        $except = $_;
        Log $($except);
    }
}

# set auditing for all mailboxes which don't have it
$non_audit = $mailboxes | where {-not $_.AuditEnabled};
ForEach ($mb in $non_audit) {
    try {
        $email = $mb.PrimarySmtpAddress;
        Log $("Adding access audit rules for {0}" -f $email);
        Invoke-command -session $session -ScriptBlock $([ScriptBlock]::Create("Set-Mailbox -Identity `"$email`" -AuditEnabled `$true -AuditAdmin 'SendAs' -AuditDelegate 'SendAs' -AuditOwner 'MailboxLogin' "));
    } catch [System.Exception] {
        Log "ERROR: Couldn't enable acces audit rules";
        $except = $_;
        Log $($except);
    }
}

# get exclusions for litigation hold (required for programmatic mailboxes with a high throughput)
$lithold_exclude = Get-MsolGroupMember -GroupObjectId $lithold_exclude;



# set litigation hold for all mailboxes which don't have it
$lithold_on = $mailboxes | where {($_.RecipientTypeDetails -eq "UserMailbox") -and (-not $_.LitigationHoldEnabled) -and ($lithold_exclude.objectid -notcontains $_.externaldirectoryobjectid)};
$lithold_off = $mailboxes | where {($_.RecipientTypeDetails -eq "UserMailbox") -and ($_.LitigationHoldEnabled) -and ($lithold_exclude.objectid -contains $_.externaldirectoryobjectid)};
ForEach ($mb in $lithold_on) {
    try {
        $email = $mb.PrimarySmtpAddress;
        Log $("Adding litigation hold rule for {0}" -f $email);
        Invoke-command -session $session -ScriptBlock $([ScriptBlock]::Create("Set-Mailbox -Identity `"$email`" -LitigationHoldEnabled `$true"));
    } catch [System.Exception] {
        Log "ERROR: Couldn't enable litigation hold";
        $except = $_;
        Log $($except);
    }
}
ForEach ($mb in $lithold_off) {
    try {
        $email = $mb.PrimarySmtpAddress;
        Log $("Removing litigation hold rule for {0}" -f $email);
        Invoke-command -session $session -ScriptBlock $([ScriptBlock]::Create("Set-Mailbox -Identity `"$email`" -LitigationHoldEnabled `$false"));
    } catch [System.Exception] {
        Log "ERROR: Couldn't disable litigation hold";
        $except = $_;
        Log $($except);
    }
}



$untracked_users = $o365_users | where {$_.userprincipalname -notin $users.objects.email};

ForEach ($user in $untracked_users) {
    $simpleuser =  $user |   select @{name='ObjectGUID'; expression={$null}},
                                    @{name='azure_guid'; expression={$_.ObjectId.Guid}},
                                    @{name='EmailAddress'; expression={$_.userprincipalname}}, 
                                    @{name='DistinguishedName'; expression={$null}}, 
                                    @{name='SamAccountName'; expression={$_.userprincipalname.split('@')[0].replace('.','').replace("'", '').replace('#', '').replace(',', '')}}, 
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
        #Write-Host $("Creating a new OIM CMS object for {0} ({1})" -f $simpleuser.EmailAddress, $simpleuser.SamAccountName);
        $response = Invoke-RestMethod $user_api -Body $userjson -Method Post -ContentType "application/json" -Verbose -WebSession $oimsession;
        # Note that a change has occurred.
        $cmsusers_updated = $true;
    } catch [System.Exception] {
        # Log any failures to sync O365 data into the OIM CMS, for reference.
        Log $("ERROR: creating new OIM CMS user failed for {0}" -f $simpleuser.EmailAddress);
        Log $("Endpoint: {0}" -f $user_api);
        Log $("Payload: {0}" -f $simpleuser | ConvertTo-Json);
        $result = $_.Exception.Response.GetResponseStream();
        $reader = New-Object System.IO.StreamReader($result);
        $reader.BaseStream.Position = 0;
        $reader.DiscardBufferedData();
        $responseBody = $reader.ReadToEnd();
        Log $("Response: {0}" -f $responseBody);   
    }
}

Log "Finished";