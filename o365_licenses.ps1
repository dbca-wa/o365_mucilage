Import-Module ActiveDirectory;
Import-Module -Force 'C:\cron\creds.psm1';
$ErrorActionPreference = "Stop";

Function Log {
   Param ([string]$logstring)
   Add-content "C:\cron\o365_licences.log" -value $("{0} ({1} - {2}): {3}" -f $(Get-Date), $(GCI $MyInvocation.PSCommandPath | Select -Expand Name), $pid, $logstring);
}

try {
    # Get a list of Office 365-licenced users.
    $licencedUsers = Get-MsolUser -All | Where {$_.isLicensed -eq "True" -and (("DPaW:ENTERPRISEPACK" -in $_.licenses.accountSkuId) -or ("DPaW:ENTERPRISEPREMIUM" -in $_.licenses.accountSkuId))} | Select UserPrincipalName;

    # Dump the licensed users to a CSV file, for reference.
    try {
        Del C:\cron\o365_licensed.csv;
        $licencedUsers | Export-Csv C:\cron\o365_licensed.csv
    } catch {
        Log "ERROR: unable to dump O365-licensed users to a CSV";
    }

    # Get a list of users from the CMS via the REST API.
    $cmsUsers = Invoke-RestMethod ("{0}fast/?all=true&compact=true" -f $user_api) -WebSession $oimsession;
    if (-Not $cmsUsers.objects) {
        [void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions");
        $json = New-Object -TypeName System.Web.Script.Serialization.JavaScriptSerializer;
        $json.MaxJsonLength = 104857600;
        $cmsUsers = $json.Deserialize($cmsUsers, [System.Object]);
    }

    # Iterate through each CMS user, updating as required.
    foreach ($cmsUser in $cmsUsers.objects) {
        $updateUser = $False;
        $body = @{email=$cmsUser.email};

        # FIXME: skip #ext# users until we have an API endpoint that uses itassets IDs instead of emails
        If ($cmsUser.email -like "*#EXT#@*") {
            Continue;
        }
        If ($cmsUser.email -like "*@dpaw.onmicrosoft.com") {
            Continue;
        }
        If ($cmsUser.email -like "*@dec.wa.gov.au") {
            Continue;
        }
        If ($cmsUser.email -like "*@der.wa.gov.au") {
            Continue;
        }
        # Case 0: if the CMS user O365 licence status is currently unknown, consider it to be "False" and flag for an update.
        If ($cmsUser.o365_licence -eq $null) {
            #echo ("{0} license status is currently 'Unknown' in the CMS; set it to 'False' (update needed)" -f $cmsUser.email);
            Log ("{0} license status is currently 'Unknown' in the CMS; set it to 'False' (update needed)" -f $cmsUser.email);
            $body.o365_licence = $False;
            $updateUser = $True;
        } Else {
            $body.o365_licence = $cmsUser.o365_licence;
        }

        # Case 1: if the CMS user IS marked "licensed" and IS NOT present in the list of licenced users, flag an update to "False".
        If ($cmsUser.o365_licence) {
            #echo ("{0} is marked as licensed" -f $cmsUser.email)
            $matches = $licencedUsers -imatch $cmsUser.email
            if ($matches.count -eq 1) {
                #echo ("{0} is present once in the list; skipping it (no change needed)" -f $cmsUser.email)
                # pass
            } else {
                #echo ("{0} is present in the list {1} times; should be marked as 'not licensed' (update needed)" -f $cmsUser.email, $matches.count)
                Log ("{0} is present in the list {1} times; should be marked as 'not licensed' (update needed)" -f $cmsUser.email, $matches.count)
                $body.o365_licence = $False;
                $updateUser = $True;
            }
        }

        # Case 2: if the CMS user IS NOT marked "licenced" and IS present in the list of licenced users, flag an update to "True".
        If (-Not $cmsUser.o365_licence) {
            #echo ("{0} is marked as not licensed" -f $cmsUser.email)
            $matches = $licencedUsers -imatch $cmsUser.email
            if ($matches.count -eq 1) {
                #echo ("{0} should be marked as O365 licensed in the CMS (update needed)" -f $cmsUser.email)
                Log ("{0} should be marked as O365 licensed in the CMS (update needed)" -f $cmsUser.email);
                $body.o365_licence = $True;
                $updateUser = $True;
            } else {
                #echo ("{0} is present in the list {1} times; skipping it (no change needed)" -f $cmsUser.email, $matches.count)
            }
        }

        # If the CMS user is flagged for an update, invoke the API and update the object.
        If ($updateUser) {
            $jsonbody = $body | ConvertTo-Json;
            if ($cmsUser.ad_guid) {
                $user_update_api = $user_api + '{0}/' -f $cmsUser.ad_guid;
            } else {
                $user_update_api = $user_api + '{0}/' -f [uri]::EscapeDataString($cmsUser.email);
            }
        
            try {
                # Invoke the API.
                $response = Invoke-RestMethod $user_update_api -Method Put -Body $jsonbody -ContentType "application/json" -WebSession $oimsession -Verbose;
                Log $("INFO: updated OIM CMS user {0} O365 licence status: {1}" -f $cmsUser.email, $body.o365_licence);
            } catch [System.Exception] {
                # Log any failures to sync AD data into the OIM CMS, for reference.
                Log $("ERROR: failed to update OIM CMS user {0} O365 licence status" -f $cmsUser.email);
                Log $($jsonbody);
            }
        }
    }

    Log "Finished";
} catch [System.Exception] {
    Log "ERROR: Exception caught, dying =(";
    $except = $_;
    Log $($except | convertto-json);
}

# Final clean up.
Get-PSSession | Remove-PSSession;