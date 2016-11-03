Import-Module ActiveDirectory;
Import-Module -Force 'C:\cron\creds.psm1';
$ErrorActionPreference = "Stop";

Function Log {
   Param ([string]$logstring)
   Add-content "C:\cron\o365_licences.log" -value $("{0} ({1} - {2}): {3}" -f $(Get-Date), $(GCI $MyInvocation.PSCommandPath | Select -Expand Name), $pid, $logstring);
}

try {
    # Get a list of Office 365-licenced users.
    # TODO: confirm/audit this query.
    $licencedUsers = Get-MsolUser -All | Where {$_.isLicensed -eq "True" -and ("DPaW:ENTERPRISEPACK" -in $_.licenses.accountSkuId)} | select signinname;
    # Get a list of users from the CMS via the REST API.
    $cmsUsers = Invoke-RestMethod ("{0}" -f $user_api) -WebSession $oimsession;
    # Do a workaround to vault PowerShell's dumb 10mb JSON limit.
    if (-not $cmsUsers.objects) {
        [void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions");
        $json = New-Object -TypeName System.Web.Script.Serialization.JavaScriptSerializer;
        $json.MaxJsonLength = 104857600;
        $cmsUsers = $json.Deserialize($cmsUsers, [System.Object]);
    }
    # Iterate through each CMS user, updating as required.
    foreach ($cmsUser in $cmsUsers.objects) {
        $body = @{email=$cmsUser.email; o365_licence=$cmsUser.o365_licence};
        # Case 1: if the CMS user IS present in the list of licenced users and IS NOT marked "licenced".
        if ((-Not $cmsUser.o365_licence) -And ($licencedUsers -Match $cmsUser.email)) {
            $body.o365_licence = $True;
        }
        # Case 2: if the CMS user is not present in the list of licenced users and is marked "licenced":
        elseif (($cmsUser.o365_licence) -And (-Not $licencedUsers -Match $cmsUser.email)) {
            $body.o365_licence = $False;
        }
        # If the CMS user data differs here, invoke the API and update the object.
        If ($body.o365_licence -ne $cmsUser.o365_licence) {
            $jsonbody = $body | ConvertTo-Json;
            try {
                # Invoke the API.
                $response = Invoke-RestMethod $user_api -Method Post -Body $jsonbody -ContentType "application/json" -WebSession $oimsession -Verbose;
                Log $("INFO: updated OIM CMS user {0} as having an O365 licence" -f $cmsUser.email);
            } catch [System.Exception] {
                # Log any failures to sync AD data into the OIM CMS, for reference.
                Log $("ERROR: failed to update OIM CMS user {0}" -f $cmsUser.email);
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