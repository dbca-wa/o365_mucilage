Import-Module -Force 'C:\cron\creds.psm1';
$ErrorActionPreference = "Stop";

Function Log {
   Param ([string]$logstring)
   $output = $("{0} ({1} - {2}): {3}" -f $(Get-Date), $(GCI $MyInvocation.PSCommandPath | Select -Expand Name), $pid, $logstring);
   Write-Host $output;
   Add-content "C:\cron\directory_wrangler.log" -value $output;
}

try {
    Log "Starting directory_wrangler script";

    # Store the domain max password age in days.
    $DefaultmaxPasswordAgeDays = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge.Days;

    # Read the full user DB from IT Assets (all DepartmentUser objects) via the API.
    # NOTE: $user_api is set in C:\cron\creds.psm1
    $users = Invoke-RestMethod ("{0}?all" -f $user_api) -WebSession $oimsession -TimeoutSec 300;
    # Deserialise response into JSON (bypass the MaxJsonLength property of 2 MB).
    if (-not $users.objects) {
        [void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions");
        $json = New-Object -TypeName System.Web.Script.Serialization.JavaScriptSerializer;
        $json.MaxJsonLength = 104857600;
        $users = $json.Deserialize($users, [System.Object]);
    }
    $user_guid = @{};
    ForEach($user in $users.objects | where {$_.ad_guid}) {
        $user_guid[$user.ad_guid] = $user;
    }
    
    # Define user object attributes that we care about.
    $keynames = @("Title", "DisplayName", "GivenName", "Surname", "Company", "physicalDeliveryOfficeName", "StreetAddress", 
        "Division", "Department", "Country", "State", "wWWHomePage", "Manager", "EmployeeID", "EmployeeNumber", "HomePhone",
        "telephoneNumber", "Mobile", "Fax", "employeeType");
    $adprops = $keynames + @("EmailAddress", "UserPrincipalName", "Modified", "AccountExpirationDate", "Info", "pwdLastSet", 
                             "targetAddress", "msExchRemoteRecipientType", "msExchRecipientTypeDetails", "proxyAddresses");
    
    # Read the user list from AD. Apply a rough filter for accounts we want to load into IT Assets:
    # - email address is *.wa.gov.au or dpaw.onmicrosoft.com
    # - DN contains a sub-OU called "Users"
    # - DN does not contain a sub-OU with "Administrators" in the name
    $adusers = @();
    ForEach ($ou in $user_ous) {
        $adusers += Get-ADUser -server $adserver -Filter {EmailAddress -like "*@*wa.gov.au"} -Properties $adprops -SearchBase $ou;
        $adusers += Get-ADUser -server $adserver -Filter {EmailAddress -like "*@rottnestisland.com"} -Properties $adprops -SearchBase $ou;
    }
    $adusers += Get-ADUser -server $adserver -Filter {EmailAddress -like "*@dpaw.onmicrosoft.com"} -Properties $adprops;
    Log $("Processing {0} users" -f $adusers.Length);
    $cmsusers_updated = $false;



    ###################################
    ##### PUSH AD EMAIL CHANGES TO CMS
    ###################################

    #Write-Output "UPDATING IT ASSETS FROM AD DATA";
    
    ForEach ($aduser in $adusers) {
        # NOTE: we no longer create new CMS DepartmentUser objects here, that happens in cloud_wrangler
        $cmsUser = $user_guid[$aduser.ObjectGUID];
        if ($cmsUser) {
            # Find any cases where the AD user's email has been changed, and update the CMS user.
            if (-Not ($cmsUser.email -like $aduser.EmailAddress)) {
                $simpleuser = $aduser | select ObjectGUID, @{name="Modified";expression={Get-Date $_.Modified -Format o}}, info, DistinguishedName, Name, Title, GivenName, Surname, EmailAddress, Enabled, AccountExpirationDate, pwdLastSet, proxyAddresses;
                $simpleuser | Add-Member -type NoteProperty -name PasswordMaxAgeDays -value $DefaultmaxPasswordAgeDays;
                if ($aduser.AccountExpirationDate) {
                    $simpleuser.AccountExpirationDate = Get-Date $aduser.AccountExpirationDate -Format o;
                }
                # only write back username if enabled for this directory. avoids collisions in IT Assets.
                if ($dw_writeusername) {
                    $simpleuser | Add-Member -type NoteProperty -name SamAccountName -value $aduser.SamAccountName;
                }
                # ...convert the whole lot to JSON and push to IT Assets via the REST API.
                $userjson = [System.Text.Encoding]::UTF8.GetBytes($($simpleuser | ConvertTo-Json));
                $user_update_api = $user_api + '{0}/' -f $simpleuser.ObjectGUID;
                try {
                    # Invoke the API.
                    #Log $("Updating IT Assets data for {0}" -f $cmsUser.email);
                    #$response = Invoke-RestMethod $user_update_api -Body $userjson -Method Put -ContentType "application/json" -WebSession $oimsession;
                    # Note that a change has occurred.
                    $cmsusers_updated = $true;
                } catch [System.Exception] {
                    # Log any failures to sync AD data into the IT Assets, for reference.
                    Log $("ERROR: updating IT Assets database failed for {0}" -f $cmsUser.email);
                    Log $("Endpoint: {0}" -f $user_update_api);
                    Log $("Payload: {0}" -f $simpleuser | ConvertTo-Json);
                    $result = $_.Exception.Response.GetResponseStream();
                    $reader = New-Object System.IO.StreamReader($result);
                    $reader.BaseStream.Position = 0;
                    $reader.DiscardBufferedData();
                    $responseBody = $reader.ReadToEnd();
                    Log $("Response: {0}" -f $responseBody);
                }
            }
        }
    }

    ###########################
    ##### IT Assets 2-WAY UPDATE
    ###########################

    # Get the list of users from the CMS again (if required, following any additions/updates).
    if ($cmsusers_updated) {
        $users = Invoke-RestMethod ("{0}?all" -f $user_api) -WebSession $oimsession -TimeoutSec 300;
        # Deserialise response into JSON (bypass the MaxJsonLength property of 2 MB).
        if (-not $users.objects) {
            [void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions");
            $json = New-Object -TypeName System.Web.Script.Serialization.JavaScriptSerializer;
            $json.MaxJsonLength = 104857600;
            $users = $json.Deserialize($users, [System.Object]);
        }
    }

    Write-Output "TIME TO UPDATE AD FROM IT ASSETS DATA";
    # filter IT Assets DepartmentUsers by whitelisted OrgUnit
    # this is to allow multiple directory writeback

    #$department_users = $users.objects | Where {$_.org_data.units.id | Where {$org_whitelist -contains $_} };
    
    $department_users = $users.objects | Where {$_.ad_dn -like $domain_dn} | Where {$_.org_data.units.id | Where {$org_global -contains $_} };
    $department_users += $users.objects | Where {-not $_.ad_dn } | Where {$_.org_data.units.id | Where {$org_whitelist -contains $_} };

    # For each IT Assets DepartmentUser...
    foreach ($user in $department_users) {
        # ...find the equivalent Active Directory Object.
        If (-Not $user.ad_guid) {
            $aduser = $adusers | where EmailAddress -eq $user.email;
            If ($aduser) {
                Log $('Found a match for {0}, adding GUID {1}' -f $user.email,$aduser.ObjectGUID);
                $simpleuser = $aduser | select ObjectGUID, DistinguishedName;
                $userjson = [System.Text.Encoding]::UTF8.GetBytes($($simpleuser | ConvertTo-Json));
                $user_update_api = $user_api + '{0}/' -f $user.email;
                try {
                    # Invoke the API.
                    #$response = Invoke-RestMethod $user_update_api -Body $userjson -Method Put -ContentType "application/json" -WebSession $oimsession;
                    # Note that a change has occurred.
                    $cmsusers_updated = $true;
                } catch  {
                    # Log any failures to sync AD data into the IT Assets db, for reference.
                    Log $("ERROR: updating IT Assets db failed for {0}" -f $user.email);
                    Log $("Endpoint: {0}" -f $user_update_api);
                    Log $("Payload: {0}" -f $simpleuser | ConvertTo-Json);
                    $result = $_.Exception.Response.GetResponseStream();
                    $reader = New-Object System.IO.StreamReader($result);
                    $reader.BaseStream.Position = 0;
                    $reader.DiscardBufferedData();
                    $responseBody = $reader.ReadToEnd();
                    Log $("Response: {0}" -f $responseBody);
                }
            }
            continue;
        }

        $aduser = $adusers | where ObjectGUID -eq $($user.ad_guid);
        
        If ($aduser) {
            # If the IT Assets user object was modified in the last hour...
            if (($(Get-Date) - (New-TimeSpan -Minutes 60)) -lt $(Get-Date $user.date_updated) -and ($aduser.Modified -lt $(Get-Date $user.date_updated) - $(New-TimeSpan -Minutes 5) )) {
                #Write-Output $("Looks like {0} was modified in the last hour, updating" -f $user.email);
                # ...set all the properties on the AD object to match the IT Assets object
                #Log $("modAD: {0}, modCMS: {1}" -f $aduser.Modified, $(Get-Date $user.date_updated));

                $aduser.Title = $user.title;
                $aduser.DisplayName, $aduser.GivenName, $aduser.Surname = $user.name, $user.given_name, $user.surname;
                $aduser.Company = $user.org_data.cost_centre.code;
                $aduser.physicalDeliveryOfficeName = $user.org_unit__location__name;
                $aduser.StreetAddress = $user.org_unit__location__address;
                #if ($user.org_data.units) {
                #    $aduser.Division = $user.org_data.units[1].name;
                #    $aduser.Department = $user.org_data.units[0].name;
                #}
                $aduser.Department = $user.gal_department;
                $aduser.Country, $aduser.State = "AU", "Western Australia";
                $aduser.wWWHomePage = "https://oim.dbca.wa.gov.au/address-book/user-details?email=" + $user.email;
                $aduser.EmployeeNumber, $aduser.EmployeeID = $user.employee_id, $user.employee_id;
                $aduser.telephoneNumber, $aduser.Mobile = $user.telephone, $user.mobile_phone;
                $aduser.Fax = $user.org_unit__location__fax;
                $aduser.employeeType = $user.account_type + " " + $user.position_type;
                if ($user.org_data.cost_centre) {
                    $aduser.Description = $user.org_data.cost_centre.code + " - " + $user.org_data.cost_centre.name;
                }
                # If the user has an account expiry date set that in AD.
                if ($user.expiry_date) {
                    $aduser.AccountExpirationDate = $(Get-Date $user.expiry_date);
                }
                if (-not ($user.parent__email -like ($adusers | where distinguishedname -like $aduser.Manager).emailaddress)) {
                    $aduser.Manager = ($adusers | where emailaddress -like $($user.parent__email)).DistinguishedName;
                }
                # ...make all of the undefined properties the string "N/A"
                foreach ($prop in $aduser.ModifiedProperties) { 
                    if ((-not $aduser.$prop) -and ($prop -notlike "manager")) {
                        $aduser.$prop = "N/A";
                    } 
                }
                # ...push changes back to AD
                try {
                    #Log $("Updating AD data with IT Assets data (newer) for {0}" -f $aduser.EmailAddress);
                    Set-ADUser -verbose -server $adserver -instance $aduser;
                    # (thumbnailPhoto isn't added as a property of $aduser for some dumb reason, so we have to push it seperately)
                    # FIXME: photo_ad is broken in IT assets
                    #if ($user.photo_ad -and $user.photo_ad.startswith('http')) {
                    #    Set-ADUser -verbose -server $adserver $aduser -replace @{thumbnailPhoto=$(Invoke-WebRequest $user.photo_ad -WebSession $oimsession).content};
                    #}
                    #else {
                    #    Set-ADUser -verbose -server $adserver $aduser -clear thumbnailPhoto;
                    #}
                    
                    # If there's an email address change, use Outlook
                    If ($user.email -and ($aduser.emailaddress -ne $user.email)) {
                        Log $("Updating email address for {0} to {1}" -f $aduser.EmailAddress, $user.email);
                        Set-ADUser -verbose -server $adserver $aduser -EmailAddress $user.email -UserPrincipalName $user.email;

                        # scrub older mentions of new primary SMTP
                        ForEach ($existing in $aduser.proxyAddresses | Where {$_ -like "smtp:"+$user.email}) {
                            Set-ADUser -verbose -server $adserver $aduser -Remove @{'proxyAddresses'=($existing)};
                        }

                        # find the current primary SMTP
                        # it'll either be nothing, or something other than the new primary SMTP
                        $proxyAddresses = $aduser.proxyAddresses | Where {$_ -notlike "smtp:"+$user.email}
                        $current_primary = $proxyAddresses | Where {$_ -clike "SMTP:*"} | Select -First 1;
                        if ($current_primary) {
                            # move the current primary to a secondary
                            $current_primary_email = $($current_primary -split 'SMTP:', 2)[1];
                            if ($current_primary_email) {
                                Set-ADUser -verbose -server $adserver $aduser -Remove @{'proxyAddresses'=($current_primary)};
                                Set-ADUser -verbose -server $adserver $aduser -Add @{'proxyAddresses'=('smtp:'+$current_primary_email)};
                            }
                        }

                        # add the new primary
                        Set-ADUser -verbose -server $adserver $aduser -Add @{'proxyAddresses'=('SMTP:'+$user.email)};
                    }

                    # force a timestamp update
                    Set-ADUser -server $adserver -identity $aduser.ObjectGUID -replace @{ExtensionAttribute15="test"};
                    Set-ADUser -server $adserver -identity $aduser.ObjectGUID -clear ExtensionAttribute15;

                } catch [System.Exception] {
                    Log $("ERROR: set-aduser failed on {0}" -f $user.email);
                    Log $($aduser | select $($aduser.ModifiedProperties) | convertto-json);
                    $except = $_;
                    Log $($except | convertto-json);
                }
            
            # If the AD object was modified after the IT Assets object, sync back to the CMS...
            } ElseIf ((-not $user.ad_data.Modified) -or ($aduser.Modified -gt $(Get-Date $user.ad_data.Modified) + (New-Timespan -Minutes 5))) {
                #Log $("modAD: {0}, modCMS: {1}" -f $aduser.Modified, $(Get-Date $user.ad_data.Modified));
                #Write-Output $("Looks like {0} was updated in AD after the CMS, updating" -f $user.email);
                # ...glom the mailbox object onto the AD object
                $simpleuser = $aduser | select ObjectGUID, @{name="Modified";expression={Get-Date $_.Modified -Format o}}, info, DistinguishedName, Name, Title, GivenName, Surname, EmailAddress, Enabled, AccountExpirationDate, pwdLastSet;
                $simpleuser | Add-Member -type NoteProperty -name PasswordMaxAgeDays -value $DefaultmaxPasswordAgeDays;
                if ($aduser.AccountExpirationDate) {
                    $simpleuser.AccountExpirationDate = Get-Date $aduser.AccountExpirationDate -Format o;
                }
                # only write back username if enabled for this directory. avoids collisions in IT Assets
                if ($dw_writeusername) {
                    $simpleuser | Add-Member -type NoteProperty -name SamAccountName -value $aduser.SamAccountName;
                }
                # ...convert the whole lot to JSON and push to IT Assets via the REST API.
                $userjson = [System.Text.Encoding]::UTF8.GetBytes($($simpleuser | ConvertTo-Json));
                $user_update_api = $user_api + '{0}/' -f $simpleuser.ObjectGUID;
                try {
                    # Invoke the API.
                    #Log $("Updating IT Assets data for {0} from AD data (newer)" -f $user.email);
                    #$response = Invoke-RestMethod $user_update_api -Body $userjson -Method Put -ContentType "application/json" -WebSession $oimsession;
                } catch [System.Exception] {
                    # Log any failures to sync AD data into the IT Assets, for reference.
                    Log $("ERROR: updating IT Assets failed for {0}" -f $cmsUser.email);
                    Log $("Endpoint: {0}" -f $user_update_api);
                    Log $("Payload: {0}" -f $simpleuser | ConvertTo-Json);
                    $result = $_.Exception.Response.GetResponseStream();
                    $reader = New-Object System.IO.StreamReader($result);
                    $reader.BaseStream.Position = 0;
                    $reader.DiscardBufferedData();
                    $responseBody = $reader.ReadToEnd();
                    Log $("Response: {0}" -f $responseBody);
                }
            }
        } Else {
            #Write-Output $("Couldn't find {0}!" -f $user.email);
            # No AD object found - mark the user as "AD deleted" in the CMS (if it's not already).
            If (-Not $user.ad_deleted) {
                $body = @{EmailAddress=$user.email; Deleted="true"};
                $jsonbody = [System.Text.Encoding]::UTF8.GetBytes($($body | ConvertTo-Json));
                try {
                    $user_update_api = $user_api + '{0}/' -f $user.ad_guid;
                    # Invoke the API.
                    $response = Invoke-RestMethod $user_update_api -Method Put -Body $jsonbody -ContentType "application/json" -WebSession $oimsession -Verbose;
                    Log $("INFO: updated IT Assets user {0} as deleted in Active Directory" -f $user.email);
                } catch [System.Exception] {
                    # Log any failures to sync AD data into the IT Assets, for reference.
                    Log $("ERROR: failed to update IT Assets user {0} as deleted in Active Directory" -f $user.email);
                    Log $("Endpoint: {0}" -f $user_update_api);
                    Log $("Payload: {0}" -f $body | ConvertTo-Json);
                    $result = $_.Exception.Response.GetResponseStream();
                    $reader = New-Object System.IO.StreamReader($result);
                    $reader.BaseStream.Position = 0;
                    $reader.DiscardBufferedData();
                    $responseBody = $reader.ReadToEnd();
                    Log $("Response: {0}" -f $responseBody);
                }
            }
        }

        # If the user is disabled in AD but still marked active in the IT Assets, update the user in the CMS.
        if ($aduser.enabled -eq $false) {
            if ($user.active) {
                $simpleuser = $aduser | select ObjectGUID,  info, DistinguishedName, Name, Title, GivenName, Surname, EmailAddress, Enabled, AccountExpirationDate, pwdLastSet;
                if ($aduser.AccountExpirationDate) { 
                    $simpleuser.AccountExpirationDate = Get-Date $aduser.AccountExpirationDate -Format o;
                }
                # only write back username if enabled for this directory. avoids collisions in IT Assets
                if ($dw_writeusername) {
                    $simpleuser | Add-Member -type NoteProperty -name SamAccountName -value $aduser.SamAccountName;
                }
                $userjson = [System.Text.Encoding]::UTF8.GetBytes($($simpleuser | ConvertTo-Json));
                
                try {
                    $user_update_api = $user_api + '{0}/' -f $simpleuser.ObjectGUID;
                    # Invoke the API.
                    $response = Invoke-RestMethod $user_update_api -Body $userjson -Method Put -ContentType "application/json" -Verbose -WebSession $oimsession;
                    Log $("Marked {0} as 'Inactive' in the IT Assets" -f $user.email);
                } catch [System.Exception] {
                    # Log any failures to sync AD data into the IT Assets, for reference.
                    Log $("ERROR: failed to update {0} as inactive in IT Assets" -f $user.email);
                    Log $("Endpoint: {0}" -f $user_update_api);
                    Log $("Payload: {0}" -f $simpleuser | ConvertTo-Json);
                    $result = $_.Exception.Response.GetResponseStream();
                    $reader = New-Object System.IO.StreamReader($result);
                    $reader.BaseStream.Position = 0;
                    $reader.DiscardBufferedData();
                    $responseBody = $reader.ReadToEnd();
                    Log $("Response: {0}" -f $responseBody);
                }
            }
        }
    }

    ####################
    ##### AZURE AD SYNC
    ####################

    if ($dw_aadsync) {
        Import-Module ADSync;
        #Write-Output "TIME TO SYNC TO O365";
        # We've done a whole pile of AD changes, so now's a good time to run AADSync to push them to O365:
        Log "Azure AD Connect Syncing with O365";
        Start-ADSyncSyncCycle -PolicyType Delta;
        # This command is not blocking and the new AAD Connect API is crap at polling for activity,
        # so let's just block for 60 seconds!
        Start-Sleep -s 60;
    }

    ############################
    ##### NEW USER AD WRITEBACK
    ############################

    # Finally, we want to do some operations on Office 365 accounts not handled by AADSync.
    # Start by reading the full user list.
    $msolusers = get-msoluser -all | select userprincipalname, lastdirsynctime, @{name="licenses";expression={[string]$_.licenses.accountskuid}}, signinname, immutableid, whencreated, displayname, firstname, lastname;
    $msolusers | convertto-json > 'C:\cron\msolusers.json';
        
    # For each "In cloud" user in Azure AD which is not directory synced and is part of the IT Assets org whitelist...
    ForEach ($msoluser in $msolusers | where lastdirsynctime -eq $null | where UserPrincipalName -in $department_users.email) {
        $username = $msoluser.FirstName + $msoluser.LastName;
        if (!$username) {
            $username = $msoluser.UserPrincipalName.Split("@", 2)[0].replace('.','').replace("'", '').replace('#', '').replace(',', '')
        }
        $username = $username.Substring(0,[System.Math]::Min(15, $username.Length));
        # ...link existing users 
        $upn = $msoluser.UserPrincipalName;
        $existing = Get-ADUser -Filter { UserPrincipalName -like $upn };
        if ($existing) {
            # if we have 365 admin credentials, update the user object in Office 365 to have the right linking ID
            if ($dw_write365) {
                $immutableid = [System.Convert]::ToBase64String($existing.ObjectGUID.tobytearray());
                Set-MsolUser -UserPrincipalName $upn -ImmutableId $immutableid;
            }
            continue;
        }
        # ...create new user
        #$sam = $msoluser.userprincipalname.split('@')[0].replace('.','').replace('#', '').replace(',', '');
        # ...assume targetAddress name is the same base as the UPN
        $rra = $msoluser.UserPrincipalName.Split("@", 2)[0]+"@dpaw.mail.onmicrosoft.com";

        Log $("About to create O365 user: New-ADUser $username -Verbose -Path `"$new_user_ou`" -Enabled $true -UserPrincipalName $($msoluser.UserPrincipalName) -SamAccountName $($username) -EmailAddress $($msoluser.UserPrincipalName) -DisplayName $($msoluser.DisplayName) -GivenName $($msoluser.FirstName) -Surname $($msoluser.LastName) -PasswordNotRequired $true");
        New-ADUser -server $adserver -verbose $username -Path $new_user_ou -Enabled $true -UserPrincipalName $msoluser.UserPrincipalName -SamAccountName $username -EmailAddress $msoluser.UserPrincipalName -DisplayName $msoluser.DisplayName -GivenName $msoluser.FirstName -Surname $msoluser.LastName -PasswordNotRequired $true;
        # ...wait for changes to propagate
        sleep 180;

        Set-ADUser -verbose -server $adserver -Identity $username -Add @{'proxyAddresses'=('SMTP:'+$msoluser.UserPrincipalName)};
        Set-ADUser -verbose -server $adserver -Identity $username -Add @{'proxyAddresses'=('smtp:'+$rra)};
        Set-ADUser -verbose -server $adserver -Identity $username -Replace @{'targetAddress'=('SMTP:'+$rra)};
    }

    ##############
    ##### CLEANUP
    ##############

    $adusers = @();
    ForEach ($ou in $user_ous) {
        $adusers += Get-ADUser -server $adserver -Filter {EmailAddress -like "*@*wa.gov.au"} -Properties $adprops -SearchBase $ou;
    }
    $adusers += Get-ADUser -server $adserver -Filter {EmailAddress -like "*@dpaw.onmicrosoft.com"} -Properties $adprops;

    # search for records with a busted/missing RRA
    $busted = $adusers | where {($_.enabled -and -not $_.targetAddress)};
    foreach ($aduser in $busted) {
        $ms = $msolusers | where {$_.immutableId -eq [System.Convert]::ToBase64String($aduser.ObjectGUID.tobytearray())};
        if ($ms -and ($ms.licenses)) {
            $rra = $aduser.EmailAddress.Split("@", 2)[0]+"@dpaw.mail.onmicrosoft.com";
            Log $("Fixing TargetAddress for {0} {1} {2}" -f $aduser.EmailAddress, $aduser.userprincipalname, $rra);
            $aduser | Set-ADUser -verbose -server $adserver -Replace @{'targetAddress'=('SMTP:'+$rra)};
            if (-not (('smtp:'+$rra) -in $aduser.proxyAddresses)) {
                $aduser | Set-ADUser -verbose -server $adserver -Add @{'proxyAddresses'=('smtp:'+$rra)};
            }
        }
    }

    # Rig the UPN for each user account so that it matches the primary SMTP address.
    foreach ($aduser in $adusers | where {$_.emailaddress -and ($_.emailaddress -ne $_.userprincipalname)}) {
        Log $("Changing UPN from {0} to {1}" -f $aduser.UserPrincipalName,$aduser.emailaddress);
        Set-ADUser -verbose -server $adserver $aduser -UserPrincipalName $aduser.emailaddress;
    }

    foreach ($aduser in $adusers) {
        # Iterate over CMS DepartmentUsers and call Disable-ADAccount for any that have expired.
        if (($aduser.Enabled -eq $true) -and ($aduser.AccountExpirationDate) -and ($aduser.AccountExpirationDate -lt $(Get-Date))) {
            Log $("Disabling AD account {0}" -f $aduser.EmailAddress);
            Disable-ADAccount -server $adserver $aduser;
        }

        # For each AD-managed Exchange Online mailbox that doesn't have it, add an archive mailbox:
#        if ($aduser.msExchRemoteRecipientType -in @(1, 4)) {
#            Log $("Adding archive mailbox for {0}" -f $aduser.userprincipalname);
#            Set-ADUser -verbose -server $adserver $aduser -Replace @{msExchRemoteRecipientType=$($aduser.msExchRemoteRecipientType -bor 2)};
#        }
    }


    # Quick loop to fix targetAddress; previously some RemoteMailbox objects were provisioned manually with the wrong one.
    ForEach ($aduser in $adusers | Where {-not ($_.targetAddress -like "*@dpaw.mail.onmicrosoft.com" )} ) {
        $rra = $aduser.proxyAddresses | Where {$_ -like "*@dpaw.mail.onmicrosoft.com"} | Select -First 1;
        If ($rra) {
            $rra = $($rra -split 'smtp:', 2)[1];
            if ($rra) {
                Log $("Fixing target address for {0} to {1}" -f $aduser.EmailAddress, $rra);
                $aduser | Set-ADUser -verbose -server $adserver -Replace @{'targetAddress'=('SMTP:'+$rra)};
            }
        }
    }
    Log "Finished";
} catch [System.Exception] {
    Log "ERROR: Exception caught, dying =(";
    $except = $_;
    Log $($except);#| convertto-json);
    Exit(1);
}

# Final clean up.
Get-PSSession | Remove-PSSession;
