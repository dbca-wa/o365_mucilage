Import-Module -Force 'C:\cron\creds.psm1'
$ErrorActionPreference = "Stop"

Function Log {
   Param ([string]$logstring)
   Add-content "C:\cron\directory_wrangler.log" -value $("{0} ({1} - {2}): {3}" -f $(Get-Date), $(GCI $MyInvocation.PSCommandPath | Select -Expand Name), $pid, $logstring)
}

try {
    $mailboxes = $(Get-Mailbox -ResultSize unlimited | select userprincipalname, primarysmtpaddress, recipienttypedetails) + $(Get-RemoteMailbox -ResultSize unlimited | select userprincipalname, primarysmtpaddress, recipienttypedetails)
    $mailboxes | convertto-json > 'C:\cron\mailboxes.json';
    
    $users = Invoke-RestMethod ("{0}?all" -f $user_api)
    if (-not $users.objects) {
        [void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")        
        $json = New-Object -TypeName System.Web.Script.Serialization.JavaScriptSerializer 
        $json.MaxJsonLength = 104857600
        $users = $json.Deserialize($users, [System.Object])
    }
    $keynames = @("Title", "DisplayName", "GivenName", "Surname", "Company", "physicalDeliveryOfficeName", "StreetAddress", "Division", "Department", "Country", "State",
        "wWWHomePage", "Manager", "EmployeeID", "EmployeeNumber", "HomePhone", "telephoneNumber", "Mobile", "Fax")
    $adprops = $keynames + @("EmailAddress", "UserPrincipalName", "Modified", "AccountExpirationDate", "Info")
    $adusers = Get-ADUser -server $adserver -Filter {EmailAddress -like "*@*wa.gov.au" -and Surname -ne $false} -Properties $adprops | where distinguishedName -Like "*OU=Users*" | where distinguishedName -NotLike "*Administrators*"
    $adusers += Get-ADUser -server $adserver -Filter {EmailAddress -like "*@dpaw.onmicrosoft.com"} -Properties $adprops
    Log $("Processing {0} users" -f $adusers.Length)

    ForEach ($aduser in $adusers | where { $_.EmailAddress -notin $users.objects.email }) {
        $simpleuser = $aduser | select ObjectGUID, DistinguishedName, Name, Title, SamAccountName, GivenName, Surname, EmailAddress, Modified, Enabled, AccountExpirationDate
        $simpleuser.Modified = Get-Date $aduser.Modified -Format s
        if ($aduser.AccountExpirationDate) { $simpleuser.AccountExpirationDate = Get-Date $aduser.AccountExpirationDate -Format s }
        $userjson = $simpleuser | ConvertTo-Json
        (Invoke-RestMethod $user_api -Body $userjson -Method Post -ContentType "application/json" -Verbose).ad_data
    }

    foreach ($user in $users.objects) {
        $aduser = $adusers | where EmailAddress -like $($user.email)
        If ($aduser) {
            if ($aduser.Modified -lt $((Get-Date $user.date_updated) - (New-TimeSpan -Minutes 60))) {
                $aduser.Title = $user.title
                $aduser.DisplayName, $aduser.GivenName, $aduser.Surname = $user.name, $user.given_name, $user.surname
                $aduser.Company = $user.org_data.cost_centre.code
                $aduser.physicalDeliveryOfficeName = $user.org_unit__location__name
                $aduser.StreetAddress = $user.org_unit__location__address
                if ($user.org_data.units) {
                    $aduser.Division = $user.org_data.units[1].name
                    $aduser.Department = $user.org_data.units[0].name }
                $aduser.Country, $aduser.State = "AU", "Western Australia"
                $aduser.wWWHomePage = "https://oim.dpaw.wa.gov.au/userinfo?email=" + $user.email
                $aduser.EmployeeNumber, $aduser.EmployeeID = $user.employee_id, $user.employee_id
                $aduser.telephoneNumber, $aduser.Mobile = $user.telephone, $user.mobile_phone
                $aduser.Fax = $user.org_unit__location__fax
                if ($user.parent__email -ne ($adusers | where distinguishedname -like $aduser.Manager).emailaddress) {
                    $aduser.Manager = ($adusers | where emailaddress -like $($user.parent__email)).DistinguishedName
                }
                foreach ($prop in $aduser.ModifiedProperties) { if ((-not $aduser.$prop) -and ($prop -notlike "manager")) {$aduser.$prop = "N/A"} }
                try {
                    set-aduser -verbose -server $adserver -instance $aduser
                } catch [System.Exception] {
                    Log $("ERROR: set-aduser failed on {0}" -f $user.email)
                    Log $($aduser | select $($aduser.ModifiedProperties) | convertto-json)
                    $except = $_
                    Log $($except | convertto-json)
                }
            }
            if ($aduser.Modified -gt $(Get-Date $user.ad_data.Modified)) {
                $mb = $mailboxes | where userprincipalname -like $user.email
                $simpleuser = $aduser | select ObjectGUID, @{name="mailbox";expression={$mb}}, @{name="Modified";expression={Get-Date $_.Modified -Format s}}, info, DistinguishedName, Name, Title, SamAccountName, GivenName, Surname, EmailAddress, Enabled, AccountExpirationDate
                if ($aduser.AccountExpirationDate) { $simpleuser.AccountExpirationDate = Get-Date $aduser.AccountExpirationDate -Format s }
                $userjson = [System.Text.Encoding]::UTF8.GetBytes($($simpleuser | ConvertTo-Json))
                try {
                    $ad_data = (Invoke-RestMethod $user_api -Body $userjson -Method Post -ContentType "application/json").ad_data
                } catch [System.Exception] {
                    Log $("ERROR: update cms failed on {0}" -f $user.email)
                    Log $($userjson)
                }
            }
        } 

        if ((-not $aduser) -or ($aduser.enabled -eq $false)) {
            if (-not $user.ad_deleted) {
                $userjson = [System.Text.Encoding]::UTF8.GetBytes($(@{EmailAddress = $user.email;Deleted = $true} | convertto-json))
                $ad_data = (Invoke-RestMethod $user_api -Body $userjson -Method Post -ContentType "application/json" -Verbose).ad_data
            }
        }
    }

    Log "Azure AD Connect Syncing with O365"
    .'C:\Program Files\Microsoft Azure AD Sync\Bin\DirectorySyncClientCmd.exe' delta

    $msolusers = get-msoluser -all | select userprincipalname, lastdirsynctime, @{name="licenses";expression={[string]$_.licenses.accountskuid}}, signinname, immutableid, whencreated, displayname, firstname, lastname
    $msolusers | convertto-json > 'C:\cron\msolusers.json';
    foreach ($aduser in $adusers | where {$_.emailaddress -ne $_.userprincipalname}) {
        $immutableid = [System.Convert]::ToBase64String($aduser.ObjectGuid.toByteArray());
        $msoluser = $msolusers | where immutableid -eq $immutableid
        If ($msoluser) {
            Set-MsolUserPrincipalName -UserPrincipalName $msoluser.UserPrincipalName -NewUserPrincipalName $aduser.emailaddress -Verbose
            Set-ADUser $aduser -UserPrincipalName $aduser.emailaddress -Verbose
        } Else {
            Log $("Warning: MSOL object not found for {0}" -f $aduser.UserPrincipalName)
        }
    }
    $mailboxes | where recipienttypedetails -like remoteusermailbox | where { $_.userprincipalname -ne $_.primarysmtpaddress } | foreach { Set-RemoteMailbox $_.userprincipalname -PrimarySmtpAddress $_.userprincipalname -EmailAddressPolicyEnabled $false -Verbose }

    ForEach ($msoluser in $msolusers | where lastdirsynctime -eq $null | where licenses) {
        $username = $msoluser.FirstName + $msoluser.LastName
        $username = $username.Substring(0,[System.Math]::Min(15, $username.Length))
        # link existing users
        $upn = $msoluser.UserPrincipalName
        $existing = Get-ADUser -Filter { UserPrincipalName -like $upn }
        if ($existing) {
            $immutableid = [System.Convert]::ToBase64String($existing.ObjectGUID.tobytearray())
            Set-MsolUser -UserPrincipalName $upn -ImmutableId $immutableid
            continue
        }
        # Create new user
        Log $("About to create O365 user: New-ADUser $username -Verbose -Path `"OU=Users,OU=DPaW,dc=corporateict,dc=domain`" -Enabled $true -UserPrincipalName $($msoluser.UserPrincipalName) -EmailAddress $($msoluser.UserPrincipalName) -DisplayName $($msoluser.DisplayName) -GivenName $($msoluser.FirstName) -Surname $($msoluser.LastName) -PasswordNotRequired $true")
        New-ADUser $username -Verbose -Path "OU=Users,OU=DPaW,dc=corporateict,dc=domain" -Enabled $true -UserPrincipalName $msoluser.UserPrincipalName -EmailAddress $msoluser.UserPrincipalName -DisplayName $msoluser.DisplayName -GivenName $msoluser.FirstName -Surname $msoluser.LastName -PasswordNotRequired $true
        sleep 10
        Set-ADUser -Identity $username -Add @{'proxyAddresses'='SMTP:'+$msoluser.UserPrincipalName}
        # add remotemailbox object, RemoteRoutingAddress starts out wrong! needs to be fixed to dpaw.mail.onmicrosoft.com, once proxyaddresses updates
        Enable-RemoteMailbox -Identity $msoluser.UserPrincipalName -RemoteRoutingAddress $msoluser.UserPrincipalName
    }

    ForEach ($mb in Get-RemoteMailbox -ResultSize Unlimited | Where {-not ($_.RemoteRoutingAddress -like "*@dpaw.mail.onmicrosoft.com" )}) {
        $remote = $mb.EmailAddresses.SmtpAddress | Where {$_ -like "*@dpaw.mail.onmicrosoft.com"} | Select -First 1;
        If ($remote) {
            $mb | Set-RemoteMailbox -RemoteRoutingAddress $remote;
        }
    }

    Log "Finished"
} catch [System.Exception] {
    Log "ERROR: Exception caught, dying =("
    $except = $_
    Log $($except | convertto-json)
}

# cleanup
Get-PSSession | Remove-PSSession
