<#

.Synopsis
   getLicenseGT is a small PowerShell Scipt which can be used in O365 to fetch the licenses assigned to users. Currently the script can be used to get details of
   E1, E3, K1 and EMS Licenses Assigned. Also, the script can handle cases in which users are assigned with multiple flavours of licenses like E1 clubbed with EMS, E3 clubbed with EMS etc.
   Note: If you have more flavours of licenses, please let me know the AccountSkuId or plan name so that the script can be updated.

   Developed by: Noble K Varghese

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Version 1.1, 26 June 2015
		#Initial Release
	In Future
		#Workload based License Assignment Details (Lync, SharePoint, Yammmer, Intune, AzureRMS etc.)
.DESCRIPTION
   getLicenseGT.ps1 is a PowerShell Sciprt for Office365. It helps the Admin in collecting details of licenses assigned to users. On completion, the Script creates a CSV report 
   as the output in the current working directory. This scripts supports PowerShell 2.0 & 3.0. I am using 3.0 though. You needn't connect to Exchange Online to run this script.
   A connection to MsolService is enough.

.getLicenseGT.ps1
   To Run the Script go to PowerShell and Start It. Eg: PS E:\PowerShellWorkshop> .\getLicenseGT.ps1

.Output Logs
   The Script creates a CSV report as the output in the present working directory in the format LicesneStatus_%Y%m%d%H%M%S.csv

 #>

Import-Module ActiveDirectory;
Import-Module -Force 'C:\cron\creds.psm1';

#Connect-MsolService

$Header = "UserPrincipalName, DisplayName, LicenseAssigned, Office"
$Data = @()
$OutputFile = "LicenseStatus_$((Get-Date -uformat %Y%m%d%H%M%S).ToString()).csv"
Out-File -FilePath $OutputFile -InputObject $Header -Encoding UTF8 -append

$users = Get-MSolUser -All

foreach($user in $users)
{
	$UPN = $User.UserPrincipalName
	$DisplayName = $User.DisplayName
	$Licenses = $User.Licenses.accountskuid
	
	$AccSkId = (Get-MsolAccountSku).accountskuid[0].split(":")[0]
	
	$E1Lic = $AccSkId+":STANDARDPACK"
	$E3Lic = $AccSkId+":ENTERPRISEPACK"
	$EMSLic = $AccSkId+":EMS"
	$K1Lic = $AccSkId+":EXCHANGEDESKLESS"
	
	if($Licenses.Count -eq "0")
	{
		$InLic = "User Not Licensed"
	}
	else
	{
		foreach($License in $Licenses)
		{
			if($Licenses.Count -gt 1)
			{
				if($Licenses -contains ($E1Lic) -and ($Licenses -contains ($EMSLic)))
				{
					$InLic = "E1 & EMS"
				}
				elseif($Licenses -contains ($E3Lic) -and ($Licenses -contains ($EMSLic)))
				{
					$InLic = "E3 & EMS"	
				}
				elseif($Licenses -contains ($K1Lic) -and ($Licenses -contains ($EMSLic)))
				{
					$InLic = "K1 & EMS"
				}
				else
				{
					$InLic = "Unknown License Combination"
				}
			}
			else
			{
				if($Licenses -contains ($E1Lic))
				{
					$InLic = "E1"
				}
				elseif($Licenses -contains ($E3Lic))
				{
					$InLic = "E3"
				}
				elseif($Licenses -contains ($EMSLic))
				{
					$InLic = "EMS"
				}
				elseif($Licenses -contains ($K1Lic))
				{
					$InLic = "K1"
				}
				else
				{
					$InLic = "Unknown License"
				}
			}
		}
	}
	$Office = $User.Office
	
	$Data = ($UPN + "," + $DisplayName + "," + $InLic + "," + $Office)
	
	Out-File -FilePath $OutputFile -InputObject $Data -Encoding UTF8 -append
}