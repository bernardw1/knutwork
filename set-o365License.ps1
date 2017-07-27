# Parameter help description
Param(
[Parameter(Mandatory=$true, ValueFromPipeline = $true)] $UPN,
[ValidateSet("Default","Field","ServiceAccount","ConferenceCall","FullE3","RemoveConferenceCall","RemoveE3","RemoveK1")] $licenseLevel = "Default")

$TranscriptFile = "c:\temp\O365License.txt"
$UsageLocation = "US"
$LicenseToApply = "knutsonconstruction:ENTERPRISEPACK"


if (!(Test-path -path c:\temp))
{
    new-item -Path c:\ -Value temp -ItemType Directory
}
Start-Transcript -Path $TranscriptFile



<#
.SYNOPSIS
Check for installed modules and install them
.DESCRIPTION
This is a function that allows checking of the modules that are installed. If the module is not installed then it will install it

.PARAMETER mod
Pass the module name to check to see if it is installed

.EXAMPLE
checkmodlue ('AzureAD')

.NOTES
created 7/25/2017
#>
Function CheckModule ($mod)
#Check if module is installed
{    if(-not(Get-Module -name $mod)) 	{ 
        write-progress -Status "Getting all the pieces ready to start working" -activity "loading the $mod module" -Id 50
        if(Get-Module -ListAvailable | Where-Object { $_.name -eq $mod }) { 
            Import-Module -Name $mod 
        } #end if module available then import 
        else { 
            Write-Warning -Message "Please install the $mod module. This will be attempted now."
            Install-module  -name $mod -Repository PSGallery -Scope CurrentUser -confirm:$False
            if(Get-Module -ListAvailable | Where-Object { $_.name -eq $mod }) { 
                Import-Module -Name $mod 
                Write-progress -Status 'The module has now been installed correctly' -Activity "" -Id 50
            } #end if module available then import 
        } #module not available 
    } # end if not module 
}


#$cred = Get-Credential -UserName "bwelmers@knutsonconstruction.com" -message "login information to Office 365"
$Cred = \\kcs.com\it\installs\1itscripts\Get-MyCredential.ps1 

CheckModule ("msonline")
CheckModule ("azuread")

#connect to MS Online
Import-Module MsOnline
Import-module azuread
$z = Connect-MsolService -Credential $cred
$z = Connect-AzureAD -Credential $cred

<# if (!(Get-PSSession | Where { ($_.ConfigurationName -eq "Microsoft.Exchange") -and ($_.ComputerName -ne "mex2.kcs.com")  }))
{
    #$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $Cred -Authentication Basic –AllowRedirection -prefix cc
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Cred -Authentication Basic -AllowRedirection
    Import-PSSession $Session  -AllowClobber
} #>

try {$a = Get-AzureADUser -ObjectId $upn}
catch {Write-Error "User does not exist, please try again"
exit
}

<#
AccountSkuId                                  ActiveUnits WarningUnits ConsumedUnits
------------                                  ----------- ------------ -------------
knutsonconstruction:ENTERPRISEPACK            197         0            11
knutsonconstruction:DESKLESSPACK              30          0            0
knutsonconstruction:MICROSOFT_BUSINESS_CENTER 10000       0            3
knutsonconstruction:MCOMEETADV                5           0            0
knutsonconstruction:EXCHANGEARCHIVE           26          0            0
knutsonconstruction:EOP_ENTERPRISE            30          0            0
#>


$disabledPlans= @()

Switch ($licenseLevel)
{
    FullE3
        {
        #Give the user access the full E3 License
        $LicenseToApply = "knutsonconstruction:ENTERPRISEPACK"
        If ((Get-MsolUser -UserPrincipalName $upn).licenses.accountskuid -contains "knutsonconstruction:DESKLESSPACK") {
            Set-MsolUserLicense -UserPrincipalName $upn -RemoveLicenses "knutsonconstruction:DESKLESSPACK"
        }

        $disabledPlans = $Null
        $ServiceLevel = New-msollicenseOptions -AccountSkuId $LicenseToApply -DisabledPlans $disabledPlans
        }
    Field
        {
        #set the user to have all the additional "small" licesnes
        $LicenseToApply = @("knutsonconstruction:DESKLESSPACK")
        If ((Get-MsolUser -UserPrincipalName $upn).licenses.accountskuid -contains "knutsonconstruction:ENTERPRISEPACK") {
            Set-MsolUserLicense -UserPrincipalName $upn -RemoveLicenses "knutsonconstruction:ENTERPRISEPACK"
        }


        $disabledPlans += "FORMS_PLAN_K"
        $disabledPlans += "STREAM_O365_K"
        $disabledPlans += "FLOW_O365_S1"
        $disabledPlans += "POWERAPPS_O365_S1"
        $disabledPlans += "TEAMS1"
        $disabledPlans += "Deskless"
        $disabledPlans += "MCOIMP"
        $disabledPlans += "SHAREPOINTWAC"
        $disabledPlans += "SWAY"
        $disabledPlans += "INTUNE_O365"
        $disabledPlans += "YAMMER_ENTERPRISE"
        $disabledPlans += "SHAREPOINTDESKLESS"
        #$disabledPlans += "EXCHANGE_S_DESKLESS"
        $ServiceLevel = New-msollicenseOptions -AccountSkuId "knutsonconstruction:DESKLESSPACK" -DisabledPlans $disabledPlans
        }
    ServiceAccount
        {
        #Give the account access to the K1 license but turn off everything but email
        $LicenseToApply = "knutsonconstruction:DESKLESSPACK"
        If ((Get-MsolUser -UserPrincipalName $upn).licenses.accountskuid -contains "knutsonconstruction:ENTERPRISEPACK") {
            Set-MsolUserLicense -UserPrincipalName $upn -RemoveLicenses "knutsonconstruction:ENTERPRISEPACK"
        }

        $disabledPlans += "FORMS_PLAN_K"
        $disabledPlans += "STREAM_O365_K"
        $disabledPlans += "FLOW_O365_S1"
        $disabledPlans += "POWERAPPS_O365_S1"
        $disabledPlans += "TEAMS1"
        $disabledPlans += "Deskless"
        $disabledPlans += "MCOIMP"
        $disabledPlans += "SHAREPOINTWAC"
        $disabledPlans += "SWAY"
        $disabledPlans += "INTUNE_O365"
        $disabledPlans += "YAMMER_ENTERPRISE"
        $disabledPlans += "SHAREPOINTDESKLESS"
        #$disabledPlans += "EXCHANGE_S_DESKLESS"
        $ServiceLevel = New-msollicenseOptions -AccountSkuId $LicenseToApply -DisabledPlans $disabledPlans
        }
    ConferenceCall
        {
        #Give the user access to the Skype for Business Conference calling abilities
        $LicenseToApply = "knutsonconstruction:MCOMEETADV"
        $disabledPlans = $Null
        $ServiceLevel = New-msollicenseOptions -AccountSkuId $LicenseToApply -DisabledPlans $disabledPlans
        }        
    Default
        {
        $LicenseToApply = "knutsonconstruction:ENTERPRISEPACK"
        If ((Get-MsolUser -UserPrincipalName $upn).licenses.accountskuid -contains "knutsonconstruction:DESKLESSPACK") {
            Set-MsolUserLicense -UserPrincipalName $upn -RemoveLicenses "knutsonconstruction:DESKLESSPACK"
        }

        #This is a list of all the current (6/20/2017) Service Plan options for an E3 license
        $disabledPlans += "FORMS_PLAN_E3"   #create quizes and surveys
        $disabledPlans += "STREAM_O365_E3"  #Video
        $disabledPlans += "Deskless"        #staff hub
        $disabledPlans += "FLOW_O365_P2"    #workflow
        $disabledPlans += "POWERAPPS_O365_P2" 
        $disabledPlans += "TEAMS1"          
        $disabledPlans += "PROJECTWORKMANAGEMENT" #MS Planner
        $disabledPlans += "SWAY"            #website
        $disabledPlans += "INTUNE_O365"     #MDM for O365
        $disabledPlans += "YAMMER_ENTERPRISE"
        $disabledPlans += "RMS_S_ENTERPRISE"    #Azure Rights Mangement
        #$disabledPlans += "OFFICESUBSCRIPTION"  #Office 365 pro plus
        $disabledPlans += "MCOSTANDARD"         #Skype for Business online (Plan2)
        $disabledPlans += "SHAREPOINTWAC"       #Office online must also enable SharepoineEnterprise
        $disabledPlans += "SHAREPOINTENTERPRISE"    #Sharepoint Online (plan 2) - also One Drive for Business
        #$disabledPlans += "EXCHANGE_S_ENTERPRISE"   Exchange online (plan 2)
        #end
        $ServiceLevel = New-msollicenseOptions -AccountSkuId $LicenseToApply -DisabledPlans $disabledPlans
        Write-verbose "minimum plan set"
        }
    RemoveConferenceCall
        {
            If ((Get-MsolUser -UserPrincipalName $upn).licenses.accountskuid -notcontains "knutsonconstruction:MCOMEETADV") {
                Set-MsolUserLicense -UserPrincipalName $upn -RemoveLicenses "knutsonconstruction:MCOMEETADV" 
                }
            exit
        }
    RemoveE3
        {
            If ((Get-MsolUser -UserPrincipalName $upn).licenses.accountskuid -notcontains "knutsonconstruction:ENTERPRISEPACK") {
                Set-MsolUserLicense -UserPrincipalName $upn -RemoveLicenses "knutsonconstruction:ENTERPRISEPACK"
            }
            exit
        }
    RemoveK1
        {
            If ((Get-MsolUser -UserPrincipalName $upn).licenses.accountskuid -notcontains "knutsonconstruction:DESKLESSPACK") {
                Set-MsolUserLicense -UserPrincipalName $upn -RemoveLicenses "knutsonconstruction:DESKLESSPACK"
            }
            exit
        }
}

# check to see if user is currently licensed for the license that is selected to apply to them. If no then add the license
if ((!(Get-MsolUser -UserPrincipalName $upn).Licenses) -or ((get-msoluser -UserPrincipalName $upn).IsLicensed -eq $false) -or ((Get-MsolUser -UserPrincipalName $upn).licenses.accountskuid -notcontains $LicenseToApply))
{
    Set-MsolUser -UserPrincipalName $upn -UsageLocation "US"
    Set-MsolUserLicense -UserPrincipalName $upn  -LicenseOptions $ServiceLevel -AddLicenses $LicenseToApply 
    Write-verbose "added license - $LicenseToApply"
} elseif ((Get-MsolUser -UserPrincipalName $upn).licenses.accountskuid -contains $LicenseToApply) {
    #If user is already licensed then just change the service levels
    Set-MsolUserLicense -UserPrincipalName $upn  -LicenseOptions $ServiceLevel
    Write-verbose 'service level changed (was already licensed)'
}


if (!(Get-PSSession | Where { ($_.ConfigurationName -eq "Microsoft.Exchange") -and ($_.ComputerName -ne "mex2.kcs.com") })) 
{
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Cred -Authentication Basic -AllowRedirection
    Import-PSSession $Session  -AllowClobber
}


#If the user is licensed for Exchange Online Plan 2 then enable the litigation hold.
$LicenseList = (Get-MsolUser -UserPrincipalName $upn).licenses.ServiceStatus
foreach ($b in $LicenseList) {
    if (($b.ServicePlan.ServiceName -eq "EXCHANGE_S_ENTERPRISE") -and ($b.ProvisioningStatus -eq "Success" )) {
    #start litigation hold if account has Exchange online Plan 2
        Get-Mailbox -ResultSize Unlimited -Filter {RecipientTypeDetails -eq "UserMailbox" -and LitigationHoldDate -eq $null} | Set-Mailbox -LitigationHoldEnabled $true -LitigationHoldDuration unlimited
    } 
}   

Write-Output "user account applied $upn"
Write-Output "final license level "  
(Get-MsolUser -UserPrincipalName $UPN).Licenses
Write-Output "final service levels "  
(Get-MsolUser -UserPrincipalName $upn).licenses.servicestatus |FT


Stop-Transcript


# SIG # Begin signature block
# MIIJXQYJKoZIhvcNAQcCoIIJTjCCCUoCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUBa5FAkxnLPUWiA1PkyiCm0Tu
# 2K2gggbOMIIGyjCCBbKgAwIBAgITGgAAAAtIEqyWvVA2zAADAAAACzANBgkqhkiG
# 9w0BAQsFADBBMRMwEQYKCZImiZPyLGQBGRYDY29tMRMwEQYKCZImiZPyLGQBGRYD
# S0NTMRUwEwYDVQQDEwxLQ1MtTUlULTItQ0EwHhcNMTcwNDExMTM1MTEyWhcNMTkw
# NDExMTQwMTEyWjCBoTETMBEGCgmSJomT8ixkARkWA2NvbTETMBEGCgmSJomT8ixk
# ARkWA0tDUzEMMAoGA1UECxMDS0NTMQ0wCwYDVQQLEwRNcGxzMQ0wCwYDVQQLEwRU
# ZWNoMRgwFgYDVQQDEw9CZXJuYXJkIFdlbG1lcnMxLzAtBgkqhkiG9w0BCQEWIGJ3
# ZWxtZXJzQEtudXRzb25Db25zdHJ1Y3Rpb24uY29tMIIBIjANBgkqhkiG9w0BAQEF
# AAOCAQ8AMIIBCgKCAQEAysr13O2miHzeNCjHCXDxgHbooJ6rjytjCiaYn6COz8Th
# aJtxZpdnnPUn3DNazFhYrs8rh972MgB+Rp+CZOSbqf1NVEpUqLmoU9kB4unfv/yV
# Lmo9IH7JgFjL29cdvDASM2pQsw37x3YvnTNCLf2Vz+HIaLhvyMwoGD3JZ/XamBmb
# FqgEyIHJcn8OclmY1X2+zk9hIsq/AzMVTcTLLMIwV9VtNTfjekSQXjUGxFFmpJUr
# GtBoXWXOVjwN0Smc0jxbvzUFqwFoLZQlBesfOW4ZhqgwbwmBQiMV33M8O/lY3SsG
# 8GAgw3mwAqskVXwNQrbxGHpAfaR0qQZlP7KqzyUYnQIDAQABo4IDWDCCA1QwOwYJ
# KwYBBAGCNxUHBC4wLAYkKwYBBAGCNxUIh8udUMeqLIPtgxyEycw4g/z/N2WJkAaH
# 0ddNAgFkAgEIMIGABgNVHSUEeTB3BgorBgEEAYI3CgMBBgorBgEEAYI3CgMEBggr
# BgEFBQcDBAYIKwYBBQUHAwIGBFUdJQAGCisGAQQBgjcKAwsGCSsGAQQBgjcVBgYI
# KwYBBQUHAwMGCisGAQQBgjcKBQEGCSsGAQQBgjcVBQYLKwYBBAGCNwoDBAEwDgYD
# VR0PAQH/BAQDAgWgMIGeBgkrBgEEAYI3FQoEgZAwgY0wDAYKKwYBBAGCNwoDATAM
# BgorBgEEAYI3CgMEMAoGCCsGAQUFBwMEMAoGCCsGAQUFBwMCMAYGBFUdJQAwDAYK
# KwYBBAGCNwoDCzALBgkrBgEEAYI3FQYwCgYIKwYBBQUHAwMwDAYKKwYBBAGCNwoF
# ATALBgkrBgEEAYI3FQUwDQYLKwYBBAGCNwoDBAEwgZQGCSqGSIb3DQEJDwSBhjCB
# gzALBglghkgBZQMEASowCwYJYIZIAWUDBAEtMAsGCWCGSAFlAwQBFjALBglghkgB
# ZQMEARkwCwYJYIZIAWUDBAECMAsGCWCGSAFlAwQBBTAKBggqhkiG9w0DBzAHBgUr
# DgMCBzAOBggqhkiG9w0DAgICAIAwDgYIKoZIhvcNAwQCAgIAMB0GA1UdDgQWBBSc
# 75qxq/MSvwpMIB6+ZvTsUWKVfzAfBgNVHSMEGDAWgBQM20wsOgiMIely8XbQ7KPV
# z1RkEzCBugYIKwYBBQUHAQEEga0wgaowgacGCCsGAQUFBzAChoGabGRhcDovLy9D
# Tj1LQ1MtTUlULTItQ0EsQ049QUlBLENOPVB1YmxpYyUyMEtleSUyMFNlcnZpY2Vz
# LENOPVNlcnZpY2VzLENOPUNvbmZpZ3VyYXRpb24sREM9S0NTLERDPWNvbT9jQUNl
# cnRpZmljYXRlP2Jhc2U/b2JqZWN0Q2xhc3M9Y2VydGlmaWNhdGlvbkF1dGhvcml0
# eTBNBgNVHREERjBEoCAGCisGAQQBgjcUAgOgEgwQYndlbG1lcnNAS0NTLmNvbYEg
# YndlbG1lcnNAS251dHNvbkNvbnN0cnVjdGlvbi5jb20wDQYJKoZIhvcNAQELBQAD
# ggEBAB3UFY+lHU+O3UvYg7CeB7Z7CMk5OloAW99AIzUw+aQPgG/4q7FyXdqtrqvR
# uBlQ37AIAct/+DxRpxDTgI/NKaAsBLMz391IAdUWXprAqPTccgRWTpsgWZs4a5E0
# Xyvqcv9Ld8ryi0b2PByjB9zWvgJKeAcwtElIZptA/PY0of52KFr9TXH2b34RyA2G
# ChGaGO0OCykTRFB7l+HZpWZiyzdIygvtN2E/KH7HS5+hIzt1HL2Mop5iao5D1a1C
# vHlE5q5TOOJx3XgqDbsxLKwKyYcLYw+JVALtiyUxX0SX5+3RqHkXub/mzsbGHY6A
# T7WrPW9gpYxCzzqU/XLGsLPjL+AxggH5MIIB9QIBATBYMEExEzARBgoJkiaJk/Is
# ZAEZFgNjb20xEzARBgoJkiaJk/IsZAEZFgNLQ1MxFTATBgNVBAMTDEtDUy1NSVQt
# Mi1DQQITGgAAAAtIEqyWvVA2zAADAAAACzAJBgUrDgMCGgUAoHgwGAYKKwYBBAGC
# NwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgor
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUbYOtAb21
# UlKc/Vf6TJMsEPCLw7owDQYJKoZIhvcNAQEBBQAEggEAyfCwpdzGy40V//EUEV5g
# LI30fMfBp94V+59H5GV4NOfnIIHpjnkG1lr/v1XVk2ATfsi6DT54tI0q8FLGZpWe
# W/1OlkuuG9/3zbtqxFhPOFyBKuw4clBUJkDVau2HGSqfMocZlh5Y5kkg7veDqfEj
# L/cR9oLfCKRKywzICYyuL2XQMO7WBm6socESMbhNrQVZ39mKsCzxrxKMdn9dr4qd
# OF3taXVhb8x6wYsG8FQRjdRw8n0BjBmlqhLLWkLiM1u1USrzUKcP+UG++2NKsbB2
# MHEPfqjk1KD3oTOzKKT4pfUQ8JnqPlogyRiZLkBu81dCMrVmhOYi4MRmh0c8x02x
# Iw==
# SIG # End signature block
