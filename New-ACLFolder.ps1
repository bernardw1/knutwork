<#

.SYNOPSIS
 Created by Bernard Welmers
 On March 20, 2017

 This script was created to automate the creation of Folders on networks share and the associated ACL list.

.DESCRIPTION
.EXAMPLE
.NOTES
 March 20, 2017 - Initial creation
 #>

param 
(
        [Parameter(Mandatory=$true)] [String]$GroupName,
        [String]$GroupDescription,
        [String]$GroupDisplayName,
        [ValidateSet("Rochester","Iowa City","Minneapolis","Corporate")] [string]$Office,
        [Parameter(Mandatory=$true)] [ValidateScript ({Test-Path $_})] [String]$ShareLocation,
        [Parameter(Mandatory=$true)] [ValidateScript ({Get-ADUser $_})] [String]$GroupOwner
)

$FolderPath = $ShareLocation + "\" + $GroupName

if(!(Test-Path -Path $FolderPath)) {New-item -ItemType directory -Path $FolderPath }

if ($Office -eq "Rochester")
    {
    $ACLOffice = "Roch"
    }
Elseif ($Office -eq "Iowa City")
    {
    $ACLOffice = "Iowa"
    }
Else
    {
    $ACLOffice = "MPLS"
    }

if (!$GroupDisplayName) {$GroupDisplayName = $GroupName}
$ACLName = (Get-Culture).TextInfo.ToTitleCase($GroupName) -replace " ",''

Add-NTFSAccess -Path $FolderPath -Account "kcs\bwelmers" -AccessRights FullControl


$ACLName1 = "ACL-"+$ACLOffice +"_"+$ACLName+"-Full"
New-adGroup -name $ACLName1 -SamAccountName $ACLName1 -GroupCategory Security -GroupScope Universal -Description $GroupDescription  -DisplayName $GroupDisplayName -Path 'OU=SecurityGroups,OU=Distro and Security Groups,DC=KCS,DC=com' -ManagedBy $GroupOwner
Add-NTFSAccess -Path $FolderPath -Account "kcs\$ACLName1" -AccessRights FullControl
Add-ADGroupMember -Identity $ACLName1 -Members "Domain Admins","IT Support"


$ACLName1 = "ACL-"+$ACLOffice +"_"+$ACLName+"-Write"
New-adGroup -name $ACLName1 -SamAccountName $ACLName1 -GroupCategory Security -GroupScope Universal -Description $GroupDescription  -DisplayName $GroupDisplayName -Path 'OU=SecurityGroups,OU=Distro and Security Groups,DC=KCS,DC=com' -ManagedBy $GroupOwner
Add-NTFSAccess -Path $FolderPath -Account "kcs\$ACLName1" -AccessRights Modify
Add-ADGroupMember -Identity $ACLName1 -Members $GroupOwner

$ACLName1 = "ACL-"+$ACLOffice +"_"+$ACLName+"-Read"
New-adGroup -name $ACLName1 -SamAccountName $ACLName1 -GroupCategory Security -GroupScope Universal -Description $GroupDescription  -DisplayName $GroupDisplayName -Path 'OU=SecurityGroups,OU=Distro and Security Groups,DC=KCS,DC=com' -ManagedBy $GroupOwner
Add-NTFSAccess -Path $FolderPath -Account "kcs\$ACLName1" -AccessRights Read

$acl = get-acl $FolderPath
    $acl.setowner([System.Security.Principal.NTAccount] $GroupOwner)
    set-acl $FolderPath $acl

Disable-NTFSAccessInheritance -Path $FolderPath -RemoveInheritedAccessRules
# SIG # Begin signature block
# MIIJXQYJKoZIhvcNAQcCoIIJTjCCCUoCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUgFBFvLRfQBNi5Pnar4KutN3K
# EjGgggbOMIIGyjCCBbKgAwIBAgITGgAAAAtIEqyWvVA2zAADAAAACzANBgkqhkiG
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
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUdx2LGXVX
# wFxBpxDNmaPn+CU19acwDQYJKoZIhvcNAQEBBQAEggEARz19Zg5zpl7/pZu1fKtO
# ecHEL0eeKixqc7CVEiWxbT6Y1R5P3OIuJjPC9/Xzdr9OHiT4mwS/YnsbBfB0fVxF
# JCigyCoxponLj2g5J9p6WVh6kFsZ7D1DMWhTWlhZ2nhcwKoEY6iaCNkNp80ER4f6
# eKPXiXUf1bcdFgusbRPJfmzI+MsR20H4na2tXHvtpJe7QB6V6veWb9tHbC//Uo78
# Li2F9sLLyp9K6fJ72mP32wWg9bC8xs/xXDIn/cOwVd1Mp3UxXGSujdHr5IxFNd1c
# HW81nJ1Me85DIsXamm2evwVvfBDYPSyRAhww6m9nconiy+XivEhimyJRy0Clo4qd
# vQ==
# SIG # End signature block
