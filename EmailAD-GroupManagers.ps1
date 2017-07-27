#this script will locate all the AD groups (Secuirty and distrbution) that are located under the Security and distribution OU. The script will gather all the groups, then all the managers of the groups, and then it will email all those managers the members of their groups. At this time it will all be individual emails. Converting it into single emails with all the group information could be a future improvement of the process.

<#
This first section gives some methods to make all the groups editable by the manageby person

get-adgroup -searchbase 'OU=SecurityGroups,OU=Distro and Security Groups,DC=KCS,DC=com' -properties managedby -filter {name -notlike '*Full'} | select name,managedby |out-file c:\temp\acl.txt

$mgr = get-aduser bwelmers
$group = get-adgroup -searchbase 'OU=SecurityGroups,OU=Distro and Security Groups,DC=KCS,DC=com' -properties managedby -filter {name -like '*Full'}
$group | Set-ADGroup -managedby "$($mgr.DistinguishedName)"
get-adgroup -searchbase 'OU=SecurityGroups,OU=Distro and Security Groups,DC=KCS,DC=com' -properties managedby -filter {name -like '*Full'} | select name,managedby

$groupmanagers = get-adgroup -searchbase 'OU=SecurityGroups,OU=Distro and Security Groups,DC=KCS,DC=com' -properties managedby -filter {name -notlike '*Full'} | select name,managedby
#>

<#
#Check to see if a connection MS Exchagne (MEX2) has already been established in this PowerShell Session
if (!(Get-PSSession | Where { $_.ConfigurationName -eq "Microsoft.Exchange" })) 
{ 
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://mex2.kcs.com/PowerShell/ -Authentication Kerberos
    
    Import-PSSession $Session -ErrorAction Stop 
}
#>

# when converting array to string enter items on new liens
$ofs = "`r`n"

$ListofGroups = get-adgroup -searchbase 'OU=Distro and Security Groups,DC=KCS,DC=com' -properties managedby,description -filter {name -like '*Certified*'} | select name,managedby,description

foreach ($groupName in $listofgroups)
{
	#get managedby email address
	$groupLeader = Get-ADUser $groupname.managedby -property emailaddress | select name,emailaddress
	if (-not $groupLeader.emailaddress){$emailLeader = 'bwelmers@knutsonconstruction.com'} Else {$emailLeader = $groupLeader.emailaddress + ' ; bwelmers@knutsonconstruction.com'}
	
	#get the group members
	$groupmembers = Get-ADGroupMember -Identity $groupname.name -Recursive | select name

	#generate an email address to the group manager
	$subject = "Email/Security group membership checkup - " + $groupname.Name 
	$messagebody = "Hello, You have been identified as the manager of the " + $groupname.Name + " group. As the manger of this group the IT department is asking you to please look over the following list of people that are members of the group and make sure that only the correct people have access to this group. If you find anyone in the list that should not have access to these resources please let us know via an email to helpdesk@knutsonconstruction.com. If the group name starts with ACL that means the group is a scurity group and determins who can have read or write/modify access to that resource otherwise it is an email distribution group.  `r`n `r`n The description of the group is `r`n " + $groupname.description + "`r`n `r`n The group is setup to be managed by " + $groupleader.name + "`r`n `r`n The members of the group are: `r`n"+ [string]$groupmembers.name

	send-mailmessage -to $emailLeader -subject $subject -body $messagebody -smtpserver mail.knutsonconstruction.com -from helpdesk@knutsonconstruction.com
	
	#clear varriables for next run
	$emailLeader = ""
	$groupLeader = ""
	$groupmembers = ''
	$subject = ''
	$messagebody = ''
	
	<# Possible addition for future use when switching over to Office 365
	
	## Define Email Settings

	$smtpServer = "smtprelay.yours.com"
	$smtpFrom = "anyemailaddress@something.com"
	$smtpTo = "helpdesk@somewhere.com"
	$TimestampEndofReport = get-date -Format g
	$subject = "Weekend Update Results: $TimestampEndofReport"
	#
	## Format Email Body
	$Emailbody1 = "Servers that Require Reboot: $CautionArrayCount , $CautionArray"
	$Emailbody2 = "|"
	$Emailbody3 = "Servers that DNR Reboot: $PassArrayCount , $PassArray"
	$Emailbody4 = "|"
	$Emailbody5 = "Servers that could not be pinged: $FailArrayCount , $FailArray"
	$Emailbody6 = "|"
	$Emailbody7 = "Log in to $env:COMPUTERNAME to execute or terminate script."
	$Emailbody += $Emailbody1, $Emailbody2, $Emailbody3, $Emailbody4, $Emailbody5, $Emailbody6, $Emailbody7
	#
	## Send Email
	$smtp = new-object Net.Mail.SmtpClient($smtpServer)
	$smtp.Send($smtpFrom, $smtpTo, $subject, $Emailbody)
	Write-Host "Email sent to helpdesk@somewhere.com " -BackgroundColor DarkGreen

	#>
	
	
}
# SIG # Begin signature block
# MIIJXQYJKoZIhvcNAQcCoIIJTjCCCUoCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUXEILZrgIDHuBArtlVu3kVU5e
# dGGgggbOMIIGyjCCBbKgAwIBAgITGgAAAAtIEqyWvVA2zAADAAAACzANBgkqhkiG
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
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU40amkX2O
# jKod7fO5UUUswOh/ANUwDQYJKoZIhvcNAQEBBQAEggEAXW/7zw+r6PfsZcUP+kVe
# CgIImMHCIvhcqR+lw4O9zkmbnJ3PwDkYCRQD6ytC7n+JhdEnED9p/a8AHt7/+d3P
# G9Df3rdd41ixpNqHghXz3J/3U21HpLaM3UfA9QYWfRrFxBM5QmtZqq2XcBlvr/mO
# CRsK1aDaKmnt4sl30br7UeCPZd7ATbHagjKCJbmDT0eOLsfhOx70bcT/wIxY1oiY
# bmhxe4Yh1vLdWMs804AH3QbvBETgUmS3OeYLDoJFHOSqAJuOlFHQBkjUvqH2MQsF
# ZmQ49PavVaa5QWH/zihP1B6DmqvviI9jKaKjfM60iji/0PQ7TzBngZEF+vu0rblu
# 2w==
# SIG # End signature block
