<#figure out a way to check and make sure every folder in TSHOME45 is owned by the correct person

#>


$DirCollections = @()

$dirs = Get-ChildItem -Directory -filter 'awhip'
$start = Get-Date 

Foreach ($Dir in $dirs) {
	$i +=1
	$Duration = (New-TimeSpan -Start ($start) -End (Get-Date)).totalseconds
	$TimeLeft = ($Duration/$i)*($Dirs.count - $i)

	Write-Progress -Status "Working on $($Dir.name)" -Activity "Changing ACL Owner for Directory $i of $($Dirs.Count)" -PercentComplete ($i/$($dirs.count)*100) -SecondsRemaining $timeleft -Id 100
	
	$DirCollection = "" | Select Path, OldOwner, NewOwner
	$DirCollection.path = $dir.Fullname
	$DirCollection.OldOwner = (Get-Acl $dir.FullName).owner
    $acl = Get-Acl $Dir.FullName
    $acl.SetOwner([System.Security.Principal.NTAccount]"$DIR")
    set-acl $Dir.FullName $acl
	$DirCollection.NewOwner = (Get-Acl $dir.FullName).owner
	
	$DirCollections +=  $DirCollection
}

$DirCollections

<#
Foreach ($dir in $Dirs)
{
	$DirCollection = "" | Select Path, Name, MKOwner
	$DirCollection.path = $dir.Fullname
	$DirCollection.name = $dir.name
	$DirCollection.MKOwner = (Get-Acl $dir.FullName).owner
	
	$DirCollections +=  $DirCollection
}
$DirCollections

Write-host 'Start Do'

ForEach ($Dir in $DirCollections)
{
	#Set-acl
	$Dir.DirACL.SetOwner([System.Security.Principal.NTAccount]$Dir.MKOwner) 
	$Dir.DirACL
}

<#
$group = @()
$total = 10
Foreach ($i in $total)
{
	$Group+=$i
}
#>

<#
https://community.spiceworks.com/scripts/show/1070-export-folder-permissions-to-csv-file
$OutFile = "C:\Permissions.csv"
$Header = "Folder Path,IdentityReference,AccessControlType,IsInherited,InheritanceFlags,PropagationFlags"
Del $OutFile
Add-Content -Value $Header -Path $OutFile 

$RootPath = "C:\Test"

$Folders = dir $RootPath -recurse | where {$_.psiscontainer -eq $true}

foreach ($Folder in $Folders){
	$ACLs = get-acl $Folder.fullname | ForEach-Object { $_.Access  }
	Foreach ($ACL in $ACLs){
	$OutInfo = $Folder.Fullname + "," + $ACL.IdentityReference  + "," + $ACL.AccessControlType + "," + $ACL.IsInherited + "," + $ACL.InheritanceFlags + "," + $ACL.PropagationFlags
	Add-Content -Value $OutInfo -Path $OutFile
	}}


https://community.spiceworks.com/scripts/show/1722-get-fileowner
.SYNOPSIS
	Produce a simple HTML report of all files in a folder, who the owner is
	and when the file was last written.
.DESCRIPTION
	Produce a simple HTML report of all files in a folder, who the owner is
	and when the file was last written.  You can specify if you want all sub-
	folders included in the report and also narrow the report to specific
	extensions if you want.
	
	Make sure to edit the $ReportPath in the parameters section to match your
	environment (if desired).

	**CAUTION** The bigger the directory structure the longer this script will
	run.  Be patient!
	
	Script inspired by a question on Spiceworks:
	http://community.spiceworks.com/topic/283850-how-to-find-out-who-is-putting-what-on-shared-drives
.PARAMETER Path
	Specify the path where you want the report to run
.PARAMETER FileExtension
	Specify the extension that you want to narrow your search to.  Do not include
	any wildcards.  IE  jpg, xls, etc.
.PARAMETER SubFolders
	Specify whether you want the report to include all sub-folders of the 
	path.  Must use $true or $false.
.PARAMETER ReportPath
	Specify where you want the html report to be stored.  Report will be called
	Get-FileOwner.html
.OUTPUTS
	Get-FileOwner.html in the specified report path ($ReportPath parameter).
.EXAMPLE
	.\Get-FileOwner.ps1 -path s:\accounting -ReportPath s:\it
	Will create a report of all files in s:\accounting, and all sub-folders and 
	place the HTML report in the s:\it folder.
.EXAMPLE
	.\Get-FileOwner.ps1 -path s:\accounting -SubFolders $false
	Will create a report of all files in the s:\accounting folder only and place
	the report in the default location, c:\utils.
.EXAMPLE
	.\Get-FileOwner.ps1 -path s:\accounting -FileExtension xls
	Will create a report of all xls files in s:\accounting and all sub-folders.
	HTML report will be saved in the default location, c:\utils
.NOTES
	Author:       Martin Pugh
	Twitter:      @thesurlyadm1n
	Spiceworks:   Martin9700
	Blog:         www.thesurlyadmin.com
	
	Changelog:
	   1.0        Initial release
.LINK
	http://community.spiceworks.com/scripts/show/1722-get-fileowner
.LINK
	http://community.spiceworks.com/topic/283850-how-to-find-out-who-is-putting-what-on-shared-drives

Param (
	[Parameter(Mandatory=$true)]
	[string]$Path,
	[string]$FileExtension = "*",
	[bool]$SubFolders = $true,
	[string]$ReportPath = "c:\utils"
)

If ($SubFolders)
{	$SubFoldersText = "(including sub-folders)"
}
Else
{	$SubFoldersText = "(does <i>not</i> include sub-folders)"
}

$Header = @"
<style type='text/css'>
body { background-color:#DCDCDC;
}
table { border:1px solid gray;
  font:normal 12px verdana, arial, helvetica, sans-serif;
  border-collapse: collapse;
  padding-left:30px;
  padding-right:30px;
}
th { color:black;
  text-align:left;
  border: 1px solid black;
  font:normal 16px verdana, arial, helvetica, sans-serif;
  font-weight:bold;
  background-color: #6495ED;
  padding-left:6px;
  padding-right:6px;
}
td { border: 1px solid black;
  padding-left:6px;
  padding-right:6px;
}
</style>
<center>
<h1>Files by Owner</h1>
<h2>Path: $Path\*.$FileExtension $SubFoldersText</h2>
<br>
"@

If ($FileExtension -eq "*")
{	$GCIProperties = @{
		Path = $Path
		Recurse = $SubFolders
	}
}
Else
{	$GCIProperties = @{
		Path = "$Path\*"
		Include = "*.$FileExtension"
		Recurse = $SubFolders
	}
}

$Report = @()
Foreach ($File in (Get-ChildItem @GCIProperties | Where { $_.PSisContainer -eq $false }))
{
	$Report += New-Object PSObject -Property @{
		Path = $File.FullName
		Size = "{0:N2} MB" -f ( $File.Length / 1mb )
		'Created on' = $File.CreationTime
		'Last Write Time' = $File.LastWriteTime
		Owner = (Get-Acl $File.FullName).Owner
	}
}

$Report | Select Path,Size,'Created on','Last Write Time',Owner | Sort Path | ConvertTo-Html -Head $Header | Out-File $ReportPath\Get-FileOwner.html

#This line will open your web browser and display the report.  Rem it out with a # if you don't want it to
Invoke-Item $ReportPath\Get-FileOwner.html

#>
# SIG # Begin signature block
# MIIJXQYJKoZIhvcNAQcCoIIJTjCCCUoCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUqMdiIdfz9Zb9XlLfd5yW4D/B
# YDKgggbOMIIGyjCCBbKgAwIBAgITGgAAAAtIEqyWvVA2zAADAAAACzANBgkqhkiG
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
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU5qqulgxa
# 5ThVLmJAitzNjxJYOrwwDQYJKoZIhvcNAQEBBQAEggEAiqAGc0rucADl27NEfDVw
# 12x+l/+Ez5TPsQEsk5ZccJfS5wKZusU2bG4p46BDVJojBjRQA4XEibQ32Ji+GO/F
# eYakhJ8OdOug3AxwS5PxOG4c41ZASRfdCR0CChU/M0eLDFP+xvvAExQRcs9qEqKV
# RJrTkZAL8bxfED+O/aAoyN0PyCq6sFwHLn9cS6xMPvJ6xgOz0DrEkt+XtTObZlzQ
# kYdnYUKDVSJoaiFzmBIgTJz4zn1InlNm7ZfWeHPR+mvgprzv2Z4iNoppbSrQdo/m
# lZNfVb51jTSIk+HoW4b7zrmJcU1+fZu1IIET8MYOpLVItHS6zR1QLphHZ5BlISW9
# mg==
# SIG # End signature block
