<#
.SYNOPSIS
   Creates a Hash of a file and returns the hash  
.DESCRIPTION
    Uses System.Security.Cryptography.HashAlgorithm and members to create the hash
	This scriptimproves on http://pshscripts.blogspot.com/2009/01/get-hashpsm1.html
	by using file stream.
.NOTES
    File Name  : Get-Hash2.PS1
	Author     : Thomas Lee - tfl@psp.co.uk
	Requires   : PowerShell V2 CTP3
	Thanks to the #PowerShell Twitter Posse (PTP) for help figuring out
	the -verbose parameter. And thanksto the PTP for comments on the
	earlier version of this script, which now uses a file stream as
	input to the hash alghorithm.
.LINK
    Posted to         :  http://pshscripts.blogspot.com/2009/01/get-hash2ps1.html
	Based on          :  http://tinyurl.com/aycszb written by Bart De Smet
	An improvement of :  http://pshscripts.blogspot.com/2009/01/get-hashpsm1.html
.PARAMETER Algorithm
    The name of one of the hash Algorithms defined at
	http://msdn.microsoft.com/en-us/library/system.security.cryptography.hashalgorithm.aspx
.PARAMETER File
    The name of a file to provide a hash for.
.PARAMETER Verbose
    if specified, this script will produce chattier output.
.EXAMPLE
    PS C:\foo> ls *.txt | where {!$_.psiscontainer}| c:\foo\Get-Hash2.ps1 sha1 -verbose
	OK, I'll be chatty
    The sha1 hash of file C:\foo\asciifile.txt is: "55055a5c8eeb3af7fa6d426314578ee1d56df016"

    The sha1 hash of file C:\foo\log.txt is: "575f4b35e3cadb9b273095fc463bd43e9a3f5774"

    The sha1 hash of file C:\foo\sites.txt is: "8ce6663cd2b64a513cf18a831843afd98e190764"

    The sha1 hash of file C:\foo\test.txt is: "a2f26abbeeb4e6846e159ba506e07cae5496d458"

    The sha1 hash of file C:\foo\test2.txt is: "9b1baaa9077a3691f8e2685d81ffa24cdd73f71d"

    The sha1 hash of file C:\foo\unicodefile.txt is: "094ef2696d9eb3374e657eb7c227ff4c36cd0cb9"
.EXAMPLE
    PS C:\foo> ls *.txt | where {!$_.psiscontainer}| c:\foo\Get-Hash2.ps1 sha1
    55055a5c8eeb3af7fa6d426314578ee1d56df016
    575f4b35e3cadb9b273095fc463bd43e9a3f5774
    8ce6663cd2b64a513cf18a831843afd98e190764
    a2f26abbeeb4e6846e159ba506e07cae5496d458
    9b1baaa9077a3691f8e2685d81ffa24cdd73f71d
    094ef2696d9eb3374e657eb7c227ff4c36cd0cb9
.EXAMPLE
    PS C:\foo> Get-Hash  md5 asciifile.txt -verbose
    OK, I'll be chatty
    The md5 hash of file c:\foo\asciifile.txt is: "06f51e7bfced5c0623eec5f72e0999d6"
.EXAMPLE
    PS C:\foo> .\get-hash2 md5 c:\foo\asciifile.txt
    06f51e7bfced5c0623eec5f72e0999d6
#>
#[CmdletBinding()]
param (
[Parameter(Position=0, mandatory=$true)]
[string] $Algorithm,
[Parameter(Position=1, mandatory=$true, valuefrompipeline=$true)]
[string] $File
)

Begin { 
if ($VerbosePreference.Value__ -eq 0) {$verbose=$false} else {$verbose=$true}
if ($Verbose) {"OK, I'll be chatty"}
}

Process {

if ($VerbosePreference.Value__ -eq 0) {$verbose=$false} else {$verbose=$true}

	# Get the alghorthm object
	$Algo=[System.Security.Cryptography.HashAlgorithm]::Create($algorithm)
    if ($Algo){
	    # Open the file
		$Filemode = [System.IO.FileMode]::Open
		$Fs = New-Object System.Io.Filestream $File, $Filemode
		if ($fs.length -gt 0) {
		    # Now compute hash
		    $Hash = $Algo.ComputeHash($Fs)   
		    $Hashstring =""
		    foreach ($byte in $hash) {$hashstring += $byte.tostring("x2")}
		    # pass hash string on
		    if ($verbose){
			  "The {0} hash of file {1} is: `"{2}`"" -f $algorithm, $file, $hashstring
			  ""
			}
			else {
			 $Hashstring
			}
		}
		else {
		     if ($verbose) {"File {0} can not be hashed" -f $file ; ""}		
		}
		$fs.close()
		}
	else {"Algorithm {0} not found" -f $algorithm}
}
# SIG # Begin signature block
# MIIKEwYJKoZIhvcNAQcCoIIKBDCCCgACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUl75lgslw4mGhL/6dTT32D/0i
# QF+gggeNMIIHiTCCBnGgAwIBAgIKFovamwAAAAAAYDANBgkqhkiG9w0BAQUFADBB
# MRMwEQYKCZImiZPyLGQBGRYDY29tMRMwEQYKCZImiZPyLGQBGRYDS0NTMRUwEwYD
# VQQDEwxLQ1MtTUlULTItQ0EwHhcNMTEwNTE2MTkyNjQ0WhcNMTIwNTE1MTkyNjQ0
# WjCBoTETMBEGCgmSJomT8ixkARkWA2NvbTETMBEGCgmSJomT8ixkARkWA0tDUzEM
# MAoGA1UECxMDS0NTMQ0wCwYDVQQLEwRNcGxzMQ0wCwYDVQQLEwRUZWNoMRgwFgYD
# VQQDEw9CZXJuYXJkIFdlbG1lcnMxLzAtBgkqhkiG9w0BCQEWIGJ3ZWxtZXJzQEtu
# dXRzb25Db25zdHJ1Y3Rpb24uY29tMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIB
# CgKCAQEAr1cpQk1fBTCKh6OnyH+MpmGqJh2oYMj/gBQpYQttzk0E+ZwUhLhxI9d0
# T/P+VwXV3SjNUjpQo9ffULuDZe9II3cxUCvA9zi50UdxmFtuqoHkadepjgJbvmH8
# dT6vs7YgXtoHuO+ks2Vz3/oU7I7uCIroVhAszWN0mDzAvObYnFpSxgti3fIbJ8fh
# cG3/tOsKvPrrIxGKqiP3WT09QzpMuUHXLHFin5mvb+ELmxPLPtMBmdOhUNqR4Dtm
# CZtVVFZqsGRO3H2UCvvYNlMekue4qrRcckWxtqrabv7xCRc2/Ku9LOhagTYwNU4L
# XF8ZC8jTAyhwvWRaWcv9R8f6e3/OewIDAQABo4IEIDCCBBwwPAYJKwYBBAGCNxUH
# BC8wLQYlKwYBBAGCNxUIh8udUMeqLIPtgxyEycw4g/z/N2WC29sjgeD/ZgIBZgIB
# ADCBgAYDVR0lBHkwdwYKKwYBBAGCNwoDCwYJKwYBBAGCNxUGBggrBgEFBQcDAwYK
# KwYBBAGCNwoFAQYJKwYBBAGCNxUFBgsrBgEEAYI3CgMEAQYEVR0lAAYIKwYBBQUH
# AwIGCCsGAQUFBwMEBgorBgEEAYI3CgMEBgorBgEEAYI3CgMBMA4GA1UdDwEB/wQE
# AwIFoDCBngYJKwYBBAGCNxUKBIGQMIGNMAwGCisGAQQBgjcKAwswCwYJKwYBBAGC
# NxUGMAoGCCsGAQUFBwMDMAwGCisGAQQBgjcKBQEwCwYJKwYBBAGCNxUFMA0GCysG
# AQQBgjcKAwQBMAYGBFUdJQAwCgYIKwYBBQUHAwIwCgYIKwYBBQUHAwQwDAYKKwYB
# BAGCNwoDBDAMBgorBgEEAYI3CgMBMIGUBgkqhkiG9w0BCQ8EgYYwgYMwCwYJYIZI
# AWUDBAEqMAsGCWCGSAFlAwQBLTALBglghkgBZQMEARYwCwYJYIZIAWUDBAEZMAsG
# CWCGSAFlAwQBAjALBglghkgBZQMEAQUwCgYIKoZIhvcNAwcwBwYFKw4DAgcwDgYI
# KoZIhvcNAwICAgCAMA4GCCqGSIb3DQMEAgICADBNBgNVHREERjBEoCAGCisGAQQB
# gjcUAgOgEgwQYndlbG1lcnNAS0NTLkNPTYEgYndlbG1lcnNAS251dHNvbkNvbnN0
# cnVjdGlvbi5jb20wHQYDVR0OBBYEFKoJg+2H47QpTpXGPtgThnC+ggiEMB8GA1Ud
# IwQYMBaAFCnyEZpR0YWad+WZKZDnMqWeQxGVMIHEBgNVHR8EgbwwgbkwgbaggbOg
# gbCGga1sZGFwOi8vL0NOPUtDUy1NSVQtMi1DQSxDTj1NSVQtMixDTj1DRFAsQ049
# UHVibGljJTIwS2V5JTIwU2VydmljZXMsQ049U2VydmljZXMsQ049Q29uZmlndXJh
# dGlvbixEQz1LQ1MsREM9Y29tP2NlcnRpZmljYXRlUmV2b2NhdGlvbkxpc3Q/YmFz
# ZT9vYmplY3RDbGFzcz1jUkxEaXN0cmlidXRpb25Qb2ludDCBugYIKwYBBQUHAQEE
# ga0wgaowgacGCCsGAQUFBzAChoGabGRhcDovLy9DTj1LQ1MtTUlULTItQ0EsQ049
# QUlBLENOPVB1YmxpYyUyMEtleSUyMFNlcnZpY2VzLENOPVNlcnZpY2VzLENOPUNv
# bmZpZ3VyYXRpb24sREM9S0NTLERDPWNvbT9jQUNlcnRpZmljYXRlP2Jhc2U/b2Jq
# ZWN0Q2xhc3M9Y2VydGlmaWNhdGlvbkF1dGhvcml0eTANBgkqhkiG9w0BAQUFAAOC
# AQEAhbXQyosWme6+WnuQdt2rYI1VaGatX9qIfBRW7MVuRZUpefnJfp11jI0LIHl1
# aJe0/1AcGwiLTRUoNLxmCY2JuFPhNAw1o6iVI1GQ8ZMBAhSydH8tIymeyKI/ywDv
# ztPYJ0/u30Cl6evKlPg1xcfTYf1ndLw4yMjvjRNyL3cokRP0bSPFM2600WajlIOZ
# a+NIcuG95HHBnDLJ8iBX355WkFPmRJXDYiRxrVj7nwWIfxXZrlxWLNI03S/1YNg2
# nxX5I2MBCpoxF1Ga5RX5N/V0IyPzvxqaFtf/ivYXgqlTQd5cq5mTVepsQT8ns7zw
# XfckrQa4i1SMRamDlGJoR+XGvjGCAfAwggHsAgEBME8wQTETMBEGCgmSJomT8ixk
# ARkWA2NvbTETMBEGCgmSJomT8ixkARkWA0tDUzEVMBMGA1UEAxMMS0NTLU1JVC0y
# LUNBAgoWi9qbAAAAAABgMAkGBSsOAwIaBQCgeDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBT0iq9Ib7yND2bKcBM/wO+2
# 5/HxxjANBgkqhkiG9w0BAQEFAASCAQAHOjLrJPmdR8Db398PuyzEN7Std0xn1WdO
# tDJyyN5rCE0yyNzx6Ol2YCdb0SQOARw2Z+vTch4We+Y92ZC9lwwCpjelcY9N6wGl
# Zm4I67qvCFPZKZcu7/8mkPWtoSOl/wZnrRl6ICJgGBrux6ZzWMLgjArgTbTh95a5
# xtc8WPyhpkbScisz4L5zfNCxhkapRbZb7G6nuM1H/EH3hWbWfBjWX2+ZMX/VEztV
# 9hZGQepzG0eHF8HfKu8jpSBwon91xf21Yw+8e9ZNXq/AESizJZXF8hKaC/e33uVv
# z4flvnuHI/qZSQPags2NUw10R/ggojkuaBlG6rLzCDWvtKkXEYDW
# SIG # End signature block
