[CmdletBinding()]
Param (
    [ValidateScript({Test-Path $_ -PathType Container})]
    [Parameter(Mandatory=$true)]
    [string]$Path,
    [switch]$Recurse
)

Write-Verbose "$(Get-Date): Script begins!"
Write-Verbose "Getting domain name..."
$Domain = (Get-ADDomain).NetBIOSName

Write-Verbose "Getting ACLs for folder $Path"

If ($Recurse)
{   Write-Verbose "...and all sub-folders"
    Write-Verbose "Gathering all folder names, this could take a long time on bigger folder trees..."
    $Folders = Get-ChildItem -Path $Path -Recurse | Where { $_.PSisContainer }
}
Else
{   $Folders = Get-Item -Path $Path
}

Write-Verbose "Gathering ACL's for $($Folders.Count) folders..."
ForEach ($Folder in $Folders)
{   Write-Verbose "Working on $($Folder.FullName)..."
    $ACLs = Get-Acl $Folder.FullName | ForEach-Object { $_.Access }
    ForEach ($ACL in $ACLs)
    {   If ($ACL.IdentityReference -match "\\")
        {   If ($ACL.IdentityReference.Value.Split("\")[0].ToUpper() -eq $Domain.ToUpper())
            {   $Name = $ACL.IdentityReference.Value.Split("\")[1]
                If ((Get-ADObject -Filter 'SamAccountName -eq $Name').ObjectClass -eq "group")
                {   ForEach ($User in (Get-ADGroupMember $Name -Recursive | Select -ExpandProperty Name))
                    {   $Result = New-Object PSObject -Property @{
                            Path = $Folder.Fullname
                            Group = $Name
                            User = $User
                            FileSystemRights = $ACL.FileSystemRights
                            AccessControlType = $ACL.AccessControlType
                            Inherited = $ACL.IsInherited
                        }
                        $Result | Select Path,Group,User,FileSystemRights,AccessControlType,Inherited
                    }
                }
                Else
                {    $Result = New-Object PSObject -Property @{
                        Path = $Folder.Fullname
                        Group = ""
                        User = Get-ADUser $Name | Select -ExpandProperty Name
                        FileSystemRights = $ACL.FileSystemRights
                        AccessControlType = $ACL.AccessControlType
                        Inherited = $ACL.IsInherited
                    }
                    $Result | Select Path,Group,User,FileSystemRights,AccessControlType,Inherited
                }
            }
            Else
            {   $Result = New-Object PSObject -Property @{
                    Path = $Folder.Fullname
                    Group = ""
                    User = $ACL.IdentityReference.Value
                    FileSystemRights = $ACL.FileSystemRights
                    AccessControlType = $ACL.AccessControlType
                    Inherited = $ACL.IsInherited
                }
                $Result | Select Path,Group,User,FileSystemRights,AccessControlType,Inherited
            }
        }
    }
}
Write-Verbose "$(Get-Date): Script completed!"
# SIG # Begin signature block
# MIIKFgYJKoZIhvcNAQcCoIIKBzCCCgMCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUko2cVKZtfZBRvgrov74Q2dDP
# Q7agggeQMIIHjDCCBnSgAwIBAgIKG9ZY5gACAAABBzANBgkqhkiG9w0BAQUFADBB
# MRMwEQYKCZImiZPyLGQBGRYDY29tMRMwEQYKCZImiZPyLGQBGRYDS0NTMRUwEwYD
# VQQDEwxLQ1MtTUlULTItQ0EwHhcNMTYwMjE1MTgxMzI2WhcNMTcwMjE0MTgxMzI2
# WjCBoTETMBEGCgmSJomT8ixkARkWA2NvbTETMBEGCgmSJomT8ixkARkWA0tDUzEM
# MAoGA1UECxMDS0NTMQ0wCwYDVQQLEwRNcGxzMQ0wCwYDVQQLEwRUZWNoMRgwFgYD
# VQQDEw9CZXJuYXJkIFdlbG1lcnMxLzAtBgkqhkiG9w0BCQEWIGJ3ZWxtZXJzQEtu
# dXRzb25Db25zdHJ1Y3Rpb24uY29tMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIB
# CgKCAQEAsOVwqeQLGhiaSwHiK7hjkLQE1etSX0uuSIm6LM8ofJYf/hi0KDnj+7g2
# PqUvE/RukOUSrKrURX51StE4xXLNxpK0dDqIHI7E5//TUAtKPCOrNpsZr/W885gK
# Ogd037ITvpdDxCzJi8hkpAfPW32hoMUPHwPBZrMR5IZPZ96rCQeZ0gA6+Y3wU7Re
# TWpp3d1+ipr9SdfihViB1t/UJMebg9ockHTcF833v9u71xyXQNgpjPLDiMIK21mN
# gNKjvWXVeID0vxlVEEqeYOWGswUsuLuv5Ddp5VHrZZo9CV/lwrWsk83tu/BCpRqk
# f8Ius+DqUPlRQ7tyk3pr33EMEEaUdQIDAQABo4IEIzCCBB8wPAYJKwYBBAGCNxUH
# BC8wLQYlKwYBBAGCNxUIh8udUMeqLIPtgxyEycw4g/z/N2WC29sjgeD/ZgIBZgIB
# AjCBgAYDVR0lBHkwdwYLKwYBBAGCNwoDBAEGCSsGAQQBgjcVBQYKKwYBBAGCNwoF
# AQYIKwYBBQUHAwMGCSsGAQQBgjcVBgYKKwYBBAGCNwoDCwYEVR0lAAYIKwYBBQUH
# AwIGCCsGAQUFBwMEBgorBgEEAYI3CgMEBgorBgEEAYI3CgMBMA4GA1UdDwEB/wQE
# AwIFoDCBngYJKwYBBAGCNxUKBIGQMIGNMA0GCysGAQQBgjcKAwQBMAsGCSsGAQQB
# gjcVBTAMBgorBgEEAYI3CgUBMAoGCCsGAQUFBwMDMAsGCSsGAQQBgjcVBjAMBgor
# BgEEAYI3CgMLMAYGBFUdJQAwCgYIKwYBBQUHAwIwCgYIKwYBBQUHAwQwDAYKKwYB
# BAGCNwoDBDAMBgorBgEEAYI3CgMBMIGUBgkqhkiG9w0BCQ8EgYYwgYMwCwYJYIZI
# AWUDBAEqMAsGCWCGSAFlAwQBLTALBglghkgBZQMEARYwCwYJYIZIAWUDBAEZMAsG
# CWCGSAFlAwQBAjALBglghkgBZQMEAQUwCgYIKoZIhvcNAwcwBwYFKw4DAgcwDgYI
# KoZIhvcNAwICAgCAMA4GCCqGSIb3DQMEAgICADBNBgNVHREERjBEoCAGCisGAQQB
# gjcUAgOgEgwQYndlbG1lcnNAS0NTLkNPTYEgYndlbG1lcnNAS251dHNvbkNvbnN0
# cnVjdGlvbi5jb20wHQYDVR0OBBYEFGX8A+llEhnGXDAws5XEHnGqiBCMMB8GA1Ud
# IwQYMBaAFHPLrIEMeCxJy/thNCJTeqOtXTRdMIHHBgNVHR8Egb8wgbwwgbmggbag
# gbOGgbBsZGFwOi8vL0NOPUtDUy1NSVQtMi1DQSgxKSxDTj1NSVQtMixDTj1DRFAs
# Q049UHVibGljJTIwS2V5JTIwU2VydmljZXMsQ049U2VydmljZXMsQ049Q29uZmln
# dXJhdGlvbixEQz1LQ1MsREM9Y29tP2NlcnRpZmljYXRlUmV2b2NhdGlvbkxpc3Q/
# YmFzZT9vYmplY3RDbGFzcz1jUkxEaXN0cmlidXRpb25Qb2ludDCBugYIKwYBBQUH
# AQEEga0wgaowgacGCCsGAQUFBzAChoGabGRhcDovLy9DTj1LQ1MtTUlULTItQ0Es
# Q049QUlBLENOPVB1YmxpYyUyMEtleSUyMFNlcnZpY2VzLENOPVNlcnZpY2VzLENO
# PUNvbmZpZ3VyYXRpb24sREM9S0NTLERDPWNvbT9jQUNlcnRpZmljYXRlP2Jhc2U/
# b2JqZWN0Q2xhc3M9Y2VydGlmaWNhdGlvbkF1dGhvcml0eTANBgkqhkiG9w0BAQUF
# AAOCAQEAZcDFjE0Tp3M2gqEaCw+PG8fZTlsZeGZ7TBT1FPGUAlrgGrEhELiW5rQf
# jY+iu7jWoHEtqq3cT3FWg9RD4HIRLGiwg+M4X7niwTD0JOuZepemfApOkLq3ydqB
# AAS5AWFZ5mrLpqJI/2I8n77JFuOcWLbz8FYdusbDkx828ADrAXqYkAXr+tOhVBEo
# +wc0n45s92dndPRwioKOFvkQ4VlOC3nBHxOoDmtcZ+B2yCunkzYV2rz4iHdamRip
# 6xbuVpFQ3QVbA2N0nyh9dlyZ21pM8MNXCL6xai6850mwf/1ZA8p1qFd/irMOiPL8
# Nt5/gCe9Oa9K1cP5qNsBs3DFP1LehDGCAfAwggHsAgEBME8wQTETMBEGCgmSJomT
# 8ixkARkWA2NvbTETMBEGCgmSJomT8ixkARkWA0tDUzEVMBMGA1UEAxMMS0NTLU1J
# VC0yLUNBAgob1ljmAAIAAAEHMAkGBSsOAwIaBQCgeDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBTZnhrrS0qIyYO0FklR
# DAroS3JwYjANBgkqhkiG9w0BAQEFAASCAQAqsVW0LpygcJU9b8FYTWdY4YakXi1d
# dazLjJBB2Bo39YEg7tN6y2V2K8m73wfRx1p3M8gVGPVUb+6Tk2rmmGR+zL6wqIMd
# YtaOffA6zn2p0snuWLJRUqDICp5gCtFRcdNnz1EHtbZNll2lGdR+vyMybGMZuJk7
# u/1RD6Pk7C4Is5HqkVDrOJ5akxz0I2eIYf+UZOEUPqPGFpjki/8OmFbuzgusEO0C
# TOcz3sZ+4QiRs5fG3O3kbmRhdIXkShEims370xLpM3Q+8MOnCMWHeJiQJw2n2E6e
# 7vsPycXt0Hw8ldeMtUSP5Vq+DYBdQHWjMqYiWzERU4YUJZWwz0LN7rEu
# SIG # End signature block
