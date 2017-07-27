<#
.SYNOPSIS
 Created by Bernard Welmers
 On March 17, 2017

 This script was created to automate the movement of people between the old Citrix/Terminal Services environment and the new one.

 It chances the directory for the TSHomeDirectory as well as the HomeDirectory. This will also give all users access to the TSHomeDirectory outside of the Terminal Services environment

 The script does allow piping values in from a previous command so you can get all the people from a list and then send them to this script to be processed.

.DESCRIPTION
.EXAMPLE

Get-ADGroupMember "Citrix General Apps" | i:\1ITScripts\UpdateCitrixDirectory.ps1

\\kcs.com\it\installs\1ITScripts\UpdateCitrixDirectory.ps1 -user btest

.NOTES
March 17, 2017 Initial creation 

#>
[CmdletBinding()]
Param (
    [Parameter(Mandatory=$true, ValueFromPipeline = $true)] [string] $User,
    [String] $HomeDirectory = '\\kcs.com\files\UserFiles'
)


Function MoveFolders ($folderName, $Destination)
{
    Write-Host "in function testing $folderName " (Test-path $folderName)
    if(Test-path $folderName)
    {
        Write-Host "Moving to $Destination"
        move-item -Path $folderName  -Destination $Destination -Exclude '*.BIN' -force
        }
}

Function RemoveLinks ($Path){
    If (Test-path -path $path){
        remove-item $Path
    }
}

$ErrorActionPreference = "Stop"
#Prerequisites
if(!(Get-InstalledModule -Name ntfssecurity))
{
    Install-Module -Repository PSGallery -Name NTFSSecurity -Scope CurrentUser -confirm:$true
}

if(!(get-aduser $user)){
    write-error "Username does not exist in AD, please check and try again"
    exit
}
$user2 = $user.SamAccountName
write-host "User info $user2"

if($user.samAccountName){
    $user = $user.SamAccountName
}


$userinfo = get-aduser $user -Properties sid
#Take the provided home directory and then add the user name to it
$HomeDirectory = $HomeDirectory + '\'+$userinfo.SamAccountName

#Clear the old Terminal Services home directory information
set-aduser $user -clear userparameters,msTSHomeDirectory,msTSHomeDrive,HomeDirectory,HomeDrive 
#Add the new TS Home direcotry information
set-aduser $user -add @{'msTSHomeDirectory' = $HomeDirectory; 'msTSHomeDrive' = "Z:" ; HomeDirectory = $HomeDirectory ; HomeDrive = "Z:"}

#Creat the new TSHome directory if needed
 write-progress -Status "Create the TS Home45 directory" -activity "Processing AD Account"
    if(!(Test-Path -Path $HomeDirectory)) 
        {
        New-Item -ItemType Directory -Force -Path $HomeDirectory
        }

new-item -ItemType Directory -Force -Path $HomeDirectory\OutlookFiles | %{$_.Attributes = "hidden"}
new-item -ItemType Directory -Force -Path $HomeDirectory\Profile | %{$_.Attributes = "hidden"}

#Remove user from current Citrix Group and add to the new group
add-adgroupmember Citrix-2017-Desktop $User
remove-adgroupmember -identity 'Citrix General Apps' -members $user -Confirm:$false

#Move the users main file locations to the new TS Home Directory.
write-progress -Status "Moving files" -activity "Processing AD Account"
 
MoveFolders -folderName "\\mfs1\TSHOME45\$user\windows\desktop"      -Destination "$HomeDirectory\Desktop"
Movefolders -FolderName "\\mfs1\TSHOME45\$user\desktop"              -Destination "$HomeDirectory\Desktop"
Movefolders -FolderName "\\mfs1\TSHOME45\$user\windows\Favorites"    -Destination "$HomeDirectory\Favorites"
Movefolders -FolderName "\\mfs1\TSHOME45\$user\Documents\Favorites"  -Destination "$HomeDirectory\Favorites"
Movefolders -FolderName "\\mfs1\TSHOME45\$user\My Documents\Favorites"  -Destination "$HomeDirectory\Favorites"
Movefolders -FolderName "\\mfs1\TSHOME45\$user\Favorites"            -Destination "$HomeDirectory\Favorites"
Movefolders -FolderName "\\mfs1\TSHOME45\$user\Downloads"            -Destination "$HomeDirectory\Downloads"
Movefolders -FolderName "\\mfs1\TSHOME45\$user\documents\Pictures"   -Destination "$HomeDirectory\Pictures"
Movefolders -FolderName "\\mfs1\TSHOME45\$user\documents\My Music"   -Destination "$HomeDirectory\Music"
Movefolders -FolderName "\\mfs1\TSHOME45\$user\My documents\My Pictures"   -Destination "$HomeDirectory\Documents\Pictures"
Movefolders -FolderName "\\mfs1\TSHOME45\$user\windows\Music"        -Destination "$HomeDirectory\Music"
Movefolders -FolderName "\\mfs1\TSHOME45\$user\Documents"            -Destination "$HomeDirectory\Documents"
Movefolders -FolderName "\\mfs1\TSHOME45\$user\My Documents"            -Destination "$HomeDirectory\Documents"
Movefolders -FolderName "\\mfs1\TSHOME45\$user\windows\Downloads"    -Destination "$HomeDirectory\Downloads"
Movefolders -FolderName "\\mfs1\TSHOME45\$user\windows\BlueBeamConfig" -Destination " $HomeDirectory\BlueBeamConfig"
Movefolders -FolderName "\\mfs1\TSHOME45\$user\BlueBeamConfig"       -Destination "$HomeDirectory\BlueBeamConfig"
Movefolders -FolderName "\\mfs1\TSHOME45\$user\windows\Pictures"     -Destination "$HomeDirectory\Documents\Pictures"
Movefolders -FolderName "\\mfs1\TSHOME45\$user\SureTrak Projects"        -Destination "$HomeDirectory\SureTrak Projects"

new-item -ItemType Directory -Force -Path "$HomeDirectory\profile\UPM_Profile\AppData\Roaming\MPS\Prolog Manager\Version 9\Configuration"

$a = "\\Mfs1\tshome45\$user\Windows\AppData\Roaming\MPS\Prolog Manager\Version 9\Configuration\*" 
if(Test-path $a)
    {
        Write-host "Moving $a"
        copy-item -path $a -Destination "$HomeDirectory\profile\UPM_Profile\AppData\Roaming\MPS\Prolog Manager\Version 9\Configuration" -filter *.ini
    }
$a = "\\Mfs1\tshome45\$user\WINDOWS\AppData\MPS\Prolog Manager\Version 9\Configuration\*" 
if(Test-path $a)
    {
        Write-host "Moving $a"
        copy-item -path $a -Destination "$HomeDirectory\profile\UPM_Profile\AppData\Roaming\MPS\Prolog Manager\Version 9\Configuration" -filter *.ini
    }

new-item -ItemType Directory -Force -Path $HomeDirectory\Documents\OutlookFiles | %{$_.Attributes = "hidden"}

RemoveLinks $HomeDirectory\desktop\CDM.lnk
RemoveLinks "$HomeDirectory\Desktop\Invoice Router.lnk"
RemoveLinks "$HomeDirectory\Desktop\Asta Powerproject.lnk"

if (test-path "$HomeDirectory\documents\my documents"){
    Get-ChildItem -Path "$HomeDirectory\Documents\My Documents\" | Move-Item -Destination $HomeDirectory\Documents\ -force
    remove-item "$HomeDirectory\documents\my documents" -Recurse -Force
}
if (test-path "$HomeDirectory\Favorites\Favorites"){
    Get-ChildItem -Path "$HomeDirectory\Favorites\Favorites" | Move-Item -Destination $HomeDirectory\Favorites\ -force
    remove-item "$HomeDirectory\Favorites\Favorites" -Recurse -Force
}
write-progress -Status "Updating folder security" -activity "Processing AD Account"
 


#Give person full access to the Home Directory
        Add-NTFSAccess -path $homedirectory -AccessRights FullControl -Account $userinfo.sid  -AccessType Allow -InheritanceFlags ContainerInherit 
        Get-childitem -path $homedirectory -Force | Add-NTFSAccess -AccessRights FullControl -Account $userinfo.sid  -AccessType Allow -InheritanceFlags ContainerInherit 
 #       Get-ChildItem -Path $HomeDirectory\Profile\ -Force -Recurse | Add-NTFSAccess -AccessRights FullControl -Account $userinfo.sid  -AccessType Allow -InheritanceFlags ContainerInherit

#Add owner permissions to the created directory    
write-progress -Status "Updating owner information" -activity "Processing AD Account"
    Get-childitem -path $homedirectory -Recurse -Force | Set-NTFSOwner -Account $userinfo.sid 
        Set-NTFSOwner -Account $userinfo.sid -Path $HomeDirectory
 #       Get-ChildItem -Path $HomeDirectory\Profile\ -Force -Recurse | Set-NTFSOwner -Account $userinfo.sid 

#Make note in the old TSHome directory that it has been processed
#get-date | out-file \\mfs1\TSHOME45\$user\CitrixMoved.txt -Append
#$HomeDirectory | out-file \\mfs1\TSHOME45\$user\CitrixMoved.txt -Append

get-childitem -path $HomeDirectory -Filter *.bin -Recurse | Remove-Item -Recurse -Force


# SIG # Begin signature block
# MIIJXQYJKoZIhvcNAQcCoIIJTjCCCUoCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUnXGTftJXOC6dzotv8Jkx7PnS
# A46gggbOMIIGyjCCBbKgAwIBAgITGgAAAAtIEqyWvVA2zAADAAAACzANBgkqhkiG
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
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUUKIbv75l
# prelZGx5VliIakGR7fowDQYJKoZIhvcNAQEBBQAEggEAKA7hR9noj86S+jEI4rDK
# fdxi9tSjE/4YYO/ytFdDeNFbJo7rEb1YqrxGwyDr5hucxkoy+qDykSuC0b0nDuAL
# wrR8QBVik3IJBfzpoL1+T8q/BUE5bzxKuFyK+SD+y1/VWGRY8fIkpmjF9toxGDPY
# sZQDmcW3DU8BCISY7bQGcPBN3WhSN7sSunrlC1dvYlmSzOrt4lJJ7DN8/LuZCG+d
# 5Lh9pdGAAkjl9ITfXP/Rj+iKhJREyhvPACKv1E8yVCu45uE7+YBomtV5N0/9qBpj
# K20wm3WJ0lifz7fk4nyjdA3lSX0L7YafYJDqDN2LP+NycxYBYJdVkjKrdLYgf1g5
# Fg==
# SIG # End signature block
