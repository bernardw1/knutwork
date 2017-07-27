<#

.SYNOPSIS
        Disable-UserAccount
        Created By: Bernard Welmers
        Created Date: August, 2016
        Last Modified Date: January 16, 2017

        To see all the parameters and what they do please run get-help .\Disabled-UserAccounts.ps1 -detailed

.DESCRIPTION
        This script is meant to allow quick and easy disabling of Knutson users. It also will disable their AD account, forward their email, setup an out of office message etc. 
        
        The script will run in interactive mode by default. This means it will ask you questions on what you want to do with the user that you are trying to disable. THe more adviced features are aveilable when you run the script from the command line. 

    .Parameter interactive
        Instruct the script to interact with you, this is enabled by default. To disable the interactive version set the parameter to $False

    .Parameter TerminatedUserName
        Please enter the user name for the person you want to disable. The script will validate this is a good username

    .Parameter EmailToUsername
        Who do you want to forward the emails to

    .Parameter EmailForwarding
        Do you want to enable email forwarding? By default emails will be forwarded. You can stop this by setting this to $False

    .Parameter EnableOOFMessage
         This setting allows you to leave the Out of Office messsage disabled by setting this parameter to be $false

    .Parameter OOFMessage
        Setting this parameter allows you to define a custom out of office message. By default the message will say: 
        
        The default Out of Office message is: <Terminated User's Name> is no longer with Knutson Construction. Please contact <Person that email is forwarded to's name> or call <Office Phone Number> with all matters going forward.  Thank you"

    .Parameter ArchiveEmail
        By default the user will be moved to the Disabled User's OU and as a result the email will be archived. If you don't want this to happen set this to $False

    .Parameter EmailAccounting
        By default the script will email accounting to let them know that the person is being terminated and remind them to disable all their accounting related access (Set to $Fales if you don't want this email to go out)

    .Parameter EmailCosential
         By default the script will email the cosential administrator to let them know that the person is being terminated and remind them to disable cosential related access (Set to $Fales if you don't want this email to go out
 
    .Parameter CopyFiles
        By default the script will copy all the person's files from the following directories:

        //mfs1/TSHOME45/<Username>
        //MFS1/users$/<Username> or //RFS/users$/<UserName> or //ifs/users/<username>

        it will also delete the folder //mfs1/profiles$/<username>
        
    .Parameter accoutingEmail
    The person in accounting that will be emailed to notify about the employee termination will be Teresa Mason, this can be changed by entering a valid email address here

    .Parameter CosentialEmail
        The cosential administrator that will be emailed to notify about the employee termination will be Michelle Ode, this can be changed by entering a valid email address here

.EXAMPLE
    Most simple (ask interactive questions)
    
    \\kcs.com\it\installs\1ITScripts\Disable-UserAccount.ps1
    or
    \\kcs.com\it\installs\1ITScripts\Disable-UserAccount.ps1 <UserName>

    run everything by default
    \\kcs.com\it\installs\1ITScripts\Disable-UserAccount.ps1 -interactive $False -TerminatedUserName ktest -EmailtoUsername btest 

        

.NOTES
    Janary 16, 2017
        Changed parameters from Boolian to Switch which removes the possiblity of $Null entries
        Added .paramater comments to the help information.

    January 13, 2017
        Changes Edited the script to make it more interactive. Also added the comments/help statements to make things easier to understand what the script is supposed to do.

    2016
        Created the base script to do the normal steps we take when disabling users.

#>


param 
(
    [Parameter(Mandatory=$true, ValueFromPipeline = $true)] [String]$TerminatedUserName,
     [Bool]$Interactive=$true,
    [String]$EmailToUsername="None",
    [Switch]$EmailForwarding=$True,
    [Switch]$EnableOOFMessage=$true,
    [String]$OOFMessage,
    [Switch]$ArchiveEmail=$True,
    [Switch]$EmailAccounting = $True,
    [Switch]$EmailCosential = $True,
    [Switch]$CopyFiles = $True,
    $accoutingEmail = "tmb@knutsonconstruction.com",
    $CosentialEmail = "mode@knutsonconstruction.com"
)


function PreProcessEmail {
    #Write-Host "Email to username:$Emailtousername"
    
     If ("none" -eq $EmailToUsername){
        $emailtousername = Read-Host -Prompt "Please enter the username for the emails to be forwarded to"
    }

    Try{
        $emailto = Get-ADUser $EmailToUsername -Properties mail, pager
        } Catch {
        Write-Error -Message "The user you are trying to send the emails to does not exist. Please try runing the script again."
        exit
        } 
    #$emailto
    return $emailto
}

#main part of the script

#Basic varriables and things that need to be answered to start
Try{
$SAM = get-aduser $TerminatedUserName -Properties SamAccountName
} Catch {
    Write-Error -Message "The user you are trying to terminate does not exist. Please try to run the script again."
    exit
}

If ($interactive -eq $true){
    $responce = read-host -prompt "Enable Email Forwarding? (Y/N)"
    If ($Responce -eq "y"){
        $EmailForwarding=$True
        $emailto = PreProcessEmail
    }

    $responce = read-host -prompt "Enable Out of Office message? (Y/N)"
    If ($Responce -eq "N"){
        $EnableOOFMessage=$False
    }else {
        $emailto = PreProcessEmail
        Write-Host "The default Out of Office message is: $($sam.Name) is no longer with Knutson Construction. Please contact $($emailto.name) ($($emailto.mail)) or call $($emailto.pager) with all matters going forward.  Thank you"
        $responce = read-host "Do you want to specify a different Out of Office Message (Y/N)"
        If ($Responce -eq "y"){
            $OOFMessage = Read-Host -Prompt "Please enter the new Out of Office Message"
        }
    }

    $responce = read-host -prompt "Do you want to move the user to the Disabled user group and enable email archiving? (Y/N)"
    If ($Responce -eq "y"){
        $ArchiveEmail=$True
    }

}



write-progress -Activity "Processing AD Account" -Status "Removing group memberships"
$GroupMembers = Get-ADPrincipalGroupMembership $sam
foreach ($group in $GroupMembers) 
{
    if ($group.name -ne "domain users") 
    {
        remove-adgroupmember -identity $group -members $SAM -Confirm:$false
    }
}
write-progress -Activity "Processing AD Account" -Status "Check to see if anyone is reporting to this user and move them to this users's manager"
    $managerinfo = get-aduser $sam -Properties Manager
    get-aduser -Properties manager -Filter {manager -eq $managerinfo.DistinguishedName} | Set-aduser -Manager $Mangerinfo.manager


write-progress -Activity "Processing AD Account" -Status "Disable AD user account"
Disable-ADAccount $sam.SamAccountName
set-ADAccountPassword -Identity $sam.samaccountname -newpassword (ConvertTo-SecureString "Lenovo1" -AsPlainText -Force)

if ($ArchiveEmail -eq $true){
    write-progress -activity "Processing AD Account" -Status "Moveing AD user account to disabled users OU"
    if ($ArchiveEmail -eq $True){
        Move-ADObject $SAM -TargetPath "OU=To be disabled,OU=KCS,DC=KCS,DC=com"
    }
}


write-progress -Status "Making the Connection to Exchange" -activity "Processing Email"
#Check to see if a connection MS Exchagne (MEX2) has already been established in this PowerShell Session
if (!(Get-PSSession | Where { $_.ConfigurationName -eq "Microsoft.Exchange" })) 
{ 
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://mex2.kcs.com/PowerShell/ -Authentication Kerberos
    Import-PSSession $Session -ErrorAction Stop 
}
$emailto = PreProcessEmail

#setup out of office message
If ($EnableOOFMessage -eq $true){
    write-progress -Activity "Processing Email" -status "Setuping up and turn on Out of Office message"
    If (-not $OOFMessage){
        $OOFMessage = "$($sam.Name) is no longer with Knutson Construction. Please contact $($emailto.name) ($($emailto.mail)) or call $($emailto.pager) with all matters going forward.  Thank you"
    }
    Set-MailboxAutoReplyConfiguration $SAM.SamAccountName -AutoReplyState enabled -ExternalAudience all -InternalMessage $OOFMessage -ExternalMessage $OOFMessage
}

$cred = \\kcs.com\it\installs\1ITScripts\Get-MyCredential.ps1

If ($EmailForwarding = $true){
    write-progress -activity "Processing Email" -Status "Forwarding email and hiding it from email list"
    Set-Mailbox -Identity $sam.SamAccountName -DeliverToMailboxAndForward $true -ForwardingAddress $emailto.mail  -HiddenFromAddressListsEnabled $true

    $messagebody = "Hello, we just wanted to make sure you were aware that the emails sent to $($SAM.Name) will be forwarded to you for the next 30 days. At that point you will get an automated message asking if you would like to stop receiving the emails. Thank you very much, IT Department"
    send-mailmessage -to $emailto.mail -subject "Email being forward notification" -body $messagebody -smtpserver 'smtp.office365.com' -UseSsl -Port 587 -from helpdesk@knutsonconstruction.com -Credential $cred
}

if ($EmailAccounting -eq $True){
    write-progress -Activity "Finale Processing" -Status "emailing Accounting to remind them to check on user accounts"
    $messagebody = "Please make sure that the accounts for $($SAM.Name) have been disabled in the Accounting systems (such as eCMS, Invoice Router, Textura, and Ebix)"
    send-mailmessage -to $accoutingEmail -subject "Employee Termination notification" -body $messagebody -smtpserver 'smtp.office365.com' -UseSsl -Port 587 -from helpdesk@knutsonconstruction.com -Credential $cred 
}

If ($EmailCosential -eq $True){
    write-progress -Activity "Finale Processing" -Status "Sending email to remind Cosential admin to check on account"
    $messagebody = "Please make sure that the account for $($SAM.Name) has been disabled in Cosential"
    send-mailmessage -to $CosentialEmail -subject "Employee Termination notification" -body $messagebody -smtpserver 'smtp.office365.com' -UseSsl -Port 587 -from helpdesk@knutsonconstruction.com -Credential $cred
}

if ($CopyFiles = $True){

    write-progress -Activity "Copy Files" -Status "Create ExUser Directory"
    $ExUser = "\\MPLS_NAS\ITArchives\ExEmpFiles\"+$Sam.Name +"\"
    if(!(Test-Path -Path $Exuser)) {New-item -ItemType directory -Path $ExUser }

    write-progress -Activity "Copy Files" -Status "Check to see if there is a Citrix home directory and move it to Exusers Directory"
    $rempath = "\\mfs2\Userfiles$\" + $sam.SamAccountName
    If (test-path $rempath) 
    { 
        if(!(Test-Path -Path $Exuser"citrix\")) {New-item -ItemType directory -Path $Exuser"Citrix\" }
        Move-Item -Path $rempath -Destination $ExUser"citrix\" -Force 
    }

    write-progress -Activity "Copy Files" -Status "Create Users directory and checking for Q drives"
    if(!(Test-Path -Path $Exuser"users\")) {New-item -ItemType directory -Path $Exuser"Users\" }

    #Check to see if MFS1 user directory and move it to ExUser directory
    write-progress -Activity "Copy Files" -Status "MFS1 Q Drive"
    $rempath = "\\mfs1\users" + $sam.SamAccountName
    If (test-path $rempath) { Move-Item -Path $rempath -Destination $ExUser"Users" -Force }

    #Check to see if RFS user directory and move it to ExUser directory
    write-progress -Activity "Copy Files" -Status "RFS Q Drive"
    $rempath = "\\rfs\users$\"+$sam.SamAccountName
    If (test-path $rempath) { Move-Item -Path $rempath -Destination $ExUser"Users" -Force }

    #Check to see if IFS user directory and move it to ExUser directory
    write-progress -Activity "Copy Files" -Status "IFS Q Drive"
    $rempath = "\\ifs\users\"+$sam.SamAccountName
    If (test-path $rempath) { Move-Item -Path $rempath -Destination $ExUser"Users" -Force }

    write-progress -Activity "Copy Files" -Status  "check to see if there is a TSprofiles directory if so delete it"
    $rempath = "\\mfs1\tsprofiles$\"+ $sam.SamAccountName
    If (test-path $rempath) { Remove-Item -Path $rempath -Recurse}

    #add retention to the mailbox of the user being terminated
    Remove-PSSession $session
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $cred -AllowRedirection -Authentication basic
    Import-PSSession $session -AllowClobber

    $a =  Get-User $sam.NAME
    Set-RetentionCompliancePolicy -Identity "Terminated Users" -AddExchangeLocation $a.name -Comment "test terminating user policy"

    # need to figure out how to add the command "revoke-azureaduserallrefreshtoken" https://go.microsoft.com/fwlink/?linkid=841345
    #connect to Azure AD to make sure user looses acccess to all the O365 stuff
    Import-module AzureAD
    Connect-AzureAD -Credential $cred
    Revoke-AzureADUserAllRefreshToken -ObjectId $sam.UserPrincipalName


}

# SIG # Begin signature block
# MIIJXQYJKoZIhvcNAQcCoIIJTjCCCUoCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQULJZbNtx90no7hbb06tEkGz3K
# 2G6gggbOMIIGyjCCBbKgAwIBAgITGgAAAAtIEqyWvVA2zAADAAAACzANBgkqhkiG
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
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUp6H7hlZJ
# 5Se/NYz/uZEhQ9dsZn8wDQYJKoZIhvcNAQEBBQAEggEAoxFU524d7O1pwAX6g8A0
# 7aeB1CSczjmNvMSnBElQ1pH56J9RS1njqCnJbONliQNfrWSv+I68lA4rr84NSjQ2
# 3FIVM9MSegVfW9UsldIuOSdcyaDvkPmZuehsLSiLDCWQEbU5V1vMz/V20RqJNlhs
# TZ2l7Cj/OCzfpThkYRaUB2paIuiQSfnGuWUn0OHcLTKAW8XoONBcVEEI9Xt8RUos
# oZVN+L417DyepslrAiMr3rKz0t5Cs5IdW/70XSGuftInw6vRCAbfliuyORauaGFx
# VfuEylRwUNK0wHAiEZdzjLp4VwHtNgmOgovQiBOyOMTnv0knZVBjBVoh+1Jwj7yd
# Aw==
# SIG # End signature block
