<#

.SYNOPSIS
 This script was created to automate the user account creation process at Knutson Construction.

 Prerequisits
 The computer running this program needs to have the AD PowerShell module the normal source of this module is the RSAT tool kit.
 
 If you are running the Excel import then you need to have Excel installed on the computer

 NOTE - To run the script with paramaters you MUST include a UserFirstName otherwise it will look for the excel file and run from there.

.DESCRIPTION
.EXAMPLE
.NOTES
    Created by Bernard Welmers
    March 11, 2016 - Initial creation
    2016 further updates to make things work correctly
    Also added the ability to run the import from the command line rather then just using the excel spreadsheet
January 2017
    Added more comments to make the script easier to understand.
March 2017
    Changed the TSHome paths to be pointed to the new citrix environment
June/July 2017
    Updated to run against Office 365
        
#>
[CmdletBinding()]
param 
(
        $UserFirstname,           
        $UserLastname,
        $UserName,    
        $Description,            
        $Password = 'knutson',
        $Manager,
        [ValidateSet("Rochester","Iowa City","Minneapolis","Corporate")]$Office,
        $Title,
        $EmployeeID,
        $Department,
        $OfficePhone,
        $MobilePhone,
        [bool]$CannotChangePassword,
        [bool]$PasswordNeverExpires,
        [bool]$ChangePasswordAtLogon,
        [ValidateSet("Default","Field","ServiceAccount","ConferenceCall","FullE3","RemoveConferenceCall","RemoveE3","RemoveK1")] $O365licenseLevel = "Default"
)
# Stop script on errors (such as user already exists) so that it does not try and add groups or mailbox to existing user
$ErrorActionPreference = "Stop"
$TranscriptFile = "Z:\Scripts\NewUser.txt"
$CredPathAndFile = "z:\Scripts\ENCRemoteCreds.txt"

if (-not $PSBoundParameters.ContainsKey('Office')) {
                Remove-Variable Office
}

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
                Write-progress -Status 'The module has now been installed correctly' -Id 50
            } #end if module available then import 
        } #module not available 
    } # end if not module 
}


#This creates a list box that allows people to select items from the list.
Function ListBox ([array]$ListofGroups)
{
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form 
    $form.Text = "Group Selection form"
    $form.Size = New-Object System.Drawing.Size(600,800) 
    $form.StartPosition = "CenterScreen"

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(75,720)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = "OK"
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(150,720)
    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
    $CancelButton.Text = "Cancel"
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20) 
    $label.Size = New-Object System.Drawing.Size(280,20) 
    $label.Text = "Please make a selection from the list below:"
    $form.Controls.Add($label) 

    $listBox = New-Object System.Windows.Forms.ListBox 
    $listBox.Location = New-Object System.Drawing.Point(10,40) 
    $listBox.Size = '460,670' 
    $listbox.ColumnWidth = 225
    $listBox.MultiColumn = $True
    $listBox.SelectionMode = "MultiExtended"


    foreach ($ListofGroups in $ListofGroups)
        {
        [void] $listBox.Items.Add($ListofGroups.name)
        }


    $form.Controls.Add($listBox) 
    $form.Topmost = $True

    $result = $form.ShowDialog()



    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $ListofGroups = $listBox.SelectedItems
    }

    return ,$ListofGroups

}

Function CreateADAccount {
    # Do the following for everyone
    
    #If SAM (Username) is empty then create a derived SAM        
    if (!$UserName) 
    {
        $UserName = $userfirstname.substring(0,1) + $Userlastname
    }
    $Displayname = $UserFirstname + " " + $UserLastname 
    
    #UPN is derived from the SAM (newer type of user account)
    $UPN = $UserName + "@knutsonconstruction.com"   

    #Check if UPN is already used
    Try {$ExistingUser = Get-Aduser $UPN
    }Catch {}
    if ($ExistingUser){
        write-error -Message "The username $UserName already exists. Please update the username and try again"
        continue 
    }


    #Setup user's home directory and "Z" path
    $TSPath = "\\kcs.com\files\userfiles\"+$UserName
    $ExpireDate = $user.AccountExpirationDate -as [datetime]
    if (!$ExpireDate)
    {  
        $ExpireDate = $null
    }

    #Add basic office address information
	if ($Office -eq "Rochester")
		{
		$StreetAddress = "5985 Bandel Road NW"
		$City = "Rochester"
		$State = "Minnesota"
		$Country = "USA"
		$PostalCode = "55901"
		$Fax = "507.280.9797"
		$pager = "507.280.9788"
		$OU = "OU=People,OU=Roch,OU=KCS,DC=KCS,DC=com"
        $Qpath = "\\rfs\users$\"+$UserName

		}
	Elseif ($Office -eq "Iowa City")
		{
		$StreetAddress = "2351 Scott Blvd SE"
		$City = "Iowa City"
		$State = "Iowa"
		$Country = "USA"
		$PostalCode = "52240"
		$Fax = "319.351.7163"
		$pager = "319.351.2040"
		$OU = "OU=People,OU=Iowa,OU=KCS,DC=KCS,DC=com"
        $Qpath = "\\ifs\users\"+$UserName
		}
	Else
		{
		$StreetAddress = "7515 Wayzata Boulevard"
		$City = "Minneapolis"
		$State = "Minnesota"
		$Country = "USA"
		$PostalCode = "55426"
		$Fax = "763.546.2226"
		$pager = "763.546.1400"
		$OU = "OU=People,OU=Mpls,OU=KCS,DC=KCS,DC=com"
        $Qpath = "\\kcs.com\files\userfiles\"+ $UserName
		}

    try {
        If ($Manager -ne $Null){
            $ADManager = Get-ADUser $Manager
        }
    }Catch {
        Write-error -Message "The manager that was entered <$manager> was not found - as a result no manager will be added to $User"
        Clear-Variable Manager
    }

    write-progress -Status "Processing AD Account" -activity "creating the user"
        New-ADUser -Name "$Displayname" `
        -DisplayName "$Displayname" `
        -SamAccountName $UserName `
        -UserPrincipalName $UPN `
        -GivenName "$UserFirstname" `
        -Surname "$UserLastname" `
        -Description "$Description" `
        -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -Force) `
        -Enabled $true  `
        -Path "$OU" `
        -CannotChangePassword $CannotChangePassword `
        -PasswordNeverExpires $PasswordNeverExpires `
        -ChangePasswordAtLogon $ChangePasswordAtLogon `
        -AccountExpirationDate $ExpireDate `
        -Company "Knutson Construction" `
        -StreetAddress $StreetAddress `
        -City $City `
        -State $State `
        -Country $Country `
        -PostalCode $PostalCode `
        -Fax $Fax `
        -Office $Office `
        -Title $Title `
        -EmployeeID $EmployeeID `
        -Department $Department `
        -OfficePhone $OfficePhone `
        -MobilePhone $MobilePhone `
        -Manager $Manager `
        -OtherAttributes @{'Pager'= $Pager; 'msTSHomeDirectory' = $TSPath; 'msTSHomeDrive' = "Z:"} `
        -HomeDirectory $TSPath `
        -HomeDrive "Z:"`
        -HomePage "www.knutsonconstruction.com" `
        -server mdc4.kcs.com

    #Other Attributes is the way to get "non standard" attributes added to the profile. It is possible to add more items using a ";" as delimiter

    write-progress -Status "Processing AD Account" -activity "Creating the mailbox for the user"
	Enable-remoteMailbox -identity $UserName -DisplayName $Displayname  -RemoteRoutingAddress "$username@knutsonconstruction.com"
    #$MigrationEndpointOnPrem = New-MigrationEndpoint -ExchangeRemoteMove -Name OnpremEndpoint -Autodiscover -EmailAddress $upn -Credentials $Cred
    #$OnboardingBatch = New-MigrationBatch -Name RemoteOnBoarding1 -SourceEndpoint $MigrationEndpointOnprem.Identity -TargetDeliveryDomain knutsonconstruction.mail.onmicrosoft.com 
    #Start-MigrationBatch -Identity $OnboardingBatch.Identity

    #start an delta syncronization cycle with Office 365 to get this user up to office 365
    Invoke-Command -ComputerName mit-2 -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}

    write-progress -Status "Processing AD Account" -activity "Create the TS Home45 directory"
    if(!(Test-Path -Path $TSPath)) 
        {
        New-Item -ItemType Directory -Force -Path $TSPath
        }

    Write-Verbose "get AD Groups"
    write-progress -Status "Processing AD Account" -activity "Create a list of all the Access Control List AD Groups and display them for selection"
    $ADGroupList = ListBox (Get-ADGroup -Filter {name -like "*"} -SearchBase "OU=SecurityGroups,OU=Distro and Security Groups,DC=kcs,DC=com" | sort)

    Write-Verbose "add user to AD groups"
    write-progress -Status "Processing AD Account" -activity "add the user to each group that was selected"
    foreach ($item in $ADGroupList) 
    {
        add-adgroupmember $item $UserName
    }

    $ADGroupList = $Null

    write-progress -Status "Processing AD Account" -activity "Create a list of all the Distribution Groups and display them for selection"
    $ADGroupList = ListBox (Get-ADGroup -Filter {name -like "*"} -SearchBase "OU=Distro-Departmental,OU=Distro and Security Groups,DC=kcs,DC=com" | sort)

    write-progress -Status "Processing AD Account" -activity "add the user to each group that was selected"
    foreach ($item in $ADGroupList) 
    {
        add-adgroupmember $item $UserName
    }

    
    write-progress -Status "Processing AD Account" -activity "Give newly created person access to their TS Home 45 directory"
    #$acl = get-acl $tspath
    #$acl.setowner([System.Security.Principal.NTAccount] $UserName)
    #set-acl $tspath $acl

    Write-Verbose "give full access to files in their Z drive"
        $userinfo = get-aduser $username


    #Give person full access to the Home Directory
        Add-NTFSAccess -path $tspath -AccessRights FullControl -Account $userinfo.sid  -AccessType Allow -InheritanceFlags ContainerInherit 
        Get-childitem -path $tspath -Force | Add-NTFSAccess -AccessRights FullControl -Account $userinfo.sid  -AccessType Allow -InheritanceFlags ContainerInherit 
 
    Write-Verbose "Make owner of files in their Z drive"
    #Add owner permissions to the created directory    
        write-progress -Status "Processing AD Account" -activity "Updating owner information"
        Get-childitem -path $tspath -Recurse -Force | Set-NTFSOwner -Account $userinfo.sid 
        Set-NTFSOwner -Account $userinfo.sid -Path $tspath
 
    <#
    write-progress -Status "Processing AD Account" -activity "Give newly created person access to their Q drive directory"
    if(!(Test-Path -Path $Qpath)) 
        {
            New-Item -ItemType Directory -Force -Path $Qpath
        }
    $acl = get-acl $Qpath
    $acl.setowner([System.Security.Principal.NTAccount] $UserName)
    set-acl $Qpath $acl
    #>

    try {$O365user = get-azureaduser -objectid $upn}
    catch {$O365user = 'none'}
    $counter = 0
    do {
        $counter++
        Start-Sleep -s 5 
        try {$O365user = get-azureaduser -objectid $upn}
        catch {$O365user = 'none'}
        Write-Progress - "still waiting $Counter"
     }
    while ($O365user -eq 'none') 

     \\kcs.com\it\installs\1ITScripts\set-o365License.ps1 -UPN $upn -licenseLevel $O365licenseLevel 


    #Make sure email mailbox is created and licensed in O365 (this section checks all the O365 stuff)
    remove-pssession $session
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Cred -Authentication Basic -AllowRedirection
    Import-PSSession $Session  -AllowClobber
<#
New-MoveRequest -Identity alias -remote -RemoteHostName hybridURL.company.com -TargetDeliveryDomain company.mail.onmicrosoft.com -RemoteCredential $onprem -BadItemLimit 50 –SuspendWhenReadyToComplete
New-MoveRequest -Identity o365t1 -RemoteHostName mail.knutsonconstruction.com -TargetDeliveryDomain knutsonconstruction.mail.onmicrosoft.com -RemoteCredential $cred -BatchName O365t1 -BadItemLimit 50 -remote
#>

    Get-Mailbox -Identity $upn | Set-Mailbox -RetainDeletedItemsFor 30

    write-progress -Status "Processing AD Account" -activity "The user $Displayname has been completly setup"
    Clear-Variable UserFirstname,UserLastname,Username,Description,Password,Manager,Office, Title, EmployeeID, Department, OfficePhone, MobilePhone, CannotChangePassword, PasswordNeverExpires, ChangePasswordAtLogon
}


#Check Powershell version
If ($PSVersionTable.PSVersion.Major -lt 5){
	write-error -message 'You must install at least Powershell version 5 to run this script. You can get it from https://www.microsoft.com/en-us/download/details.aspx?id=50395'
	Exit
}

$Cred = Get-MyCredential -CredPath $CredPathAndFile

<#Check to see if a connection MS Exchagne (MEX2) has already been established in this PowerShell Session
if (!(Get-PSSession | Where { $_.ConfigurationName -eq "Microsoft.Exchange" })) 
{ 
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://mex2.kcs.com/PowerShell/ -Authentication Kerberos
    
    Import-PSSession $Session -ErrorAction Stop 
} #>

if (!(Get-PSSession | Where { ($_.ConfigurationName -eq "Microsoft.Exchange") -and ($_.ComputerName -ne "mex2.kcs.com") })) 
{
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://mex2.kcs.com/PowerShell/ -Authentication Kerberos
    Import-PSSession $Session -AllowClobber
}

#connect to MS Online
CheckModule ('AzureAD')
Import-module AzureAD
Connect-AzureAD -Credential $cred

#Prerequisites
#Check for the NTFS security module and if it does not exist install it
CheckModule ('NTFSSecurity')

<# This is the original way to check and install.
if(!(Get-InstalledModule -Name ntfssecurity))
{
    Install-Module -Repository PSGallery -Name NTFSSecurity -Scope CurrentUser -confirm:$true
}
#>

#Process all the people from the command line
If ($UserFirstname -ne $Null) {
    CreateADAccount
}

#Process people from the excel sheet
If ((-not $UserFirstname)) {
    
    #Check if ImportExcel is installed
    CheckModule ('ImportExcel')

    write-progress -Status "Importing User Information" -activity "Geting user information from Excel File and starting to process all users in the excel file"
    $Users = Import-Excel -Path "\\kcs.com\it\installs\1ITScripts\UserCreation.xlsx" -WorkSheetname "UserList"

    # Now run through every user in the sheet and create their account, update their user directory, and create their mailbox
    foreach ($User in $Users)            
    {            
        #Process the data in the sheet (get it ready to create the account)
        $UserFirstname = $User.Firstname            
        $UserLastname = $User.Lastname
        $UserName = $User.SamAccountName    
        $Description = $User.Description            
        $Password = $User.Password            
        $Manager= $User.Manager
        $Office = $user.Office
        $Title = $User.Title
        $EmployeeID = $User.EmployeeID
        $Department = $User.Department
        $OfficePhone = $User.OfficePhone
        $MobilePhone = $User.MobilePhone
        $CannotChangePassword = ([System.Convert]::ToBoolean($User.CannotChangePassword))
        $PasswordNeverExpires = ([System.Convert]::ToBoolean($User.PasswordNeverExpires))
        $ChangePasswordAtLogon = ([System.Convert]::ToBoolean($User.ChangePasswordAtLogon))
        
        CreateADAccount
    }
}

stop-Transcript
#other retention Policy name 'Delete old deleted items
# SIG # Begin signature block
# MIIJXQYJKoZIhvcNAQcCoIIJTjCCCUoCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUj8Oa010l/P42dh+G2C7KStwr
# +KGgggbOMIIGyjCCBbKgAwIBAgITGgAAAAtIEqyWvVA2zAADAAAACzANBgkqhkiG
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
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU0BDiSbQO
# IrMxAPNihf0soyk4LicwDQYJKoZIhvcNAQEBBQAEggEAMHidSsXjiwA8z9ci3sK9
# ne/Zp/Z3lRiLQWakhW4EpjVZ819iI6pIwfWfUIOpKd6O04ZgB5T6N37wJe1PCxnn
# Va7CZaYD1kV09GoyIN5ByRgx5XpRk/y2saWnin4JWYVQsIWcBXCfSNTZ945rsUeH
# MHYSAJgYSDZDOnPsGqBkWJ/Dh68/P2KHjbL+CvmKcg6L0ZoyXGOl+hZcWv/Hbv/c
# LETFW/B9Xcn0oOTYmnHecSRKBq9Fuyrne0zTUwg6cCyUv3pORIYpydyqaaijJ4P6
# bH0sxNNcPfv4QFVHaHga9f+Kbv2oD6KrC/zVKQeWLlrIgSxfY1t5ih6BZQ4aSAHN
# Ew==
# SIG # End signature block
