<#
    This module is all the reoccuring functions that are used in multiple scripts

#>


<#
.SYNOPSIS
Gets credentials from the user and saves them to the C:\temp folder. This allows running multiple scripts without prompting for credentials

.DESCRIPTION
Long description

.PARAMETER CredPath
Path to the credential file. By default it goes to c:\temp\

.PARAMETER Help
shows the help file

.EXAMPLE
An example

.NOTES
    Get-MyCredential
    Usage:
    Get-MyCredential -CredPath $CredPath

    If a credential is stored in $CredPathAndFile, it will be used.
    If no credential is found, Export-Credential will start and offer to
    Store a credential at the location specified.
#>
function Get-MyCredential
{
param($CredPath= "c:\temp\EncryptedO365Creds.txt",[switch]$Help)
$HelpText = @"

    Get-MyCredential
    Usage:
    Get-MyCredential -CredPath `$CredPath

    If a credential is stored in $CredPathAndFile, it will be used.
    If no credential is found, Export-Credential will start and offer to
    Store a credential at the location specified.

"@

    if (!(Test-path -path c:\temp))
    {
        new-item -Path c:\ -Value temp -ItemType Directory
    }

    if($Help -or (!($CredPath))){write-host $Helptext; Break}
    if (!(Test-Path -Path $CredPath -PathType Leaf)) {
        Export-Credential (Get-Credential -Message "Please enter your Office 365 credentials") $CredPath
    }
    $cred = Import-Clixml $CredPath
    $cred.Password = $cred.Password | ConvertTo-SecureString
    $Credential = New-Object System.Management.Automation.PsCredential($cred.UserName, $cred.Password)
    Return $Credential
}

<#
.SYNOPSIS
Exports the credentials to a file

.DESCRIPTION
Long description

.PARAMETER cred
Parameter description

.PARAMETER path
Parameter description

.EXAMPLE
Export-Credential $CredentialObject $FileToSaveTo

.NOTES
This was received from Jason Hargrove from Lofler
#>
function Export-Credential($cred, $path) {
      $cred = $cred | Select-Object *
      $cred.password = $cred.Password | ConvertFrom-SecureString
      $cred | Export-Clixml $path
}


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
Function Check-Module ($mod)
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

<#
.SYNOPSIS
Script to connect to Office 365 exchange

.DESCRIPTION
Long description

.EXAMPLE
An example

.NOTES
General notes
#>
function Connect-ExchangeOnLine {
    if (!(Get-PSSession | Where { ($_.ConfigurationName -eq "Microsoft.Exchange") -and ($_.ComputerName -ne "mex2.kcs.com") })) 
    {
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Cred -Authentication Basic -AllowRedirection
        Import-PSSession $Session  -AllowClobber
    }

}

function Connect-Compliance {
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $Cred -Authentication Basic –AllowRedirection -prefix cc
    Import-PSSession $Session  -AllowClobber

}

Function Connect-Exchange {
    if (!(Get-PSSession | Where { ($_.ConfigurationName -eq "Microsoft.Exchange") -and ($_.ComputerName -eq "mex2.kcs.com") })) 
{ 
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://mex2.kcs.com/PowerShell/ -Authentication Kerberos
    Import-PSSession $Session -ErrorAction Stop 
}

}

export-modulemember -function Get-MyCredential
export-modulemember -function Check-Module
export-modulemember -function Connect-ExchangeOnLine
export-modulemember -function Connect-Compliance
export-modulemember -function Connect-Exchange
export-modulemember -function 
export-modulemember -function 
export-modulemember -function 
export-modulemember -function 
