<#
.SYNOPSIS
        New-XenAppServers.ps1
        Created By: Bernard Welmers
        Created Date: August, 2016
        Last Modified Date: January 16, 2017

        To see all the parameters and what they do please run get-help .\New-XenAppServers.ps1 -detailed

.DESCRIPTION
        This script is meant to allow for a consistant creatation of the XenApp Servers for Knutson Construction.

		It checks if you want to create a template and what servers you want to replace on each run of the program.

    .Parameter CreateTemplate
		Include this paramater if you want to create a new template to deploy the servers from. Note you need to create a new template for each month that servers are deployed. The system automatically looks for a template for the current month.
    .Parameter CreateVMs
		Which of the 4 XenApp servers do you want to replace with this run of the script, Note you need to use the VMWare VM name for this field. Choose from the list of: "XenApp01","XenApp02","XenApp03", or "Xenapp04"

.EXAMPLE
    
.NOTES
	February 8, 2017
	Updated the initial comments to allow get-help to work correclty.

	Things that still need to be done:
	register bluebeam automatically

	Feb 27, 2016
	Updated the script to run using PowerCLi 6.5
	Fixed some issues that were happening with creating the XenApp servers where they did not get created with a transformation template.
	Added some more comments to the script.
#>

param 
(
	[Parameter(Mandatory=$True)][ValidateSet("XenApp01","XenApp02","XenApp03","Xenapp04")] $CreateVMs,
	[Switch]$CreateTemplate	
)

	

function WaitForSysprep (
	[Parameter(Mandatory=$True)]
	$vm
	){
#Much of the following code was taken from http://vxpertise.net/2013/07/powercli-waiting-for-guest-customization-to-finish/
	Write-output "WaitForSysprep Processing " $vm

	#$vm = Start-VM $vmRef -Confirm:$False -ErrorAction:Stop

	
	
	# wait until VM has started
	Write-output "Waiting for VM to start ..."
	while ($True)
	{
		$vmEvents = Get-VIEvent -Entity $vm
 
		$startedEvent = $vmEvents | Where { $_.GetType().Name -eq "VMStartingEvent" }
 
		if ($startedEvent) 
		{
			break
		}
		else
		{
			Start-Sleep -Seconds 2	
		}	
	}
 
	# wait until customization process has started	
	Write-output "Waiting for Customization to start ..."
	while($True)
	{
		$vmEvents = Get-VIEvent -Entity $vm 
		$startedEvent = $vmEvents | Where { $_.GetType().Name -eq "CustomizationStartedEvent" }
 
		if ($startedEvent)
		{
			break	
		}
		else 	
		{
			Start-Sleep -Seconds 2
		}
	}
 
	# wait until customization process has completed or failed
	Write-output "Waiting for customization ..."
	while ($True)
	{
		$vmEvents = Get-VIEvent -Entity $vm
		$succeedEvent = $vmEvents | Where { $_.GetType().Name -eq "CustomizationSucceeded" }
		$failEvent = $vmEvents | Where { $_.GetType().Name -eq "CustomizationFailed" }
 
		if ($failEvent)
		{
			Write-output "Sysprep Customization failed!"
			return $False
		}
 
		if($succeedEvent)
		{
			Write-output "WaitforSysprep Customization succeeded!"
			return $True
		}
 
		Start-Sleep -Seconds 2			
	}
}

function WaitForToolsToReportIP ($vmRef)
{
#Much of the following code was taken from http://vxpertise.net/2013/07/powercli-waiting-for-guest-customization-to-finish/
	$ping = New-Object System.Net.NetworkInformation.Ping
     Write-output "Waiting for IP address"
	while ($True) 
	{	
		$vm = Get-VM $vmRef
 
		$guestInfo = $vm.guest
 
		if($guestInfo.IPAddress -and $guestInfo.IPAddress[0])
		{
			return $True
		}
		else
		{
			Start-Sleep -Seconds 2
		}
	}	
    Write-output "Ip Address is present"
}

Function TemplateProgress
    { Param ($Percnt, $cmplt, $curOp)
        Write-Progress -Activity "Creating new XenAppTemplate" -Completed -currentoperation $CurOp -Id 100 -PercentComplete $Percnt 
    }


#Check if PowerCLI is installed and what version is installed

# This section insures that the PowerCLI PowerShell Modules are currently active. The pipe to Out-Null can be removed if you desire additional
# Console output.
if (!(Get-Module -Name VMware*)){ Get-Module -ListAvailable VMware* | Import-Module | Out-Null }
<#
	$PowerCLIinstalled = 'No'
	if('C:\Program Files (x86)\VMware\Infrastructure\vSphere PowerCLI\Scripts\Initialize-PowerCLIEnvironment.ps1')
		{$PowerCLIinstalled = '6.3'
		Add-PSSnapin VMware.VimAutomation.Core
		& 'C:\Program Files (x86)\VMware\Infrastructure\vSphere PowerCLI\Scripts\Initialize-PowerCLIEnvironment.ps1'	
		}
		
		
	If ('C:\Program Files (x86)\VMware\Infrastructure\PowerCLI\Scripts\Initialize-PowerCLIEnvironment.ps1\' )
		{$PowerCLIinstalled = '6.5'
		&'C:\Program Files (x86)\VMware\Infrastructure\PowerCLI\Scripts\Initialize-PowerCLIEnvironment.ps1'
	}
}

If ($PowerCLIinstalled = 'No') 
#>

#if PowerCli is installed is should now be running. So check again and if it is not running then error out the script
if (!(Get-Module -Name VMware*)){ Write-Error "PowerCLI is not installed. Please install it before running this script."}


$vcenter = "vcenter2.kcs.com"
Write-verbose "Connecting to the default VCenter Server " 
Connect-VIServer $vcenter -WarningAction SilentlyContinue 

$Month = get-date -UFormat %b
$vmTemplate = "XATemp-"+$Month
$VMHost = Get-Cluster "Production Cluster" |Get-VMHost | Get-Random
$vmOrigTempl = get-vm XenAppTemplate
$lastmonth = (get-date).AddMonths(-1) | get-date -uformat %b

#get a list of all the running "XenApp" servers where the name is like last month | sort
$VMNames = get-VM * | where-object {($_.name -like "XenApp0*")}
# -and ($_.name -like ("*"+$lastmonth))}  
#other option -and ($_.PowerState -eq "PoweredON")}
$NewVMTask = @()
$NewVMName = @()


#Create the new XenApp Template for the current month
If ($CreateTemplate -eq $true)
{
Write-Progress -Activity "Creating new XenAppTemplate" -Completed -currentoperation "Remove old XenApp templates" -Id 100 -PercentComplete 1 
write-verbose "Remove old XenApp templates"
$CurrentTemplates = get-template * | Where-object {($_.name -like "XATemp-*")}
Remove-Template $currenttemplates -DeletePermanently -confirm:$False -ErrorAction SilentlyContinue

$CurrentTemplates = get-vm "XenAppTemplate" 
Write-Progress -Activity "Creating new XenAppTemplate" -Completed -currentoperation "Turn off XenAppTemplate so that it can be copied into a new template file" -Id 100 -PercentComplete 10 
Write-verbose "Turn off XenAppTemplate so that it can be copied into a new template file"
if ($CurrentTemplates.powerstate -eq "PoweredON")
	{
	stop-VMGuest -VM $CurrentTemplates -Confirm:$false
	start-sleep -seconds 60
	}
	
Write-Progress -Activity "Creating new XenAppTemplate" -Completed -currentoperation "Create the new template $vmTemplate" -Id 100 -PercentComplete 40 
Write-verbose ("Create the new template" + $vmTemplate)
$template = get-template $vmTemplate -ErrorAction SilentlyContinue

if (!$template) {
		New-Template -VM XenAppTemplate -Name $vmTemplate -Location Citrix -Datastore VMFS-XenApp
	}
else {
	$vmTemplate = $vmTemplate+"1"
	write-verbose "original template already existed"
	New-Template -VM XenAppTemplate -Name $vmTemplate -Location Citrix -Datastore VMFS-XenApp
	}

Write-Progress -Activity "Creating new XenAppTemplate" -Completed -currentoperation "Template Creation is done" -Id 100 -PercentComplete 100 
#write-output "Template Creation is done"
}


#If not creating a template then check to see if the template exists
If ($CreateTemplate -eq $false)
{
Write-verbose ("Validating that " +$vmtemplate +" exists")

$VMName = get-template -name $vmtemplate -ErrorAction SilentlyContinue
If (!($vmname))
	{
		Write-output $vmTemplate "does not exist. Checking to see if any templates exist"
		$vmtemplate = get-template * | Where-object {($_.name -like "XATemp-*")}
		If (!($vmTemplate)) 
		{
			throw "no template exists"
		}
	}
}


#Now creat the XenApp VMs
If ($CreateVMs.length -gt 0)
{
Write-Progress -Activity "Creating New XenApp servers" -Completed -Id 200 -PercentComplete 0

foreach ($CreateVM in $CreateVMs)

{

#make sure that the vm that you are creating has a name in the approved structure
	If ($createVM -notlike '*0*')
		{
			write-error "You must enter the VMWare name for the servers to create, Such as XenApp01"
			Exit
		}

Write-Progress -Activity "Creating New XenApp servers" -CurrentOperation "Begining to process $createvm" -Id 200
	
	
	Write-verbose "Processing vm $createVM"
	

#check to see if there is VM with the name that we are planning to use. If so 
	$VMName = get-VM * | where-object {($_.name -like $createVM+"*")} #-and ($_.PowerState -eq "PoweredON")}
	If ($vmname -like '*0*')
		{
		write-verbose "VMName -like *0* $vmname"
		write-verbose $VMName.name.substring(7,1)
		$NewName = $VMName.name.substring(0,9)+$month
		}	Else		{
		$NewName = $createVM + '-' + $Month
		$VMName = $NewName
		$message = 'VMName is ' + $newName
		Write-verbose $message
		}
	

	#check if powered on - if it is turn it off
	if ($VMName.powerstate -eq "PoweredON")
		{
			Write-host ( "turning off VM" + $VMName)
			stop-VMGuest -vm $VMName -server vcenter2.kcs.com -confirm:$False
		} 
	
$OSCustSpec = Get-OSCustomizationSpec -Name $CreateVM


	#create new VM
Write-Progress -Activity "Creating New XenApp servers" -CurrentOperation "creating new VM from template. This process can take 15 minutes to complete"  -Id 200
	Write-verbose "creating new VM from template. This process can take 15 minutes to complete"
	$task = New-VM -name $NewName -Template $vmTemplate -Datastore VMFS-XenApp -OSCustomizationSpec $OSCustSpec  -ResourcePool "Citrix" -location Citrix -RunAsync -ErrorAction:Stop
	$NewVMName += $NewName
	

$timetaken = 1
	while($task.ExtensionData.Info.State -eq "running") 
		{
			#keep this as output to show people things are still happening
Write-Progress -Activity "Creation of $createvm" -CurrentOperation "creating new VM from template." -Completed -Id 210 -ParentId 200 -Status "Please be paitent" -PercentComplete ([int32]$timetaken/1200)
			Write-verbose "still working"
			sleep 30
			$timetaken = $TimeTaken+30
			$task.ExtensionData.UpdateViewData('Info.State')
		}
Write-Progress -Activity "Creation of $createvm" -CurrentOperation "creating new VM from template is now complete" -Completed -Id 210 -ParentId 200 -Status "Thanks for being paitent" -PercentComplete 100
	Write-Verbose "Vm Creation finished"  
	Write-verbose "Starting VM $newname"
	Start-VM $NewName -Confirm:$False -ErrorAction:Stop
	$message = "finished processing $VMName"
	Write-verbose $message

Write-Progress -Activity "Creating New XenApp servers" -CurrentOperation "Finished creating $createvm" -Id 200 

}


foreach ($VMName in $NewVMName)
	{
	Write-Progress -Activity "Creating New XenApp servers" -CurrentOperation "Customizing $createvm" -Id 200 

		Write-verbose "Waiting for VM to Finish Customization $VMName"
		$VM = get-vm $VMName
	Write-Progress -Activity "Creating New XenApp servers" -CurrentOperation "Checking to see if $createvm has finished being syspreped" -Id 200 
		
		
		Write-verbose "check for sysprep on $vm"
		
		WaitForSysprep ($vm)

	Write-Progress -Activity "Creating New XenApp servers" -CurrentOperation "waiting for Customization to complete for $createvm" -Id 200 
		Write-verbose "Customization is complete for $VMName"
		WaitForToolsToReportIP ($VMName)
		
		#use this area to connect to the new XenApp server and run the bluebeam default PDF reader command
		#Write-output "Making BlueBeam default PDF reader"
		
		#Invoke-Command -ComputerName (($vm).guest.hostname) -scriptblock {& 'c:\administrators\bluebeam.cmd'}
	#Write-Progress -Activity "Creating New XenApp servers" -CurrentOperation "$createvm is ready to be used. Don't forget to register Bluebeam" -Id 200 
	#	Write-output $VMName "Is ready to go. Don't forget to register Bluebeam"

	Wait-tools -vm $VMName
	invoke-VMScript -vm xenapp03-feb -scripttext "&'C:\Program Files\Bluebeam Software\Bluebeam Revu\2016\Pushbutton PDF\PbMngr5.exe' /register"
	}
}

# SIG # Begin signature block
# MIIJXQYJKoZIhvcNAQcCoIIJTjCCCUoCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUe86Z9tAAszKmCZWZUB/3cSJk
# ME6gggbOMIIGyjCCBbKgAwIBAgITGgAAAAtIEqyWvVA2zAADAAAACzANBgkqhkiG
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
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUffcos8Bm
# RkYTCduSyu2GEmtpM+UwDQYJKoZIhvcNAQEBBQAEggEAN2QOuDJwTOYNw5z19hfe
# thuln+lgq7379afqL30I3wgnkPq1LKsiKKbcb2LiBP9pCcvnr79loKaSGkGM7oxm
# pthMbnWHGx3ghzzdy5pvxoA0UHFJAoM2/47x7RxLiGPIhAvJpgfjTUsGi5Y6+ZX6
# pcdWkctn8rNbJ+W/WnrmUGc5Y8khM/Jg/TAW25YjPjpPCZiEjRol6wEf7PSgnJS2
# g3lZf+9vztng8+tkrmKYVniJV62LsSYPBsnD1Cisjhj9Z1OA/oJkPH4Iolp4Le1p
# nskfLJvUTbeZt85ct4p65O4rmzRDJkpEsl1k6EkTRNUKd87y5LNwk4KdOHan6KoL
# Bw==
# SIG # End signature block
