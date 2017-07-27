<#
http://stackoverflow.com/questions/10921563/extract-the-report-of-room-calendar-from-exchage-server-using-powershell-scripti

Save as Get-RoomUsage.ps1 optional parameters are -StartDate and -Enddate


#>

param 
(

[DateTime]$StartDate = (Get-Date).addDays(-13),
[DateTime]$EndDate = (Get-date),
[System.Management.Automation.CredentialAttribute()]$psCred = (Get-Credential -username 'kcs admin' -message "Please provide the credentials of an account that has access to all the room mailboxes")
)


#Check to see if a connection MS Exchagne (MEX2) has already been established in this PowerShell Session
if (!(Get-PSSession | Where { $_.ConfigurationName -eq "Microsoft.Exchange" })) 
{ 
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://mex2.kcs.com/PowerShell/ -Authentication Kerberos
    
    Import-PSSession $Session -ErrorAction Stop 
}

## Load Managed API dll  
###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
$EWSDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
if (Test-Path $EWSDLL)
    {
    Import-Module $EWSDLL
    }
else
    {
    "$(get-date -format yyyyMMddHHmmss):"
    "This script requires the EWS Managed API 1.2 or later."
    "Please download and install the current version of the EWS Managed API from"
    "http://go.microsoft.com/fwlink/?LinkId=255472"
    ""
    "Exiting Script."
    exit
    } 

$Mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails RoomMailbox #| select -First 3


$rptcollection = @()


## Load Managed API dll  
#Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll"  
## Set Exchange Version  
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013

## Create Exchange Service Object  
$service = New-Object    Microsoft.Exchange.WebServices.Data.ExchangeService


## Set Credentials to use two options are availible Option1 to use explict credentials or Option 2 use the Default (logged On) credentials  

#Credentials Option 1 using UPN for the windows Account  
#Added to Paramater area
#$psCred = Get-Credential -username 'kcs admin' -message "Please provide the credentials of an account that has access to all the room mailboxes"

$creds = New-Object    System.Net.NetworkCredential($psCred.GetNetworkCredential().username.ToString(),$psCred.GetNetworkCredential().password.ToString(),$psCred.GetNetworkCredential().domain.ToString())  
$service.Credentials = $creds   

#Credentials Option 2  

#$service.UseDefaultCredentials = $true  

## Choose to ignore any SSL Warning issues caused by Self Signed Certificates  



## Code From http://poshcode.org/624
## Create a compilation environment
$Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
$Compiler=$Provider.CreateCompiler()
$Params=New-Object System.CodeDom.Compiler.CompilerParameters
$Params.GenerateExecutable=$False
$Params.GenerateInMemory=$True
$Params.IncludeDebugInformation=$False
$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

$TASource=@'
  namespace Local.ToolkitExtensions.Net.CertificatePolicy{
    public class TrustAll : System.Net.ICertificatePolicy {
      public TrustAll() { 
      }
      public bool CheckValidationResult(System.Net.ServicePoint sp,
        System.Security.Cryptography.X509Certificates.X509Certificate cert, 
        System.Net.WebRequest req, int problem) {
        return true;
      }
    }
  }
'@ 
$TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
$TAAssembly=$TAResults.CompiledAssembly

## We now create an instance of the TrustAll and attach it to the ServicePointManager
$TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
[System.Net.ServicePointManager]::CertificatePolicy=$TrustAll

## end code from http://poshcode.org/624

## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use  

#CAS URL Option 1 Autodiscover  
#$service.AutodiscoverUrl($MailboxName,{$true})  
#"Using CAS Server : " + $Service.url   

#CAS URL Option 2 Hardcoded  

$uri=[system.URI] "https://mail.knutsonconstruction.com/ews/exchange.asmx"  
$service.Url = $uri    

$obj = @{}
$start = Get-Date 
$i = 0
foreach($Mailbox in $Mailboxes)
{
	$i +=1
	$Duration = (New-TimeSpan -Start ($start) -End (Get-Date)).totalseconds
	$TimeLeft = ($Duration/$i)*($mailboxes.count - $i)
	Write-Progress -Status "$($Mailbox.DisplayName)" -Activity "Mailbox $i of $($Mailboxes.Count)" -PercentComplete ($i/$($mailboxes.count)*100) -SecondsRemaining $timeleft -Id 100


	$WorkingDays = ($Mailbox | Get-mailboxCalendarConfiguration).WorkDays.ToString()
	$WorkingHoursStartTime = ($Mailbox |    Get-mailboxCalendarConfiguration).WorkingHoursStartTime
	$WorkingHoursEndTime = ($Mailbox | Get-mailboxCalendarConfiguration).WorkingHoursEndTime

	if($WorkingDays -eq "Weekdays"){$WorkingDays = "Monday,Tuesday,Wednesday,Thursday,Friday"}
	if($WorkingDays -eq "AllDays"){$WorkingDays = "Monday,Tuesday,Wednesday,Thursday,Friday,Saturday,Sunday"}
	if($WorkingDays -eq "WeekEndDays"){$WorkingDays = "Saturday,Sunday"}
	
	## Optional section for Exchange Impersonation  
	$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $Mailbox.PrimarySMTPAddress) 

	$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$Mailbox.PrimarySMTPAddress.tostring())   
	$Calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service,$folderid)

	if($Calendar.TotalCount -gt 0){
		$cvCalendarview = new-object    Microsoft.Exchange.WebServices.Data.CalendarView($StartDate,$EndDate,2000)
	$cvCalendarview.PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)

	$frCalendarResult = $Calendar.FindAppointments($cvCalendarview)
	$inPolicy = New-TimeSpan 
	$OutOfPolicy = New-TimeSpan
	$TotalDuration = New-timespan
	$BookableTime = New-TimeSpan
	$totalItems = $frCalendarResult.Items.count
	$c = 0
	[DateTime]$dtStart = $dtEnd = '1/1/2015'
	foreach ($apApointment in $frCalendarResult.Items){
		 $c +=1
		 Write-Progress -Status "$($Mailbox.DisplayName)" -Activity "Mailbox $i of $($Mailboxes.Count)" -PercentComplete ($i/$($mailboxes.count)*100) -SecondsRemaining $timeleft -Id 100 -CurrentOperation "Processing calendarItem $c of $($frCalendarResult.Items.count)"
		 $psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
		 $apApointment.load($psPropset)
		 if($apApointment.IsAllDayEvent -eq $false)
		 {

			 if($apApointment.Duration)
			 {
					if($WorkingDays.split(",") -contains ($apApointment.start).dayofweek)
					{
						$TotalDuration = $TotalDuration.add((new-timespan -End $apApointment.End.tolongTimeString() -start $apApointment.start.tolongTimeString()))

						#Only count to inPolicy if within the workinghours time
						if("{0:HH:mm}" -f [datetime]$apApointment.start.tolongTimeString() -lt $WorkingHoursStartTime)
						{   
							$tStart = $WorkingHoursStartTime.ToString()
						
						}   
						else
						{
							$tStart = $apApointment.start.ToLongTimeString()
						}
						
						if("{0:HH:mm}" -f [datetime]$apApointment.End.tolongTimeString() -gt $WorkingHoursEndTime)
						{   
							$tEnd = $WorkingHoursEndTime.ToString()
						}   
						else
						{
							$tEnd = $apApointment.End.ToLongTimeString()
						}
	
						
						$Duration = New-TimeSpan -Start $tStart -End $tEnd
						$inPolicy = $inPolicy.add($Duration)
	
		$tend = $null
					}
				}
		 }

	}

	#Calculate to total hours of bookable time between the 2 dates
	for ($d=$Startdate;$d -le $Enddate;$d=$d.AddDays(1)){
	  if ($WorkingDays.split(",") -contains $d.DayOfWeek) 
	  {
		$BookableTime += $WorkingHoursEndTime - $WorkingHoursStartTime
	  }
	  } #for
 
	#Save result....
	$rptobj = "" | Select samAccountName,DisplayName,Number-of-Meetings,inPolicy,Out-Of-Policy,TotalDuration,BookableTime,BookedPersentage
	$rptobj.samAccountName = $Mailbox.samAccountName
	$rptobj.DisplayName = $Mailbox.DisplayName
	$rptobj."Number-of-Meetings" = $totalItems
	$rptobj.inPolicy =  '{0:f2}' -f ($inPolicy.TotalHours)
	$rptobj."Out-Of-Policy" =  '{0:f2}' -f (($TotalDuration.TotalHours - $inPolicy.TotalHours))
	$rptobj.TotalDuration =  '{0:f2}' -f ($TotalDuration.TotalHours)
	$rptobj.BookableTime =  '{0:f2}' -f ($BookableTime.TotalHours)
	$rptobj.BookedPersentage =  '{0:f2}' -f (($TotalDuration.TotalHours / $BookableTime.TotalHours) * 100)
	$rptcollection += $rptobj

	} #ForEach
}

$rptcollection
# SIG # Begin signature block
# MIIJXQYJKoZIhvcNAQcCoIIJTjCCCUoCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU+BITt2PsyejfHlTOdzqZzW84
# tUSgggbOMIIGyjCCBbKgAwIBAgITGgAAAAtIEqyWvVA2zAADAAAACzANBgkqhkiG
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
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUK8xf2HjA
# lIrQuYEav0o/Tt2Ekt8wDQYJKoZIhvcNAQEBBQAEggEAIj61Dy/Kjrs6vAKYe5j0
# R4iPk1HuNqu/bvDAWLhl7vHS79Fm8AcbBJpqtVyB2QVDaWoAFIpenTQACU2pfHir
# N9Hy7Ro76Dk6L8OlgCFrZCdRJhaRLBNG7si8YyefmKMOeBkI2Pv18+vqf9Xp7rq1
# jV64haF3ktOGbNPC8c/M+DwL4ANOGlZvzK4ZV+TA/vWkAQjikFLx3qGpBcf0M9lS
# BwKXePJBXBH0di0QKwmMIpWZX4avdgcVb8hJHYV/cgWUavEoPOCdifOcd+CfAtHr
# TxzSNvLrEU+0aF1OhqtVDrmF64k2VZ1bDqjPvnFiZ9/BdPuSNR68NS3plLdPV4gb
# 1w==
# SIG # End signature block
