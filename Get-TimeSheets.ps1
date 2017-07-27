<#

Script Name: TimeSheets to eCMS
Created: 10/26/2016
Created By: Bernard Welmers

This script will pull all the Excel files from a specified directory and process the data. Once the data is processed it will then output a CSV file with information to be uploaded to eCMS



#>

param 
(
[String]$Path = (Get-Location).path, #'C:\Users\bwelmers\Documents\Timesheet\',
[DateTime]$WeekEndDate = @(@(-7..0) | % {$(Get-Date).AddDays($_)} | ? {$_.DayOfWeek -ieq "Sunday"})[0]
)

Function GetDetails {
	$TimeSheetDetail.Batchno = $Batchno
	#$TimeSheetDetail.EmpName = $CurXLSX[1].Col4
	$TimeSheetDetail.EmpNo = $CurXLSX[2].Col4
	$TimeSheetDetail.EmpCompany = $CurXLSX[1].Col1.substring(0,2)
	Switch ($CurXLSX[$row].Col7)	{
		'R' {
			$TimeSheetDetail.OtherHoursType = $null
			$TimeSheetDetail.RegHours = $CurXLSX[$row].$Col
			}
		'VA'{
			$TimeSheetDetail.OtherHoursType = 'VA'
			$TimeSheetDetail.OtherHours = $CurXLSX[$row].$Col
			}
		'S' {
			$TimeSheetDetail.OtherHoursType = 'S'
			$TimeSheetDetail.OtherHours = $CurXLSX[$row].$Col
			}
		'HL' {
			$TimeSheetDetail.OtherHoursType = 'HL'
			$TimeSheetDetail.OtherHours = $CurXLSX[$row].$Col
			}
		'OV' {
			$TimeSheetDetail.OtherHoursType = 'OV'
			$TimeSheetDetail.OtherHours = $CurXLSX[$row].$Col
			}

		}
	$TimeSheetDetail.DayOfWeek = $ColNo - 7
	$TimeSheetDetail.WeekNo = 1
	$TimeSheetDetail.WeekEndingDate = $WeekEndDate
}

Function ErrorMess {Param ($Message)
	Write-Error $Message
	$Message | Out-File ($Path + '\' + $BatchNo + '.log') -append
}

Function AddLog  {Param ($Message)
	$Message | Out-File ($Path + '\' + $BatchNo + '.log') -append
}


<# Check for Prerequisists
Checks for Powershell version 5, if it is not found then it will prompt with a location to run it from
Check to see if the Import Excel module has been installed and if it is currently loaded
code taken from https://blogs.technet.microsoft.com/heyscriptingguy/2010/07/11/hey-scripting-guy-weekend-scripter-checking-for-module-dependencies-in-windows-powershell/
#>
IF ($path.Substring($path.lENGTH - 1,1)= '/'){$path.Substring(0,$path.lENGTH - 1)}

$BatchNo = ([string]([DateTime]::now.ToOADate())).substring(0,5) + [string](get-date).Hour + [string](get-date).Minute  
			#([string](get-date).Year).substring(2,2) + [string](get-date).Month  + [string](get-date).Day + [string](get-date).Minute
AddLog "This batch was run on $(Get-Date)"
AddLog "The batch number is $BatchNo"
			
			
#Check Powershell version
If ($PSVersionTable.PSVersion.Major -lt 5){
	ErrorMess 'You must install at least Powershell version 5 to run this script. You can get it from https://www.microsoft.com/en-us/download/details.aspx?id=50395'
	Exit
}

#Check if ImportExcel is installed
if(-not(Get-Module -name 'ImportExcel')) 	{ 
	write-progress -Status "Getting all the pieces ready to start working" -activity "loading the Import-Excel module"
	if(Get-Module -ListAvailable | 
		Where-Object { $_.name -eq 'ImportExcel' }) { 
		Import-Module -Name 'ImportExcel' 
	} #end if module available then import 
	else { 
		ErrorMess 'Please install the ImportExcel module. This will be attempted automatically now.'
		Install-module ImportExcel -scope CurrentUser
		if(Get-Module -ListAvailable | 
			Where-Object { $_.name -eq 'ImportExcel' }) { 
			Import-Module -Name 'ImportExcel' 
			Write-host 'The module has now been installed correctly'
		} #end if module available then import 
	} #module not available 
} # end if not module 

#Pre Processing information (this is the same for ALL Timesheets)

$FileNames = (Get-ChildItem $path *.xlsx).fullname
$WeekNo = Get-date ($weekenddate) -uformat %V	
$TimeSheet = @()
$ItemNo = 1
$start = Get-Date 
$CurrentVersion = 1

ForEach ($FileName in $FileNames){
	$Duration = (New-TimeSpan -Start ($start) -End (Get-Date)).totalseconds
	$TimeLeft = ($Duration/$ItemNo)*($FileNames.count - $ItemNo)
	If ($TimeLeft -lt 1){$TimeLeft = 1}
	
	Write-Progress -Status "Working on $FileName" -Activity "Importing timesheet $itemno of $($FileNames.Count)" -PercentComplete ($itemno/$($FileNames.count)*100) -SecondsRemaining $timeleft -Id 100
	
	#Import the Excel spreadsheet and identify the particular tab
	Try {
		$CurXLSX = (Import-excel -path $FileName -WorkSheetname 'Office Timesheet') | Select-Object -First 50
	}
	Catch{
		ErrorMess ("The Excel file <$FileName > does not have the worksheet <Office Timesheet> ")
		Continue
	}

	If ($CurXLSX[0].COl15 -gt ($CurrentVersion + 1) -OR $CurXLSX[0].COl15 -lt ($CurrentVersion)){
		ErrorMess "The Excel file $FileName has a timesheet with an old version" 
		Continue
	}
	
	If ($CurXLSX[2].Col4 -lt 0){
		ErrorMess "The Excel file $FileName does not have an Employee Number" 
		Continue
	}
	
	
	#Get the current week end date (that is specified in the spreadsheet and check if it is this week
	$TSWeekEnd = [datetime]::fromOADate($CurXLSX[1].col15)

	if ($TSWeekEnd -ne $WeekEndDate)
		{
			#stop processing if this is not the current week end date
			$TSFolder = [string]$TSWeekEnd.year + [string]$TSWeekEnd.Month + [string]$TSWeekEnd.Day
			AddLog "This was not the current processing week end date for $FileName. It has been moved to the folder $TSFolder"

			#Check if Week end date folder exists, If it does not then create a folder
			if (-not(test-path ($path + '\' +$TSFolder))){
				New-item ($Path + '\' +$TSFolder + '\') -type Directory
			}
			
			#Move the time sheet to the week end folder
			Move-item $FileName -destination $TSFolder
		
			#Now go do the next item
			continue
		}

	#Some static details for the timesheet
		$HomeDpt = $CurXLSX[2].Col15.substring(0,3)
		$HomeComp = $CurXLSX[1].Col1.substring(0,2)

	#Define the variables being used	
		$row = 5
		
		:GetTimesheetDetails Do {
			If ($CurXLSX[$row].Col15 -gt 0)
				{
				$ColNo = 8
				
				:GetDaylyHours Do {
					$col = 'Col'+$ColNo
						If ($CurXLSX[$row].$col)
						{
							$TimeSheetDetail = ""| Select Batchno,CostType,DistComp,Dept,WeekEndingDate,DayOfWeek,EmpCompany,EmpNo,JobNo,CostCode,OtherHours,OtherHoursType,RegHours,SubJobNo,WeekNo    #,EmpName
							$TimeSheetDetail.DistComp = $CurXLSX[$row].Col1#.substring(0,2)
							$TimeSheetDetail.JobNo = $CurXLSX[$row].Col2
							$TimeSheetDetail.SubJobNo = $CurXLSX[$row].Col3
							$TimeSheetDetail.CostCode = ([string]$CurXLSX[$row].Col4).Substring(0,9)
							$TimeSheetDetail.CostType = $CurXLSX[$row].Col5
							GetDetails
							$TimeSheet += $TimeSheetDetail
						}
					$ColNo+=1
					}While ($ColNo -lt 15)
				
				}
			$row+=1
		} While ($CurXLSX[$row].Col2 -ne 'Overhead')


	#Get the overhead/departmental charges
		$stopRow = 1

		:GetTimesheetDetails Do {
			$ColNo = 8
				
			:GetDaylyHours Do {
				[string]$col = 'Col'+$ColNo
					If ($CurXLSX[$row].$col)
					{
						$TimeSheetDetail = ""| Select Batchno,CostType,DistComp,INDIDV,Dept,WeekEndingDate,DayOfWeek,EmpCompany,INDEEDV,EmpNo,JobNo,CostCode,OtherHours,OtherHoursType,OVTime,RegHours,SubJobNo,WeekNo
						$TimeSheetDetail.Dept = $HomeDpt
						$TimeSheetDetail.DistComp = $HomeComp
						$TimeSheetDetail.CostType = $CurXLSX[$row].Col5
						GetDetails
						$TimeSheet += $TimeSheetDetail
					}
				$ColNo+=1
				}While ($ColNo -lt 15)
				
			$row+=1
			$Stoprow +=1
		} While ($stopRow -le 4)

	#Move row marker to Payroll Adjustment line
	$row+=3
	$ColNo = 8

	:GetDaylyHours Do {
		[string]$col = 'Col'+$ColNo
			If ($CurXLSX[$row].$col)
			{
				$TimeSheetDetail = ""| Select Batchno,CostType,DistComp,INDIDV,Dept,WeekEndingDate,DayOfWeek,EmpCompany,INDEEDV,EmpNo,JobNo,CostCode,OtherHours,OtherHoursType,OVTime,RegHours,SubJobNo,WeekNo
				$TimeSheetDetail.Dept = $HomeDpt
				$TimeSheetDetail.DistComp = $HomeComp
				$TimeSheetDetail.CostType = $CurXLSX[$row].Col5
				GetDetails
				$TimeSheet += $TimeSheetDetail
			}
		$ColNo+=1
		}While ($ColNo -lt 15)

	#Move row marker to Overtime Hours
	<# Overtime Hourse are not going to be processed by this script/upload routine at this time
		$row+=2
		$ColNo = 8

		Do {
			[string]$col = 'Col'+$ColNo
				If ($CurXLSX[$row].$col)
				{
					$TimeSheetDetail = ""| Select Batchno,CostType,DistComp,INDIDV,Dept,WeekEndingDate,DayOfWeek,EmpCompany,INDEEDV,EmpNo,JobNo,CostCode,OtherHours,OtherHoursType,OVTime,RegHours,SubJobNo,WeekNo
					$TimeSheetDetail.Dept = $HomeDpt
					$TimeSheetDetail.CostType = $CurXLSX[$row].Col5
					GetDetails
					$TimeSheetDetail.OVtime = $CurXLSX[$row].$Col
					$TimeSheet += $TimeSheetDetail
				}
			$ColNo+=1
			}While ($ColNo -lt 15)
	#>
	
# Moving is not currently working
	
	#Move processed time sheet to week end folder
	$TSFolder = [string]$TSWeekEnd.year + [string]$TSWeekEnd.Month + [string]$TSWeekEnd.Day
		
	#Check if Week end date folder exists, If it does not then create a folder
	if (-not(test-path ($path + '\' +$TSFolder))){
		New-item ($Path + '\' +$TSFolder + '\') -type Directory
	}
	
	#Move the time sheet to the week end folder
	Move-item $FileName -destination $TSFolder
#>	
	AddLog "Completed processing $FileName at $(Get-Date)"
}
	
$TimeSheet | Export-csv ($Path + '\' + $BatchNo + '.csv') -NoTypeInformation

# SIG # Begin signature block
# MIIJXQYJKoZIhvcNAQcCoIIJTjCCCUoCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUtHOLWe7yD0rpWNaM583ul3Xn
# F4ygggbOMIIGyjCCBbKgAwIBAgITGgAAAAtIEqyWvVA2zAADAAAACzANBgkqhkiG
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
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUxap1ECZj
# pvwY6rFtWoN4vAZRX6MwDQYJKoZIhvcNAQEBBQAEggEAPEirV86fdxNEary0jj45
# YeguCrzsikE99KzSVdBGkffVCLM654vETPzdEfazqzvIqou0mvBIqyHkTe08lhSe
# DRT0TeMm2u18bGwnpuKs9u1VGOiXr3Se/0irtYjRslTZcfrE3tttarLIof1dZbRG
# cq746cZ65lGs8g9b3v5nH4s/QcqwTSJ55b4qH5GJzeKbgnTht6JbeSA8qF9KLbVW
# bcydi8httbWCx4kG3kaNGHAZK8Bvn2gKoty9sOJXaNaMAIIuhrQv4i4fWz80QvqI
# 3nr/CHhxVkmi+U1Y0pdTCz/QSIzzdhhpS6pnmphiDQX7FB5C3tr5TrpZDjPCrPRE
# Ww==
# SIG # End signature block
