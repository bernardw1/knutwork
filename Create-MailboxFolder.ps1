#
# Create-MailboxFolder.ps1
#
# By David Barrett, Microsoft Ltd. 2016. Use at your own risk.  No warranties are given.
#
#  DISCLAIMER:
# THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
# MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR
# A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL
# MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS,
# BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE
# SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION
# OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.

# blog with comments on how script works https://blogs.msdn.microsoft.com/emeamsgdev/2013/10/20/powershell-create-folders-in-users-mailboxes/
# got script from https://gallery.technet.microsoft.com/office/Create-folders-in-users-4630c241

param (
	[Parameter(Position=0,Mandatory=$False,HelpMessage="Specifies the mailbox to be accessed")]
	[ValidateNotNullOrEmpty()]
	[string]$Mailbox,
	
	[Parameter(Position=1,Mandatory=$True,HelpMessage="Specifies the folder(s) that should be checked/created.  For multiple folders, separate using semicolon")]
	[ValidateNotNullOrEmpty()]
	[string]$RequiredFolders,
	
	[Parameter(Position=2,Mandatory=$False,HelpMessage="The folder that should contain the subfolders (default is Inbox)")]
	[string]$ParentFolder,
	
	[Parameter(Position=3,Mandatory=$False,HelpMessage="The folder class (default is IPF.note)")]
	[string]$FolderClass = "IPF.Note",
	
	[Parameter(Mandatory=$False,HelpMessage="When specified, the archive mailbox is accessed (instead of the main mailbox)")]
	[switch]$Archive,
		
	[Parameter(Mandatory=$False,HelpMessage="Credentials used to authenticate with EWS")]
    [System.Management.Automation.PSCredential]$Credentials,
				
	[Parameter(Mandatory=$False,HelpMessage="Username used to authenticate with EWS")]
	[string]$Username,
	
	[Parameter(Mandatory=$False,HelpMessage="Password used to authenticate with EWS")]
	[string]$Password,
	
	[Parameter(Mandatory=$False,HelpMessage="Domain used to authenticate with EWS")]
	[string]$Domain,
	
	[Parameter(Mandatory=$False,HelpMessage="Whether we are using impersonation to access the mailbox")]
	[switch]$Impersonate,
	
	[Parameter(Mandatory=$False,HelpMessage="EWS Url (if omitted, then autodiscover is used)")]	
	[string]$EwsUrl,
	
	[Parameter(Mandatory=$False,HelpMessage="Path to managed API (if omitted, a search of standard paths is performed)")]	
	[string]$EWSManagedApiPath,
	
	[Parameter(Mandatory=$False,HelpMessage="Whether to ignore any SSL errors (e.g. invalid certificate)")]	
	[switch]$IgnoreSSLCertificate,
	
	[Parameter(Mandatory=$False,HelpMessage="Whether to allow insecure redirects when performing autodiscover")]	
	[switch]$AllowInsecureRedirection,

	[Parameter(Mandatory=$False,HelpMessage="If specified, no changes will be applied")]	
	[switch]$WhatIf,

	[Parameter(Mandatory=$False,HelpMessage="Log file - activity is logged to this file if specified")]	
	[string]$LogFile = "",

	[Parameter(Mandatory=$False,HelpMessage="Trace file - if specified, EWS tracing information is written to this file")]	
	[string]$TraceFile
)


Function Log([string]$Details, [ConsoleColor]$Colour)
{
    if ($Colour -eq $null)
    {
        $Colour = [ConsoleColor]::White
    }
    Write-Host $Details -ForegroundColor $Colour
	if ( $LogFile -eq "" ) { return	}
	$Details | Out-File $LogFile -Append
}

Function LogVerbose([string]$Details)
{
    Write-Verbose $Details
	if ( $LogFile -eq "" ) { return	}
	$Details | Out-File $LogFile -Append
}

Function LoadEWSManagedAPI()
{
	# Find and load the managed API
	
	if ( ![string]::IsNullOrEmpty($EWSManagedApiPath) )
	{
		if ( { Test-Path $EWSManagedApiPath } )
		{
			Add-Type -Path $EWSManagedApiPath
			return $true
		}
		Write-Host ( [string]::Format("Managed API not found at specified location: {0}", $EWSManagedApiPath) ) -ForegroundColor Yellow
	}
	
	$a = Get-ChildItem -Recurse "C:\Program Files (x86)\Microsoft\Exchange\Web Services" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ( $_.Name -eq "Microsoft.Exchange.WebServices.dll" ) }
	if (!$a)
	{
		$a = Get-ChildItem -Recurse "C:\Program Files\Microsoft\Exchange\Web Services" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ( $_.Name -eq "Microsoft.Exchange.WebServices.dll" ) }
	}
	
	if ($a)	
	{
		# Load EWS Managed API
		Write-Host ([string]::Format("Using managed API {0} found at: {1}", $a.VersionInfo.FileVersion, $a.VersionInfo.FileName)) -ForegroundColor Gray
		Add-Type -Path $a.VersionInfo.FileName
        $script:EWSManagedApiPath = $a.VersionInfo.FileName
		return $true
	}
	return $false
}

Function CreateTraceListener($service)
{
    # Create trace listener to capture EWS conversation (useful for debugging)
    if ($script:Tracer -eq $null)
    {
        $Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
        $Params=New-Object System.CodeDom.Compiler.CompilerParameters
        $Params.GenerateExecutable=$False
        $Params.GenerateInMemory=$True
        $Params.IncludeDebugInformation=$False
	    $Params.ReferencedAssemblies.Add("System.dll") | Out-Null
        $Params.ReferencedAssemblies.Add($EWSManagedApiPath) | Out-Null

        $traceFileForCode = $traceFile.Replace("\", "\\")

        if (![String]::IsNullOrEmpty($TraceFile))
        {
            LogVerbose "Tracing to: $TraceFile"
        }

        $TraceListenerClass = @"
		    using System;
		    using System.Text;
		    using System.IO;
		    using System.Threading;
		    using Microsoft.Exchange.WebServices.Data;
		
            namespace TraceListener {
		        class EWSTracer: Microsoft.Exchange.WebServices.Data.ITraceListener
		        {
			        private StreamWriter _traceStream = null;
                    private string _lastResponse = String.Empty;

			        public EWSTracer()
			        {
				        try
				        {
					        _traceStream = File.AppendText("$traceFileForCode");
				        }
				        catch { }
			        }

			        ~EWSTracer()
			        {
                        Close();
			        }

                    public void Close()
			        {
				        try
				        {
					        _traceStream.Flush();
					        _traceStream.Close();
				        }
				        catch { }
			        }


			        public void Trace(string traceType, string traceMessage)
			        {
                        if ( traceType.Equals("EwsResponse") )
                            _lastResponse = traceMessage;

                        if ( traceType.Equals("EwsRequest") )
                            _lastResponse = String.Empty;

				        if (_traceStream == null)
					        return;

				        lock (this)
				        {
					        try
					        {
						        _traceStream.WriteLine(traceMessage);
						        _traceStream.Flush();
					        }
					        catch { }
				        }
			        }

                    public string LastResponse
                    {
                        get { return _lastResponse; }
                    }
		        }
            }
"@

        $TraceCompilation=$Provider.CompileAssemblyFromSource($Params,$TraceListenerClass)
        $TraceAssembly=$TraceCompilation.CompiledAssembly
        $script:Tracer=$TraceAssembly.CreateInstance("TraceListener.EWSTracer")
    }

    # Attach the trace listener to the Exchange service
    $service.TraceListener = $script:Tracer
}

function CreateService($targetMailbox)
{
    # Creates and returns an ExchangeService object to be used to access mailboxes

    # First of all check to see if we have a service object for this mailbox already
    if ($script:services -eq $null)
    {
        $script:services = @{}
    }
    if ($script:services.ContainsKey($targetMailbox))
    {
        return $script:services[$targetMailbox]
    }

    # Create new service
    $exchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1)

    # We enable tracing so that we can retrieve the last response (and read any throttling information from it - this isn't exposed in the EWS Managed API)
    CreateTraceListener $exchangeService
    $exchangeService.TraceFlags = [Microsoft.Exchange.WebServices.Data.TraceFlags]::All
    $exchangeService.TraceEnabled = $True

    # Set credentials if specified, or use logged on user.
    if ($Credentials -ne $Null)
    {
        LogVerbose "Applying given credentials"
        $exchangeService.Credentials = $Credentials.GetNetworkCredential()
    }
    elseif ($Username -and $Password)
    {
	    LogVerbose "Applying given credentials for $Username"
	    if ($Domain)
	    {
		    $exchangeService.Credentials = New-Object  Microsoft.Exchange.WebServices.Data.WebCredentials($Username,$Password,$Domain)
	    } else {
		    $exchangeService.Credentials = New-Object  Microsoft.Exchange.WebServices.Data.WebCredentials($Username,$Password)
	    }
    }
    else
    {
	    LogVerbose "Using default credentials"
        $exchangeService.UseDefaultCredentials = $true
    }

    # Set EWS URL if specified, or use autodiscover if no URL specified.
    if ($EwsUrl)
    {
    	$exchangeService.URL = New-Object Uri($EwsUrl)
    }
    else
    {
    	try
    	{
		    LogVerbose "Performing autodiscover for $targetMailbox"
		    if ( $AllowInsecureRedirection )
		    {
			    $exchangeService.AutodiscoverUrl($targetMailbox, {$True})
		    }
		    else
		    {
			    $exchangeService.AutodiscoverUrl($targetMailbox)
		    }
		    if ([string]::IsNullOrEmpty($exchangeService.Url))
		    {
			    Log "$targetMailbox : autodiscover failed" Red
			    return $Null
		    }
		    LogVerbose "EWS Url found: $($exchangeService.Url)"
    	}
    	catch
    	{
            Log "$targetMailbox : error occurred during autodiscover: $($Error[0])" Red
            return $null
    	}
    }
 
    if ($Impersonate)
    {
		$exchangeService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $targetMailbox)
	}

    $script:services.Add($targetMailbox, $exchangeService)
    return $exchangeService
}

Function Throttled()
{
    # Checks if we've been throttled.  If we have, we wait for the specified number of BackOffMilliSeconds before returning

    if ([String]::IsNullOrEmpty($script:Tracer.LastResponse))
    {
        return $false # Throttling does return a response, if we don't have one, then throttling probably isn't the issue (though sometimes throttling just results in a timeout)
    }

    $lastResponse = $script:Tracer.LastResponse.Replace("<?xml version=`"1.0`" encoding=`"utf-8`"?>", "")
    $lastResponse = "<?xml version=`"1.0`" encoding=`"utf-8`"?>$lastResponse"
    $responseXml = [xml]$lastResponse

    if ($responseXml.Trace.Envelope.Body.Fault.detail.MessageXml.Value.Name -eq "BackOffMilliseconds")
    {
        # We are throttled, and the server has told us how long to back off for

        # Increase our throttling delay to try and avoid throttling (we only increase to a maximum delay of 15 seconds between requests)
        if ( $script:throttlingDelay -lt 15000)
        {
            if ($script:throttlingDelay -lt 1)
            {
                $script:throttlingDelay = 2000
            }
            else
            {
                $script:throttlingDelay = $script:throttlingDelay * 2
            }
            if ( $script:throttlingDelay -gt 15000)
            {
                $script:throttlingDelay = 15000
            }
        }
        LogVerbose "Updated throttling delay to $($script:throttlingDelay)ms"

        # Now back off for the time given by the server
        Log "Throttling detected, server requested back off for $($responseXml.Trace.Envelope.Body.Fault.detail.MessageXml.Value."#text") milliseconds" Yellow
        Sleep -Milliseconds $responseXml.Trace.Envelope.Body.Fault.detail.MessageXml.Value."#text"
        Log "Throttling budget should now be reset, resuming operations" Gray
        return $true
    }
    return $false
}

function ThrottledFolderBind()
{
    param (
        [Microsoft.Exchange.WebServices.Data.FolderId]$folderId,
        $propset = $null)

    LogVerbose "Attempting to bind to folder $folderId"
    try
    {
        if ($propset -eq $null)
        {
            $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:service, $folderId)
        }
        else
        {
            $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:service, $folderId, $propset)
        }
        Sleep -Milliseconds $script:throttlingDelay
        if (-not ($folder -eq $null))
        {
            LogVerbose "Successfully bound to $($folderId): $($folder.DisplayName)"
        }
        return $folder
    }
    catch
    {
        $Error[0]
    }

    if (Throttled)
    {
        try
        {
            if ($propset -eq $null)
            {
                $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:service, $folderId)
            }
            else
            {
                $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:service, $folderId, $propset)
            }
            return $folder
        }
        catch {}
    }
    else
    {
        # We weren't throttled, so report the error
        Log "Error occurred attempting to bind to folder: $($Error[0])" Red
    }

    # If we get to this point, we have been unable to bind to the folder
    return $null
}

function GetFolderPath($Folder)
{
    # Return the full path for the given folder

    # We cache our folder lookups for this script
    if (!$script:folderCache)
    {
        # Note that we can't use a PowerShell hash table to build a list of folder Ids, as the hash table is case-insensitive
        # We use a .Net Dictionary object instead
        $script:folderCache = New-Object 'System.Collections.Generic.Dictionary[System.String,System.Object]'
    }

    $propset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, [Microsoft.Exchange.WebServices.Data.FolderSchema]::ParentFolderId)

    if ($Folder -eq "\")
    {
        # Special handling for root folder
        if  ($script:folderCache.ContainsKey("\"))
        {
            return $script:folderCache["\"]
        }
        $mbx = New-Object Microsoft.Exchange.WebServices.Data.Mailbox( $Mailbox )
        $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$wkf, $mbx )
        $rootFolder = ThrottledFolderBind $folderId $propset
        if ($rootFolder)
        {
            $folderPath = "\$($rootFolder.DisplayName)"
            $script:folderCache.Add("\", $folderPath)
            $script:FolderCache.Add($rootFolder.Id.UniqueId, $rootFolder)
            return $folderPath
        }
        return ""
    }
    else
    {
        $parentFolder = ThrottledFolderBind $Folder.Id $propset
        $folderPath = $Folder.DisplayName
        $parentFolderId = $Folder.Id
    }

    while ($parentFolder.ParentFolderId -ne $parentFolderId)
    {
        if ($script:folderCache.ContainsKey($parentFolder.ParentFolderId.UniqueId))
        {
            $parentFolder = $script:folderCache[$parentFolder.ParentFolderId.UniqueId]
        }
        else
        {
            $parentFolder = ThrottledFolderBind $parentFolder.ParentFolderId $propset
            $script:FolderCache.Add($parentFolder.Id.UniqueId, $parentFolder)
        }
        $folderPath = "$($parentFolder.DisplayName)\$folderPath"
        $parentFolderId = $parentFolder.Id
    }
    return $folderPath
}

Function GetFolder()
{
	# Return a reference to a folder specified by path
	
	$RootFolder, $FolderPath, $Create = $args[0]
	
    if ( $RootFolder -eq $null )
    {
        LogVerbose "GetFolder called with null root folder"
        return $null
    }

    $s = GetFolderPath $RootFolder
    LogVerbose "GetFolder: root folder is $s"
    LogVerbose "GetFolder: requested folder path is $FolderPath"

	$Folder = $RootFolder
	if ($FolderPath -ne '\')
	{
		$PathElements = $FolderPath -split '\\'
		For ($i=0; $i -lt $PathElements.Count; $i++)
		{
			if ($PathElements[$i])
			{
				$View = New-Object  Microsoft.Exchange.WebServices.Data.FolderView(2,0)
				$View.PropertySet = [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly
						
				$SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $PathElements[$i])
				
                $FolderResults = $Null
                try
                {
				    $FolderResults = $Folder.FindFolders($SearchFilter, $View)
                    Sleep -Milliseconds $script:throttlingDelay
                }
                catch {}
                if ($FolderResults -eq $Null)
                {
                    if (Throttled)
                    {
                    try
                    {
				        $FolderResults = $Folder.FindFolders($SearchFilter, $View)
                    }
                    catch {}
                    }
                }
                if ($FolderResults -eq $null)
                {
                    return $null
                }

				if ($FolderResults.TotalCount -gt 1)
				{
					# We have more than one folder returned... We shouldn't ever get this, as it means we have duplicate folders
					$Folder = $null
					Write-Host "Duplicate folders ($($PathElements[$i])) found in path $FolderPath" -ForegroundColor Red
					break
				}
                elseif ( $FolderResults.TotalCount -eq 0 )
                {
                    if ($Create)
                    {
                        # Folder not found, so attempt to create it
					    $subfolder = New-Object Microsoft.Exchange.WebServices.Data.Folder($script:service)
					    $subfolder.DisplayName = $PathElements[$i]
                        try
                        {
					        $subfolder.Save($Folder.Id)
                            LogVerbose "Created folder $($PathElements[$i])"
                        }
                        catch
                        {
					        # Failed to create the subfolder
					        $Folder = $null
					        Log "Failed to create folder $($PathElements[$i]) in path $FolderPath" Red
					        break
                        }
                        $Folder = $subfolder
                    }
                    else
                    {
					    # Folder doesn't exist
					    $Folder = $null
					    Log "Folder $($PathElements[$i]) doesn't exist in path $FolderPath" Red
					    break
                    }
                }
                else
                {
				    $Folder = ThrottledFolderBind $FolderResults.Folders[0].Id
                }
			}
		}
	}
	
	return $Folder
}

Function CreateFolders()
{
	$pFolder = $args[0]
	if (!$pFolder) { return }

    $pFolderPath = GetFolderPath $pFolder
	
	foreach ($requiredFolder in $FolderCheckList)
	{
		LogVerbose "Checking for existence of $requiredFolder in $pFolderPath"
		$rf = GetFolder( $pFolder, $requiredFolder, $false )
		if ( $rf )
		{
			Log "$requiredFolder already exists" Green
		}
		Else
		{
			# Create the folder
			if (!$WhatIf)
			{
				$rf = New-Object Microsoft.Exchange.WebServices.Data.Folder($service)
				$rf.DisplayName = $requiredFolder
                $rf.FolderClass = $FolderClass
				$rf.Save($pFolder.Id)
				if ($rf.Id.UniqueId)
				{
					Log "$requiredFolder created successfully" Green
				}
			}
			Else
			{
				Log "$requiredFolder would be created" Yellow
			}
		}
	}
}

Function ProcessMailbox()
{
    # Process the mailbox
    if ( [string]::IsNullOrEmpty($Mailbox) )
    {
        Log "ProcessMailbox called with no mailbox set" Red
        return
    }
    Write-Host ([string]::Format("Processing mailbox {0}", $Mailbox)) -ForegroundColor Gray
	$script:service = CreateService($Mailbox)
	if ($script:service -eq $Null)
	{
		Log "Failed to create ExchangeService" Red
        return
	}

    $script:throttlingDelay = 0

    # Bind to root folder	
    $mbx = New-Object Microsoft.Exchange.WebServices.Data.Mailbox( $Mailbox )
    $Folder = $Null
    if ([String]::IsNullOrEmpty($ParentFolder))
    {
        $ParentFolder = "wellknownfoldername.Inbox"
    }
    if ($ParentFolder.ToLower().StartsWith("wellknownfoldername."))
    {
        # Well known folder specified (could be different name depending on language, so we bind to it using WellKnownFolderName enumeration)
        $wkf = $ParentFolder.SubString(20)
        LogVerbose "Attempting to bind to well known folder: $wkf"
        $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$wkf, $mbx )
        $Folder = ThrottledFolderBind($folderId)
    }
    else
    {
        $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $mbx )
        $Folder = ThrottledFolderBind($folderId)
        if ($Folder -and ($ParentFolder -ne "\"))
        {
	        $Folder = GetFolder($Folder, $ParentFolder, $false)
        }
    }

	if (!$Folder)
	{
		Log "Failed to find folder $ParentFolder" Red
		return
	}

	CreateFolders $Folder
}

Function CurrentUserPrimarySmtpAddress()
{
    # Attempt to retrieve the current user's primary SMTP address
    $searcher = [adsisearcher]"(samaccountname=$env:USERNAME)"
    $result = $searcher.FindOne()

    if ($result -ne $null)
    {
        $mail = $result.Properties["mail"]
        return $mail
    }
    return $null
}

Function TrustAllCerts() {
    <#
    .SYNOPSIS
    Set certificate trust policy to trust self-signed certificates (for test servers).
    #>

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
        namespace Local.ToolkitExtensions.Net.CertificatePolicy {
        public class TrustAll : System.Net.ICertificatePolicy {
            public TrustAll()
            { 
            }
            public bool CheckValidationResult(System.Net.ServicePoint sp,
                                                System.Security.Cryptography.X509Certificates.X509Certificate cert, 
                                                System.Net.WebRequest req, int problem)
            {
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
}


# The following is the main script

if ( [string]::IsNullOrEmpty($Mailbox) )
{
    $Mailbox = CurrentUserPrimarySmtpAddress
    if ( [string]::IsNullOrEmpty($Mailbox) )
    {
	    Write-Host "Mailbox not specified.  Failed to determine current user's SMTP address." -ForegroundColor Red
	    Exit
    }
    else
    {
        Write-Host ([string]::Format("Current user's SMTP address is {0}", $Mailbox)) -ForegroundColor Green
    }
}

# Check if we need to ignore any certificate errors
# This needs to be done *before* the managed API is loaded, otherwise it doesn't work consistently (i.e. usually doesn't!)
if ($IgnoreSSLCertificate)
{
	Write-Host "WARNING: Ignoring any SSL certificate errors" -foregroundColor Yellow
    TrustAllCerts
}

# Load EWS Managed API
if (!(LoadEWSManagedAPI))
{
	Write-Host "Failed to locate EWS Managed API, cannot continue" -ForegroundColor Red
	Exit
}
  

if ($RequiredFolders.Contains(";"))
{
	# Have more than one folder to check, so convert to array
	$FolderCheckList = $RequiredFolders -split ';'
}
else
{
	$FolderCheckList = $RequiredFolders
}


# Check whether we have a CSV file as input...
$FileExists = Test-Path $Mailbox
If ( $FileExists )
{
	# We have a CSV to process
    Write-Verbose "Reading mailboxes from CSV file"
	$csv = Import-CSV $Mailbox -Header "PrimarySmtpAddress"
	foreach ($entry in $csv)
	{
        Write-Verbose $entry.PrimarySmtpAddress
        if (![String]::IsNullOrEmpty($entry.PrimarySmtpAddress))
        {
            if (!$entry.PrimarySmtpAddress.ToLower().Equals("primarysmtpaddress"))
            {
		        $Mailbox = $entry.PrimarySmtpAddress
			    ProcessMailbox
            }
        }
	}
}
Else
{
	# Process as single mailbox
	ProcessMailbox
}

if ($script:Tracer -ne $null)
{
    $script:Tracer.Close()
}
# SIG # Begin signature block
# MIIJXQYJKoZIhvcNAQcCoIIJTjCCCUoCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUAeQU3AEs8SRsW7aCDxo3w2Jr
# gGSgggbOMIIGyjCCBbKgAwIBAgITGgAAAAtIEqyWvVA2zAADAAAACzANBgkqhkiG
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
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUtWfseJX5
# 5LUJSaf5gR0U2P5lsWswDQYJKoZIhvcNAQEBBQAEggEAVY4jW3x6wjW5gdiYiivr
# 9KWQNrvmU02PObNxc2Vd39rg2hmw5Af7rvGioXhGy0iMojPERJoo/UjsSZjFIR7o
# uDB1euEjzX8oqK4K6jcwtEgJDOh9f4BXD7zdR2wQ/dwOcjg8j9heyv1OS8oDUqX8
# K+QW+0TiGrmykI84skopFcEf+qFSjOJ7VestJH22Mg7c9eI/jRYarVpbZ34gb7bW
# uzgnqhpg9nU7plAPRLQlwIPJA5WFiMa/m4AjhtvJCJWY/UplbTrI8FahDyEf0ibJ
# mAeMbRsDB3O5aGUjI4Njb/fUJgpMq2RozMNf+Xv3Wx55bcYOB+BZSSnfGUEJycy4
# yA==
# SIG # End signature block
