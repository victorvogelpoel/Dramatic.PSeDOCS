[CmdletBinding()]
param
()

# Dramatic.PSeDOCS.psm1
# Module for working with OpenText eDOCS DM
#
# Jul 2017
# Copyright 2017 Dramatic Development
# If this works, it was written by Victor Vogelpoel (victor@victorvogelpoel.nl).
# If it doesn't, I don't know who wrote it.
#
# TODO: Copyright, license and disclaimer
#

# PREREQUISITES
# - eDOCS and the eDOCS server API must have been installed on the workstation.
# - eDOCS is 32-bit software, which will NOT work in 64-bit PowerShell session; COM errors will occur. Unless...
#   - On 32-bit Windows, PowerShell is 32-bit as well and this module will work fine.
#   - When Windows is 64-bit and you don't have the eDOCS 64-bit API installed, you can only use this module in 32-bit PowerShell (x86).
#   - When Windows is 64-bit and you have installed the eDOCS 64-bit API, you can use this module in either 32-bit or 64-bit PowerShell.


#requires -version 4
Set-PSDebug -Strict
Set-StrictMode -Version Latest

#----------------------------------------------------------------------------------------------------------------------
# Set variables
$script:thisModuleDirectory			= $PSScriptRoot								# Directory \Path\Modules\Dramatic.PSeDOCS\
$script:modulesDirectory			= $PSScriptRoot | Split-Path				# Directory \Path\Modules\
$script:rootDirectory				= $Script:modulesDirectory | Split-Path 	# Directory \Path\


# Add /Modules/ directory to the environment PSModulePath, if not present
if ($env:PSModulePath.Split(";") -notcontains $script:modulesDirectory)
{
	$env:PSModulePath = $script:modulesDirectory + ";" + $env:PSModulePath
}


#----------------------------------------------------------------------------------------------------------------------





#----------------------------------------------------------------------------------------------------------------------
# INSTRUCTIONS
# - Use Write-MFDebug in functions for writing DEBUG messages for this module only.
# - Use Write-MFVerbose in functions for writing Verbose messages for this module only.

# Define the global variable DMVerbosePreference, which controls if Write-DMVerbose outputs anything
[System.Management.Automation.ActionPreference]$DMVerbosePreference	= 'SilentlyContinue'

function Write-DMVerbose
{
	[CmdLetBinding(SupportsShouldProcess=$false)]
	param
	(	
		[Parameter(Mandatory=$true, position=0, ValueFromPipeline=$true, ValueFromPipeLineByPropertyName=$true, HelpMessage="Specify the message to display. This parameter is required. You can also pipe a message string to Verbose-Message" )]
		[string]$Message
	)
	
	process
	{
		$verbose = ($PSBoundParameters['Verbose'] -eq $true)
		
		if ($DMVerbosePreference -eq 'Continue' -or $VerbosePreference -eq 'Continue' -or $verbose)
		{
			# Call powershell's own function, regardles overrides of Write-Verbose
			Microsoft.PowerShell.Utility\Write-Verbose -Message $Message -Verbose
		}
	}
	
	# TODO: Write-DMVerbose Help
}


# Define the global variable DMDebugPreference
[System.Management.Automation.ActionPreference]$DMDebugPreference	= "SilentlyContinue"

function Write-DMDebug
{
	[CmdLetBinding(SupportsShouldProcess=$false)]
	param
	(	
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipeLineByPropertyName=$true, HelpMessage="Specifies the message to display. This parameter is required. You can also pipe a message string to Verbose-Message" )]
		[string]$Message
	)
	
	process
	{
		$debug = ($PSBoundParameters['debug'] -eq $true)
		
		if ($DMDebugPreference -eq 'Continue' -or ('Continue', 'Inquire' -contains $DebugPreference) -or $debug)
		{
            # Reset DebugPreference for now to 'Continue', because using '-Debug' sets DebugPreference to 'Inquire' 
            # and halts execution to ask the user for continuation
            $DebugPreference = 'Continue'

			# Call powershell's own function, regardles overrides of Write-Debug
			Microsoft.PowerShell.Utility\Write-Debug -Message $Message
		}
	}
	
	# TODO: Write-DMDebug Help
}






#----------------------------------------------------------------------------------------------------------------------
# Enumerate all PS1 files in this module directory and dot-source it into the module space.

# But ignore these files:
$ignoreCommandsForDotSourcing = @(
	'install.ps1'
)

# Expose these 
$exposingVariables  = @(
    'DMDebugPreference', 
    'DMVerbosePreference',
    'eDOCSTokens',
    'eDOCSColumns'
)


$publicFunctions   = @(Get-ChildItem -Path "$script:thisModuleDirectory\public\*.ps1"   -ErrorAction SilentlyContinue)
$internalFunctions = @(Get-ChildItem -Path "$script:thisModuleDirectory\internal\*.ps1" -ErrorAction SilentlyContinue)

#Write-Host $publicFunctions
#Write-Host $internalFunctions

foreach($file in @($publicFunctions + $internalFunctions))
{
	if ($ignoreCommandsForDotSourcing -notcontains $file.Name -and $file.Name -notlike '*.Tests.ps1' -and $file.Name -notlike '*_TEMP.ps1' )
	{
        try
        {
		    Write-Verbose "Importing functions from file '$($file.Name)' by dotsourcing `"$($file.Fullname)`""
		    . $file.Fullname
        }
        catch
        {
            Write-Error "Failed to import function $($file.FullName): $_"
        }

	}
	else
	{
		Write-Verbose "Ignoring file '$($_.Name)'"
	}
}

# Load 'Hummingbird.DM.Server.Interop.PCDClient.dll'
Add-DMPCDClientType

Export-ModuleMember -Variable $exposingVariables -Function ($publicFunctions | select -ExpandProperty BaseName)

<#
    TODO: Additional module help.
#>
