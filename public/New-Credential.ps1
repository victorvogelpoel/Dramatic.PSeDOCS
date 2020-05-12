# New-Credential.ps1
# Create a new PSCredential
#
# Dec 2014
# If this works, Victor Vogelpoel <victor@victorvogelpoel.nl> wrote this.
# If it doesn't, I don't know who wrote this.

#requires -version 2
Set-PSDebug -Strict
Set-StrictMode -Version Latest


function New-Credential
{ 
	[CmdletBinding()]
	param
	( 
		[Parameter(Position=0, Mandatory=$true, HelpMessage="User name, must be in format DOMAIN\USER")]
		[string]$UserName, 
		
		[Parameter(Position=1, Mandatory=$true, HelpMessage="Password, cannot be empty")]
		[string]$Password
	) 
		
	# Create the credential 
	$spwd = ConvertTo-SecureString -AsPlainText $Password -Force  
	$cred = New-Object System.Management.Automation.PSCredential $UserName, $spwd 
	
	# Now return it to the caller 
	return $cred 
	
	
<# 
    .SYNOPSIS 
       A function to create a credential object from script. 
    .DESCRIPTION 
       Enables you to create a credential objects from stored details. 
    .NOTES 
        File Name  : New-Credential.ps1 
        Author     : Thomas Lee - tfl@psp.co.uk 
        Requires   : PowerShell Version 2.0 
    .LINK 
        This script posted to: 
            http://pshscripts.blogspot.com/2011/03/new-credentialps1.html 
    .PARAMETER UserId 
       The userid in the form of "domain\user" 
    .PARAMETER Password 
       The password for this user 
    .EXAMPLE 
       New-Credential contoso\administrator  Pa$$w0rd 
#> 
} 