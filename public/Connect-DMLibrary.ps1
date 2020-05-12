# Connect-DMLibrary.ps1
# Create a loginsession to an eDOCS library
# Jul 2017
# Copyright 2017 Dramatic Development
# If this works, it was written by Victor Vogelpoel (victor@victorvogelpoel.nl).
# If it doesn't, I don't know who wrote it.


function Connect-DMLibrary
{
    [CmdletBinding()]
    param
    (
        [Parameter(position=0, mandatory=$true, ValueFromPipeLineByPropertyName=$true, helpMessage='TODO')]
        [ValidateNotNullOrEmpty()]
        [String]$Library,

        [Parameter(position=1, mandatory=$true, ValueFromPipeLineByPropertyName=$true, HelpMessage='TODO')]
        [ValidateNotNull()]
        [PSCredential]$Credential
    )

    process
    {
        try
        {
            $PCDLogin = New-Object Hummingbird.DM.Server.Interop.PCDClient.PCDLoginClass

            # Prepare the login request
            [void]$PCDLogin.AddLogin(0, $Library, $Credential.UserName, $Credential.GetNetworkCredential().Password)
            $PCDLogin | Assert-DMOperationSuccess -ExceptionMessage ('ERROR 0x{0:X} preparing login: {1}' -f $PCDLogin.ErrNumber, $PCDLogin.ErrDescription)

            # Fire the request
            [void]$PCDLogin.Execute()
            $PCDLogin | Assert-DMOperationSuccess -ExceptionMessage ('ERROR 0x{0:X} attempting login: {1}' -f $PCDLogin.ErrNumber, $PCDLogin.ErrDescription)

            # Administer a loginsession
            $loginSession = Set-DMLoginSession -UserName $PCDLogin.GetDOCSUserName() -DST $PCDLogin.GetDST() -LoginLibrary $PCDLogin.GetLoginLibrary() -PrimaryGroup $PCDLogin.GetPrimaryGroup()

            # Selects the current library the other CmdLets will work on
            Use-DMLibrary -Library $PCDLogin.GetLoginLibrary() 

            # Return the login to the caller
            $loginSession
        }
        finally
        {
            if ($PCDLogin -ne $null)
            {
                [VOID][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($PCDLogin)
            }
        }
    }

}





# TESTING
<#

$credential = New-Object System.Management.Automation.PSCredential 'VVO', (ConvertTo-SecureString -AsPlainText 'MyVoiceIsMyPassport' -force)
Connect-DMLibrary -Library WANBETALERS -Credential $credential

#>