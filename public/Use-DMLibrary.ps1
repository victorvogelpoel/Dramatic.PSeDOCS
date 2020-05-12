# Use-DMLibrary.ps1
# Selects the library the other CmdLets are working on to a connected library.
# Aug 2017
# Copyright 2017 Dramatic Development
# If this works, it was written by Victor Vogelpoel (victor@victorvogelpoel.nl).
# If it doesn't, I don't know who wrote it.

# This is the current library Cmdlets will be working on
[string]$script:CurrentLibrary = ''



function Use-DMLibrary
{
    [CmdletBinding()]
    param
    (
        [Parameter(position=0, mandatory=$true, ValueFromPipeLineByPropertyName=$true, helpMessage='TODO')]
        [ValidateNotNullOrEmpty()]
        [String]$Library
    )

    process
    {
        $loginSession = Get-DMLoginSession -Library $Library
        if ($null -eq $loginSession)
        {
            throw "You are not logged on to library `"$Library`"; use Connect-DMLibrary -Library `"$Library`" to log in."
        }

        # If we got here, then we are logged on to this library
        $script:CurrentLibrary = $Library
    }
}
