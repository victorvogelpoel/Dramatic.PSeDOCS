# Assert-DMLibraryConnected.ps1
# throw an exception if no login session exists for the specified library or any at all
# Jul 2017
# Copyright 2017 Dramatic Development
# If this works, it was written by Victor Vogelpoel (victor@victorvogelpoel.nl).
# If it doesn't, I don't know who wrote it.



function Assert-DMLibraryConnected
{
    [CmdletBinding()]
    param
    ()

    process
    {
        if ([String]::IsNullOrEmpty((Get-DMCurrentLibrary)))
        {
            throw 'You are not logged on; use Connect-DMLibrary to log in into an eDOCS DM library.'
        }
    }
}


