# Test-DMLibraryConnectedps1
# Test if a library is connected
# Aug 2017
# Copyright 2017 Dramatic Development
# If this works, it was written by Victor Vogelpoel (victor@victorvogelpoel.nl).
# If it doesn't, I don't know who wrote it.



function Test-DMLibraryConnected
{
    [CmdletBinding()]
    param
    ()

    process
    {
        return (![String]::IsNullOrEmpty((Get-DMCurrentLibrary)))
    }
}


