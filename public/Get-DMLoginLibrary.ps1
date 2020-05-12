# Get-DMLoginLibrary.ps1
# Returns the valid eDOCS libraries.
# Jul 2017
# Copyright 2017 Dramatic Development
# If this works, it was written by Victor Vogelpoel (victor@victorvogelpoel.nl).
# If it doesn't, I don't know who wrote it.

function Get-DMLoginLibrary
{
    [CmdletBinding()]
    param
    ()

    process
    {
        try
        {
            $pcdLoginLibs = New-Object Hummingbird.DM.Server.Interop.PCDClient.PCDGetLoginLibsClass

            # Fetch the loginlibraries from the eDOCS server
            [void]$pcdLoginLibs.Execute()
            $pcdLoginLibs | Assert-DMOperationSuccess -ExceptionMessage 'ERROR 0x{0:X} while fetching eDOCS libraries: {1}'

            # Now get the library names from the fetched data
            for ($i=0; $i -lt $pcdLoginLibs.GetSize(); $i++)
            {
                $lib = $pcdLoginLibs.GetAt($i)
                $pcdLoginLibs | Assert-DMOperationSuccess -Exception 'ERROR 0x{0:X} while building eDOCS library name list: {1}'

                # Return the library name (as a string)
                $lib -as [string]
            }
        }
        finally
        {
            if ($null -ne $pcdLoginLibs)
            {
                [VOID][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($pcdLoginLibs)

                $pcdLoginLibs = $null
            }
        }

    }
}