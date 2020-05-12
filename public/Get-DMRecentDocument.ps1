# Get-DMRecentDocument.ps1
# 
# Jul 2017
# Copyright 2017 Dramatic Development
# If this works, it was written by Victor Vogelpoel (victor@victorvogelpoel.nl).
# If it doesn't, I don't know who wrote it.



function Get-DMRecentDocument
{
    [CmdLetBinding()]
    param
    ()

    begin
    {
        Assert-DMLibraryConnected
    }


    process
    {
        try
        {
            # -------------------------------------------------------------------------------
            $PCDRecentDocClass = New-object Hummingbird.DM.Server.Interop.PCDClient.PCDRecentDocClass

            # TODO: nakijken return properties
            $returnProperties = "DOCNUM", "DOCNAME", "AUTHOR_ID", "APP_ID", "STATUS", "%RECENTACTIVITYDATE", "%RECENTACTIVITYTIME"

            $loginSession = Get-DMLoginSession -Library (Get-DMCurrentLibrary)

            [void]$PCDRecentDocClass.SetDST($loginSession.DST)
            [void]$PCDRecentDocClass.AddSearchLib($loginSession.LoginLibrary)
            [void]$PCDRecentDocClass.SetSearchObject("def_qbe")  #mandatory
            $returnProperties | foreach { [void]$PCDRecentDocClass.AddReturnProperty($_) }

            [void]$PCDRecentDocClass.AddOrderByProperty("%RECENTACTIVITYDATE", $false)
            [void]$PCDRecentDocClass.AddOrderByProperty("%RECENTACTIVITYTIME", $false)

            # TODO: Filter for current user? -> Switch parameter toevoegen?
            # [void]$PCDRecentDocClass.AddSearchCriteria("AUTHOR_ID", "VVO")

            [void]$PCDRecentDocClass.Execute()
            $PCDRecentDocClass | Assert-DMOperationSuccess

            # Gather the results
            while ($PCDRecentDocClass.NextRow())
            {
                # Initialize a result object
                $resultRow = [PSCustomObject][ORDERED]@{
                    PSTypeName = 'Dramatic.eDOCS.RecentDocument'
                }

                # Propagate the result object with return properties + values
                $returnProperties | foreach {
                    $resultRow.Add($_, $PCDRecentDocClass.GetPropertyValue($_))
                }
    
                # And return the result
                $resultRow
            }
        }
        finally
        {
            if ($PCDRecentDocClass -ne $null)
            {
                [VOID][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($PCDRecentDocClass)
            }
        }
    }
}

<#

    Get-DMRecentDocument


#>