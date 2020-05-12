# Get-DMDocumentVersion.ps1
# Gets the document versions data for the specified DocNum
# Aug 2017
# Copyright 2017 Dramatic Development
# If this works, it was written by Victor Vogelpoel (victor@victorvogelpoel.nl).
# If it doesn't, I don't know who wrote it.


function Get-DMDocumentVersion
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true, position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage='TODO')]
        [Alias('DocNumber', 'DocumentNumber')] 
        [long]$DocNum
    )

    begin
    {
        Assert-DMLibraryConnected
    }

    process
    {
        $pcdSearch    = $null   # initialize variables for the final() in case the New-Object fails

        $loginSession = Get-DMLoginSession -Library (Get-DMCurrentLibrary)
        $library      = $loginSession.LoginLibrary
        $DST          = $loginSession.DST

        # And now, get the document versions
        try
        {
            # first, get the fileextensions from the document versions for the DocNum
            $versionFileNames = Invoke-DMSqlCmd -SQLCommand "SELECT PATH, VERSION_ID FROM DOCSADM.COMPONENTS WHERE DOCNUMBER=$DocNum"

            $pcdSearch = New-Object Hummingbird.DM.Server.Interop.PCDClient.PCDSearchClass

            [void]$pcdSearch.SetDST($DST)
            [void]$pcdSearch.AddSearchLib($library)
            [void]$pcdSearch.SetSearchObject('cyd_cmnversions')
            [void]$pcdSearch.AddSearchCriteria($eDOCSColumns.DOCNUMBER, $DocNum)

            [void]$pcdSearch.AddReturnProperty($eDOCSColumns.VERSION_ID)
            [void]$pcdSearch.AddReturnProperty($eDOCSColumns.VERSION)
            [void]$pcdSearch.AddReturnProperty($eDOCSColumns.SUBVERSION)
            [void]$pcdSearch.AddReturnProperty($eDOCSColumns.VERSION_LABEL)
            [void]$pcdSearch.AddReturnProperty($eDOCSColumns.COMMENTS)
            [void]$pcdSearch.AddReturnProperty($eDOCSColumns.LASTEDITDATE)
            [void]$pcdSearch.AddReturnProperty($eDOCSColumns.LASTEDITTIME)

            [void]$pcdSearch.AddOrderByProperty($eDOCSColumns.VERSION, 0)
            [void]$pcdSearch.AddOrderByProperty($eDOCSColumns.VERSION_LABEL, 0)

            [void]$pcdSearch.Execute()
            $pcdSearch | Assert-DMOperationSuccess

            $rows = $pcdSearch.GetRowsFound()
            for ($i=1; $i -le $rows; $i++)
            {
                [void]$pcdSearch.SetRow($i)

                $versionID    = [long]$pcdSearch.GetPropertyValue($eDOCSColumns.VERSION_ID).ToString()
                $versionLabel = $pcdSearch.GetPropertyValue($eDOCSColumns.VERSION_LABEL).ToString()
                $versionNum   = [int]($pcdSearch.GetPropertyValue($eDOCSColumns.VERSION).ToString())
                $subVersion   = $pcdSearch.GetPropertyValue($eDOCSColumns.SUBVERSION).ToString()

                if ($versionNum -eq 0 -or $versionLabel -eq 'PR1')
                {
                    # Version 0 indicates an attachment; VersionLabel PR1 is the Preview Attachment; DISREGARD this row
                    Continue
                }

                $lastEditDate = [DateTime]$pcdSearch.GetPropertyValue($eDOCSColumns.LASTEDITDATE)
                $lastEditTime = [DateTime]$pcdSearch.GetPropertyValue($eDOCSColumns.LASTEDITTIME)
                $comments     = $pcdSearch.GetPropertyValue($eDOCSColumns.COMMENTS).ToString()
                # Create a single LastEdited DateTime
                $lastEdited   = $lastEditDate.Add($lastEditTime.TimeOfDay)

                # Output the version object
                [PSCustomObject][ORDERED]@{
                    PSTYPENAME    = 'Dramatic.eDOCS.DocumentVersion'
                    VersionID     = $versionID
                    Version       = $versionLabel
                    # VersionNum    = $versionNum
                    # SubVersion    = $subVersion
                    LastEdited    = $lastEdited
                    Comments      = $comments
                    FileExtension = ($versionFileNames | Where VERSION_ID -eq $versionID | foreach { 
                        if (![String]::IsNullOrEmpty($_.PATH)) 
                        {
                            [System.IO.Path]::GetExtension($_.PATH)
                        }
                    })
                }
            }
        }
        finally
        {
            if ($pcdSearch -ne $null)
            {
                [VOID][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($pcdSearch)
            }         
        }
    }
}


<#

---------------------------------------
Get-DMDocumentVersion -DocNum 4051936

Returns:
Version SubVersion VersionLabel
------- ---------- ------------
      3 !          3           
      2 !          2           
      1 !          1     


---------------------------------------
Get-DMDocumentVersion -DocNum 4051932

Version SubVersion VersionLabel
------- ---------- ------------
      2 A          2A          
      2 !          2           
      1 !          1           

#>