

$DocNum = '4051936'

# Getting a list of document versions for the specified DocNum


        try
        {
            $pcdSearch = New-Object Hummingbird.DM.Server.Interop.PCDClient.PCDSearchClass

            [void]$pcdSearch.SetDST($loginSession.DST)
            [void]$pcdSearch.AddSearchLib($loginSession.LoginLibrary)
            [void]$pcdSearch.SetSearchObject('cyd_cmnversions')
            [void]$pcdSearch.AddSearchCriteria($eDOCSColumns.DOCNUMBER, $DocNum)

            [void]$pcdSearch.AddReturnProperty($eDOCSColumns.VERSION_ID)
            [void]$pcdSearch.AddReturnProperty($eDOCSColumns.VERSION)
            [void]$pcdSearch.AddReturnProperty($eDOCSColumns.SUBVERSION)
            [void]$pcdSearch.AddReturnProperty($eDOCSColumns.VERSION_LABEL)
            [void]$pcdSearch.AddReturnProperty($eDOCSColumns.LASTEDITDATE)
            [void]$pcdSearch.AddReturnProperty($eDOCSColumns.LASTEDITTIME)

            [void]$pcdSearch.AddOrderByProperty($eDOCSColumns.LASTEDITDATE, 0)
            [void]$pcdSearch.AddOrderByProperty($eDOCSColumns.LASTEDITTIME, 0)
            [void]$pcdSearch.AddOrderByProperty($eDOCSColumns.VERSION_ID, 0)

            [void]$pcdSearch.Execute()
            if ($pcdSearch.ErrNumber -ne 0)
            {
                throw 'ERROR 0x{0:X} in {2}: {1}' -f $PCDErrorCapableObject.ErrNumber, $PCDErrorCapableObject.ErrDescription, $PCDErrorCapableObject.GetType().Name
            }

            $rows = $pcdSearch.GetRowsFound()
            for ($i=1; $i -le $rows; $i++)
            {
                [void]$pcdSearch.SetRow($i)

                $versionID = $pcdSearch.GetPropertyValue($eDOCSColumns.VERSION_ID).ToString()
                if ($versionID -eq '0')
                {
                    Continue
                }

                $versionLabel = $pcdSearch.GetPropertyValue($eDOCSColumns.VERSION_LABEL).ToString()
                if ($versionLabel -eq 'PR1')
                {
                    Continue
                }

                [PSCustomObject][ORDERED]@{
                    PSTYPENAME   = 'Dramatic.eDOCS.DocumentVersion'
                    VersionID    = $pcdSearch.GetPropertyValue($eDOCSColumns.VERSION).ToString()
                    SubVersion   = $pcdSearch.GetPropertyValue($eDOCSColumns.SUBVERSION).ToString()
                    VersionLabel = $versionLabel
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