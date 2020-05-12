

$credential = New-Object System.Management.Automation.PSCredential 'VVO', (ConvertTo-SecureString -AsPlainText 'TsifutlUid#6' -force) 
$login      = Connect-DMLibrary -DMLibrary POVOOPEN -Credential $credential


$sql = "SELECT FORM_NAME, FORM_TITL from DOCSADM.FORMS"

        try
        {
            $pcdSQL = New-Object Hummingbird.DM.Server.Interop.PCDClient.PCDSQLClass

            [void]$pcdSQL.SetLibrary($login.LoginLibrary)
            [void]$pcdSQL.SetDST($login.DST)
            [void]$pcdSQL.Execute($sql)
            $pcdSQL | Assert-DMOperationSuccess -ExceptionMessage ('ERROR 0x{0:X} while executing query: {1}Native SQL error: {2}' -f $pcdSQL.ErrNumber, $pcdSQL.ErrDescription, $pcdSQL.GetSQLErrorCode())
            
            # Get the 
            $columNames = @()
            for ($col=1; $col -le $pcdSQL.GetColumnCount(); $col++)
            {
                $columNames += $pcdSQL.GetColumnName($col)
            }

            for($row=1; $row -le $pcdSQL.GetRowCount(); $row++)
            {
                $data = [PSCustomObject][ORDERED]@{ PSTYPENAME="PSeDOCS.SQLData" }

                [void]$pcdSQL.SetRow($row)
                for($col=1; $col -le $columNames.Count; $col++)
                {
                    $data | Add-Member -Name $columNames[$col-1] -Value ($pcdSQL.GetColumnValue($col)) -MemberType NoteProperty
                }
                
                # Return the data
                $data
            }
            
            
        }
        finally
        {
            if ($pcdSQL -ne $null)
            {
                try
                {
                    [void]$pcdSQL.ReleaseResults()
                }
                finally
                {
                    [VOID][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($pcdSQL)
                    $pcdSQL = $null
                }
            }
        }
