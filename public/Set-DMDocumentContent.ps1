# Set-DMDocumentContent.ps1
# Sets the document content data for the specified DocNum / version
# Aug 2017
# Copyright 2017 Dramatic Development
# If this works, it was written by Victor Vogelpoel (victor@victorvogelpoel.nl).
# If it doesn't, I don't know who wrote it.


function Set-DMDocumentContent
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true, position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage='TODO')]
        [Alias('DocNumber', 'DocumentNumber')] 
        [long]$DocNum,

        [Parameter(Mandatory=$false, position=1, ValueFromPipelineByPropertyName=$true, HelpMessage='TODO')]
        [ValidateNotNullOrEmpty()]
        [String]$FormName,

        [Parameter(Mandatory=$true, position=2, ValueFromPipelineByPropertyName=$true, HelpMessage='TODO')]
        [ValidateNotNullOrEmpty()]
        [String]$VersionID,

        [Parameter(Mandatory=$true, position=3, ValueFromPipelineByPropertyName=$true, HelpMessage='TODO')]
        [byte[]]$Content
    )

    begin
    {
        Assert-DMLibraryConnected
    }

    process
    {
        # IMPORTANT: Did you lock the document profile first?
        # Otherwise: ERROR 0x800401D1 in PCDPutDocClass: WIN-G8GEKHHS9OE: WIN-G8GEKHHS9OE: Document is unlocked


        $formNameSpecified = $PSBoundParameters.ContainsKey('FormName')

        if (!$formNameSpecified)
        {
            # FormName was not specified; retrieve the formname that the document profile was stored with.
            $resolvedFormName = Resolve-DMFormName -DocNum $DocNum

            if ([String]::IsNullOrEmpty($resolvedFormName))
            {
                # TODO: better exception
                throw "Cannot resolve Form for DocNum $DocNum."
            }

            $FormName = $resolvedFormName
        }

        Write-Verbose "Using Form `"$FormName`" for DocNum $DocNum."

        $loginSession = Get-DMLoginSession -Library (Get-DMCurrentLibrary)
        $library      = $loginSession.LoginLibrary
        $DST          = $loginSession.DST

        try
        {
            $pcdPutDoc    = $null  # initialize variables for the final() in case the New-Object fails
            $pcdPutStream = $null

            $pcdPutDoc    = New-Object Hummingbird.DM.Server.Interop.PCDClient.PCDPutDocClass
            #$pcdPutStream = New-Object Hummingbird.DM.Server.Interop.PCDClient.PCDPutStreamClass

            [void]$pcdPutDoc.SetDST($loginSession.DST)
            [void]$pcdPutDoc.AddSearchCriteria($eDOCSTokens.TARGET_LIBRARY,  $library)
            [void]$pcdPutDoc.AddSearchCriteria($eDOCSTokens.DOCUMENT_NUMBER, $DocNum.ToString())
            [void]$pcdPutDoc.AddSearchCriteria($eDOCSTokens.VERSION_ID,      $VersionID.ToString())

            [void]$pcdPutDoc.Execute()
            $pcdPutDoc | Assert-DMOperationSuccess

            $pcdPutDoc.NextRow()
            $pcdPutStream = $pcdPutDoc.GetPropertyValue($eDOCSTokens.CONTENT)
            if ($null -eq $pcdPutStream)
            {
                throw "ERROR while requesting putstream for setting document version content."
            }

            $DMConfiguration   = Get-DMConfiguration
            
            $dataSize          = $Content.Length
            $maxBytesWriteSize = $DMConfiguration.MaxComWriteSize
            if ($maxBytesWriteSize -eq 0 -or $maxBytesWriteSize -gt $dataSize)
            {
                $maxBytesWriteSize = $dataSize
            }

            $bytesWritten = 0
            do
            {
                # Take a part of the content bytes, by size of $maxBytesWriteSize
                $bytesPart        = $Content[$bytesWritten..($bytesWritten+$maxBytesWriteSize-1)]

                $bytesPartWritten = 0
                [void]$pcdPutStream.Write($bytesPart, $maxBytesWriteSize, [ref]$bytesPartWritten)
                $pcdPutStream | Assert-DMOperationSuccess

                if ($bytesPartWritten -ne $maxBytesWriteSize)
                {
                    Write-Warning "Number of bytes written to putstream ($bytesPartWritten) is not equal to number of bytes that should have been written: $maxBytesWriteSize"
                }

                $bytesWritten += $bytesPartWritten

            } while ($bytesWritten -lt $dataSize)


            [void]$pcdPutStream.SetComplete()
            $pcdPutStream | Assert-DMOperationSuccess
        }
        finally
        {
            if ($null -ne $pcdPutStream)
            {
                [VOID][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($pcdPutStream)
            }

            if ($null -ne $pcdPutDoc)
            {
                [VOID][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($pcdPutDoc)
            }            
        }
    }
}



<#
    Set-DMDocumentContent -DocNum 4051933 -VersionID 1526079 -Content 128,129,130
#>
