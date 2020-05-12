# Get-DMDocumentContent.ps1
# Gets the document content data for the specified DocNum / version, reads it into memory and returns it to the caller.
# Aug 2017
# Copyright 2017 Dramatic Development
# If this works, it was written by Victor Vogelpoel (victor@victorvogelpoel.nl).
# If it doesn't, I don't know who wrote it.


function Get-DMDocumentContent
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
        [String]$VersionID
    )

    begin
    {
        Assert-DMLibraryConnected
    }

    process
    {
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
            $pcdGetDocObject = $null  # initialize variables for the final() in case the New-Object fails
            $pcdGetStream    = $null

            $pcdGetDocObject = New-Object Hummingbird.DM.Server.Interop.PCDClient.PCDGetDocClass

            [void]$pcdGetDocObject.SetDST($loginSession.DST)
            [void]$pcdGetDocObject.SetSearchObject($FormName)
            [void]$pcdGetDocObject.AddSearchCriteria($eDOCSTokens.TARGET_LIBRARY, $library)
            [void]$pcdGetDocObject.AddSearchCriteria($eDOCSTokens.DOCUMENT_NUMBER, $DocNum.ToString())
            [void]$pcdGetDocObject.AddSearchCriteria($eDOCSTokens.VERSION_ID, $VersionID.ToString())

            [void]$pcdGetDocObject.Execute()
            $pcdGetDocObject | Assert-DMOperationSuccess

            # Find the %CONTENT property
            $rows = $pcdGetDocObject.GetRowsFound() # should be 1
            [void]$pcdGetDocObject.SetRow(1)
            $returnProps   = $pcdGetDocObject.GetReturnProperties()                  # Apparently to get the CONTENT property, GetReturnProperties() must be called first.
            $contentStream = $pcdGetDocObject.GetPropertyValue($eDOCSTokens.CONTENT)
            $pcdGetDocObject | Assert-DMOperationSuccess

            $pcdGetStream  = $contentStream
            # Get the size of the stream
            $streamSizeProp = $pcdGetStream.GetPropertyValue($eDOCSTokens.ISTREAM_LOWPART)
            $streamSize     = [int]$streamSizeProp.ToString()

            # Now read the bytes
            [byte[]]$bytes = @()
            if ($streamSize -gt 0)
            {
                [int]$bytesRead = 0
                $bytes = $pcdGetStream.Read($streamSize, [ref]$bytesRead)

                if ($bytesRead -ne $streamSize)
                {
                    throw "ERROR while reading stream for DocNum $DocNUm versionID $VersionID`: number of bytes read $bytesRead is less than bytes to be read $streamSize`."
                }
            }

            # Output the bytes to the caller
            $bytes
        }
        finally
        {
            if ($null -ne $pcdGetStream)
            {
                [VOID][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($pcdGetStream)
            }

            if ($null -ne $pcdGetDocObject)
            {
                [VOID][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($pcdGetDocObject)
            }            
        }
    }
}



<#
    Get-DMDocumentContent -DocNum 4051936 -VersionID 1526068

    Returns bytes[]
#>
