# Get-DMDocument.ps1
# Gets the document registration data for the specified DocNum
# Jul 2017
# Copyright 2017 Dramatic Development
# If this works, it was written by Victor Vogelpoel (victor@victorvogelpoel.nl).
# If it doesn't, I don't know who wrote it.


function Get-DMDocument
{
    [CmdLetBinding(DefaultParameterSetName='DocumentProfile')]
    param
    (
        # Parameterset 'DocumentProfile': DocNum, FormName and DestinationDirectory
        [Parameter(Mandatory=$true, position=0, ParameterSetName='DocumentProfile', ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage='TODO')]
        [Alias('DocNumber', 'DocumentNumber')] 
        [long]$DocNum,

        [Parameter(Mandatory=$false, position=2, ParameterSetName='DocumentProfile', ValueFromPipelineByPropertyName=$true, HelpMessage='TODO')]
        [ValidateNotNullOrEmpty()]
        [String]$FormName,

        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, HelpMessage='TODO')]
        [Alias('GetContent', 'Content')]
        [Switch]$GetLatestVersionContent
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


        # Define the DMDocument object
        $DMDocument = [PSCustomObject][ORDERED]@{
            PSTypeName          = 'Dramatic.eDOCS.Document'
            DocNum              = $DocNum
            FormName            = $FormName
            Library             = $library
            MetaData            = [ORDERED]@{}
            Versions            = @()
            LatestVersionID     = [int]0
            LatestVersion       = ''
            LatestVersionEdited = [DateTime]::MinValue
            LatestVersionContent= [Byte[]]@()
        }

        # Using a dynamically compiled C# class to get the document/metadata from eDOCS (a PowerShell version was crashing with the COM object properyList iterator)
        $tempDocument = [Dramatic.eDOCS.Client]::GetDocument($DST, $library, $FormName, $DocNum)

        # Copy the metaData
        $tempDocument.MetaData | Sort name | foreach {
            $DMDocument.MetaData[$_.Name] = $_.Value

            # TODO: Meta data als member properties toevoegen aan DMDocument PSCustomObject?
            #if ('' -eq $DMDocument.PSObject.Members.Match($_.Name))
            #{
            #    $DMDocument | Add-Member -MemberType NoteProperty -Name $_.Name -Value $_.Value
            #}
        }

        # Initialize file extenion for file to empty string
        $latestVersionFileExtension = ''
        $latestVersionLabel         = ''

        # And now, get the document versions
        $DMDocument.Versions      = Get-DMDocumentVersion -DocNum $DocNum

        # Retrieve the 'latest and greatest' version.
        if ($null -ne $DMDocument.Versions -and @($DMDocument.Versions).Count -gt 0)
        {
            $DMDocument.LatestVersionID      = $DMDocument.Versions[0].VersionID
            $DMDocument.LatestVersion        = $DMDocument.Versions[0].Version
            $DMDocument.LatestVersionEdited  = $DMDocument.Versions[0].LastEdited
                                             
            $latestVersionFileExtension      = $DMDocument.Versions[0].FileExtension
            $latestVersionLabel              = $DMDocument.Versions[0].Version

            if ($GetLatestVersionContent)
            {
                $DMDocument.LatestVersionContent = ( Get-DMDocumentContent -DocNum $DMDocument.DocNum -Form $FormName -VersionID $DMDocument.LatestVersionID )

<###########################################################################################################
                # EXAMPLE: construct a filename from meta data and write content bytes to disk

                # "[DOCNAME] #4051938 v3.DOC" or "[LIBRARY] #4051938 v3.DOC"
                $DMDocumentDocName               = $DMDocument.MetaData[$eDOCSColumns.DOCNAME]
                if ([String]::IsNullOrEmpty($DMDocumentDocName)) { $DMDocumentDocName = $library }
                $documentFileName                = "$docName (#$DocNum v$latestVersionLabel)$latestVersionFileExtension"
                $filePath                        = Join-Path -Path 'C:\PATH' -ChildPath $documentFileName
                # And write it to a file on disk
                $DMDocument.LatestVersionContent | Set-Content -Path $filePath -Encoding Byte
############################################################################################################>
            }
        }

        # And return the DMDocument to the caller
        $DMDocument




<# This PowerShell code crashes at $propList.BeginIter(); the C# version has no problems

        try
        {
            $pcdDocObject = New-Object Hummingbird.DM.Server.Interop.PCDClient.PCDDocObjectClass

            [void]$pcdDocObject.SetDST($loginSession.DST)
            [void]$pcdDocObject.SetObjectType($FormName)
            [void]$pcdDocObject.SetProperty($eDOCSTokens.TARGET_LIBRARY,    $loginSession.LoginLibrary)
            #[void]$pcdDocObject.SetProperty($eDOCSTokens.OBJECT_IDENTIFIER, $DocNum.ToString())
            [void]$pcdDocObject.SetProperty($eDOCSTokens.DOCUMENT_NUMBER, $DocNum.ToString())                          # ISE crashes when $DocNum argument is not a string

            # Get the profile information
            [void]$pcdDocObject.Fetch()
            $pcdDocObject | Assert-DMOperationSuccess

            # Define the DMDocument object
            $DMDocument = [PSCustomObject][ORDERED]@{
                PSTypeName          = 'Dramatic.eDOCS.Document'
                DocNum              = $DocNum
                Form                = $FormName
                MetaData            = @{}
                LatestVersionID     = ''
                LatestVersionLabel  = ''
                Versions            = @{}
            }

            # And now gather the Document properties
            $propList = $pcdDocObject.GetReturnProperties()

            $nameProp = $propList.GetPropertyValue('Name')
            $docNumProp= $propList.GetPropertyValue($eDOCSTokens.DOCUMENT_NUMBER)

            $currentPropIndex = $propList.BeginIter()
            do
            {
                $propName = $propList.GetCurrentPropertyName()
                # if ($propName -Notlike '%*')
                {
                    $propValue = $propList.GetCurrentPropertyValue()
                    $DMDocument.MetaData.Add($propName, $propValue)
                }

                $currentPropIndex = $propList.NextProperty()
            }
            while ($currentPropIndex -eq 0)


            # And return the DMDocument
            $DMDocument
        }
        finally
        {
            if ($pcdDocObject -ne $null)
            {
                [VOID][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($pcdDocObject)
                $pcdDocObject = $null
            }            
        }
#>
    }
}


<#
    library, form, metadata, verify, paper (paper without document file?)

    Nieuwe profiel registratie: New-DMDocument
    
    New-DMDocumentVersion?

    Update document/profile registratie: Set-DMDocument metadata, documentfile 

    Ophalen document versoe/Profiel registratie: Get-DMDocument -DMDocNum 

    Checkin / checkout:
    Lock-DMDocument
    Unlock-DMDocument


    --------------------------------------------------------------------
    Internal:

    Invoke-PCDSearch   (zie GetDocumentProfileVersions)
      -library
      -SearchObject
      -ReturnProperties
      -OrderByProperties
#>

# the method returned a com variant type that is not interop compatible
# De methode heeft een COM Variant-type geretourneerd dat niet compatibel is met Interop.

<#
# CAK:  
$doc = Get-DMDocument -DocNum 2017061588

# IVHO: 
$doc = Get-DMDocument -DocNum 4051936
$doc = Get-DMDocument -DocNum 4051936 -GetLatestVersionContent


#>