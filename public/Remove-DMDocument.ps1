# Remove-DMDocument.ps1
# Gets the document registration data for the specified DocNum
# Aug 2017
# Copyright 2017 Dramatic Development
# If this works, it was written by Victor Vogelpoel (victor@victorvogelpoel.nl).
# If it doesn't, I don't know who wrote it.


function Remove-DMDocument
{
    [CmdLetBinding(DefaultParameterSetName='DocumentProfile')]
    param
    (
        [Parameter(Mandatory=$true, position=0, ParameterSetName='DocumentProfile', ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage='TODO')]
        [Alias('DocNumber', 'DocumentNumber')] 
        [long]$DocNum,

        [Parameter(Mandatory=$false, position=2, ParameterSetName='DocumentProfile', ValueFromPipelineByPropertyName=$true, HelpMessage='TODO')]
        [ValidateNotNullOrEmpty()]
        [String]$FormName
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
            $pcdDocObject = $null   # initialize variables for the final() in case the New-Object fails

            $pcdDocObject = New-Object Hummingbird.DM.Server.Interop.PCDClient.PCDDocObjectClass

            [void]$pcdDocObject.SetDST($loginSession.DST)
            [void]$pcdDocObject.SetObjectType($FormName)
            [void]$pcdDocObject.SetProperty($eDOCSTokens.TARGET_LIBRARY,  $loginSession.LoginLibrary)
            [void]$pcdDocObject.SetProperty($eDOCSTokens.DOCUMENT_NUMBER, $DocNum.ToString())                          # ISE crashes when $DocNum argument is not a string

            # Get the profile information
            [void]$pcdDocObject.Fetch()
            $pcdDocObject | Assert-DMOperationSuccess
    
            # And delete the profile
            [void]$pcdDocObject.Delete()
            $pcdDocObject | Assert-DMOperationSuccess
        }
        finally
        {
            if ($null -ne $pcdDocObject)
            {
                [VOID][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($pcdDocObject)
            }            
        }
    }
}


<#



#>

