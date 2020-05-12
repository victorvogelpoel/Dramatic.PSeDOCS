# Lock-DMDocument.ps1
# Locks the document profile
# Aug 2017
# Copyright 2017 Dramatic Development
# If this works, it was written by Victor Vogelpoel (victor@victorvogelpoel.nl).
# If it doesn't, I don't know who wrote it.


function Lock-DMDocument
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true, position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage='TODO')]
        [Alias('DocNumber', 'DocumentNumber')] 
        [long]$DocNum,

        [Parameter(Mandatory=$false, position=1, ValueFromPipelineByPropertyName=$true, HelpMessage='TODO')]
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

        $lockParam = @{
            ($eDOCSTokens.STATUS) = $eDOCSTokens.LOCK
        }

        Set-DMDocument -DocNum $DocNum -FormName $FormName -ProfileProperties $unlockParam
    }
}