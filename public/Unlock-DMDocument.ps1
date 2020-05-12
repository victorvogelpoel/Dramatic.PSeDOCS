# Unlock-DMDocument.ps1
# Unlocks the document profile
# Aug 2017
# Copyright 2017 Dramatic Development
# If this works, it was written by Victor Vogelpoel (victor@victorvogelpoel.nl).
# If it doesn't, I don't know who wrote it.


function Unlock-DMDocument
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

        $unlockParam = @{
            ($eDOCSTokens.STATUS) = $eDOCSTokens.UNLOCK
        }

        Set-DMDocument -DocNum $DocNum -FormName $FormName -ProfileProperties $unlockParam

        # NOTE: Above command may return an error if the documentprofile is already unlocked:
        # ERROR 0x800401D1 in PCDDocObjectClass: WIN-G8GEKHHS9OE: WIN-G8GEKHHS9OE: Document is unlocked.
    }
}


<#
   Unlock-DMDocument -DocNum 4051933

#>