



function Search-DMDocumentR
{
    [CmdletBinding()]
    param
    (
        #[Parameter(position=0, mandatory=$true, ValueFromPipeLineByPropertyName=$true, helpMessage='TODO')]
        #[ValidateNotNullOrEmpty()]
        #[String]$Library,

        [Parameter(position=1, mandatory=$true, ValueFromPipelineByPropertyName=$true, HelpMessage='TODO')]
        [ValidateNotNullOrEmpty()]
        [string]$FormName,

        [Parameter(position=2, mandatory=$true, ValueFromRemainingArguments=$true, HelpMessage='TODO')]
        [string[]]$SearchArgument,

        [Parameter(position=3, mandatory=$true, ValueFromPipeLineByPropertyName=$true, helpMessage='TODO')]
        [ValidateNotNullOrEmpty()]
        [String[]]$ReturnProperty,

        [Parameter(position=4, mandatory=$false, ValueFromPipeLineByPropertyName=$true, helpMessage='TODO')]
        [int]$MaxRows = 50
    )

    process
    {
        Write-Host "Formname: $FormName"
        Write-Host "ReturnProps: $($ReturnProperty -join ', ')"
        Write-Host "MaxRows: $maxrows"

        Write-Host "Remaining: $($SearchArgument -Join ', ')"

        $PSBoundParameters
        #Write-Host "PSBoundParameters $($PSBoundParameters.GetEnumerator() | foreach { "$($_.Name) = $($_.Value)" } )"

    }

}



Search-DMDocumentR -FormName IVHO_DEF_PROF -search2 SearchProp2 -ReturnProperty "ReturnProp1", "ReturnProp2" -search1 SearchProp1 

