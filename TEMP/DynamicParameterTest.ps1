

$LibraryForms = 'IVHO_DEF_PROF', 'QBE_PROF'

function Search-DMDocumentD
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

        #[Parameter(position=2, mandatory=$true, ValueFromPipeLineByPropertyName=$true, HelpMessage='TODO')]
        #[string[]]$SearchArgument,

        #[Parameter(position=3, mandatory=$true, ValueFromPipeLineByPropertyName=$true, helpMessage='TODO')]
        #[ValidateNotNullOrEmpty()]
        #[String[]]$ReturnProperty,

        [Parameter(position=4, mandatory=$false, ValueFromPipeLineByPropertyName=$true, helpMessage='TODO')]
        [int]$MaxRows = 50
    )

    DynamicParam {
        
        # Create dynparameters for each field in de form; GREAT for intellisense!
        $paramDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary

        $form            = Get-DMForm -FormName $FormName

        # Create a ReturnPropery parameter with validate set, using fields from the form!
        $ReturnPropertyParam = New-DynamicParameter -ParameterName 'ReturnProperty' -ParameterType 'String[]' -Mandatory -ValueFromPipelineByPropertyName -ValidateNotNullOrEmpty -ValidateSet $form.Fields -HelpMessage 'TODO'
        $paramDictionary.Add('ReturnProperty', $ReturnPropertyParam)

        # Create dynparameters for each field in de form; GREAT for intellisense! This will be our SearchArguments!
        foreach ($field in $form.Fields)
        {
            $param = New-DynamicParameter -ParameterName $field -ParameterType String -ValueFromPipelineByPropertyName 
            $paramDictionary.Add($field, $param)
        }

        return $paramDictionary
    }

    process
    {
        # Init a proper variable for the Dynamic Parameter 'ReturnProperty'
        $ReturnProperty = $PSBoundParameters.ReturnProperty

        $SearchArguments = [ORDERED]@{}
        foreach ($param in $PSBoundParameters.GetEnumerator())
        {
            Write-Host "$($param.Key) = $($param.Value)"

            if ($param.Key -notin 'FormName', 'ReturnProperty', 'MaxRows', 'Verbose', 'Debug', 'ErrorAction', 'ErrorVariable', 'WarningAction', 'WarningVariable', 'OutBuffer', 'PipelineVariable', 'OutVariable') 
            {
                $SearchArguments[$param.Key] = $param.Value
            }
        }

        Write-Host "Formname: $FormName"
        Write-Host "ReturnProps: $($ReturnProperty -join ', ')"
        Write-Host "MaxRows: $maxrows"

        Write-Host "SearchArgument:"
        $SearchArguments

        

        

        #$PSBoundParameters
        #Write-Host "PSBoundParameters $($PSBoundParameters.GetEnumerator() | foreach { "$($_.Name) = $($_.Value)" } )"

    }

}



Search-DMDocumentD -FormName IVHO_DEF_PROF -ReturnProperty APP_ID, DOCNAME -DOCNAME 'victor' -AUTHOR_ID ATSKERK




