

function New-DynamicParameter
{
    # .SYNOPSIS
    #   Create a new dynamic parameter object for use with a dynamicparam block.
    # .DESCRIPTION
    #   New-DynamicParameter allows simplified creation of runtime (dynamic) parameters.
    # .PARAMETER ParameterName
    #   The name of the parameter to create.
    # .PARAMETER ParameterType
    #   The .NET type of this parameter. 
    # .PARAMETER Mandatory
    #   Set the mandatory flag for this parameter.
    # .PARAMETER Position
    #   Define a position for the parameter.
    # .PARAMETER ValueFromPipeline
    #   The parameter can be filled from the input pipeline.
    # .PARAMETER ValueFromPipelineByPropertyName
    #   The parameter can be filled from a specific property in the input pipeline.
    # .PARAMETER ParameterSetName
    #   Assign the parameter to a specific parameter set.
    # .PARAMETER ValidateNotNullOrEmpty
    #   Disallow null or empty values for the parameter if the parameter is specified.
    # .PARAMETER ValidatePattern
    #   Test the parameter value using a regular expression.
    # .PARAMETER ValidatePatternOptions
    #   Regular expression options which dictate the behaviour of ValidatePattern.
    # .PARAMETER ValidateRange
    #   A minimum and maximum value to compare the argument to.
    # .PARAMETER ValidateScript
    #   Test the parameter value using a script.
    # .PARAMETER ValidateSet
    #   Test the parameter value against a set of values.
    # .PARAMETER ValidateSetIgnoreCase
    #   ValidateSet can be configured to be case sensitive by setting this parameter to $false. The default behaviour for ValidateSet ignores case.
    # .INPUTS
    #   System.Object
    #   System.Object[]
    #   System.String
    #   System.Type
    # .OUTPUTS
    #   System.Management.Automation.RuntimeDefinedParameter
    # .EXAMPLE
    #   New-DynamicParameter Name -DefaultValue "Test" -ParameterType "String" -Mandatory -ValidateSet "Test", "Live"
    # .EXAMPLE
    #   New-DynamicParameter Name -ValueFromPipelineByPropertyName
    # .EXAMPLE
    #   New-DynamicParameter Name -ValidateRange 1, 2
    # .NOTES
    #   Author: Chris Dent
    #
    #   Change log:
    #     24/10/2014 - Chris Dent - Added support for ValidatePattern options and ValidateSet case sensitivity.
    #     22/10/2014 - Chris Dent - First release.

    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory = $true)] 
        [ValidateNotNullOrEmpty()]
        [Alias('Name')]
        [String]$ParameterName,
    
        [Object]$DefaultValue,
    
        [Type]$ParameterType = "Object",

        [Switch]$Mandatory,

        [Int32]$Position = -2147483648,
    
        [Switch]$ValueFromPipeline,
    
        [Switch]$ValueFromPipelineByPropertyName,
    
        [String]$ParameterSetName = "__AllParameterSets",

        [String]$HelpMessage = 'TODO',
    
        [Switch]$ValidateNotNullOrEmpty,

        [ValidateNotNullOrEmpty()]
        [RegEx]$ValidatePattern,
    
        [Text.RegularExpressions.RegexOptions]$ValidatePatternOptions = [Text.RegularExpressions.RegexOptions]::IgnoreCase,

        [Object[]]$ValidateRange,
    
        [ValidateNotNullOrEmpty()]
        [ScriptBlock]$ValidateScript,
   
        [ValidateNotNullOrEmpty()]
        [Object[]]$ValidateSet,

        [Boolean]$ValidateSetIgnoreCase = $true
    )
  
    $attributeCollection = @()

    $parameterAttribute  = New-Object Management.Automation.ParameterAttribute
    $parameterAttribute.Mandatory                       = $Mandatory
    $parameterAttribute.Position                        = $Position
    $parameterAttribute.ValueFromPipeline               = $ValueFromPipeline
    $parameterAttribute.ValueFromPipelineByPropertyName = $ValueFromPipelineByPropertyName
    $parameterAttribute.HelpMessage                     = $HelpMessage

    $attributeCollection += $parameterAttribute

    if ($PSBoundParameters.ContainsKey('ValidateNotNullOrEmpty'))
    {
        $attributeCollection += New-Object Management.Automation.ValidateNotNullOrEmptyAttribute
    }

    if ($PSBoundParameters.ContainsKey('ValidatePattern')) 
    {
        $ValidatePatternAttribute = New-Object Management.Automation.ValidatePatternAttribute($ValidatePattern.ToString())
        $ValidatePatternAttribute.Options = $ValidatePatternOptions

        $attributeCollection += $ValidatePatternAttribute
    }

    if ($PSBoundParameters.ContainsKey('ValidateRange')) 
    {
        $attributeCollection += New-Object Management.Automation.ValidateRangeAttribute($ValidateRange)
    }

    if ($PSBoundParameters.ContainsKey('ValidateScript'))
    {
        $attributeCollection += New-Object Management.Automation.ValidateScriptAttribute($ValidateScript)
    }

    if ($PSBoundParameters.ContainsKey('ValidateSet'))
    {
        $ValidateSetAttribute = New-Object Management.Automation.ValidateSetAttribute($ValidateSet)
        $ValidateSetAttribute.IgnoreCase = $ValidateSetIgnoreCase

        $attributeCollection += $ValidateSetAttribute
    }

    $Parameter = New-Object Management.Automation.RuntimeDefinedParameter($ParameterName, $ParameterType, $attributeCollection)
    if ($PSBoundParameters.ContainsKey('DefaultValue'))
    {
        $Parameter.Value = $DefaultValue
    }

    return $Parameter
}