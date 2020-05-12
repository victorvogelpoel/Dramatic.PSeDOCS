# Get-DMForm.ps1
# 
# Jul 2017
# Copyright 2017 Dramatic Development
# If this works, it was written by Victor Vogelpoel (victor@victorvogelpoel.nl).
# If it doesn't, I don't know who wrote it.
#
# Get-DMForm returns all form names and description, but not the definition
# Get-DMForm -FormName 'N_DOCUMENT' returns form name, descripton and raw, required and normal fields



function Get-DMForm
{
    [CmdLetBinding()]
    param
    (
        # FormName is a dynamic optional parameter (below) which ValidateSet is populated from the Form names in the eDOCS database.
    )

    dynamicparam 
    {
        $paramDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary

        $ValidateSetArgs = @{}
        if (Test-DMLibraryConnected)
        {
            # If a library is connected, we can get the formnames from the eDOCS database.
            $allforms        = (Invoke-DMSQLCmd -SQLCommand "SELECT FORM_NAME, FORM_TITLE FROM DOCSADM.FORMS ORDER BY FORM_NAME" | select -ExpandProperty FORM_NAME)

            if ($null -ne $allforms -and @($allforms).Count -gt 0)
            {
                # Construct "ValidateSet" parameter that will be splatted later
                $ValidateSetArgs.Add('ValidateSet', $allforms)
            }
        }

        # Now create the "FormName" parameter 
        $formNameparam = New-DynamicParameter -ParameterName 'FormName' -ParameterType 'String[]' -ValueFromPipeline -ValueFromPipelineByPropertyName -ValidateNotNullOrEmpty @ValidateSetArgs -Position 0 -HelpMessage 'TODO'
        $paramDictionary.Add('FormName', $formNameparam)
        return $paramDictionary
    }
 
    begin
    {
        Assert-DMLibraryConnected
    }

    process
    {
        $formNameSpecifed = ($PSBoundParameters.ContainsKey('FormName'))
        $library          = Get-DMCurrentLibrary
   
        # Construct the SELECT SQL Command 
        $formSQLCommand = "SELECT FORM_NAME, FORM_TITLE FROM DOCSADM.FORMS ORDER BY FORM_NAME"
 
        if ($formNameSpecifed)
        {
            $FormName = $PSBoundParameters.FormName

            # Construct the filter to something like
            #   select FORM_NAME, FORM_TITLE, FORM_TYPE from DOCSADM.FORMS where FORM_NAME = 'OCW_DOCUMENT'
            if ($FormName -and @($FormName).Count -gt 0)
            {
                $formSQLCommand = "SELECT FORM_NAME, FORM_TITLE, FORM_TYPE, FORM_DEFINITION FROM DOCSADM.FORMS"

                $filter = @()

                foreach ($name in $FormName)
                {
                    # TODO: toepassen vermijden SQL injectie via formnaam
                    $filter += "FORM_NAME = '$($name.ToUpperInvariant())'" 
                }
 
                $formSQLCommand += ' WHERE ' + ($filter -join ' or ') + ' ORDER BY FORM_NAME'
            }
        }
 
        # Now issue the SQL command.
        $results = Invoke-DMSqlCmd -SQLCommand $formSQLCommand
 
        foreach ($result in $results)            
        {
            if (!$formNameSpecifed)
            {
                # Return an simple Form object
                [PSCustomObject]@{
                    PSTypeName     = 'Dramatic.eDOCS.Form'
                    FormName       = $result.FORM_NAME
                    Description    = $result.FORM_TITLE
                    Library        = $library
                }

            }
            else
            {
                $formdefinition = [PSCustomObject]@{
                    PSTypeName     = 'Dramatic.eDOCS.Form'
                    FormName       = $result.FORM_NAME
                    Description    = $result.FORM_TITLE
                    FormType       = $result.FORM_TYPE
                    Library        = $library
                    Fields         = @()
                    RequiredFields = @()
                    FieldProperties= @{}
                    FieldsRaw      = @{}
                }
 
                # ---------------------------------------------------------------------------------------------
                # Parse the Form definition INI data for fields and required fields
                $iniFormData       = @{}
                $sectionArrayIndex = 0
 
                switch -regex ($result.FORM_DEFINITION -split '\n')
                {
                    # Parse the INI line by line

                    # recognize a SECTION
                    '^\[(.+)\]\s*$' {
 
                        $iniSection = $matches[1]
 
                        # Add the section name to the iniFormData hashtable if not there already.
                        # Initialize with an empty array.
                        if (!$iniFormData.ContainsKey($iniSection))
                        {
                            $iniFormData.Add($iniSection, @())
                        }
 
                        # Add hashtable to section array
                        $iniFormData[$iniSection] += @{}
                        $sectionArrayIndex        = $iniFormData[$iniSection].Count - 1
                    }
 
 
                    # recognize a PROPERTY
                    '^\s*([^#]+?)\s*=\s*(.*)' {

                        $name,$value = $matches[1..2]

                        # Add the name/value to the last section array item
                        (($iniFormData[$iniSection])[$sectionArrayIndex])[$name] = $value.trim()
                    }
                }
 
                # ---------------------------------------------------------------------------------------------
                # Now transform the Form definition INI data to useful objects
                if ($iniFormData.ContainsKey('CMSEdit'))
                {
                    # Build the 'Fields' and 'RequiredFields' collections
                    $iniFormData.CMSEdit | foreach {

                        $rawField = $_
 
                        # Store the CMSEDit under its Name in the FieldsRaw hashtable
                        $formdefinition.FieldsRaw[$rawField.Name] = $rawField

                        # Copy specific fields from the $rawField to the FieldProperties
                        $formdefinition.FieldProperties[$rawField.Name] = @{}
                        $rawField.GetEnumerator() | where { $_.Name -in ('Name', 'Prompt', 'MaxChars', 'KeyType', 'Type', 'DataType', 'Lookup') } | foreach {
                            # TODO: translate DataType, KeyType and Type to meaningful values
                            $formdefinition.FieldProperties[$rawField.Name][$_.Name] = $_.Value
                        }

                        # Now if the field is a REQUIRED field, where bit 6 of the 5th number of the Info field is set, 
                        $formdefinition.FieldProperties[$rawField.Name]['Required'] = (([Convert]::ToInt32(($rawField.Info -split ',')[4], 16) -band 0x40) -ne 0)
                    }

                    # Sort the data
                    $formdefinition.Fields         = $formdefinition.FieldProperties.GetEnumerator() | Select -ExpandProperty Name | Sort
                    $formdefinition.RequiredFields = $formdefinition.FieldProperties.GetEnumerator() | Where { $_.Value.Required } | Select -ExpandProperty Name | Sort
                }
                             
                # Output the form data object
                $formdefinition
            }
        }
    }

}

<#

    Get-DMForm

    Get-DMForm -FormName 'DEF_PROF'

#>