# Search-DMDocument.ps1
# 
# Jul 2017
# Copyright 2017 Dramatic Development
# If this works, it was written by Victor Vogelpoel (victor@victorvogelpoel.nl).
# If it doesn't, I don't know who wrote it.



function Search-DMDocument
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
<#
        [Parameter(position=2, mandatory=$true, ValueFromPipelineByPropertyName=$true, HelpMessage='TODO')]
        [Hashtable]$SearchArgument,

        [Parameter(position=3, mandatory=$true, ValueFromPipeLineByPropertyName=$true, helpMessage='TODO')]
        [ValidateNotNullOrEmpty()]
        [String[]]$ReturnProperty,
#>
        [Parameter(position=4, mandatory=$false, ValueFromPipeLineByPropertyName=$true, helpMessage='TODO')]
        [int]$MaxRows = 50
    )

   DynamicParam {
        
        # Create dynparameters for each field in de form; GREAT for intellisense!
        $paramDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        # Initialize the validateset for non-connected scenario
        $ValidateSetProps = @{ }

        if (Test-DMLibraryConnected)
        {
            $form = Get-DMForm -FormName $FormName

            if ($null -ne $form)
            {
                # Construct a ValidateSet for ReturnProperty
                $ValidateSetProps = [ORDERED]@{ ValidateSet = $form.Fields }

                # Create dynparameters for each field in de form; GREAT for intellisense! This will be our SearchArguments!
                foreach ($field in $form.Fields)
                {
                    $param = New-DynamicParameter -ParameterName $field -ParameterType String -ValueFromPipelineByPropertyName 
                    $paramDictionary.Add($field, $param)
                }
            }
        }

        # Add the "ReturnProperty" DynParameter (for both the connected and non-connected scenarios, with or without a validateset with fieldnames)
        $ReturnPropertyParam = New-DynamicParameter -ParameterName 'ReturnProperty' -ParameterType 'String[]' -Mandatory -ValueFromPipelineByPropertyName -ValidateNotNullOrEmpty @ValidateSetProps -HelpMessage 'TODO'
        $paramDictionary.Add('ReturnProperty', $ReturnPropertyParam)

        return $paramDictionary
    }

    begin
    {
        Assert-DMLibraryConnected
    }

    process
    {
        try
        {
            $pcdSearch = $null     # initialize variables for the final() in case the New-Object fails

            # Init a proper variable for the Dynamic Parameter 'ReturnProperty'
            $ReturnProperty = $PSBoundParameters.ReturnProperty

            # Construct searcharguments from the dynamic other parameters.
            $SearchArgument = [ORDERED]@{}
            foreach ($param in $PSBoundParameters.GetEnumerator())
            {
                if ($param.Key -notin 'FormName', 'ReturnProperty', 'MaxRows', 'Verbose', 'Debug', 'ErrorAction', 'ErrorVariable', 'WarningAction', 'WarningVariable', 'OutBuffer', 'PipelineVariable', 'OutVariable') 
                {
                    $SearchArgument[$param.Key] = $param.Value
                }
            }

            # Validate
            if ($SearchArgument.Count -eq 0) { throw "You need to specify 1 or more search arguments." }
            if ($ReturnProperty.Count -eq 0) { throw "You need to specify 1 or more return properties." }

            # ----------------------------------------------------------------------------------------------------------
            # Prepare
            # Add 'DOCNUM' and 'FORM' to the return properties if it is not there
            # Note: 'FORM' is the FormID, not FormName
            foreach ($column in ($eDOCSColumns.DOCNUM, $eDOCSColumns.FORM))
            {
                $ReturnProperty = @($column) + $ReturnProperty
            }

            $loginSession = Get-DMLoginSession -Library (Get-DMCurrentLibrary)
            $library      = $loginSession.LoginLibrary
            $DST          = $loginSession.DST

            $pcdSearch = New-Object Hummingbird.DM.Server.Interop.PCDClient.PCDSearchClass

            # Initialize the basics
            [void]$pcdSearch.SetDST($loginSession.DST)
            [void]$pcdSearch.AddSearchLib($loginSession.LoginLibrary)
            [void]$pcdSearch.SetSearchObject($FormName)
            
            [void]$pcdSearch.SetMaxRows($MaxRows)

            # Add search criteria
            foreach ($aSearchArgument in $SearchArgument.GetEnumerator())
            {
                if ($null -ne $aSearchArgument.Value)
                {
                    [void]$pcdSearch.AddSearchCriteria($aSearchArgument.Name, $aSearchArgument.Value.ToString())
                }
            }

            # Add return properties
            foreach ($aReturnProperty in $ReturnProperty) 
            {
                [void]$pcdSearch.AddReturnProperty($aReturnProperty)
            }

            [void]$pcdSearch.AddOrderByProperty($eDOCSColumns.DOCNUM, 0)

            # ----------------------------------------------------------------------------------------------------------
            # Do the work and get results
            [void]$pcdSearch.Execute()
            $pcdSearch | Assert-DMOperationSuccess

            $formsCache = @{}

            # For each row create a DocumentSearchResult object 
            for ($row=1; $row -le $pcdSearch.GetRowsFound(); $row++)
            {
                [void]$pcdSearch.SetRow($row)

                $docPropList = [ORDERED]@{
                    PSTypeName = 'Dramatic.eDOCS.DocumentSearchResult'
                    DocNum     = 0
                    FormName   = $FormName
                }

                foreach ($aReturnProperty in $ReturnProperty)
                {
                    $propValue = $pcdSearch.GetPropertyValue($aReturnProperty)
                    if ($null -eq $propValue -or $propValue -eq 'UNKNOWN_PROPERTY')
                    {
                        $propValue = ''
                    }
                    
                    if ($aReturnProperty -eq $eDOCSColumns.FORM)
                    {
                        # Add the NAME for the FORM ID to the returned properties.
                        # Resolve the ID to NAME first
                        if (!$formsCache.ContainsKey($propValue))
                        {
                            # Query for the formname using the form ID and add it to the cache
                            $formsCache[$propValue] = (Resolve-DMFormName -FormID $propValue)
                        }

                        $docPropList.FormName = $formsCache[$propValue]
                    }
                    else
                    {
                        # Add the property to the proplist
                        $docPropList[$aReturnProperty] = $propValue
                    }
                }
                
                # Return the found properties as a PowerShell custom object
                [PSCustomObject]$docPropList
            }
        }
        finally
        {
            if ($pcdSearch -ne $null)
            {
                try
                {
                    [void]$pcdSearch.ReleaseResults()
                }
                finally
                {
                    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($pcdSearch)
                }
            }         
        }
    }
}


# TESTING
<#

$searchArguments = @{
    #TYPIST_ID = 'ATSKERK'
    DOCUMENTTYPE = 2554
}

Search-DMDocument -FormName "IVHO_DEF_PROF" -SearchArgument $searchArguments -ReturnProperty 'APP_ID', 'TYPIST_ID' | ft -a



Search-DMDocument -FormName "IVHO_DEF_PROF" -DOCUMENTTYPE 2554 -ReturnProperty 'APP_ID', 'TYPIST_ID' | ft -a

Search-DMDocument -FormName "IVHO_DEF_PROF" -TYPIST_ID 'ATSKERK' -returnproperty 'ABSTRACT', 'APP_ID', 'ATTACH_NUM', 'AUTHOR_ID', 'BRON', 'CREATION_DATE', 'DOC_STATUS', 'DOCNAME', 'DOCNUM', 'DOCTYPE_FULLTEXT', 'DOCTYPE_RETENTION', 'DOCTYPE_STORAGE', 'EMAIL_RECEIVED', 'EMAIL_SENT', 'FULL_NAME', 'FULLTEXT', 'IVHO_ADRES', 'IVHO_ADRES_NR', 'IVHO_ADRES_TOEV', 'IVHO_AFZENDER', 'IVHO_CONT_PERS', 'IVHO_DAT_ANTW', 'IVHO_DAT_ARCH', 'IVHO_DAT_DOC', 'IVHO_DAT_ONV', 'IVHO_DAT_VERS', 'IVHO_ELEMENT', 'IVHO_EMAIL', 'IVHO_KENMERK', 'IVHO_MAV_STATUS', 'IVHO_MAV_VASTDAT', 'IVHO_MAV_VASTSTL', 'IVHO_NAW_STATUS', 'IVHO_ONDERZOEK', 'IVHO_PL_DAT_AFD', 'IVHO_POSTCODE', 'IVHO_REL_APP', 'IVHO_TOEG_CUS', 'IVHO_TREFWOORD', 'IVHO_WOONPL', 'LAND_CODE', 'LAST_ACCESS_DATE', 'LAST_EDIT_ID', 'LASTEDITDATE', 'LEVERANCIER', 'MAIL_ID', 'ORG_AFD', 'PARENTMAIL_ID', 'PD_ADDRESSEE', 'PD_EMAIL_BCC', 'PD_EMAIL_CC', 'PD_EMAIL_DATE', 'PD_FILE_DATE', 'PD_FILE_NAME', 'PD_FILEPT_NO', 'PD_LOCATION_CODE', 'PD_OBJ_TYPE', 'PD_ORGANIZATION', 'PD_ORIGINATOR', 'PD_PART_NAME', 'PD_SUSPEND', 'PD_TITLE', 'PD_VREVIEW_DATE', 'PROJ_CODE', 'PROJ_NAAM', 'READONLY_DATE', 'RECHTENWAARDE', 'REGIO', 'RELEASE_NR', 'RETENTION', 'RICHTING_ID', 'STATUS', 'STORAGE', 'THREAD_NUM', 'TOEG_NIVEAU', 'TYPE_ID', 'TYPIST_ID', 'X1125'

#>
