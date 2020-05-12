# Set-DMDocument.ps1
# Update the document profile
# Aug 2017
# Copyright 2017 Dramatic Development
# If this works, it was written by Victor Vogelpoel (victor@victorvogelpoel.nl).
# If it doesn't, I don't know who wrote it.


function Set-DMDocument
{
    [CmdLetBinding(DefaultParameterSetName = 'DynamicDocumentProperties')]
    param
    (
        [Parameter(Mandatory=$true, position=0, ParameterSetName='DynamicDocumentProperties', ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage='TODO')]
        [Parameter(Mandatory=$true, position=0, ParameterSetName='DocumentProperties', ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage='TODO')]
        [Alias('DocNumber', 'DocumentNumber')] 
        [long]$DocNum,

        [Parameter(Mandatory=$true, position=1, ParameterSetName='DynamicDocumentProperties', ValueFromPipelineByPropertyName=$true, HelpMessage='TODO')]
        [Parameter(Mandatory=$true, position=1, ParameterSetName='DocumentProperties', ValueFromPipelineByPropertyName=$true, HelpMessage='TODO')]
        [ValidateNotNullOrEmpty()]
        [String]$FormName,


        [Parameter(Mandatory=$true, position=1, ParameterSetName='DocumentProperties', ValueFromPipelineByPropertyName=$true, HelpMessage='TODO')]
        [Alias('Properties', 'DocumentProperties')]
        [Hashtable]$ProfileProperties


        # Content / File / NewVersion/NewSubversion
    )

   DynamicParam {
        
        # Create dynparameters for each field in de form; GREAT for intellisense!
        $paramDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary

        if ((Test-DMLibraryConnected))
        {
            $form = Get-DMForm -FormName $FormName
            
            if ($null -ne $form)
            {
                # Create dynparameters for each field in de form; GREAT for intellisense! This will be our propertyArguments!
                foreach ($field in ($form.Fields | where { $_ -notin 'DocNum', 'FormName', 'ProfileProperties' }))
                {
                    $param = New-DynamicParameter -ParameterName $field -ParameterType String -ValueFromPipelineByPropertyName -ParameterSetName 'DynamicDocumentProperties'
                    $paramDictionary.Add($field, $param)
                }
            }
        }

        return $paramDictionary
    }

    begin
    {
        Assert-DMLibraryConnected
    }

    process
    {
        Write-Verbose "Using Form `"$FormName`" for DocNum $DocNum."

        $loginSession = Get-DMLoginSession -Library (Get-DMCurrentLibrary)
        $library      = $loginSession.LoginLibrary
        $DST          = $loginSession.DST

        if ($PSCmdlet.ParameterSetName -eq 'DynamicDocumentProperties')
        {
            # Construct properties from the dynamic parameters.
            $propertyArguments = [ORDERED]@{}
            foreach ($param in $PSBoundParameters.GetEnumerator())
            {
                if ($param.Key -notin 'FormName', 'DocNum', 'Verbose', 'Debug', 'ErrorAction', 'ErrorVariable', 'WarningAction', 'WarningVariable', 'OutBuffer', 'PipelineVariable', 'OutVariable') 
                {
                    $propertyArguments[$param.Key] = $param.Value
                }
            }
        }
        else
        {
            $propertyArguments = $ProfileProperties
        }


        try
        {
            $pcdDocObject = $null   # initialize variables for the final() in case the New-Object fails

            $pcdDocObject = New-Object Hummingbird.DM.Server.Interop.PCDClient.PCDDocObjectClass

            [void]$pcdDocObject.SetDST($DST)
            [void]$pcdDocObject.SetObjectType($FormName)
            [void]$pcdDocObject.SetProperty($eDOCSTokens.TARGET_LIBRARY, $library)
            [void]$pcdDocObject.SetProperty($eDOCSTokens.OBJECT_IDENTIFIER, $DocNum.ToString())

            foreach ($propertyArg in $propertyArguments.GetEnumerator())
            {
                [void]$pcdDocObject.SetProperty($propertyArg.Key, $propertyArg.Value)
            }
            
            # Now execute the update to the profile
            [void]$pcdDocObject.Update()
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