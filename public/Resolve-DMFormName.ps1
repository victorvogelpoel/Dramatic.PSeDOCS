# Resolve-DMFormName.ps1
# 
# Jul 2017
# Copyright 2017 Dramatic Development
# If this works, it was written by Victor Vogelpoel (victor@victorvogelpoel.nl).
# If it doesn't, I don't know who wrote it.



function Resolve-DMFormName
{
    [CmdLetBinding(DefaultParameterSetname='DocNum')]
    param
    (
        [Parameter(Mandatory=$true, position=0, ParameterSetName='DocNum', ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage='TODO')]
        [Alias('DocNumber', 'DocumentNumber')] 
        [long[]]$DocNum,

        [Parameter(Mandatory=$true, position=0, ParameterSetName='FormID', ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage='TODO')]
        [Alias('SYSTEM_ID')]
        [int[]]$FormID
    )
 
    begin
    {
        Assert-DMLibraryConnected
    }

    process
    {
        if ($PSCmdlet.ParameterSetName -eq 'DocNum')
        {
            foreach ($oneDocNum in $DocNum)
            {
                $resolveFormSQLCommand = "SELECT FORM_NAME FROM DOCSADM.FORMS INNER JOIN DOCSADM.PROFILE ON DOCSADM.FORMS.SYSTEM_ID=DOCSADM.PROFILE.FORM WHERE DOCSADM.PROFILE.DOCNUMBER='$oneDocNum'"
 
                # Now issue the SQL command
                $results = Invoke-DMSqlCmd -SQLCommand $resolveFormSQLCommand

                if ($null -ne $results -and @($results).Count -gt 0)
                {
                    # And output the form name
                    $results.FORM_NAME
                }
                # Otherwise return nothing
            }
        }
        else
        {
            foreach ($aFormID in $FormID)
            {
                $resolveFormSQLCommand = "SELECT FORM_NAME FROM DOCSADM.FORMS WHERE DOCSADM.FORMS.SYSTEM_ID='$aFormID'"
 
                # Now issue the SQL command
                $results = Invoke-DMSqlCmd -SQLCommand $resolveFormSQLCommand

                if ($null -ne $results -and @($results).Count -gt 0)
                {
                    # And output the form name
                    $results.FORM_NAME
                }
                # Otherwise return nothing
            }
        }
    }
}


<# IVHO

Resolve-DMFormName -DocNum 4051936


#>
