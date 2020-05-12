# Assert-DMOperationSuccess.ps1
# Assert that the DM API operation was successful.
# Jul 2017
# Copyright 2017 Dramatic Development
# If this works, it was written by Victor Vogelpoel (victor@victorvogelpoel.nl).
# If it doesn't, I don't know who wrote it.

function Assert-DMOperationSuccess
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage='TODO')]
        [Alias('DMObject')]
        [Hummingbird.DM.Server.Interop.PCDClient.IPCDError]$PCDErrorCapableObject,

        [Parameter(Mandatory=$false, Position=1, ValueFromPipelineByPropertyName=$true, HelpMessage='TODO')]
        [ValidateNotNullOrEmpty()]
        [string]$ExceptionMessage = $null
    )

    process
    {
        if ($PCDErrorCapableObject.ErrNumber -ne 0)
        {
            if ([String]::IsNullOrEmpty($ExceptionMessage))
            {
                # eg: 'ERROR 0xA in PCDGetLoginLibsClass: some error description'
                $ExceptionMessage = 'ERROR 0x{0:X} in {2}: {1}' -f $PCDErrorCapableObject.ErrNumber, $PCDErrorCapableObject.ErrDescription, $PCDErrorCapableObject.GetType().Name
            }

            throw $ExceptionMessage
        }
    }
}
