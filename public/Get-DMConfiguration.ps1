# Get-DMConfiguration.ps1
# Gets some configuration information
# Aug 2017
# Copyright 2017 Dramatic Development
# If this works, it was written by Victor Vogelpoel (victor@victorvogelpoel.nl).
# If it doesn't, I don't know who wrote it.


$script:DMConfiguration = @{
    MaxComWriteSize = 1048576

}





function Get-DMConfiguration
{
    [CmdLetBinding()]
    param ()

    $script:DMConfiguration
}