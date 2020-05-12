# Get-DMLoginSession.ps1
# Get all loginsessions or just for the specified library
# Aug 2017
# Copyright 2017 Dramatic Development
# If this works, it was written by Victor Vogelpoel (victor@victorvogelpoel.nl).
# If it doesn't, I don't know who wrote it.


$script:libraryLoginSessions = @{}
<#
    @{
        <DST> = [PSCustomObject][ORDERED]@{
                    PSTypeName   = 'Dramatic.eDOCS.Login'
                    UserName     = ''
                    DST          = ''
                    LoginLibrary = ''
                    PrimaryGroup = ''
                }
#>


function Get-DMLoginSession
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$false, position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage='TODO' )]
        [String]$Library
    )

    process
    {
        # A $Library value of $null or empty string will act as if the Library parameter was not specified and return ALL loginsessions

        if ($PSBoundParameters.ContainsKey('Library') -and ![String]::IsNullOrEmpty($Library))
        {
            if ($script:libraryLoginSessions.ContainsKey($Library))
            {
                # return the loginSession information for the specified library
                $script:libraryLoginSessions[$Library]
            }
            # else return null
        }
        else
        {
            # Return all the login sessions
            
            $script:libraryLoginSessions.GetEnumerator() | foreach {$_.Value}  # $libraryLoginSessions is a module scope variable (not exported)
        }
    }
}




# NOTE: Function Set-DMLoginSession will not be exported outside the module and is a PRIVATE function.
function Set-DMLoginSession
{
    [CmdLetBinding()]
    param 
    (
        [Parameter(Mandatory=$true, position=0, ValueFromPipelineByPropertyName=$true, HelpMessage='TODO' )]
        [ValidateNotNullOrEmpty()]
        [string]$UserName,

        [Parameter(Mandatory=$true, position=1, ValueFromPipelineByPropertyName=$true, HelpMessage='TODO' )]
        [ValidateNotNullOrEmpty()]
        [string]$DST, 

        [Parameter(Mandatory=$true, position=2, ValueFromPipelineByPropertyName=$true, HelpMessage='TODO' )]
        [ValidateNotNullOrEmpty()]
        [string]$LoginLibrary, 
        
        [Parameter(Mandatory=$true, position=3, ValueFromPipelineByPropertyName=$true, HelpMessage='TODO' )]
        [ValidateNotNullOrEmpty()]
        [string]$PrimaryGroup
    )

    process
    {
        $loginSession = [PSCustomObject][ORDERED]@{
                            PSTypeName   = 'Dramatic.eDOCS.Login'
                            UserName     = $UserName 
                            DST          = $DST
                            LoginLibrary = $LoginLibrary
                            PrimaryGroup = $PrimaryGroup
                        }

        # Cache the login session information
        $libraryLoginSessions[$LoginLibrary] = $loginSession

        # And return the custom object
        $loginSession
    }
 }