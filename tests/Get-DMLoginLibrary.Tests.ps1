# Get-DMLoginLibrary.Tests.ps1
# Doing Pester tests on Get-DMForm.
# Mocking is used to return form data, so no connection to the library
#
# July 2017
# Victor Vogelpoel


# GRRR te ingewikkeld om werkend te krijgen... Mocks gaan niet lekker.



Set-StrictMode -Version Latest

Remove-Module 'Dramatic.PSeDOCS' -force -ErrorAction SilentlyContinue
Import-Module (Join-Path $PSScriptRoot '..\Dramatic.PSeDOCS.psd1')
Import-Module Pester      # Find-Module –Name Pester | Install-Module    # Import-Module "\\rayban\data\MMS\SCRIPTS\Pester\Pester.psd1"



    Describe 'Get-DMLoginLibrary' {

        #Mock Add-DMPCDClientType { }  # Refraining it from loading the PCDClient.DLL.
        

        Context 'Loginlibraries' {
        
            Mock New-PCDGetLoginLibsClass -ModuleName 'Dramatic.PSeDOCS' {
                
                $librariesStub = 'EDOCSLIB1', 'EDOCSLIB2'

                $PCDGetLogin = [PSCustomObject]@{
                    stubLibraries  = $librariesStub
                    ErrNumber      = 0
                    ErrDescription = ''
                }

                $PCDGetLogin | Add-Member -Name Execute -MemberType ScriptMethod -Value {
                    return 0
                }

                $PCDGetLogin | Add-Member -Name GetSize -MemberType ScriptMethod -Value {
                    return @($librariesStub).Count
                }

                $PCDGetLogin | Add-Member -Name GetAt -MemberType ScriptMethod -Value { 
                    param($Index)
                    
                    return $librariesStub[$Index]
                }

                $PCDGetLogin
            }


            Mock Assert-DMOperationSuccess -ModuleName 'Dramatic.PSeDOCS' {
                param ( $PCDErrorCapableObject, $ExceptionMessage )
                #param ( [Parameter(ValueFromPipeline=$true, position=0)][PSCustomObject]$PCDErrorCapableObject, [string]$ExceptionMessage )
                # do nothing
            }

            Mock Clear-COMObject -ModuleName 'Dramatic.PSeDOCS' { param($Object) 
                # do nothing
            }

            #----------------------------------------------------------------------------------
            # Act
            $libraries = Get-DMLoginLibrary


            #----------------------------------------------------------------------------------
            # Assert
            It "Returns eDOCS login libraries"
            {
                $libraries | ShouldBe $null

            }
        }


        Context "DMOperation fails" {

            Mock Assert-DMOperationSuccess  -ModuleName 'Dramatic.PSeDOCS' {
                throw "DM Operation failed"
            } 


            # DO the work
            #$libraries = Get-DMLoginLibraries

            # Assert
           
        }

    }
