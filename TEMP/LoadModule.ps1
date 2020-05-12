


# Stop at each error
$global:ErrorActionPreference = 'STOP'


Remove-Module Dramatic.PSeDOCS -Force -errorAction SilentlyContinue -Verbose
Import-Module (Join-Path $PSScriptRoot '..\Dramatic.PSeDOCS.psd1') -DisableNameChecking -Verbose




<#

Remove-Module Dramatic.PSeDOCS -Force -errorAction SilentlyContinue -Verbose
Import-Module (Join-Path $PSScriptRoot '..\Dramatic.PSeDOCS.psd1') -DisableNameChecking -Verbose

$credential = New-Object System.Management.Automation.PSCredential 'VVO', (ConvertTo-SecureString -AsPlainText 'MyVoiceIsMyPassport' -force) 
$login      = Connect-DMLibrary -Library EDOCSLIB -Credential $credential
$loginSession = $login


Invoke-DMSQLCmd "SELECT FORM_NAME, FORM_TITLE from DOCSADM.FORMS"


#>