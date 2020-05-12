
# Examples

Import-Module Dramatic.PSeDOCS.psd1

$credential   = Get-Credential      # or New-Credential -UserName 'User Name' -Password 'My Voice is my Passport'
$login        = Connect-DMLibrary -Library EDOCS_LIB -Credential $credential

$formArrau    = Get-DMForm
$FormDef      = Get-DMForm -FormName 'aFormName'
$FormDefArray = Get-DMForm -FormName 'aFormName', 'anotherForm'

