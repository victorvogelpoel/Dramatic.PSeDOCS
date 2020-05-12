# Get-DMForm.Tests.ps1
# Doing Pester tests on Get-DMForm.
# Mocking is used to return form data, so no connection to 
#
# July 2017
# Victor Vogelpoel

Set-StrictMode -Version Latest

Remove-Module 'Dramatic.PSeDOCS' -force -ErrorAction SilentlyContinue
Import-Module (Join-Path $PSScriptRoot '..\Dramatic.PSeDOCS.psd1')
Import-Module Pester      # Find-Module –Name Pester | Install-Module

Describe 'Get-DMForm' {

    # The mocked form definitions to use
    $aFormName        = 'DRAMATIC_DOCUMENT'
    $aFormTitle       = 'Dramatic Document registratie'

    $aSecondFormName  = 'DRAMATIC_DOCUMENT2'
    $aSecondFormTitle = 'Dramatic Document2 registratie'

    $expectedFormFields         = @('DOCNUM', 'MAIL_ID', 'PD_EMAIL_DATE', 'PROCESNAAM', 'PROCESJAAR', 'PARENTMAIL_ID', 'PD_EMAIL_BCC', 'D_DAT_VERZONDEN', 'D_DAT_DOC', 'PD_EMAIL_CC', 'DRAGER_ID', 'X2604', 'ATTACH_NUM', 'THREAD_NUM', 'PD_ORGANIZATION', 'D_KENMERK', 'DRAMATIC_ONSKENMERK', 'TYPE_ID', 'DRAMATIC_SETNUMMER', 'PD_ORIGINATOR', 'RICHTING_ID', 'DRAMATIC_BEWAARTERMIJN', 'DOCTYPE_RETENTION', 'PD_ADDRESSEE', 'DOCTYPE_STORAGE', 'LAST_EDIT_ID', 'DOCTYPE_FULLTEXT', 'DOCNAME', 'RETENTION', 'ABSTRACT', 'DRAMATIC_OBJECT', 'STORAGE', 'PD_OBJ_TYPE', 'TYPIST_ID', 'PD_VREVIEW_DATE', 'RELATIEID', 'LASTEDITDATE', 'ORGANISATIE', 'VESTIGING', 'STATUS', 'POSTADRES', 'HUISNRP', 'POSTCODEP', 'PD_SUSPEND', 'WOONPLAATSP', 'LANDP', 'DRAMATIC_CONTACTPERSOON', 'DRAMATIC_GESLACHT', 'PD_TITLE', 'EMAIL', 'DRAMATIC_BSN', 'LAST_ACCESS_DATE', 'FULL_NAME', 'X1125', 'CREATION_DATE', 'PD_ACTION_NAME', 'TELEFOONB', 'AUTHOR_ID', 'APP_ID', 'PD_LOCATION_CODE')
    $expectedFormRequiredFields = @('PROCESNAAM', 'PROCESJAAR', 'D_DAT_VERZONDEN', 'D_DAT_DOC', 'DRAGER_ID', 'X2604', 'TYPE_ID', 'RICHTING_ID', 'DOCNAME', 'TYPIST_ID', 'RELATIEID', 'DRAMATIC_GESLACHT', 'APP_ID')



    Context "Specific eDOCS form `"$aformName`"" {

        # Mock Assert-DMLibraryConnected not to check loginsession
        Mock Assert-DMLibraryConnected -ModuleName 'Dramatic.PSeDOCS' {
            # No Exception, as if library connected.
        }

        # Mock Invoke-DMSqlCmd to return specific FORM results
        Mock Invoke-DMSqlCmd -ModuleName 'Dramatic.PSeDOCS' { 

            # Stubbed eDOCS form data
            [PSCustomObject][ORDERED]@{
                PSTypeName      = 'PSeDOCS.SQLData'
                FORM_NAME       = 'DRAMATIC_DOCUMENT'
                FORM_TITLE      = 'Dramatic Document registratie'
                FORM_TYPE       = 'P'
                FORM_DEFINITION =  @'
[CMSForm]
Info=626,186,545,557,40000a05
Name=DRAMATIC_DOCUMENT
Prompt=Dramatic Document registratie
ForeColorName=WorkAreaText
BackColor=c0c0c0
BackColorName=WorkArea
HelpID=IDH_The_Profile_Fields
GradientName=WorkArea
Table=DOCSADM.PROFILE
clientWidth=539
clientHeight=532

[CMSBuffer]
Info=0,0,0,0,0
Name=APPLICATION
KeyType=2
Type=4
SQLInfo=APPLICATION

[CMSBuffer]
Info=0,0,0,0,0
Name=AUTHOR
KeyType=2
Type=4
SQLInfo=AUTHOR

[CMSBuffer]
Info=0,0,0,0,0
Name=DOCUMENTTYPE
KeyType=2
Type=4
SQLInfo=DOCUMENTTYPE

[CMSBuffer]
Info=0,0,0,0,0
Name=LAST_EDITED_BY
KeyType=2
Type=4
SQLInfo=LAST_EDITED_BY

[CMSBuffer]
Info=0,0,0,0,0
Name=TYPIST
KeyType=2
Type=4
SQLInfo=TYPIST

[CMSBuffer]
Info=0,0,0,0,0
Name=PD_FILE_PART
KeyType=2
Type=4
SQLInfo=PD_FILE_PART

[CMSBuffer]
Info=0,0,0,0,0
Name=D_RICHTING
KeyType=2
Type=4
SQLInfo=D_RICHTING

[CMSBuffer]
Info=0,0,0,0,0
Name=PD_PTTOBCODE_LINK
KeyType=2
Type=4
SQLInfo=PD_FILE_PART.PD_PTTOBCODE_LINK;DOCSADM.PD_FILE_PART.SYSTEM_ID

[CMSBuffer]
Info=0,0,0,0,0
Name=X4976
KeyType=2
Type=4
SQLInfo=PD_FILE_PART.PD_PTTOBCODE_LINK.PD_WITH_LINK;DOCSADM.PD_BARCODE.SYSTEM_ID

[CMSBuffer]
Info=0,0,0,0,0
Name=D_DRAGER
KeyType=2
Type=4
SQLInfo=D_DRAGER

[CMSBuffer]
Info=0,0,0,0,0
Name=D_BEVEILIGING
KeyType=2
Type=4
SQLInfo=D_BEVEILIGING

[CMSBuffer]
Info=0,0,0,0,0
Name=PD_PRTO_ACTION
KeyType=2
Type=4
SQLInfo=PD_PRTO_ACTION

[CMSBuffer]
Info=0,0,0,0,0
Name=DRAMATIC_PROCES_NAAM
KeyType=2
Type=4
SQLInfo=DRAMATIC_PROCES_NAAM

[CMSBuffer]
Info=0,0,0,0,0
Name=DRAMATIC_PROCES_JAREN
KeyType=2
Type=4
SQLInfo=DRAMATIC_PROCES_JAREN

[CMSBuffer]
Info=0,0,0,0,0
Name=DRAMATIC_RBS_NAW
KeyType=2
Type=4
SQLInfo=DRAMATIC_RBS_NAW

[CMSText]
Info=301,0,69,25,0
Prompt=Static Text
ForeColor=ff
Font=Arial,16,1,0
Border=0
Value=T E S T

[CMSBox]
Info=0,3,535,165,0
Name=DOCUMENT_BOX
Prompt=Document
ForeColorName=WorkAreaText
Font=Arial,11,3,0
Border=15

[CMSEdit]
Info=449,3,80,20,84004
Name=DOCNUM
Prompt=nr.
ForeColorName=WorkAreaText
BackColor=c0c0c0
BackColorName=WorkArea
Font=Arial,11,0,0
Border=0
KeyType=4
Type=4
SQLInfo=DOCNUMBER
PAlignment=2
PromptWidth=17
EditBorder=0
EditFont=Arial,11,0,0
EBackColor=c0c0c0
EForeColorName=WorkAreaText
EBackColorName=WorkArea
DataType=4

[CMSEdit]
Info=875,5,69,21,1000004
Name=MAIL_ID
Prompt=Mail ID
ForeColor=ff0000
Font=Arial,11,1,0
Border=0
SQLInfo=MAIL_ID
PAlignment=2
PromptWidth=43
EditBorder=8
EditFont=Arial,11,0,0

[CMSEdit]
Info=955,5,69,21,1080004
Name=PD_EMAIL_DATE
Prompt=Email Date
ForeColorName=WorkAreaText
BackColorName=Transparent
Font=Arial,11,0,0
Border=0
SQLInfo=PD_EMAIL_DATE
PAlignment=2
PromptWidth=44
EditBorder=8
EditFont=Arial,11,0,0
EForeColorName=WorkAreaText
EBackColorName=Transparent

[CMSEdit]
Info=4,28,255,20,46
Name=PROCESNAAM
Prompt=Procesnaam
Font=Arial,11,0,0
Border=0
KeyType=4
SQLInfo=DRAMATIC_PROCES_NAAM.PROCESNAAM;DOCSADM.DRAMATIC_PROCES_NAAM.SYSTEM_ID
PAlignment=2
PromptWidth=84
EditBorder=8
EditFont=Arial,11,0,0
EBackColor=80ffff
Lookup=DRAMATIC_PROCES_NAAM

[CMSEdit]
Info=265,28,262,21,46
Name=PROCESJAAR
Prompt=Procesjaar
Font=Arial,11,0,0
Border=0
KeyType=4
SQLInfo=DRAMATIC_PROCES_JAREN.PROCESJAAR;DOCSADM.DRAMATIC_PROCES_JAREN.SYSTEM_ID
PAlignment=2
PromptWidth=84
EditBorder=8
EditFont=Arial,11,0,0
EBackColor=80ffff
Lookup=DRAMATIC_PROCES_JAREN

[CMSEdit]
Info=875,30,69,21,1000004
Name=PARENTMAIL_ID
Prompt=Parent Mail ID
ForeColor=ff0000
Font=Arial,11,1,0
Border=0
SQLInfo=PARENTMAIL_ID
PAlignment=2
PromptWidth=43
EditBorder=8
EditFont=Arial,11,0,0

[CMSEdit]
Info=955,30,69,21,1080004
Name=PD_EMAIL_BCC
Prompt=Email Bcc
ForeColorName=WorkAreaText
BackColorName=Transparent
Font=Arial,11,0,0
Border=0
SQLInfo=PD_EMAIL_BCC
PAlignment=2
PromptWidth=44
EditBorder=8
EditFont=Arial,11,0,0
EForeColorName=WorkAreaText
EBackColorName=Transparent

[CMSEdit]
Info=4,48,255,21,44
Name=D_DAT_VERZONDEN
Prompt=Da&t.verz/ontv
Font=Arial,11,0,0
Border=0
Type=1
SQLInfo=D_DAT_VERZONDEN
PAlignment=2
PromptWidth=84
EditBorder=8
EditFont=Arial,11,0,0
EBackColor=80ffff
DataType=1

[CMSEdit]
Info=265,48,262,21,44
Name=D_DAT_DOC
Prompt=Dat&um document
Font=Arial,11,0,0
Border=0
Type=1
SQLInfo=D_DAT_DOC
PAlignment=2
PromptWidth=84
EditBorder=8
EditFont=Arial,11,0,0
EBackColor=80ffff
DataType=1

[CMSEdit]
Info=955,55,69,21,1080004
Name=PD_EMAIL_CC
Prompt=Email CC
ForeColorName=WorkAreaText
BackColorName=Transparent
Font=Arial,11,0,0
Border=0
SQLInfo=PD_EMAIL_CC
PAlignment=2
PromptWidth=44
EditBorder=8
EditFont=Arial,11,0,0
EForeColorName=WorkAreaText
EBackColorName=Transparent

[CMSEdit]
Info=4,69,255,21,46
Name=DRAGER_ID
Prompt=Dra&ger
Font=Arial,11,0,0
Border=0
KeyType=4
SQLInfo=D_DRAGER.DRAGER_ID;DOCSADM.D_DRAGER.SYSTEM_ID
PAlignment=2
PromptWidth=84
EditBorder=8
EditFont=Arial,11,0,0
EBackColor=80ffff
Lookup=D_DRAGER

[CMSEdit]
Info=265,69,262,21,46
Name=X2604
Prompt=Beveiliging
Font=Arial,11,0,0
Border=0
KeyType=4
SQLInfo=D_BEVEILIGING.BEVEILIGING_ID;DOCSADM.D_BEVEILIGING.SYSTEM_ID
PAlignment=2
PromptWidth=84
EditBorder=8
EditFont=Arial,11,0,0
EBackColor=80ffff
Lookup=D_BEVEILIGING

[CMSEdit]
Info=875,70,69,21,1000004
Name=ATTACH_NUM
Prompt=Attachment Number
ForeColor=ff0000
Font=Arial,11,1,0
Border=0
SQLInfo=ATTACH_NUM
PAlignment=2
PromptWidth=43
EditBorder=8
EditFont=Arial,11,0,0

[CMSEdit]
Info=875,85,69,21,1000004
Name=THREAD_NUM
Prompt=Thread Number
ForeColor=ff0000
Font=Arial,11,1,0
Border=0
Type=4
SQLInfo=THREAD_NUM
PAlignment=2
PromptWidth=43
EditBorder=8
EditFont=Arial,11,0,0
DataType=4

[CMSEdit]
Info=955,85,69,21,4
Name=PD_ORGANIZATION
Prompt= Or&ganisatie zender
ForeColor=0
ForeColorName=WorkAreaText
BackColor=c0c0c0
BackColorName=Transparent
Font=Arial,11,0,0
Border=0
SQLInfo=PD_ORGANIZATION
PAlignment=2
PromptWidth=44
MaxChars=100
EditBorder=8
EditFont=Arial,11,0,0
EForeColor=0
EBackColor=ffffff
EForeColorName=InputText
EBackColorName=Input

[CMSEdit]
Info=4,90,255,20,4
Name=D_KENMERK
Prompt=Uw &Kenmerk
Font=Arial,11,0,0
Border=0
SQLInfo=D_KENMERK
PAlignment=2
PromptWidth=84
EditBorder=8
EditFont=Arial,11,0,0

[CMSEdit]
Info=265,90,262,21,80004
Name=DRAMATIC_ONSKENMERK
Prompt=Ons kenmerk
ForeColor=0
BackColor=c0c0c0
Font=Arial,11,0,0
Border=0
SQLInfo=DRAMATIC_ONSKENMERK
PAlignment=2
PromptWidth=84
MaxChars=30
EditBorder=8
EditFont=Arial,11,0,0
EForeColor=0
EBackColor=c0c0c0

[CMSCheckBox]
Info=875,100,69,16,1000000
Name=MSG_ITEM
Prompt=Message Item
ForeColor=ff0000
Font=Arial,11,1,0
Border=0
Value=0
Type=4
SQLInfo=MSG_ITEM

[CMSEdit]
Info=4,110,255,21,4e
Name=TYPE_ID
Prompt=DocumentT&ype
ForeColorName=WorkAreaText
BackColor=c0c0c0
BackColorName=Transparent
Font=Arial,11,0,0
Border=0
KeyType=4
SQLInfo=DOCUMENTTYPE.TYPE_ID;DOCSADM.DOCUMENTTYPES.SYSTEM_ID
PAlignment=2
PromptWidth=84
MaxChars=10
EditBorder=8
EditFont=Arial,11,0,0
EBackColor=80ffff
EForeColorName=InputText
EBackColorName=Input
Lookup=DOCUMENTTYPES

[CMSEdit]
Info=265,110,262,21,4
Name=DRAMATIC_SETNUMMER
Prompt=Setnummer
ForeColor=0
BackColor=c0c0c0
Font=Arial,11,0,0
Border=0
SQLInfo=DRAMATIC_SETNUMMER
PAlignment=2
PromptWidth=84
MaxChars=20
EditBorder=8
EditFont=Arial,11,0,0
EForeColor=0
EBackColor=c0c0c0

[CMSEdit]
Info=955,110,69,21,4
Name=PD_ORIGINATOR
Prompt=&Zender
ForeColor=0
ForeColorName=WorkAreaText
BackColor=c0c0c0
BackColorName=Transparent
Font=Arial,11,0,0
Border=0
SQLInfo=PD_ORIGINATOR
PAlignment=2
PromptWidth=44
MaxChars=50
EditBorder=8
EditFont=Arial,11,0,0
EForeColor=0
EBackColor=ffffff
EForeColorName=InputText
EBackColorName=Input

[CMSEdit]
Info=4,131,255,21,46
Name=RICHTING_ID
Prompt=&Richting
Font=Arial,11,0,0
Border=0
KeyType=4
SQLInfo=D_RICHTING.RICHTING_ID;DOCSADM.D_RICHTING.SYSTEM_ID
PAlignment=2
PromptWidth=84
EditBorder=8
EditFont=Arial,11,0,0
EBackColor=80ffff
Lookup=D_RICHTING

[CMSEdit]
Info=265,131,262,21,4
Name=DRAMATIC_BEWAARTERMIJN
Prompt=Vernietigen
Font=Arial,11,0,0
Border=0
SQLInfo=DRAMATIC_BEWAARTERMIJN
PAlignment=2
PromptWidth=84
MaxChars=5
EditBorder=8
EditFont=Arial,11,0,0
EBackColor=ffffff

[CMSEdit]
Info=820,135,67,21,1080004
Name=DOCTYPE_RETENTION
ForeColor=ff0000
BackColor=c0c0c0
BackColorName=Transparent
Border=0
SQLInfo=DOCUMENTTYPE.RETENTION_DAYS;DOCSADM.DOCUMENTTYPES.SYSTEM_ID
PromptWidth=26
EForeColorName=WorkAreaText
EBackColorName=Transparent

[CMSEdit]
Info=955,135,69,21,4
Name=PD_ADDRESSEE
Prompt=Gead&resseerde
ForeColorName=WorkAreaText
BackColorName=Transparent
Font=Arial,11,0,0
Border=0
SQLInfo=PD_ADDRESSEE
PAlignment=2
PromptWidth=44
MaxChars=100
EditBorder=8
EditFont=Arial,11,0,0
EForeColorName=InputText
EBackColorName=Input

[CMSEdit]
Info=820,160,67,21,1080004
Name=DOCTYPE_STORAGE
ForeColor=ff0000
BackColor=c0c0c0
BackColorName=Transparent
Border=0
SQLInfo=DOCUMENTTYPE.STORAGE_TYPE;DOCSADM.DOCUMENTTYPES.SYSTEM_ID
PromptWidth=26
EForeColorName=WorkAreaText
EBackColorName=Transparent

[CMSEdit]
Info=955,165,69,21,85004
Name=LAST_EDIT_ID
Prompt=-
ForeColor=ff0000
BackColor=c0c0c0
BackColorName=Transparent
Font=Arial,11,0,0
Border=0
SQLInfo=LAST_EDITED_BY.FULL_NAME;DOCSADM.PEOPLE.SYSTEM_ID
PAlignment=2
PLocation=0
PromptWidth=44
MaxChars=20
EditBorder=0
EditFont=Arial,11,0,0
EForeColorName=WorkAreaText
EBackColorName=Transparent

[CMSBox]
Info=0,173,535,100,0
Prompt=Documentinhoud
Font=Arial,11,3,0
Border=15

[CMSEdit]
Info=820,180,67,21,1080004
Name=DOCTYPE_FULLTEXT
ForeColor=ff0000
BackColor=c0c0c0
BackColorName=Transparent
Border=0
SQLInfo=DOCUMENTTYPE.FULL_TEXT;DOCSADM.DOCUMENTTYPES.SYSTEM_ID
PromptWidth=26
EForeColorName=WorkAreaText
EBackColorName=Transparent

[CMSEdit]
Info=4,188,525,20,544
Name=DOCNAME
Prompt=Info
ForeColorName=WorkAreaText
BackColor=c0c0c0
BackColorName=Transparent
Font=Arial,11,0,0
Border=0
SQLInfo=DOCNAME
PAlignment=2
PromptWidth=84
MaxChars=240
EditBorder=8
EditFont=Arial,11,0,0
EForeColor=0
EBackColor=80ffff
EForeColorName=InputText
EBackColorName=Input

[CMSCheckBox]
Info=945,195,100,16,1000000
CheckedTrigger=Y
UncheckedTrigger=N
Name=PD_SUPSEDES
Prompt=Supersedes
ForeColor=ff0000
Font=Arial,11,1,0
Border=0
Value=N
SQLInfo=PD_SUPSEDES

[CMSEdit]
Info=820,205,67,21,5000004
Name=RETENTION
Prompt=R&etention Days
ForeColor=ff0000
BackColor=c0c0c0
BackColorName=Transparent
Font=Arial,11,0,0
Border=0
Type=4
SQLInfo=RETENTION
PAlignment=2
PromptWidth=26
EditBorder=8
EditFont=Arial,11,0,0
EForeColorName=InputText
EBackColorName=Input
DataType=4
Validation=Range 0 999

[CMSEdit]
Info=4,208,525,40,584
Name=ABSTRACT
Prompt=Omschr&ijving
ForeColorName=WorkAreaText
BackColor=c0c0c0
BackColorName=Transparent
Font=Arial,11,0,0
Border=0
SQLInfo=ABSTRACT
PAlignment=2
PromptWidth=84
MaxChars=254
EditBorder=8
EditFont=Arial,11,0,0
EForeColorName=InputText
EBackColorName=Input

[CMSCheckBox]
Info=945,215,100,16,1000000
CheckedTrigger=Y
UncheckedTrigger=N
Name=PD_SUPSEDED
Prompt=Superceded
ForeColor=ff0000
Font=Arial,11,1,0
Border=0
Value=N
SQLInfo=PD_SUPSEDED

[CMSCheckBox]
Info=820,230,120,16,0
CheckedTrigger=Y
UncheckedTrigger=N
Name=FULLTEXT
Prompt=Zoeken op &inhoud
ForeColor=0
Font=Arial,11,0,0
Border=0
Value=N
SQLInfo=FULLTEXT

[CMSCheckBox]
Info=945,240,100,16,1000000
Name=DELIVER_REC
Prompt=Deliver Rec
ForeColor=ff0000
Font=Arial,11,1,0
Border=0
Value=0
Type=4
SQLInfo=DELIVER_REC

[CMSEdit]
Info=4,248,525,20,4
Name=DRAMATIC_OBJECT
Prompt=Object
Font=Arial,11,0,0
Border=0
SQLInfo=DRAMATIC_OBJECT
PAlignment=2
PromptWidth=84
MaxChars=254
EditBorder=8
EditFont=Arial,11,0,0

[CMSEdit]
Info=820,250,78,21,1184004
Name=STORAGE
Prompt=T&ype
ForeColor=ff0000
BackColor=c0c0c0
BackColorName=Transparent
Font=Arial,11,0,0
Border=0
SQLInfo=STORAGETYPE
PAlignment=2
PromptWidth=32
MaxChars=1
EditBorder=8
EditFont=Arial,11,0,0
EForeColorName=WorkAreaText
EBackColorName=Transparent
Format=[=A]Archive;[=D]Delete;[=K]Keep;[=O]Optical;[=T]Template
LInfo=0,0,112,74,40000080
LFont=Arial,11,0,0
LBorder=16
Column0=,108,0,0,0,[=A]Archive;[=D]Delete;[=K]Keep;[=O]Optical;[=T]Template,
Row0=A
Row1=D
Row2=K
Row3=O
Row4=T

[CMSEdit]
Info=945,260,100,21,1184004
Name=PD_OBJ_TYPE
Prompt=Item Type
ForeColor=ff0000
BackColorName=Transparent
Font=Arial,11,1,0
Border=0
SQLInfo=PD_OBJ_TYPE
PAlignment=2
PromptWidth=25
MaxChars=2
EditBorder=8
EditFont=Arial,11,0,0
EForeColorName=WorkAreaText
EBackColorName=Transparent
Format=[=0]Document;[=1]Record;[=4]File;[=5]Box
LInfo=0,0,187,150,40000080
LFont=Arial,11,0,0
LBorder=16
Column0=,183,0,0,0,[=0]Document;[=1]Record;[=4]File;[=5]Box,
Row0=0
Row1=1
Row2=4
Row3=5

[CMSEdit]
Info=820,275,107,21,100004e
Name=TYPIST_ID
Prompt=&Entered By
ForeColor=ff0000
BackColor=c0c0c0
BackColorName=Transparent
Font=Arial,11,1,0
Border=0
KeyType=4
SQLInfo=TYPIST.USER_ID;DOCSADM.PEOPLE.SYSTEM_ID
PAlignment=2
PromptWidth=65
MaxChars=27
EditBorder=8
EditFont=Arial,11,0,0
EForeColorName=InputText
EBackColorName=Input
Lookup=PEOPLE

[CMSBox]
Info=0,278,535,145,0
Prompt=Relatiegegevens
Font=Arial,11,3,0
Border=15

[CMSEdit]
Info=945,290,100,21,80004
Name=PD_VREVIEW_DATE
Prompt=Herzieningsdatum
ForeColorName=WorkAreaText
BackColorName=Transparent
Font=Arial,11,0,0
Border=0
Type=1
SQLInfo=PD_VREVIEW_DATE
PAlignment=2
PromptWidth=25
EditBorder=8
EditFont=Arial,11,0,0
EBackColor=c0c0c0
EForeColorName=WorkAreaText
EBackColorName=Transparent
DataType=1

[CMSEdit]
Info=4,293,523,20,1046
Name=RELATIEID
Prompt=Relatie ID
Font=Arial,11,0,0
Border=0
KeyType=4
SQLInfo=DRAMATIC_RBS_NAW.RELATIEID;DOCSADM.DRAMATIC_RBS_NAW.SYSTEM_ID
PAlignment=2
PromptWidth=84
EditBorder=8
EditFont=Arial,11,0,0
EBackColor=80ffff
Lookup=DRAMATIC_RBS_NAW

[CMSEdit]
Info=825,300,70,21,1084004
Name=LASTEDITDATE
Prompt=Edited:
ForeColor=ff0000
BackColor=c0c0c0
BackColorName=Transparent
Font=Arial,11,1,0
Border=0
Type=1
SQLInfo=LAST_EDIT_DATE
PAlignment=2
PromptWidth=37
EditBorder=0
EditFont=Arial,11,0,0
EForeColorName=WorkAreaText
EBackColorName=Transparent
DataType=1

[CMSEdit]
Info=4,313,280,20,81004
Name=ORGANISATIE
Prompt=Organisatie
Font=Arial,11,0,0
Border=0
SQLInfo=DRAMATIC_RBS_NAW.ORGANISATIE;DOCSADM.DRAMATIC_RBS_NAW.SYSTEM_ID
PAlignment=2
PromptWidth=84
MaxChars=255
EditBorder=8
EditFont=Arial,11,0,0

[CMSEdit]
Info=289,313,240,20,80004
Name=VESTIGING
Prompt=aanvullend
Font=Arial,11,0,0
Border=0
SQLInfo=DRAMATIC_RBS_NAW.VESTIGING;DOCSADM.DRAMATIC_RBS_NAW.SYSTEM_ID
PAlignment=2
PromptWidth=58
MaxChars=200
EditBorder=8
EditFont=Arial,11,0,0

[CMSEdit]
Info=945,320,100,21,84004
Name=STATUS
Prompt=Status
ForeColorName=WorkAreaText
BackColor=c0c0c0
BackColorName=Transparent
Font=Arial,11,0,0
Border=0
Type=4
SQLInfo=STATUS
PAlignment=2
PromptWidth=25
EditBorder=8
EditFont=Arial,11,0,0
EBackColor=c0c0c0
EForeColorName=WorkAreaText
EBackColorName=Transparent
DataType=4
Format=[=0]Available;[=1]Document Being Edited;[=2]Profile Being Edited;[=3]Checked out;[=4]Not Available;[=5]Being Indexed;[=6]Archived;[=16]Being Archived;[=18]Deleted;[=19]Read Only;[>6]Not Available

[CMSEdit]
Info=4,333,280,20,81004
Name=POSTADRES
Prompt=Straat
Font=Arial,11,0,0
Border=0
SQLInfo=DRAMATIC_RBS_NAW.POSTADRES;DOCSADM.DRAMATIC_RBS_NAW.SYSTEM_ID
PAlignment=2
PromptWidth=84
MaxChars=255
EditBorder=8
EditFont=Arial,11,0,0

[CMSEdit]
Info=289,333,119,20,80004
Name=HUISNRP
Prompt=nr
Font=Arial,11,0,0
Border=0
SQLInfo=DRAMATIC_RBS_NAW.HUISNRP;DOCSADM.DRAMATIC_RBS_NAW.SYSTEM_ID
PAlignment=2
PromptWidth=58
MaxChars=50
EditBorder=8
EditFont=Arial,11,0,0

[CMSEdit]
Info=409,333,120,20,80004
Name=POSTCODEP
Prompt=Postcode
Font=Arial,11,0,0
Border=0
SQLInfo=DRAMATIC_RBS_NAW.POSTCODEP;DOCSADM.DRAMATIC_RBS_NAW.SYSTEM_ID
PAlignment=2
PromptWidth=55
MaxChars=50
EditBorder=8
EditFont=Arial,11,0,0

[CMSEdit]
Info=945,350,100,21,80004
Name=PD_SUSPEND
Prompt=Uitgesteld
ForeColorName=WorkAreaText
BackColorName=Transparent
Font=Arial,11,0,0
Border=0
SQLInfo=PD_FILE_PART.PD_SUSPEND;DOCSADM.PD_FILE_PART.SYSTEM_ID
PAlignment=2
PromptWidth=25
EditBorder=8
EditFont=Arial,11,0,0
EBackColor=c0c0c0
EForeColorName=WorkAreaText
EBackColorName=Transparent

[CMSEdit]
Info=4,353,280,20,80004
Name=WOONPLAATSP
Prompt=Woonplaats
Font=Arial,11,0,0
Border=0
SQLInfo=DRAMATIC_RBS_NAW.WOONPLAATSP;DOCSADM.DRAMATIC_RBS_NAW.SYSTEM_ID
PAlignment=2
PromptWidth=84
MaxChars=100
EditBorder=8
EditFont=Arial,11,0,0

[CMSEdit]
Info=289,353,240,20,80004
Name=LANDP
Prompt=Land
Font=Arial,11,0,0
Border=0
SQLInfo=DRAMATIC_RBS_NAW.LANDP;DOCSADM.DRAMATIC_RBS_NAW.SYSTEM_ID
PAlignment=2
PromptWidth=58
MaxChars=40
EditBorder=8
EditFont=Arial,11,0,0

[CMSEdit]
Info=4,373,280,20,4
Name=DRAMATIC_CONTACTPERSOON
Prompt=Contactpersoon
Font=Arial,11,0,0
Border=0
SQLInfo=DRAMATIC_CONTACTPERSOON
PAlignment=2
PromptWidth=84
MaxChars=80
EditBorder=8
EditFont=Arial,11,0,0

[CMSEdit]
Info=289,373,240,20,184044
Name=DRAMATIC_GESLACHT
Prompt=Geslacht
Font=Arial,11,0,0
Border=0
SQLInfo=DRAMATIC_GESLACHT
PAlignment=2
PromptWidth=58
MaxChars=1
EditBorder=8
EditFont=Arial,11,0,0
EBackColor=80ffff
Format=[=M]Man;[=V]Vrouw;[=O]Onbekend
LInfo=0,0,187,150,40000080
LFont=Arial,11,0,0
LBorder=16
Column0=,183,0,0,0,[=M]Man;[=V]Vrouw;[=O]Onbekend,
Row0=M
Row1=V
Row2=O

[CMSEdit]
Info=945,390,100,21,81404
Name=PD_TITLE
Prompt=Categorie
ForeColorName=WorkAreaText
BackColorName=Transparent
Font=Arial,11,0,0
Border=0
SQLInfo=PD_FILE_PART.PD_TITLE;DOCSADM.PD_FILE_PART.SYSTEM_ID
PAlignment=2
PromptWidth=25
MaxChars=255
EditBorder=8
EditFont=Arial,11,0,0
EBackColor=c0c0c0
EForeColorName=WorkAreaText
EBackColorName=Transparent

[CMSEdit]
Info=4,393,280,20,81004
Name=EMAIL
Prompt=E-mail
Font=Arial,11,0,0
Border=0
SQLInfo=DRAMATIC_RBS_NAW.EMAIL;DOCSADM.DRAMATIC_RBS_NAW.SYSTEM_ID
PAlignment=2
PromptWidth=84
EditBorder=8
EditFont=Arial,11,0,0

[CMSEdit]
Info=289,393,240,20,4
Name=DRAMATIC_BSN
Prompt=BSN
Font=Arial,11,0,0
Border=0
SQLInfo=DRAMATIC_BSN
PAlignment=2
PromptWidth=58
MaxChars=9
EditBorder=8
EditFont=Arial,11,0,0

[CMSBox]
Info=0,423,80,70,0
Prompt=Toegang
ForeColorName=WorkAreaText
Font=Arial,11,3,0
Border=15

[CMSBox]
Info=84,423,285,70,0
Prompt=Historie
ForeColor=0
ForeColorName=WorkAreaText
Font=Arial,11,3,0
Border=15

[CMSEdit]
Info=945,425,100,21,80004
Name=LAST_ACCESS_DATE
Prompt=Bewerkt op
ForeColorName=WorkAreaText
BackColorName=Transparent
Font=Arial,11,0,0
Border=0
Type=1
SQLInfo=LAST_ACCESS_DATE
PAlignment=2
PromptWidth=25
EditBorder=8
EditFont=Arial,11,0,0
EBackColor=c0c0c0
EForeColorName=WorkAreaText
EBackColorName=Transparent
DataType=1

[CMSCheckBox]
Info=396,435,127,16,1000001
UncheckedTrigger=
Name=DRAMATIC_COPY_DATA
Prompt=Kopieer metadata set
Font=Arial,11,0,0
Border=0
SQLInfo=DRAMATIC_COPY_DATA

[CMSCheckBox]
Info=4,443,70,15,0
Name=SECURITY
Prompt=Beveilig
Font=Arial,11,0,0
Border=0
Value=0
Type=4
SQLInfo=DEFAULT_RIGHTS

[CMSEdit]
Info=89,443,275,20,80004
Name=FULL_NAME
Prompt=Geregistreerd door
ForeColorName=WorkAreaText
BackColor=c0c0c0
BackColorName=Transparent
Font=Arial,11,0,0
Border=0
SQLInfo=TYPIST.FULL_NAME;DOCSADM.PEOPLE.SYSTEM_ID
PAlignment=2
PromptWidth=101
MaxChars=60
EditBorder=8
EditFont=Arial,11,0,0
EBackColor=c0c0c0
EForeColorName=WorkAreaText
EBackColorName=Transparent

[CMSEdit]
Info=945,455,100,21,80004
Name=X1125
Prompt=Uitgeleend aan
ForeColorName=WorkAreaText
BackColorName=Transparent
Font=Arial,11,0,0
Border=0
KeyType=4
SQLInfo=PD_FILE_PART.PD_PTTOBCODE_LINK.PD_WITH_LINK.USER_ID;DOCSADM.PEOPLE.SYSTEM_ID
PAlignment=2
PromptWidth=25
EditBorder=8
EditFont=Arial,11,0,0
EBackColor=c0c0c0
EForeColorName=WorkAreaText
EBackColorName=Transparent

[CMSPush]
Info=9,463,65,23,4000001
Name=TRUSTEES
Prompt=Bewerk
BackColor=c0c0c0
Font=Arial,11,0,0
Border=0
PictureLoc=1

[CMSEdit]
Info=89,463,275,20,80004
Name=CREATION_DATE
Prompt=Geregistreerd op
ForeColorName=WorkAreaText
BackColor=c0c0c0
BackColorName=Transparent
Font=Arial,11,0,0
Border=0
Type=1
SQLInfo=CREATION_DATE
PAlignment=2
PromptWidth=101
EditBorder=8
EditFont=Arial,11,0,0
EBackColor=c0c0c0
EForeColorName=WorkAreaText
EBackColorName=Transparent
DataType=1

[CMSPush]
Info=374,468,80,26,4000001
Name=OK
Prompt=&OK

[CMSPush]
Info=455,468,80,26,4000001
Name=CANCEL
Prompt=A&fbreken

[CMSPush]
Info=455,468,80,26,4000001
Name=CLOSE
Prompt=&Sluiten

[CMSCheckBox]
Info=945,480,100,16,0
CheckedTrigger=Y
UncheckedTrigger=N
Name=PD_VITAL
Prompt=&Belangrijk stuk
ForeColor=0
Font=Arial,11,0,0
Border=0
Value=N
SQLInfo=PD_VITAL

[CMSEdit]
Info=947,510,145,20,6
Name=PD_ACTION_NAME
Prompt=Verwijderacties
Font=Arial,11,1,0
Border=0
KeyType=4
SQLInfo=PD_PRTO_ACTION.PD_ACTION_NAME;DOCSADM.PD_ACTION.SYSTEM_ID
PAlignment=2
PromptWidth=93
EditBorder=8
EditFont=Arial,11,0,0
Lookup=PD_ACTION

[CMSEdit]
Info=185,570,215,21,1080004
Name=TELEFOONB
Prompt=Telefoonnr
Font=Arial,11,0,0
Border=0
SQLInfo=DRAMATIC_RBS_NAW.TELEFOONB;DOCSADM.DRAMATIC_RBS_NAW.SYSTEM_ID
PAlignment=2
PromptWidth=84
MaxChars=100
EditBorder=8
EditFont=Arial,11,0,0

[CMSEdit]
Info=285,605,215,21,100000e
Name=AUTHOR_ID
Prompt=&Auteur
ForeColorName=WorkAreaText
BackColor=c0c0c0
BackColorName=Transparent
Font=Arial,11,0,0
Border=0
KeyType=4
SQLInfo=AUTHOR.USER_ID;DOCSADM.PEOPLE.SYSTEM_ID
PAlignment=2
PromptWidth=84
MaxChars=20
EditBorder=8
EditFont=Arial,11,0,0
EBackColor=c0c0c0
EForeColorName=InputText
EBackColorName=Input
Lookup=PEOPLE

[CMSEdit]
Info=105,505,215,21,4e
Name=APP_ID
Prompt=Toepassing
ForeColor=0
ForeColorName=WorkAreaText
BackColor=c0c0c0
BackColorName=Transparent
Font=Arial,11,0,0
Border=0
KeyType=4
SQLInfo=APPLICATION.APPLICATION;DOCSADM.APPS.SYSTEM_ID
PAlignment=2
PromptWidth=84
EditBorder=8
EditFont=Arial,11,0,0
EForeColor=0
EBackColor=80ffff
EForeColorName=InputText
EBackColorName=Input
Lookup=_APPS

[CMSEdit]
Info=945,370,100,20,6
Name=PD_LOCATION_CODE
Prompt=Naam lokatie
Font=Arial,11,0,0
Border=0
KeyType=4
SQLInfo=PD_FILE_PART.PD_PT2LOC_LINK.PD_LOCATION_CODE;DOCSADM.PD_LOCATION.SYSTEM_ID
PromptWidth=32
EditBorder=8
EditFont=Arial,11,0,0
Lookup=PD_LOCATION

[CMSBuffer]
Info=0,0,0,0,0
Name=PD_PT2LOC_LINK
KeyType=2
Type=4
SQLInfo=PD_FILE_PART.PD_PT2LOC_LINK;DOCSADM.PD_FILE_PART.SYSTEM_ID
'@ -split '\r\n'             
            }
        }

        # Let Get-Form to its work
        $form = Get-DMForm -formname $aformName

        # Now assert
        It "Returns name `"$aFormName`"" {
            $form.FormName| Should BeExactLy $aFormName
        }

        It "Returns description `"$aFormTitle`"" {
            $form.Description | Should BeExactLy $aFormTitle
        }

        It 'Returns type "P"' {
            $form.FormType | Should BeExactLy 'P'
        }

        It 'Returns the correct fields' {
            @(Compare-Object ($form.Fields | Sort) ($expectedFormFields | Sort)).length | Should Be 0
        }

        It 'Returns the correct required fields' {
            @(Compare-Object ($form.RequiredFields | Sort) ($expectedFormRequiredFields | Sort)).length | Should Be 0
        }
    }


    Context "All eDOCS forms" {

        # Mock Assert-DMLibraryConnected not to check loginsession
        Mock Assert-DMLibraryConnected -ModuleName 'Dramatic.PSeDOCS' {
            # No Exception, as if library connected.
        }

        # Mock Invoke-DMSqlCmd to return specific forms
        Mock Invoke-DMSqlCmd -ModuleName 'Dramatic.PSeDOCS' { 

            # Stubbed eDOCS form data
            @(
                [PSCustomObject][ORDERED]@{
                    PSTypeName      = 'PSeDOCS.SQLData'
                    FORM_NAME       = 'DRAMATIC_DOCUMENT'
                    FORM_TITLE      = 'Dramatic Document registratie'
                    FORM_TYPE       = 'P'
                    FORM_DEFINITION = (@'
[CMSForm]
Info=626,186,545,557,40000a05
Name=DRAMATIC_DOCUMENT
Prompt=Dramatic Document registratie
ForeColorName=WorkAreaText
BackColor=c0c0c0
BackColorName=WorkArea
HelpID=IDH_The_Profile_Fields
GradientName=WorkArea
Table=DOCSADM.PROFILE
clientWidth=539
clientHeight=532

[CMSEdit]
Info=449,3,80,20,84004
Name=DOCNUM
Prompt=nr.
ForeColorName=WorkAreaText
BackColor=c0c0c0
BackColorName=WorkArea
Font=Arial,11,0,0
Border=0
KeyType=4
Type=4
SQLInfo=DOCNUMBER
PAlignment=2
PromptWidth=17
EditBorder=0
EditFont=Arial,11,0,0
EBackColor=c0c0c0
EForeColorName=WorkAreaText
EBackColorName=WorkArea
DataType=4
'@ -split '\r\n')
                },

                [PSCustomObject][ORDERED]@{
                    PSTypeName      = 'PSeDOCS.SQLData'
                    FORM_NAME       = 'DRAMATIC_DOCUMENT2'
                    FORM_TITLE      = 'Dramatic Document2 registratie'
                    FORM_TYPE       = 'P'
                    FORM_DEFINITION = (@'
[CMSForm]
Info=626,186,545,557,40000a05
Name=DRAMATIC_DOCUMENT2
Prompt=Dramatic Document2 registratie
ForeColorName=WorkAreaText
BackColor=c0c0c0
BackColorName=WorkArea
HelpID=IDH_The_Profile_Fields
GradientName=WorkArea
Table=DOCSADM.PROFILE
clientWidth=539
clientHeight=532

[CMSEdit]
Info=449,3,80,20,84004
Name=DOCNUM
Prompt=nr.
ForeColorName=WorkAreaText
BackColor=c0c0c0
BackColorName=WorkArea
Font=Arial,11,0,0
Border=0
KeyType=4
Type=4
SQLInfo=DOCNUMBER
PAlignment=2
PromptWidth=17
EditBorder=0
EditFont=Arial,11,0,0
EBackColor=c0c0c0
EForeColorName=WorkAreaText
EBackColorName=WorkArea
DataType=4
'@ -split '\r\n') 
                }
            )

        }

        # Let Get-Form to its work
        $forms = Get-DMForm

        It "Returns 2 forms" {
            @($forms).Count | Should BeExactly 2
        }

        It "Returns name `"$aFormName`" for the first form" {
            $forms[0].FormName | Should BeExactLy $aFormName
        }

        It "Returns name `"$aSecondFormName`" for the second form" {
            $forms[1].FormName | Should BeExactLy $aSecondFormName
        }
    }
}