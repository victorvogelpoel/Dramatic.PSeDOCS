# Invoke-DMSqlCmd.ps1
# Execute a SQL command through the DM interface
# Jul 2017
# Copyright 2017 Dramatic Development
# If this works, it was written by Victor Vogelpoel (victor@victorvogelpoel.nl).
# If it doesn't, I don't know who wrote it.



function Invoke-DMSqlCmd
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage='TODO')]
        [ValidateNotNullOrEmpty()]
        [String]$SQLCommand
    )

    begin
    {
        Assert-DMLibraryConnected
    }

    process
    {
        try
        {
            $pcdSQL = $null # this is necessary for the final() when the PCDSQLClass fails to initialize

            $pcdSQL = New-Object Hummingbird.DM.Server.Interop.PCDClient.PCDSQLClass

            $library      = Get-DMCurrentLibrary
            $loginSession = Get-DMLoginSession -Library $library
            [void]$pcdSQL.SetLibrary($loginSession.LoginLibrary)
            [void]$pcdSQL.SetDST($loginSession.DST)
            [void]$pcdSQL.Execute($SQLCommand)
            $pcdSQL | Assert-DMOperationSuccess -ExceptionMessage ('ERROR 0x{0:X} while executing query: {1}Native SQL error: {2}' -f $pcdSQL.ErrNumber, $pcdSQL.ErrDescription, $pcdSQL.GetSQLErrorCode())
            
            if ($pcdSQL.GetColumnCount() -gt 0 -and $pcdSQL.GetRowCount() -gt 0)
            {
                # Get the column names from the resultset
                $columNames = @()
                for ($col=1; $col -le $pcdSQL.GetColumnCount(); $col++)
                {
                    $columNames += $pcdSQL.GetColumnName($col)
                }

                for ($row=1; $row -le $pcdSQL.GetRowCount(); $row++)
                {
                    $data = [PSCustomObject][ORDERED]@{
                        PSTYPENAME = "PSeDOCS.SQLData"
                    }

                    [void]$pcdSQL.SetRow($row)
                    # For each column, add member to the PS customobject $data, with columnname and columnvalue
                    for($col=1; $col -le $columNames.Count; $col++)
                    {
                        $data | Add-Member -Name $columNames[$col-1] -Value ($pcdSQL.GetColumnValue($col)) -MemberType NoteProperty
                    }
                
                    # Return the data
                    $data
                }
            }
        }
        finally
        {
            # Release the PCDSQLClass if it was instantiated.
            if ($pcdSQL -ne $null)
            {
                try
                {
                    # Release the results first
                    [void]$pcdSQL.ReleaseResults()
                }
                finally
                {
                    # And then release the COM object
                    [VOID][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($pcdSQL)
                    $pcdSQL = $null
                }
            }
        }
    }
}


<#
# TESTING

$credential = New-Object System.Management.Automation.PSCredential 'VVO', (ConvertTo-SecureString -AsPlainText 'MyVoiceIsMyPassport' -force) 
$login      = Connect-DMLibrary -DMLibrary POVOOPEN -Credential $credential
Invoke-DMSQL "SELECT FORM_NAME, FORM_TITLE from DOCSADM.FORMS"

#>


<# Sample result

FORM_NAME          FORM_TITLE                              
---------          ----------                              
PD_SEARCH          Zoekformulier RM                        
P_SOORTREGISTR     SOORT REGISTRATIE                       
JA_ADVIEZEN        JA Adviezen                             
JA_FUNCTIES        JA Functies                             
P_LAND             LAND                                    
Z_ORGANISATIE      Organisatie                             
PD_EVENT_MNT       Gebeurtenis                             
PD_FILE_SERIES_MNT Dossiercodering                         
C_PLAATSING        Plaatsings kode's                       
_N_TELEFOON        Telefoonregistratie (<16-11-09)         
PD_DISP_ACTION_MNT Verwijderactie                          
PD_PPROF           Papieren document RM                    
DEF_QBE39          STANDAARD ZOEKFORMULIER 39              
DUMMY              NIET GEAUTORISEERD                      
PEOPLE             Gebruikers                              
DEF_PROF           Standaard registratieformulier          
PD_EPROF           Elektronisch document RM                
DEF_HITDM5_ALG     Zoekresultaten DM5 Algemeen             
_N_WERKREGISTRATIE Werkregistratie (<16-11-09)             
DEF_QBE            Zoekformulier                           
N_BB               Bestuurlijke boete vertraagde aanmelding
_N_INDICATIE       Indicatiegeschillen                     
UBZ_RUBRIEK4       UBZ Rubriek(4)                          
PD_EPROF_IMAGE     Elektronische afbeelding RM             
N_DUIDEN           Duiden / Adviseren (project 96)         
N_Z_ZOEKEN_UITGEBR Zoeken uitgebreid CAK                   
N_DOCUMENT_VERKORT Verkorte documentregistratie CAK        
N_Z_BB             Zoeken bestuurlijke boetes              
N_Z_NUMMERS        Zoeken nummers                          
N_UNIT             Verstrekkingengeschillen                
WIJZIG_SCHERM      Wijzigen registratiescherm              
PD_MPROF           E-mailbericht RM                        
PD_HITLIST         Resultatenformulier RM                  
_N_UNIT            Verstrekkingengeschillen                
DEF_HITDM5_WEB     Zoekresultaten DM5 Webtop               
C_DELETED          Verwijderde gegevens                    
DOCUMENT_LIST      Standaard zoekresultaat                 
SCHERMEN           Schermen                                
PD_FILE_PART_PROF  Dossiermapregistratie                   
DEF_MPROF          Standaard e-mailregistratieformulier    
_N_PO_VACATURE     P&O Vacaturestelling                    
UBZ_ONDERWERPEN    UBZ ONDERWERPEN                         
UBZ_AFWIJS         JA Afwijzings gronden                   
DS_NIBG            DS NIET IN BEHANDELING GENOMEN          
TERMINOLOGY        Terminologie                            
GROUP_DEF          Document Profile groepsinstellingen     
DOCUMENTTYPES      Document Types                          
UBZ_RUBRIEK3       UBZ Rubriek(3)                          
PD_DESTROY         Vernietigen                             
UBZ_ADVIES         UBZ Adviezen                            
Z_FR_PROJ          ASB PROJECTEN                           
Z_PROJECT          PROJECTEN                               
PD_AUTHORITY_MNT   Verwijderbron                           
PD_BOX_MNT3        Archiefdoos                             
PD_TERM_MNT        Term                                    
KA_SETTING         PowerPOVO settings                      
Z_PRODUKT          Produkten JA/DF/DZ                      
N_PRIVE            Prive document                          
UBZ_RUBRIEK2       UBZ Rubriek(2)                          
OLDPEOPLE          PERSONEN                                
_N_Z_KLACHTEN      Zoeken klachten                         
C_TYPE             Type                                    
N_INDICATIE        Indicatiegeschillen                     
N_Z_INDICATIE2     Zoeken Indicatiegeschillen              
N_Z_ZOEKEN         Zoeken CAK                              
_N_KLACHT          Klachtregistratie                       
PD_TRANSFER_TO     Overbrengen                             
N_Z_UNIT2          Zoeken Verstrekkingengeschillen         
N_Z_DUIDEN         Zoeken Duiden / Adviseren (project 96)  
UBZ_RUBRIEK1       UBZ Rubriek(1)                          
FOLDER             Folder registratie                      
_N_UNIT_ZVW        Verstrekkingengeschillen art.114 ZVW    
N_MODELPOLISSEN    Modelpolissen                           
_N_DOCUMENT_VERKOR Verkorte documentregistratie            
PD_LOCATION_MNT    Lokatie                                 
PD_SECTION_MNT     Afdeling                                
PD_SCHEDULE_MNT    Verplaatsactie                          
A_RECHTENMATRIX    Rechtenmatrix                           
N_UNIT_ZVW         Verstrekkingengeschillen art.114 ZVW    
PD_BOX_PROF        Archiefdoosregistratie                  
PD_ROLLOVER_MNT    Afsluit/aanmaakactie                    
_N_MODELPOLISSEN   Modelpolissen (<16-11-09)               
_N_BB              Bestuurlijke boete vertraagde aanmelding
_N_DOCUMENT        Documentregistratie                     
C_OPENBAARHEID     Openbaarheid                            
N_PO_VACATURE      P&O Vacaturestelling                    
PD_FILE_PART_MNT3  Dossiermap                              
N_DOCUMENT         Documentregistratie CAK                 
DEF_PROF39         Standaard Documentregistratie 39        
_N_DUIDEN          Duiden / Adviseren (project 96)         
N_Z_UNIT_ZVW2      Zoeken Verstrekkingengeschillen 114ZVW 

#>
