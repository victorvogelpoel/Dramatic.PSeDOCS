# Add-DMPCDClientType.ps1
# Load the eDOCS PCDClient API interop, getting location from the Registry
# Jul 2017
# Copyright 2017 Dramatic Development
# If this works, it was written by Victor Vogelpoel (victor@victorvogelpoel.nl).
# If it doesn't, I don't know who wrote it.

function Add-DMPCDClientType
{
    [CmdletBinding()]
    param
    ()

    Write-Verbose 'Initializing eDOCS API'

    $is32BitPowerShell = ([IntPtr]::Size -eq 4)
    Write-Verbose "Running 32-bit PowerShell: $is32BitPowerShell"

    if ($is32BitPowerShell)
    {
        # We're in 32 bit PowerShell
        $hummingbirdInstallPath = (Get-ItemProperty 'HKLM:\SOFTWARE\Hummingbird\Hummingbird DM API').InstallPath
    }
    else
    {
        # We're in 64 bit PowerShell. Read from 32-bit node in 64 bit Registry
        $hummingbirdInstallPath = (Get-ItemProperty 'HKLM:\SOFTWARE\Wow6432Node\Hummingbird\Hummingbird DM API').InstallPath
    }
    
    $PCDClientDLLFilePath   = (Join-Path $hummingbirdInstallPath 'Hummingbird.DM.Server.Interop.PCDClient.dll')

    if (Test-Path -Path $PCDClientDLLFilePath)
    {
        add-type -Path $PCDClientDLLFilePath
        Write-Verbose "Loaded eDOCS PCDClient interop DLL from `"$PCDClientDLLFilePath`"."

        if (!$is32BitPowerShell)
        {
            # Okay, we're on a 64bit Windows machine in 64bit PowerShell. 
            # The eDOCS API is normally 32bit COM (even on a 64bit Windows).
            # There is a 64bit COM eDOCS API; however, this must have been installed separately.

            Write-Verbose 'Testing if the eDOCS x64 API is installed, because we are running on x64 PowerShell.'
            # Test if is it installed.
            try
            {
                # This will fail on Windows x64 without eDOCS x64 API installed
                # "Retrieving the COM class factory for component with CLSID {BAE80C14-D2AC-11D0-8384-00A0C92018F4} failed due to the following error: 80040154 Class not registered (Exception from HRESULT: 0x80040154 (REGDB_E_CLASSNOTREG))."
                $pcdLoginLibsTest = New-Object Hummingbird.DM.Server.Interop.PCDClient.PCDGetLoginLibsClass
            }
            catch [System.Management.Automation.MethodInvocationException]
            {
                # "Retrieving the COM class factory for component with CLSID {BAE80C14-D2AC-11D0-8384-00A0C92018F4} failed due to the following error: 80040154 Class not registered (Exception from HRESULT: 0x80040154 (REGDB_E_CLASSNOTREG))."

                Write-Verbose 'RESULT: eDOCS x64 API is installed is NOT installed.'
                throw 'ERROR initializing eDOCS API: it seems we are running 64-bit PowerShell and the eDOCS 64-bit API is not installed... You may want to start this module in an x86 PowerShell session or install the x64 eDOCS API.'
            }
            finally
            {
                # if instantiation of the COM object succeeded, then the variable is set and the COM object must be released.
                if ((Get-Variable 'pcdLoginLibsTest' -ErrorAction SilentlyContinue) -and ($null -ne $pcdLoginLibsTest))
                {
                    [VOID][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($pcdLoginLibsTest)
                }
            }            
        }
    }
    else
    {
        $message = "ERROR initializing eDOCS: the PCDClient.dll was not found at `"$PCDClientDLLFilePath`". Is the eDOCS API installed?"
        if (!$is32BitPowerShell)
        {
            $message += " Don't forget to install the 64-Bit API as well, if you want to use eDOCS from 64-bit PowerShell"
        }
        
        throw $message
    }


    # =============================================================================================================
    # Some C# code to get the document and its metadata.
    # 
    # Rationale: Somehow, if we write this eDOCS client code in pure PowerShell, it crashes
    # crashes due to some COM culprits:
    #  - GetDocument: when setting "%OBJECT_IDENTIFIER" property with a long; a string DocNum will work though
    #  - GetDocument: when starting the iterator of the ReturnProperties propertyList: "propList.BeginIter()"

    # The C# version below is working without problems and can be called from PowerShell.

    $GetDocumentSource = @"

        using System;
        using System.Collections.Generic;
        using System.Text;
        using Hummingbird.DM.Server.Interop.PCDClient;
        using System.Runtime.InteropServices;

        namespace Dramatic.eDOCS
        {
            public class MetaData
            {
                public MetaData(string name)
                {
                    this.Name = name;
                }

                public MetaData(string name, object value) : this(name)
                {
                    Value = value;
                }

        
                public string Name { get; set; }
                public object Value { get; set; }

                public override string ToString()
                {
                    return string.Format("Metadata, Name:[{0}][{1}]", Name, Value);
                }
            }

            public class Document
            {
                public Document(string library, long docNum, string formName)
                {
                    this.Library  = library;
                    this.DocNum   = docNum;
                    this.FormName = formName;
                    MetaData      = new List<MetaData>();
                }
   
                public long DocNum { get; private set; }
                public string Library { get; set; }
                public string FormName { get; set; }
                public List<MetaData> MetaData { get; set; }
        

                public override string ToString()
                {
                    //return string.Format("Document:{0}, {1}, {2}, Metadata:{3}, Versions:{4}", DocNum, Library, FormName, MetaData.Count, Versions.Count);
                    return string.Format("Document:{0}, {1}, {2}, Metadata:{3}", DocNum, Library, FormName, MetaData.Count);
                }
            }



            //==================================================================================================================================
            // The static Client class with some eDOCS DAL functions

            public static class Client
            {
                public static Document GetDocument(string DST, string library, string formName, long docNum)
                {
                    PCDDocObject pcdDocObject = null;
                    PCDPropertyList propList  = null;
                    Document document         = null;

                    try
                    {
                        pcdDocObject = new PCDDocObject();
                        pcdDocObject.SetDST(DST);
                        pcdDocObject.SetObjectType(formName);
                        pcdDocObject.SetProperty("$($eDOCSTokens.TARGET_LIBRARY)", library);
                        pcdDocObject.SetProperty("$($eDOCSTokens.OBJECT_IDENTIFIER)", docNum);

                        pcdDocObject.Fetch();
                        if (pcdDocObject.ErrNumber != 0)
                        { 
                            throw new ApplicationException(String.Format("ERROR {1} while fetching document {0}: {2}", docNum, pcdDocObject.ErrNumber, pcdDocObject.ErrDescription));
                        }

                        // Create a new Document model instance and fill it with metadata
                        document = new Document(library, docNum, formName);

                        propList = pcdDocObject.GetReturnProperties();
                        int currentIndex = propList.BeginIter();
                        do
                        {
                            string propName = propList.GetCurrentPropertyName();

                            //Skip system-properties starting with %
                            if (propName.StartsWith("%"))
                            {
                                currentIndex = propList.NextProperty();
                                continue;
                            }

                            dynamic propValue = propList.GetCurrentPropertyValue();
                            var metaData      = new MetaData(propName, propValue as object);
                            document.MetaData.Add(metaData);

                            currentIndex = propList.NextProperty();
                        }
                        while (currentIndex == 0);

                        return document;
                    }
                    finally
                    {
                        if (propList != null)
                        {
                            Marshal.FinalReleaseComObject(propList);
                        }
                        if (pcdDocObject != null)
                        {
                            Marshal.FinalReleaseComObject(pcdDocObject);
                        }
                    }
                }
            }
        }
"@

    # reference the PCDClient DLL:
    Add-Type -ReferencedAssemblies $PCDClientDLLFilePath -TypeDefinition $GetDocumentSource -Language CSharp
}
