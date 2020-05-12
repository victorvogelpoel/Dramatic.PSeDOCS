

# Rationale: Somehow, if we write this eDOCS client code in pure PowerShell, it crashes
# crashes due to some COM culprits:
#  - GetDocument: when setting "%OBJECT_IDENTIFIER" property with a long; a string DocNum will work
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

$referencedAssemblies = 'C:\Program Files (x86)\Open Text\DM API\Hummingbird.DM.Server.Interop.PCDClient.dll'

Add-Type -ReferencedAssemblies $referencedAssemblies -TypeDefinition $GetDocumentSource -Language CSharp

# IVHO
$DST      = "some_dst";
$docNum   = 4051936;
$formName = "IVHO_DEF_PROF";
$library  = "IVHO_EDOCS";


$document = [Dramatic.eDOCS.Client]::GetDocument($DST, $library, $formName, $docNum)
$document.MetaData 

