using System.IO;
using Aspose.Words;
using DocumentFormat.OpenXml.Packaging;

namespace SampleNamespace
{
    public static class SampleClass
    {
        public static Microsoft.Office.Interop.Word.Document wordDocument { get; set; }
        public static void GenerateDocument()
        {
            string rootPath = @"D:\Temp";
            string newTemplate = rootPath + @"\item1.xml";
            string templateDocument = rootPath + @"\Test1_updated.docx";
            string outputDocument = rootPath + @"\MyGeneratedDocument.docx";
            string outputDocumentPdf = rootPath + @"\MyGeneratedDocument.pdf";


            File.Create(outputDocument).Close();

            File.Copy(templateDocument, outputDocument, true);

            using (WordprocessingDocument actualContract = WordprocessingDocument.Open(outputDocument, true))
            {
                //get the main part of the document which contains CustomXMLParts
                MainDocumentPart mainPart = actualContract.MainDocumentPart;

                //delete all CustomXMLParts in the document. If needed only specific CustomXMLParts can be deleted using the CustomXmlParts IEnumerable
                mainPart.DeleteParts<CustomXmlPart>(mainPart.CustomXmlParts);

                //var parts = mainPart.GetPartsOfType<CustomXmlPart>();

                //add new CustomXMLPart with data from new XML file
                CustomXmlPart myXmlPart = mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                using (FileStream stream = new FileStream(newTemplate, FileMode.OpenOrCreate))
                {
                    myXmlPart.FeedData(stream);

                }

                //mainPart.AddPart(myXmlPart);               
            }
        }
        public static void Main()
        {
            GenerateDocument();
        }
    }
}