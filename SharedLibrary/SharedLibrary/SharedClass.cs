using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Office.Interop.Word;
using System.Text;
using System.Xml.Linq;
using Range = Microsoft.Office.Interop.Word.Range;

namespace SharedLibrary
{
    public class SharedClass
    {
        public static string ReplaceParagraphText(Request request)
        {

            //using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(@"C:\NewDocument.docx", WordprocessingDocumentType.Document))
            //{
            //    wordDocument.AddMainDocumentPart();
            //    wordDocument.MainDocumentPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document
            //    {
            //        Body = new Body
            //        {
            //            InnerXml = request.InnerXml
            //        }
            //    };
            //    wordDocument.MainDocumentPart.Document.Save();
            //    wordDocument.Close();
            //    wordDocument.Dispose();
            //}

            //using (MemoryStream ms = new MemoryStream(Encoding.ASCII.GetBytes(request.WordOpenXML)))
            //{
            //    using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            //    {
            //        wordDocument.AddMainDocumentPart();
            //        wordDocument.SaveAs(@"C:\DocFromAPI.docx");
            //        wordDocument.Close();
            //        wordDocument.Dispose();
            //    }
            //}


            //Application application = new();
            //Microsoft.Office.Interop.Word.Document wordDoc = application.Documents.Open(@"C:\DocFromAPI.docx");
            //Range rngPara = wordDoc.Paragraphs[1].Range;
            //object unitCharacter = WdUnits.wdCharacter;
            //object backOne = -1;
            //rngPara.MoveEnd(ref unitCharacter, ref backOne);
            //rngPara.Text = "replacement text";
            //wordDoc.SaveAs2(@"C:\NewDocument_Updated.docx");

            var docXml = XDocument.Parse(request.WordOpenXML);
            docXml.Save(@"C:\DocFromAPI.docx");

            var numbering = docXml.Descendants().Where(x => (string)x.Attribute(Pkg.name)! == "/word/numbering.xml");

            List<AbstractNum>  abstractNumList = new List<AbstractNum>();

            var abstractNums = numbering.Descendants(W.abstractNum);

            foreach (var abstractNum in abstractNums)
                abstractNumList.Add(new AbstractNum(abstractNum.ToString()));


            var mDoc = MDocumentCreator.CreateMDocument(request.WordOpenXML);


            return string.Empty;
        }
    }
}