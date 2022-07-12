using Microsoft.AspNetCore.Mvc;
using Microsoft.Office.Interop.Word;
using System.Linq;
using System.Xml.Linq;
using Task = System.Threading.Tasks.Task;

namespace RestApi.Controllers
{
   public class BADocument
    {
        public Document Document { get; set; }
    }

    [Route("api/startup")]
    [ApiController]
    public class StartupController : ControllerBase
    {
        [HttpPost("HasNullHeaderFooter")]
        public IActionResult HasNullHeaderFooter([FromBody] XDocument xDocument) //
        {
            //bool result = DocumentHelper.HasNullHeaderFooter(Document);
            
            //return new JsonResult(result);
            return null;
        }


        //[HttpPost("HasNullHeaderFooter")]
        //public IActionResult HasNullHeaderFooter([FromBody] Document Document) //
        //{
        //    //bool result = DocumentHelper.HasNullHeaderFooter(Document);

        //    //return new JsonResult(result);
        //    return null;
        //}

        //[HttpPost("HasNullHeaderFooter")]
        //public async Task<IActionResult> HasNullHeaderFooter([FromBody] byte[] byteArray) //
        //{
        //    //System.IO.File.WriteAllBytes(@"C:\ReceivedDoc1.docx", byteArray);
        //    //Application ap = new();
        //    //Document document = ap.Documents.Open(@"C:\NewDoc1.docx");
        //    //var pageCount = document.ActiveWindow.Panes[1].Pages.Count;
        //    //ap.Quit();

        //    //Application ap1 = new();
        //    //Document document1 = ap.Documents.Open(@"C:\NewDoc2.docx");
        //    //var pageCount1 = document1.ActiveWindow.Panes[1].Pages.Count;
        //    //ap1.Quit();
        //    //string path1 = @"C:\NewDoc1.docx";
        //    //string path2 = @"C:\NewDoc2.docx";
        //    //string path3 = @"C:\NewDoc3.docx";

        //    List<Task<int>> t = new();
        //    for (int i = 1; i <= 3; i++)
        //    {
        //        t.Add(GetTask(i));
        //    }
        //    var result1 = await System.Threading.Tasks.Task.WhenAll(t);
        //    var result = result1.Select(x => x);
        //    // return null;
        //    //return System.Threading.Tasks.Task.FromResult(null);

        //    //BADocument bADocument = new() { WordDocument = document };
        //    //bool result = DocumentHelper.HasNullHeaderFooter(bADocument);
        //    return new JsonResult(result);
        //}

        //private Task<int> GetTask(int i)
        //{
        //    Application ap = new();
        //    Document document = ap.Documents.Open(@$"C:\NewDoc{i}.docx", null, true);
        //    var pageCount = document.ActiveWindow.Panes[1].Pages.Count;

        //    return Task.FromResult(pageCount);
        //}


        //[HttpPost("HasNullHeaderFooter")]
        //public IActionResult HasNullHeaderFooter([FromBody] byte[] byteArray) //
        //{
        ////System.IO.File.WriteAllBytes(@"C:\ReceivedDoc.docx", byteArray);

        //using (WordprocessingDocument doc = WordprocessingDocument.Create("C:\\ReceivedDoc1.docx", false))
        //{

        //    // Add a main document part.
        //    //doc.MainDocumentPart.Document.Body.
        //    //// Create the document structure and add some text.
        //    //mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
        //    //Body body = mainPart.Document.AppendChild(new Body());
        //    //DocumentFormat.OpenXml.Drawing.Paragraph para = body.AppendChild(new DocumentFormat.OpenXml.Drawing.Paragraph());
        //    //Run run = para.AppendChild(new Run());

        //    //// String msg contains the text, "Hello, Word!"
        //    //run.AppendChild(new Text("New text in document"));
        //}
        //return null;

        //    using (MemoryStream ms = new())
        //    {
        //        ms.Write(byteArray, 0, byteArray.Length);
        //        //FileStream fileStream = new FileStream(@"C:\NewDoc.docx", FileMode.Create, FileAccess.ReadWrite);
        //        //ms.WriteTo(fileStream);
        //        //fileStream.Close();
        //        using (WordprocessingDocument document = WordprocessingDocument.Open(ms, false))
        //        {
        //            int pageCount = Convert.ToInt32(value: document.ExtendedFilePropertiesPart.Properties.Pages.Text);
        //            var paragraphs = document.MainDocumentPart.Document.Body.OfType<Paragraph>().ToList();
        //        }
        //    }

        //    return Ok();

        //}
    }
}
