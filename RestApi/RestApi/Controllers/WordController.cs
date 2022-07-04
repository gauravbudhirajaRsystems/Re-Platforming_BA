using DocumentFormat.OpenXml.Packaging;
using Microsoft.AspNetCore.Mvc;
using SharedLibrary;

namespace RestApi.Controllers
{
    [ApiController]
    [Route("api/word")]
    public class WordController : ControllerBase
    {
        [HttpPost("addparagraph")]
        public IActionResult AddParagraph([FromBody] Request request)
        {
            string receivedXml = SharedClass.ReplaceParagraphText(request);

            return new JsonResult(receivedXml);
        }
    }
}
