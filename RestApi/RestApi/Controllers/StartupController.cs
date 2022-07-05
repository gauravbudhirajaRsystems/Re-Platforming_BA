using Microsoft.AspNetCore.Mvc;
using SharedLibrary;

namespace RestApi.Controllers
{
    [Route("api/startup")]
    [ApiController]
    public class StartupController : ControllerBase
    {
        [HttpPost("HasNullHeaderFooter")]
        public IActionResult HasNullHeaderFooter([FromBody] BADocument Document)
        {
            bool result = DocumentHelper.HasNullHeaderFooter(Document);

            return new JsonResult(result);
        }
    }
}
