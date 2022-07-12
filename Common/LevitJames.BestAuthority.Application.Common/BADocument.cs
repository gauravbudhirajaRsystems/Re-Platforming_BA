using LevitJames.AddinApplicationFramework.Common.App;
using LevitJames.MSOfficeInterop.Common.CompilerExtensions;
using Microsoft.AspNetCore.Http;
using Microsoft.Office.Interop.Word;

namespace LevitJames.BestAuthority.Application.Common
{
    public class BADocument : AddinAppDocument
    {
        public Document WordDocument { get; set; }

        public bool HasNullHeaderFooter() => DocumentViewHelper.RunFuncFromPrintView(this, () => WordDocument.HasNullHeaderFooter());
    }
}
