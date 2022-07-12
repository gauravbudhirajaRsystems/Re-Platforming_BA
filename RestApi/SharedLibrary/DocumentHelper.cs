using LevitJames.BestAuthority.Application.Common;
using Microsoft.Office.Interop.Word;

namespace SharedLibrary
{
    public class DocumentHelper
    {
        public static bool HasNullHeaderFooter(BADocument Document)
        {
            return Document.HasNullHeaderFooter();
            //return false;
        }


    }
}
