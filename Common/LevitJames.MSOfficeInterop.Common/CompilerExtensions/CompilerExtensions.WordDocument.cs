using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace LevitJames.MSOfficeInterop.Common.CompilerExtensions
{
    public static partial class Extensions
    {

        public static bool HasNullHeaderFooter(this Document document)
        {
            var pages = document.ActiveWindow.Panes[1].Pages;
            var pageCount = pages.Count;  // For troubleshooting
            var curPage = 0;              // For troubleshooting

            while (true)
            {
                curPage++;
                if (curPage > pages.Count)
                    break;
                var pageHasHdr = false;
                var pageHasFtr = false;
                Page page;
                try
                {
                    page = pages[curPage];
                }
                catch (COMException ex) when (ex.ErrorCode == -2146822347)
                {
                    // If the document is doing odd/even pages and there's a section break/next page,
                    // Word will insert a blank page. However, when accessing the Pages collection for that blank
                    // page, Word will throw an error that the member of the collection does not exist. In this
                    // case, just skip and go to the next page.
                    continue;
                }
                var rects = page.Rectangles;

                curPage++; // For troubleshooting
                for (var i = 1; i <= rects.Count; i++)
                {
                    var rct = rects[i];
                    var rng = rct.RangeLJ();
                    try
                    {
                        switch (rng?.StoryType)
                        {
                            case null:
                                break;

                            case WdStoryType.wdPrimaryHeaderStory:
                            case WdStoryType.wdFirstPageHeaderStory:
                            case WdStoryType.wdEvenPagesHeaderStory:
                                pageHasHdr = true;
                                break;
                            case WdStoryType.wdPrimaryFooterStory:
                            case WdStoryType.wdFirstPageFooterStory:
                            case WdStoryType.wdEvenPagesFooterStory:
                                pageHasFtr = true;
                                break;
                        }

                        if (pageHasHdr && pageHasFtr)
                            break;
                    }
                    finally
                    {
                        if (rng != null)
                            Marshal.ReleaseComObject(rng);

                        if (rct != null)
                            Marshal.ReleaseComObject(rct);
                    }
                }

                Marshal.ReleaseComObject(page);

                if (!pageHasHdr || !pageHasFtr)
                    return true;

            }

            return false;
        }
    }
}
