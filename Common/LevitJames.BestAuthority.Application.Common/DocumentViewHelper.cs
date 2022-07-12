using LevitJames.MSOffice.Common.WordExtensions;
using Microsoft.Office.Interop.Word;
using System;

namespace LevitJames.BestAuthority.Application.Common
{
    public class DocumentViewHelper
    {
        public static T RunFuncFromPrintView<T>(BADocument document, Func<T> func)
        {
            return func();
            //var vw = document.WordDocument.ActiveWindow.View;
            //var origView = vw.Type;
            //try
            //{
            //    if (origView != WdViewType.wdPrintView)
            //    {
            //        WordExtensions.LockScreenUpdating();
            //        vw.Type = WdViewType.wdPrintView;
            //    }

            //    return func();
            //}
            //finally
            //{
            //    if (origView != WdViewType.wdPrintView)
            //    {
            //        vw.Type = origView;
            //        WordExtensions.UnLockScreenUpdating();
            //    }
            //}
        }
    }
}
