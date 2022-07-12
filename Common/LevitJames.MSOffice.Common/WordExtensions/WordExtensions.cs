using LevitJames.Core.Common;
using LevitJames.Core.Common.Diagnostics;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Text;

namespace LevitJames.MSOffice.Common.WordExtensions
{
    /// <summary>
    ///     A singleton class containing extensions for Microsoft
    /// </summary>

    [SuppressMessage("Microsoft.Maintainability", "CA1506:AvoidExcessiveClassCoupling")]
    [DesignerCategory(@"Code")]
    public sealed class WordExtensions : IDisposable
    {
        private static LockCounter _wordScreenUpdateLock;

        /// <summary>
        ///     Returns a Microsoft Word Application object
        /// </summary>

        /// <returns>A Application instance</returns>
        public static Application WordApplication { get; private set; }

        public void Dispose()
        {
            throw new NotImplementedException();
        }



        /// <remarks>
        ///     This method uses reference counting so ScreenUpdating is not turned back on until the same number of
        ///     UnLockScreenUpdating has been called or UnLockScreenUpdating is called with the reset set to true. It is
        ///     recommended that Locking and unlocking is done using a Try Finally Block to guarantee the LockScreenUpdating are
        ///     always balanced with the same number of UnLockScreenUpdating calls.
        /// </remarks>
        public static bool LockScreenUpdating()
        {
            if (_wordScreenUpdateLock.Lock())
            {
                if (AppDiagnostics.GetOption(AppDiagnosticOptions.SuppressScreenLocking) == false)
                {
                    WordApplication.ScreenUpdating = false;
                }

                return true;
            }

            return false;
        }

        /// <summary>
        ///     UnLocks Words ScreenUpdating.
        /// </summary>
        /// <param name="reset">
        ///     Resets the screen locking, and turns painting back on. This member should only be used in rare
        ///     cases, such as unhandled exceptions.
        /// </param>

        /// <remarks>
        ///     ScreenUpdating is not turned back on until UnLockScreenUpdating has been called the same number of times as
        ///     the LockScreenUpdating call or UnLockScreenUpdating is called with the reset set to true.
        /// </remarks>
        public static bool UnLockScreenUpdating(bool reset = false)
        {
            if (reset)
            {
                WordApplication.ScreenUpdating = true;
                WordApplication.ScreenRefresh();
                _wordScreenUpdateLock.Reset();
                return true;
            }

            if (_wordScreenUpdateLock.Unlock())
            {
                if (AppDiagnostics.GetOption(AppDiagnosticOptions.SuppressScreenLocking) == false)
                {
                    WordApplication.ScreenUpdating = true;
                    WordApplication.ScreenRefresh();
                }

                return true;
            }

            return false;
        }
    }
}
