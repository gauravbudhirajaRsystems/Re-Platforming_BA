// © Copyright 2018 Levit & James, Inc.

namespace LevitJames.MSOffice.MSWord
{
    /// <summary>
    ///     Defines the standard dialog return results for a Word ShowDialog call.
    /// </summary>
    public enum WordDialogResult
    {
        /// <summary>
        ///     The user canceled the dialog.
        /// </summary>
        Cancel = 0,

        /// <summary>
        ///     The dialog operation succeeded.
        /// </summary>
        Ok = -1,

        /// <summary>
        ///     The dialog was closed.
        /// </summary>
        Close = -2

        //   -2: Close (cancel when printer is changed)
        //   -1: OK
        //    0: Cancel
        //   >0: Button number
    }
}