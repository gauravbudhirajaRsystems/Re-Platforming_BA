// © Copyright 2018 Levit & James, Inc.

using System.ComponentModel;
using System.Runtime.InteropServices.ComTypes;

namespace LevitJames.TextServices
{
    /// <summary>
    ///     A helper class for setting and retrieving the rtf text from a TextRange or TextSelection.
    /// </summary>

    [EditorBrowsable(EditorBrowsableState.Advanced)]
    internal sealed class TextRangeDataObject
    {
        /// <summary>
        ///     The dataobject used to get the rtf from the TextRange or SelectionRange.
        /// </summary>



        public IDataObject DataObject { get; internal set; }


        ///// <summary>
        ///// Called by the TextRangeVariantMarshaller instance to set the dataobject retrieved from the TextRange or SelectionRange.
        ///// </summary>
        ///// <param name="ido"></param>
        //
        //internal void SetValue(System.Runtime.InteropServices.ComTypes.IDataObject ido)
        //{
        //    if (ido != null)
        //    {
        //        Data = new DataObject(ido);
        //    }
        //}
    }
}