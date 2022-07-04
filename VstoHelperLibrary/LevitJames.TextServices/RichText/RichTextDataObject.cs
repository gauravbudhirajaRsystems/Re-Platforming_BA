// © Copyright 2018 Levit & James, Inc.

using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LevitJames.Interop;

namespace LevitJames.TextServices
{
    /// <summary>
    ///     Internal IDataObject implementation used for transferring rtf to the TextRange objects.
    /// </summary>
    internal class RichTextDataObject : IDataObject
    {
        // ReSharper disable once InconsistentNaming
        private const int DV_E_FORMATETC = unchecked((int)0x80040064);

        public static short RichTextFormat = (short)-Convert.ToInt16((NativeMethods.RegisterClipboardFormat("Rich Text Format") ^ 0xFFFF) + 1);

        private readonly string _rtf;
        private Stream _stream;


        public RichTextDataObject(string rtf)
        {
            _rtf = rtf ?? string.Empty;
        }
        public RichTextDataObject(Stream stream)
        {
            _stream = stream;
        }

        public void GetData(ref FORMATETC format, out STGMEDIUM medium)
        {
            medium = new STGMEDIUM();
            if (!IsFormatEtcValid(format))
                throw new COMException(nameof(GetData), DV_E_FORMATETC);

            if (_stream != null)
            {
                var converter = new RichTextConverter();
                medium = converter.StreamToStgMedium(_stream);
                return;
            }

            medium.tymed = TYMED.TYMED_HGLOBAL;
            medium.unionmember = Marshal.StringToHGlobalAnsi(_rtf);
        }

        public void GetDataHere(ref FORMATETC format, ref STGMEDIUM medium)
        {
            throw new NotImplementedException();
        }


        public int QueryGetData(ref FORMATETC format)
        {
            if (!IsFormatEtcValid(format))
                return DV_E_FORMATETC;
            return 0;
        }

        public int GetCanonicalFormatEtc(ref FORMATETC formatIn, out FORMATETC formatOut)
        {
            throw new NotImplementedException();
        }

        public void SetData(ref FORMATETC formatIn, ref STGMEDIUM medium, bool release)
        {
            throw new NotImplementedException();
        }

        public IEnumFORMATETC EnumFormatEtc(DATADIR direction)
        {
            throw new NotImplementedException();
        }

        public int DAdvise(ref FORMATETC pFormatetc, ADVF advf, IAdviseSink adviseSink, out int connection)
        {
            throw new NotImplementedException();
        }

        public void DUnadvise(int connection)
        {
            throw new NotImplementedException();
        }

        public int EnumDAdvise(out IEnumSTATDATA enumAdvise)
        {
            throw new NotImplementedException();
        }

        bool IsFormatEtcValid(FORMATETC formatetc) =>
            formatetc.cfFormat == RichTextFormat && (formatetc.tymed & TYMED.TYMED_HGLOBAL) != 0 && formatetc.dwAspect == DVASPECT.DVASPECT_CONTENT && formatetc.lindex == -1;
    }
}