// © Copyright 2018 Levit & James, Inc.

using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security;
using System.Text;
using LevitJames.Core;
using LevitJames.Interop;

namespace LevitJames.TextServices
{
    /// <summary>
    ///     A static class for validating and extracting rtf to a byte array.
    /// </summary>

    public class RichTextConverter
    {
        //Private constructor' All members are Static.
        internal static readonly short RichTextFormat;

        static RichTextConverter()
        {
            RichTextFormat =
                (short)-Convert.ToInt16((NativeMethods.RegisterClipboardFormat("Rich Text Format") ^ 0xFFFF) + 1);
        }


        /// <summary>
        ///     Returns a <see cref="byte">Byte</see> array of ASCII formatted rtf from the supplied
        ///     <see cref="Stream">IO.Stream</see>.
        /// </summary>
        /// <param name="stream">An <see cref="Stream">IO.Stream</see> containing the formatted rtf to copy.</param>
        /// <param name="disposeStream">True if the method should Dispose of the stream after conversion, false otherwise.</param>
        /// <returns>A byte array representing the ASCII rtf characters.</returns>
        /// <remarks>
        ///     May throw an <see cref="RtfReaderException">RtfReaderException</see> if the stream does not contain valid rtf
        ///     data.
        /// </remarks>
        public byte[] ToBytes(Stream stream, bool disposeStream)
        {
            Check.NotNull(stream, "stream");

            try
            {
                var fileSizeInt32 = Convert.ToInt32(stream.Length);
                byte[] bytes = { };

                if (fileSizeInt32 > 0)
                {
                    bytes = new byte[fileSizeInt32];
                    stream.Read(bytes, offset: 0, count: fileSizeInt32);
                }

                InvalidRtfGuard(bytes);
                return bytes;
            }
            finally
            {
                if (disposeStream)
                {
                    stream.Dispose();
                }
            }
        }

        ///// <summary>
        ///// Returns a <see cref="byte">Byte</see> array of ASCII formatted rtf from the supplied <see cref="Stream">IO.Stream</see>.
        ///// </summary>
        ///// <param name="source">An <see cref="TextRange">TextRange</see> containing the formatted rtf to copy.</param>
        ///// <returns>A byte array representing the ASCII rtf characters.</returns>
        ///// <remarks>May throw an <see cref="RtfReaderException">RtfReaderException</see> if the stream does not contain valid rtf data.</remarks>

        //public byte[] ToBytes(TextRange source)
        //{
        //    Check.NotNull(source, "source");

        //    var trDo = new TextRangeDataObject();
        //    source.Copy(trDo);

        //    var bytes = new byte[0];

        //    if (trDo.DataObject != null)
        //    {
        //        bytes = ToBytes(trDo.DataObject);
        //    }

        //    return bytes;
        //}

        /// <summary>
        ///     Returns a <see cref="byte">Byte</see> array of ASCII formatted rtf from the supplied
        ///     <see cref="System.Runtime.InteropServices.ComTypes.IDataObject">IDataObject</see>.
        /// </summary>
        /// <param name="source">
        ///     An object that implements the
        ///     <see cref="System.Runtime.InteropServices.ComTypes.IDataObject">IDataObject</see> interface
        /// </param>
        /// <returns>A byte array representing the rtf ASCII characters.</returns>
        /// <remarks>
        ///     May throw an <see cref="RtfReaderException">RtfReaderException</see> if the extracted rtf does not contain valid
        ///     rtf data.
        ///     <para>
        ///         The <see cref="System.Runtime.InteropServices.ComTypes.IDataObject">IDataObject</see> must store the rtf using
        ///         the "Rich Text Format" clipboard format.
        ///     </para>
        /// </remarks>
        [SecurityCritical]
        public byte[] ToBytes(IDataObject source)
        {
            Check.NotNull(source, "source");

            byte[] rtfBytes;

            using (var safeHGlobal = DataObjectTohGlobal(source))
            {
                if (safeHGlobal == null)
                {
                    return new byte[] { };
                }

                if (safeHGlobal.Lock() == false)
                {
                    throw new Win32Exception("Failed to Lock HGLOBAL");
                }

                var size = safeHGlobal.GetSize();
                if (size > 0)
                {
                    rtfBytes = new byte[size];
                    Marshal.Copy(safeHGlobal.LockedData, rtfBytes, startIndex: 0, length: size);
                }
                else
                {
                    return new byte[] { };
                }
            }

            InvalidRtfGuard(rtfBytes);
            return rtfBytes;
        }

        /// <summary>
        ///     Returns a <see cref="byte">Byte</see> array of ASCII formatted rtf from the supplied string.
        /// </summary>
        /// <param name="rtf">A System.String containing the formatted rtf to read.</param>
        /// <returns>A <see cref="byte">Byte</see> array representing the rtf ASCII characters.</returns>
        /// <remarks>An RtfReaderException is thrown if the string passed does not contain valid rtf data.</remarks>
        public byte[] ToBytes(string rtf)
        {
            Check.NotNull(rtf, "rtf");
            InvalidRtfGuard(rtf);
            return Encoding.ASCII.GetBytes(rtf);
        }

        /// <summary>
        ///     Loads an Rtf file and returns it as a <see cref="byte">Byte</see> array
        /// </summary>
        /// <param name="fileName">The name of the rtf file to load the rtf from.</param>
        /// <returns>A byte array representing the rtf rtf ASCII characters.</returns>
        /// <remarks>
        ///     This member way return an IO exception and also a <see cref="RtfReaderException">RtfReaderException</see> if
        ///     the file is not a valid rtf file.
        /// </remarks>
        [SecurityCritical]
        public byte[] ToBytesFromFileName(string fileName)
        {
            using (var fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                return ToBytes(fs, disposeStream: true);
            }
        }


        /// <summary>
        ///     Extracts the rtf from the supplied <see cref="IDataObject">IDataObject</see> into the destination
        ///     <see cref="Stream">stream.</see>
        /// </summary>
        /// <param name="source">An <see cref="IDataObject">IDataObject</see> instance containing the rtf to copy.</param>
        /// <param name="destination">A stream to copy the ASCII formatted rtf to.</param>
        /// <remarks>
        ///     May throw an <see cref="RtfReaderException">RtfReaderException</see> if the stream does not contain valid rtf
        ///     data.
        /// </remarks>
        [SecurityCritical]
        public bool ToStream(IDataObject source, Stream destination)
        {
            Check.NotNull(source, "data");
            Check.NotNull(destination, "stream");

            var buffer = ToBytes(source);
            var ms = destination as MemoryStream;
            if (ms != null)
            {
                ms.Capacity = buffer.Length;
            }
            else
            {
                var fs = destination as FileStream;
                fs?.SetLength(buffer.Length);
            }

            ms?.Write(buffer, offset: 0, count: buffer.Length);
            return true;
        }


        /// <summary>
        ///     Transfers the rtf contained in the rtfBytes to the supplied file.
        /// </summary>
        /// <param name="rtf">A System.String containing the formatted rtf to read.</param>
        /// <param name="fileName">A string representing the file name and location to save the rtf to.</param>
        /// <remarks>
        ///     May throw an <see cref="RtfReaderException">RtfReaderException</see> if the extracted rtf does not contain valid
        ///     rtf data. May also throw any of the standard IO exceptions.
        /// </remarks>
        public void ToFile(string rtf, string fileName)
        {
            Check.NotEmpty(rtf, "rtf");
            Check.NotEmpty(fileName, "fileName");
            InvalidRtfGuard(rtf);

            File.WriteAllText(rtf, fileName);
        }

        /// <summary>
        ///     Transfers the rtf from the supplied <see cref="TextRange">TextRange</see> to the file.
        /// </summary>
        /// <param name="source">An <see cref="TextRange">TextRange</see> containing the formatted rtf to copy.</param>
        /// <param name="fileName">A string representing the file name and location to save the rtf to.</param>
        /// <returns>True if the rtf was successfully written to the file; false otherwise.</returns>
        /// <remarks>
        ///     May throw an <see cref="RtfReaderException">RtfReaderException</see> if the extracted rtf does not contain valid
        ///     rtf data. May also throw any of the standard IO exceptions.
        /// </remarks>
        public bool ToFile(TextRange source, string fileName)
        {
            Check.NotNull(source, "source");
            Check.NotEmpty(fileName, "fileName");

            if (source.End == source.Start)
            {
                return false;
            }

            var ido = source.CopyToDataObject();

            if (ido != null)
            {
                return ToFile(ido, fileName);
            }

            return false;
        }

        /// <summary>
        ///     Transfers the rtf contained in rtfBytes to the supplied file.
        /// </summary>
        /// <param name="rtf">An array of <see cref="byte">Bytes</see> containing the ASCII formatted rtf to extract.</param>
        /// <param name="fileName">A string representing the file name and location to save the rtf to.</param>
        /// <returns>True if the rtf was successfully written to the file; false otherwise.</returns>
        /// <remarks>
        ///     May throw an <see cref="RtfReaderException">RtfReaderException</see> if the extracted rtf does not contain valid
        ///     rtf data. May also throw any of the standard IO exceptions.
        /// </remarks>
        public bool ToFile(byte[] rtf, string fileName)
        {
            Check.NotNull(rtf, "rtf");
            Check.NotEmpty(fileName, "fileName");
            InvalidRtfGuard(rtf);

            File.WriteAllBytes(fileName, rtf);
            return true;
        }

        /// <summary>
        ///     Extracts the rtf from the supplied <see cref="IDataObject">IDataObject</see> into the destination fileName.
        /// </summary>
        /// <param name="source">An <see cref="IDataObject">IDataObject</see> instance containing the rtf to copy.</param>
        /// <param name="fileName">A string representing the file name and location to save the rtf to.</param>
        /// <returns>True if the rtf was successfully written to the file; false otherwise.</returns>
        /// <remarks>
        ///     May throw an <see cref="RtfReaderException">RtfReaderException</see> if the extracted rtf does not contain valid
        ///     rtf data. May also throw any of the standard IO exceptions.
        /// </remarks>
        public bool ToFile(IDataObject source, string fileName)
        {
            Check.NotNull(source, "source");
            Check.NotEmpty(fileName, "fileName");

            var bytes = ToBytes(source);
            if (bytes != null && bytes.Length > 0)
            {
                return ToFile(bytes, fileName);
            }

            return false;
        }

        /// <summary>
        ///     Extracts the rtf from the supplied <see cref="IDataObject">IDataObject</see> into the destination
        ///     <see cref="FileStream">IO.FileStream</see>.
        /// </summary>
        /// <param name="source">An <see cref="IDataObject">IDataObject</see> instance containing the rtf to copy.</param>
        /// <param name="fileStream">An <see cref="FileStream">IO.FileStream</see> to copy the rtf to.</param>
        /// <returns>True if the rtf was successfully written to the file; false otherwise.</returns>
        /// <remarks>
        ///     May throw an <see cref="RtfReaderException">RtfReaderException</see> if the extracted rtf does not contain valid
        ///     rtf data. May also throw any of the standard IO exceptions.
        /// </remarks>
        public bool ToFile(IDataObject source, FileStream fileStream)
        {
            Check.NotNull(source, "source");
            Check.NotNull(fileStream, "fileStream");

            var bytes = ToBytes(source);
            if (bytes != null && bytes.Length > 0)
            {
                fileStream.Write(bytes, offset: 0, count: bytes.Length);
                return true;
            }

            return true;
        }


        /// <summary>
        ///     Copies the rtf contained in the supplied string into the supplied <see cref="IDataObject">IO.IDataObject</see>.
        /// </summary>
        /// <param name="rtf">A System.String containing the formatted rtf to read.</param>
        /// <param name="destination">An <see cref="IDataObject">IDataObject</see> instance to copy the rtf to.</param>
        /// <returns>True if the rtf was successfully transferred; false otherwise.</returns>
        /// <remarks>
        ///     May throw an <see cref="RtfReaderException">RtfReaderException</see> if the extracted rtf does not contain valid
        ///     rtf data. May also throw com exceptions, or an OutOfMemory exception.
        /// </remarks>
        public bool ToDataObject(string rtf, IDataObject destination)
        {
            Check.NotEmpty(rtf, "rtf");
            Check.NotNull(destination, "destination");
            InvalidRtfGuard(rtf);

            using (var safeHGlobal = new SafeHGlobalHandle(Marshal.StringToHGlobalAnsi(rtf), ownsHandle: true))
            {
                return HGlobalToDataObject(safeHGlobal, destination);
            }
        }

        /// <summary>
        ///     Copies the rtf contained in the supplied string into the supplied <see cref="IDataObject">IO.IDataObject</see>.
        /// </summary>
        /// <param name="rtf">An array of <see cref="byte">Bytes</see> containing the ASCII formatted rtf to extract.</param>
        /// <param name="destination">An <see cref="IDataObject">IDataObject</see> instance to copy the rtf to.</param>
        /// <returns>True if the rtf was successfully transferred; false otherwise.</returns>
        /// <remarks>
        ///     May throw an <see cref="RtfReaderException">RtfReaderException</see> if the extracted rtf does not contain valid
        ///     rtf data. May also throw com exceptions, or an OutOfMemory exception.
        /// </remarks>
        public bool ToDataObject(byte[] rtf, IDataObject destination)
        {
            Check.NotNull(rtf, "rtf");
            Check.NotNull(destination, "destination");
            InvalidRtfGuard(rtf);

            if (rtf.Length == 0)
            {
                return false;
            }

            using (var safeHGlobal = new SafeHGlobalHandle(Marshal.AllocHGlobal(rtf.Length)))
            {
                if (safeHGlobal.Lock())
                {
                    Marshal.Copy(rtf, startIndex: 0, destination: safeHGlobal.LockedData, length: rtf.Length);
                    safeHGlobal.Unlock();
                    return HGlobalToDataObject(safeHGlobal, destination);
                }
            }

            return true;
        }
 
        /// <summary>
        ///     Copies the rtf contained in the supplied string into the supplied <see cref="IDataObject">IO.IDataObject</see>.
        /// </summary>
        /// <param name="source">An <see cref="IDataObject">IDataObject</see> instance containing the rtf to copy.</param>
        /// <param name="destination">An <see cref="IDataObject">IDataObject</see> instance to copy the rtf to.</param>
        /// <returns>True if the rtf was successfully transferred; false otherwise.</returns>
        /// <remarks>
        ///     May throw an <see cref="RtfReaderException">RtfReaderException</see> if the extracted rtf does not contain valid
        ///     rtf data. May also throw com exceptions, or an OutOfMemory exception.
        /// </remarks>
        public bool ToDataObject(IDataObject source, IDataObject destination)
        {
            Check.NotNull(source, "rtfBytes");
            Check.NotNull(destination, "destination");

            using (var safeHGlobal = DataObjectTohGlobal(destination))
            {
                return HGlobalToDataObject(safeHGlobal, destination);
            }
        }

        /// <summary>
        ///     Copies the rtf contained in the supplied string into the supplied <see cref="IDataObject">IO.IDataObject</see>.
        /// </summary>
        /// <param name="source">An <see cref="TextRange">TextRange</see> containing the formatted rtf to copy.</param>
        /// <param name="destination">An <see cref="IDataObject">IDataObject</see> instance to copy the rtf to.</param>
        /// <returns>True if the rtf was successfully transferred; false otherwise.</returns>
        /// <remarks>
        ///     May throw an <see cref="RtfReaderException">RtfReaderException</see> if the extracted rtf does not contain valid
        ///     rtf data. May also throw com exceptions, or an OutOfMemory exception.
        /// </remarks>
        public static bool ToDataObject(TextRange source, IDataObject destination)
        {
            Check.NotNull(source, "source");
            Check.NotNull(destination, "destination");

            if (source.End == source.Start)
            {
                return false;
            }

            return true;
        }


        /// <summary>
        ///     Converts the supplied rtf contained in rtfBytes to a String.
        /// </summary>
        /// <param name="rtfBytes">An array of <see cref="byte">Bytes</see> containing the ASCII formatted rtf to extract.</param>
        /// <returns>A String containing the extracted rtf.</returns>
        public string ToString(byte[] rtfBytes)
        {
            Check.NotNull(rtfBytes, "rtfBytes");
            if (rtfBytes.Length == 0)
            {
                return string.Empty;
            }

            InvalidRtfGuard(rtfBytes);
            return Encoding.ASCII.GetString(rtfBytes, index: 0, count: rtfBytes.Length);
        }

        /// <summary>
        ///     Extracts the rtf from the supplied <see cref="IDataObject">IDataObject</see>.
        /// </summary>
        /// <param name="source">An <see cref="IDataObject">IDataObject</see> instance containing the rtf to copy.</param>
        /// <returns>A String containing the extracted rtf.</returns>
        public string ToString(IDataObject source)
        {
            Check.NotNull(source, "source");

            using (var safeHGlobal = DataObjectTohGlobal(source))
            {
                if (safeHGlobal.IsInvalid)
                {
                    return null;
                }

                safeHGlobal.Lock();
                try
                {
                    return Marshal.PtrToStringAnsi(safeHGlobal.LockedData);
                }
                finally
                {
                    safeHGlobal.Unlock();
                }


            }
        }

        /// <summary>
        ///     Returns the rtf from the supplied fileName.
        /// </summary>
        /// <param name="fileName">A string representing the File to load the rtf from.</param>
        /// <returns>A String containing the extracted rtf.</returns>

        public string ToString(string fileName)
        {
            Check.NotNull(fileName, "fileName");
            using (var fileStream = new FileStream(fileName, FileMode.Open))
            {
                InvalidRtfGuard(fileStream);
                fileStream.Position = 0;
                return new StreamReader(fileStream).ReadToEnd();
            }
        }


        /// <summary>
        ///     Determines if the passed rtf <see cref="byte">Byte</see> array contains a valid Rtf header.
        ///     No other validation is done on the array other than checking for \rtf followed by a version number (currently only
        ///     1).
        /// </summary>
        /// <param name="rtf">A <see cref="byte">Byte</see> array containing the formatted rtf.</param>
        /// <returns>True if the Rtf contained in the array is valid; false otherwise.</returns>

        public bool IsRichText(byte[] rtf)
        {
            if (rtf == null)
                throw new ArgumentNullException(nameof(rtf));

            var progress = 0;
            var isValid = false;
            var exitLoop = false;
            for (var i = 0; i < rtf.Length; i++)
            {
                switch (rtf[i])
                {
                    case AsciiCodes.CharacterReturn:
                    case AsciiCodes.Linefeed:
                    case AsciiCodes.Space:
                        break;
                    case AsciiCodes.OpeningCurlyBrace:
                        progress++;
                        break;
                    case AsciiCodes.Backslash:
                        if (progress == 1)
                        {
                            if (i + 5 <= rtf.Length &&
                                AsciiCodes.ToLower(rtf[i + 1]) == AsciiCodes.LowercaseR &&
                                AsciiCodes.ToLower(rtf[i + 2]) == AsciiCodes.LowercaseT &&
                                AsciiCodes.ToLower(rtf[i + 3]) == AsciiCodes.LowercaseF)
                            {
                                isValid = AsciiCodes.IsNumeric(rtf[i + 4]);
                            }
                        }

                        exitLoop = true;
                        break;
                    default:
                        exitLoop = true;
                        break;
                }

                if (exitLoop)
                    break;
            }

            return isValid;
        }

        /// <summary>
        ///     Determines if the passed rtf <see cref="string">String</see> contains a valid Rtf header.
        ///     No other validation is done on the array other than checking for {\rtf followed by a version number (currently only
        ///     1).
        /// </summary>
        /// <param name="rtf">A <see cref="string">String</see> containing the formatted rtf.</param>
        /// <returns>True if the Rtf contained in the string is valid; false otherwise.</returns>

        public static bool IsRichText(string rtf)
        {
            if (string.IsNullOrEmpty(rtf))
            {
                return false;
            }

            var progress = 0;
            var isValid = false;
            var characterReturnChar = Convert.ToChar(AsciiCodes.CharacterReturn);
            var linefeedChar = Convert.ToChar(AsciiCodes.Linefeed);

            for (var i = 0; i < rtf.Length; i++)
            {
                if ((rtf[i] == characterReturnChar) || (rtf[i] == linefeedChar) || (rtf[i] == ' ')) // AsciiCodes.Space
                { }
                else if (rtf[i] == '{') //AsciiCodes.OpeningCurlyBrace
                {
                    progress++;
                }
                else if (rtf[i] == '\\') //AsciiCodes.Backslash
                {
                    if (progress == 1)
                    {
                        if (i + 5 >= rtf.Length)
                        {
                            break;
                        }

                        isValid = rtf.Substring(i + 1, length: 3).ToUpperInvariant() == "RTF" &&
                                  int.TryParse(rtf[i + 4].ToString(), out _);
                    }

                    break;
                }
                else
                {
                    break;
                }
            }

            return isValid;
        }


        //Internal members


        private void InvalidRtfGuard(byte[] bytes)
        {
            if (!IsRichText(bytes))
            {
                throw new RtfReaderException("The rtf data passed is not valid.");
            }
        }

        private void InvalidRtfGuard(string rtf)
        {
            if (!IsRichText(rtf))
            {
                throw new RtfReaderException("The rtf data passed is not valid.");
            }
        }


        private void InvalidRtfGuard(FileStream fileStream)
        {
            //Allow standard io Exceptions to be thrown

            var bytes = new byte[11];
            if (fileStream.Read(bytes, offset: 0, count: 10) <= 0)
                return;

            if (!IsRichText(bytes))
                throw new RtfReaderException("The rtf data passed is not valid.");
        }


        //Helper members


        private FORMATETC CreateFormatEtc()
        {
            var formatetc = new FORMATETC
            {
                cfFormat = RichTextFormat,
                dwAspect = DVASPECT.DVASPECT_CONTENT,
                lindex = -1,
                tymed = TYMED.TYMED_HGLOBAL,
                ptd = IntPtr.Zero
            };
            return formatetc;
        }

        internal STGMEDIUM StreamToStgMedium(Stream stream)
        {

            Check.NotNull(stream, "stream");
 
            var rtfBytes = ToBytes(stream, false);
 
            using (var safeHGlobal = new SafeHGlobalHandle(Marshal.AllocHGlobal(rtfBytes.Length)))
            {
                if (safeHGlobal.Lock())
                {
                    Marshal.Copy(rtfBytes, startIndex: 0, destination: safeHGlobal.LockedData, length: rtfBytes.Length);
                    safeHGlobal.Unlock();
                    var stgMedium = new STGMEDIUM();

                    stgMedium.tymed = TYMED.TYMED_HGLOBAL;
                    stgMedium.unionmember = safeHGlobal.DangerousGetHandle();

                    //If SetData did not throw an exception then it should have 
                    // freed the HGLOBAL contained in the SafeHGlobalHandle
                    //So flag that the safeHGlobal is invalid so we do not free it.
                    safeHGlobal.SetHandleAsInvalid();
                    return stgMedium;
                }
            }

            return default;
        }

        [SuppressMessage("Microsoft.Reliability", "CA2001:AvoidCallingProblematicMethods",
            MessageId = "System.Runtime.InteropServices.SafeHandle.DangerousGetHandle")]
        private bool HGlobalToDataObject(SafeHGlobalHandle handle, IDataObject destination)
        {
            if (handle.IsInvalid)
            {
                return false;
            }

            var formatEtc = CreateFormatEtc();
            var stgMedium = new STGMEDIUM();

            stgMedium.tymed = formatEtc.tymed;
            stgMedium.unionmember = handle.DangerousGetHandle();
            destination.SetData(ref formatEtc, ref stgMedium, release: true);

            //If SetData did not throw an exception then it should have 
            // freed the HGLOBAL contained in the SafeHGlobalHandle
            //So flag that the safeHGlobal is invalid so we do not free it.
            handle.SetHandleAsInvalid();
            return true;
        }


        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
        [SecurityCritical]
        private SafeHGlobalHandle DataObjectTohGlobal(IDataObject data)
        {
            Check.NotNull(data, "data");

            var formatEtc = CreateFormatEtc();

            // Do we have RichText on the clipboard
            if (0 != data.QueryGetData(ref formatEtc))
            {
                return null;
            }

            data.GetData(ref formatEtc, out var stgMedium); //Raises Com Errors
            var safeHandle = new SafeHGlobalHandle(stgMedium.unionmember);

            if (safeHandle.IsInvalid || (stgMedium.tymed != TYMED.TYMED_HGLOBAL))
            {
                safeHandle.SetHandleAsInvalid(); //Cannot free if tymed is invalid type (should never get this)
                throw new ArgumentException("Data is null or not available in the requested format");
            }

            return safeHandle;
        }
    }
}