// © Copyright 2018 Levit & James, Inc.

using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace LevitJames.TextServices
{
    /// <summary>
    ///     Marshals an IDataObject as a reference inside Com Variant
    /// </summary>

    internal sealed class TextRangeVariantMarshaler : ICustomMarshaler
    {
        private static TextRangeVariantMarshaler _instance;
        private static readonly int OffsetToTagVariantExObjRef = 8 + IntPtr.Size * 2;
        private static readonly int OffsetToTagVariantExObjVariantWrapper = 8 + IntPtr.Size * 3;


        private TextRangeVariantMarshaler() { }


        public void CleanUpManagedData(object managedObj) { }


        public void CleanUpNativeData(IntPtr pNativeData)
        {
            if (pNativeData != IntPtr.Zero)
            {
                if (Marshal.ReadIntPtr(pNativeData, OffsetToTagVariantExObjVariantWrapper) == IntPtr.Zero)
                {
                    //Paste/ 
                    var ptrUnk = Marshal.ReadIntPtr(pNativeData, OffsetToTagVariantExObjRef);
                    if (ptrUnk != IntPtr.Zero)
                    {
                        // Release the pointer to IDataObject we passed in
                        Marshal.Release(ptrUnk);
                    }
                }

                //Release the pointer we created for the VARIANT
                Marshal.FreeCoTaskMem(pNativeData);
            }
        }


        public int GetNativeDataSize()
        {
            return -1;
        }


        public IntPtr MarshalManagedToNative(object managedObject)
        {
            if (managedObject is TextRangeDataObject trdo)
            {
                //Cut/Copy

                var ptr = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(TagVariantEx)));
                var tv = new TagVariantEx
                {
                    // ReSharper disable once BitwiseOperatorOnEnumWithoutFlags
                    vt = (short)(VarEnum.VT_UNKNOWN | VarEnum.VT_BYREF),
                    ppunkVal = ptr + OffsetToTagVariantExObjRef
                };

                var handle = GCHandle.Alloc(managedObject, GCHandleType.Normal);
                tv.objVariantWrapper = GCHandle.ToIntPtr(handle);

                Marshal.StructureToPtr(tv, ptr, fDeleteOld: false);
                return ptr;
            }

            //Paste
            if (managedObject is IDataObject ido)
            {
                var ptr = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(TagVariantEx)));
                var tv = new TagVariantEx
                {
                    vt = (short)VarEnum.VT_UNKNOWN,
                    ppunkVal = Marshal.GetIUnknownForObject(ido)
                };
                Marshal.StructureToPtr(tv, ptr, fDeleteOld: false);
                return ptr;
            }

            return IntPtr.Zero;
        }


        public object MarshalNativeToManaged(IntPtr pNativeData)
        {
            if (pNativeData != IntPtr.Zero)
            {
                //Cut/Copy
                var tv = (TagVariantEx)Marshal.PtrToStructure(pNativeData, typeof(TagVariantEx));
                if (tv.objRef != IntPtr.Zero)
                {
                    var ido = (IDataObject)Marshal.GetObjectForIUnknown(tv.objRef);
                    if (tv.objVariantWrapper != IntPtr.Zero)
                    {
                        var handle = GCHandle.FromIntPtr(tv.objVariantWrapper);
                        var target = (TextRangeDataObject)handle.Target;
                        handle.Free();
                        target.DataObject = ido;
                        return target;
                    }

                    //Marshal.Release(tv.objRef);
                }
            }

            return null;
        }


        //Must keep this signature with the cookie, .net framework requirement.
        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        [SuppressMessage("Microsoft.Usage", "CA1801:ReviewUnusedParameters", MessageId = "cookie")]
        public static ICustomMarshaler GetInstance(string cookie)
        {
            if (_instance == null)
            {
                _instance = new TextRangeVariantMarshaler();
            }

            return _instance;
        }


        //This is a VARIANT Structure with two extra pointers grafted onto the end
        //The first graft is the objRef. This memory location is then stored in the data1 field
        //to create a VarEnum.VT_BYREF value
        //The Second value is a pointer to a TextRangeVariant object 
        [SuppressMessage("Microsoft.Design", "CA1049:TypesThatOwnNativeResourcesShouldBeDisposable")]
        [Serializable]
        [StructLayout(LayoutKind.Sequential)]
        private struct TagVariantEx
        {
            [MarshalAs(UnmanagedType.I2)] public short vt; // // VARTYPE vt;

            //<FieldOffset(2), MarshalAs(UnmanagedType.I2)> Public wReserved1 As Short '; // WORD wReserved1;
            [MarshalAs(UnmanagedType.I4)] public readonly int wReserved2_3; //; // WORD wReserved2;

            //<FieldOffset(6), MarshalAs(UnmanagedType.I2)> Public wReserved3 As Short '; // WORD wReserved3; 
            public IntPtr ppunkVal;
            public IntPtr data2;
            public readonly IntPtr objRef;
            public IntPtr objVariantWrapper;
        }
    }
}