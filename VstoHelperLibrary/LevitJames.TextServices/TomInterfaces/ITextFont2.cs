using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace LevitJames.TextServices
{
    #region      ITextFont2 (Interface) 

    [ComImport, Guid("C241F5E3-7206-11D8-A2C7-00A0D1D6C6B3"),
     TypeLibType((TypeLibTypeFlags.FDispatchable | TypeLibTypeFlags.FNonExtensible | TypeLibTypeFlags.FDual))]
    internal interface ITextFont2
    {

        [DispId(0)]
        ITextFont2 Duplicate { get; set; }

        [return: MarshalAs(UnmanagedType.I4)]
        [DispId(0x301)]
        bool CanChange();

        [return: MarshalAs(UnmanagedType.I4)]
        [DispId(770)]
        bool IsEqual([In, MarshalAs(UnmanagedType.Interface)] ITextFont2 font);

        [DispId(0x303)]
        void Reset([In] int value);

        [DispId(0x304)]
        BuiltInStyles Style { get; set; }

        [DispId(0x305)]
        int AllCaps { get; set; } // Standard Boolean

        [DispId(0x306)]
        FontAnimation Animation { get; set; }

        [DispId(0x307)]
        int BackColor { get; set; }

        [DispId(0x308)]
        int Bold { get; set; } // Standard Boolean

        [DispId(0x309)]
        int Emboss { get; set; } // Standard Boolean

        [DispId(0x310)]
        int ForeColor { get; set; }

        [DispId(0x311)]
        int Hidden { get; set; } // Standard Boolean

        [DispId(0x312)]
        int Engrave { get; set; } // Standard Boolean

        [DispId(0x313)]
        int Italic { get; set; } // Standard Boolean

        [DispId(0x314)]
        float Kerning { get; set; }

        [DispId(0x315)]
        int LanguageID { get; set; }

        [DispId(790)]
        string Name { get; set; }

        [DispId(0x317)]
        int Outline { get; set; }

        [DispId(0x318)]
        float Position { get; set; }

        [DispId(0x319)]
        int Protected { get; set; } // Standard Boolean

        [DispId(800)]
        int Shadow { get; set; } // Standard Boolean

        [DispId(0x321)]
        float GetSize();

        [DispId(0x321), MethodImpl(MethodImplOptions.PreserveSig)]
        int SetSize(float newSize);

        [DispId(0x322)]
        int SmallCaps { get; set; } // Standard Boolean

        [DispId(0x323)]
        float Spacing { get; set; }

        [DispId(0x324)]
        int StrikeThrough { get; set; } // Standard Boolean

        [DispId(0x325)]
        int Subscript { get; set; } // Standard Boolean

        [DispId(0x326)]
        int Superscript { get; set; } // Standard Boolean

        [DispId(0x327)]
        FontUnderline Underline { get; set; }

        [DispId(0x328)]
        FontWeight Weight { get; set; }

        #endregion // ITextFont (Interface)

        [DispId(2)]
        int Count { get; }

        [DispId(0x329)]
        int AutoLigatures { get; set; }

        [DispId(810)]
        int AutospaceAlpha { get; set; }

        [DispId(0x32B)]
        int AutospaceNumeric { get; set; }

        [DispId(0x32C)]
        int AutospaceParens { get; set; }

        [DispId(0x32D)]
        int CharRep { get; set; }

        [DispId(0x32E)]
        int CompressionMode { get; set; }

        [DispId(0x32F)]
        int Cookie { get; set; }

        [DispId(0x330)]
        int DoubleStrike { get; set; }

        [DispId(0x331)]
        ITextFont2 Duplicate2 { get; set; }

        [DispId(0x332)]
        int LinkType { get; }

        [DispId(0x333)]
        int MathZone { get; set; }

        [DispId(820)]
        int ModWidthPairs { get; set; }

        [DispId(0x335)]
        int ModWidthSpace { get; set; }

        [DispId(0x336)]
        int OldNumbers { get; set; }

        [DispId(0x337)]
        int Overlapping { get; set; }

        [DispId(0x338)]
        int PositionSubSuper { get; set; }

        [DispId(0x339)]
        int Scaling { get; set; }

        [DispId(0x33A)]
        float SpaceExtension { get; set; }

        [DispId(0x33B)]
        int UnderlinePositionMode { get; set; }

        [DispId(0x340)]
        void GetEffects(out int pValue, out int pMask);

        [DispId(0x341)]
        void GetEffects2(out int pValue, out int pMask);

        [DispId(0x342)]
        int GetProperty([In] int Type);

        [DispId(0x343)]
        void GetPropertyInfo([In] int Index, out int pType, out int pValue);

        [DispId(0x344)]
        int IsEqual2([In, MarshalAs(UnmanagedType.Interface)] ITextFont2 pFont);

        [DispId(0x345)]
        void SetEffects([In] int Value, [In] int Mask);

        [DispId(0x346)]
        void SetEffects2([In] int Value, [In] int Mask);

        [DispId(0x347)]
        void SetProperty([In] int Type, [In] int Value);
    }


}
