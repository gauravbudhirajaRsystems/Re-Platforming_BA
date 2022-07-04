using System;// © Copyright 2018 Levit & James, Inc.

using System.ComponentModel;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;

// ReSharper disable InconsistentNaming

namespace LevitJames.MSOffice.Internal
{
    [ComImport]
    [TypeLibType(flags: 0x10D0)]
    [DefaultMember("Name")]
    [Guid("00020970-0000-0000-C000-000000000046")]
    internal interface WordApplication11 //_VtblGap7_194
    {
        //7 to get past the IDispatch members
        //The ShowWindowsInTaskbar member is member number 201 so that's 201-7 gives us 194 hence _VtblGap7_194 
        //Note: REMEMBER that read/write properties take 2 vtable slots.

        [MethodImpl(MethodCodeType = MethodCodeType.Runtime)]
        [EditorBrowsable(EditorBrowsableState.Never)]
        void _VtblGap7_6();

        //  <DispId(&H3E8)> 
        //  ReadOnly Property Application() As <MarshalAs(UnmanagedType.Interface)> Application
        //  <DispId(&H3E9)> ReadOnly Property Creator() As Integer
        //  <DispId(&H3EA)> ReadOnly Property Parent() As <MarshalAs(UnmanagedType.IDispatch)> Object
        //  <DispId(0)> Default ReadOnly Property Name() As <MarshalAs(UnmanagedType.BStr)> String
        //  <DispId(6)> ReadOnly Property Documents() As <MarshalAs(UnmanagedType.Interface)> Documents
        //  <DispId(2)> ReadOnly Property Windows() As <MarshalAs(UnmanagedType.Interface)> Windows


        //  <DispId(3)> ReadOnly Property ActiveDocument() As <MarshalAs(UnmanagedType.Interface)> Document
        [DispId(dispId: 3)]
        [PreserveSig]
        int ActiveDocument([MarshalAs(UnmanagedType.Interface)] ref Document retVal);

        //  <DispId(4)> ReadOnly Property ActiveWindow() As <MarshalAs(UnmanagedType.Interface)> Window
        [DispId(dispId: 4)]
        [PreserveSig]
        int ActiveWindow([MarshalAs(UnmanagedType.Interface)] ref Window retVal);

        //Calc the gap between  ActiveWindow and UserName
        //UserName offset (52) -ActiveWindow offset (14) = (52-(14+1)) = 37
        [MethodImpl(MethodCodeType = MethodCodeType.Runtime)]
        [EditorBrowsable(EditorBrowsableState.Never)]
        void _VtblGap_37(); //UserName

        [DispId(0x34)]
        [PreserveSig]
        int UserName([MarshalAs(UnmanagedType.BStr)] ref string retVal);

        //Calc the gap between  UserName and Organizer Copy
        //OrganizerCopy offset (134) -UserName offset (52) = (134-(52+1)) = 81
        [MethodImpl(MethodCodeType = MethodCodeType.Runtime)]
        [EditorBrowsable(EditorBrowsableState.Never)]
        void _VtblGap_81();

        //  <DispId(5)> 
        //  ReadOnly Property Selection() As <MarshalAs(UnmanagedType.Interface)> Selection
        //  <DispId(1)> 
        //  ReadOnly Property WordBasic() As <MarshalAs(UnmanagedType.IDispatch)> Object
        //  <DispId(7)> 
        //  ReadOnly Property RecentFiles() As <MarshalAs(UnmanagedType.Interface)> RecentFiles
        //  <DispId(8)> 
        //  ReadOnly Property NormalTemplate() As <MarshalAs(UnmanagedType.Interface)> Template
        //  <DispId(9)> 
        //  ReadOnly Property System() As <MarshalAs(UnmanagedType.Interface)> System
        //  <DispId(10)> 
        //  ReadOnly Property AutoCorrect() As <MarshalAs(UnmanagedType.Interface)> AutoCorrect
        //  <DispId(11)> 
        //  ReadOnly Property FontNames() As <MarshalAs(UnmanagedType.Interface)> FontNames
        //  <DispId(12)> 
        //  ReadOnly Property LandscapeFontNames() As <MarshalAs(UnmanagedType.Interface)> FontNames
        //  <DispId(13)> 
        //  ReadOnly Property PortraitFontNames() As <MarshalAs(UnmanagedType.Interface)> FontNames
        //  <DispId(14)> 
        //  ReadOnly Property Languages() As <MarshalAs(UnmanagedType.Interface)> Languages
        //  <DispId(15)> 
        //  ReadOnly Property Assistant() As <MarshalAs(UnmanagedType.Interface)> Assistant
        //  <DispId(&H10)> 
        //  ReadOnly Property Browser() As <MarshalAs(UnmanagedType.Interface)> Browser
        //  <DispId(&H11)> 
        //  ReadOnly Property FileConverters() As <MarshalAs(UnmanagedType.Interface)> FileConverters
        //  <DispId(&H12)> 
        //  ReadOnly Property MailingLabel() As <MarshalAs(UnmanagedType.Interface)> MailingLabel
        //  <DispId(&H13)> 
        //  ReadOnly Property Dialogs() As <MarshalAs(UnmanagedType.Interface)> Dialogs
        //  <DispId(20)> 
        //  ReadOnly Property CaptionLabels() As <MarshalAs(UnmanagedType.Interface)> CaptionLabels
        //  <DispId(&H15)> 
        //  ReadOnly Property AutoCaptions() As <MarshalAs(UnmanagedType.Interface)> AutoCaptions
        //  <DispId(&H16)> 
        //  ReadOnly Property AddIns() As <MarshalAs(UnmanagedType.Interface)> AddIns
        //  <DispId(&H17)> 
        //  Property Visible() As Boolean
        //  <DispId(&H18)> 
        //  ReadOnly Property Version() As <MarshalAs(UnmanagedType.BStr)> String
        //  <DispId(&H1A)> 
        //  Property ScreenUpdating() As Boolean
        //  <DispId(&H1B)> 
        //  Property PrintPreview() As Boolean
        //  <DispId(&H1C)> 
        //  ReadOnly Property Tasks() As <MarshalAs(UnmanagedType.Interface)> Tasks
        //  <DispId(&H1D)> 
        //  Property DisplayStatusBar() As Boolean
        //  <DispId(30)> 
        //  ReadOnly Property SpecialMode() As Boolean
        //  <DispId(&H21)> 
        //  ReadOnly Property UsableWidth() As Integer
        //  <DispId(&H22)> 
        //  ReadOnly Property UsableHeight() As Integer
        //  <DispId(&H24)> 
        //  ReadOnly Property MathCoprocessorAvailable() As Boolean
        //  <DispId(&H25)> 
        //  ReadOnly Property MouseAvailable() As Boolean
        //  <DispId(&H2E)> 
        //  ReadOnly Property International(ByVal Index As WdInternationalIndex) As <MarshalAs(UnmanagedType.Struct)> Object
        //  <DispId(&H2F)> 
        //  ReadOnly Property Build() As <MarshalAs(UnmanagedType.BStr)> String
        //  <DispId(&H30)> 
        //  ReadOnly Property CapsLock() As Boolean
        //  <DispId(&H31)> 
        //  ReadOnly Property NumLock() As Boolean
        //  <DispId(&H34)> 
        //  Property UserName() As <MarshalAs(UnmanagedType.BStr)> String
        //  <DispId(&H35)> 
        //  Property UserInitials() As <MarshalAs(UnmanagedType.BStr)> String
        //  <DispId(&H36)> 
        //  Property UserAddress() As <MarshalAs(UnmanagedType.BStr)> String
        //  <DispId(&H37)> 
        //  ReadOnly Property MacroContainer() As <MarshalAs(UnmanagedType.IDispatch)> Object
        //  <DispId(&H38)> 
        //  Property DisplayRecentFiles() As Boolean
        //  <DispId(&H39)> 
        //  ReadOnly Property CommandBars() As <MarshalAs(UnmanagedType.Interface)> CommandBars
        //  <DispId(&H3B)> 
        //  ReadOnly Property SynonymInfo(ByVal Word As String, ByRef LanguageID As Object) As <MarshalAs(UnmanagedType.Interface)> SynonymInfo
        //  <DispId(&H3D)> 
        //  ReadOnly Property VBE() As <MarshalAs(UnmanagedType.Interface)> VBE
        //  <DispId(&H40)> 
        //  Property DefaultSaveFormat() As <MarshalAs(UnmanagedType.BStr)> String
        //  <DispId(&H41)> 
        //  ReadOnly Property ListGalleries() As <MarshalAs(UnmanagedType.Interface)> ListGalleries
        //  <DispId(&H42)> 
        //  Property ActivePrinter() As <MarshalAs(UnmanagedType.BStr)> String
        //  <DispId(&H43)> 
        //  ReadOnly Property Templates() As <MarshalAs(UnmanagedType.Interface)> Templates
        //  <DispId(&H44)> 
        //  Property CustomizationContext() As <MarshalAs(UnmanagedType.IDispatch)> Object
        //  <DispId(&H45)> 
        //  ReadOnly Property KeyBindings() As <MarshalAs(UnmanagedType.Interface)> KeyBindings
        //  <DispId(70)> 
        //  ReadOnly Property KeysBoundTo(ByVal KeyCategory As WdKeyCategory, ByVal Command As String, ByRef CommandParameter As Object) As <MarshalAs(UnmanagedType.Interface)> KeysBoundTo
        //  <DispId(&H47)> 
        //  ReadOnly Property FindKey(ByVal KeyCode As Integer, ByRef KeyCode2 As Object) As <MarshalAs(UnmanagedType.Interface)> KeyBinding
        //  <DispId(80)> 
        //  Property Caption() As <MarshalAs(UnmanagedType.BStr)> String
        //  <DispId(&H51)> 
        //  ReadOnly Property Path() As <MarshalAs(UnmanagedType.BStr)> String
        //  <DispId(&H52)> 
        //  Property DisplayScrollBars() As Boolean
        //  <DispId(&H53)> 
        //  Property StartupPath() As <MarshalAs(UnmanagedType.BStr)> String
        //  <DispId(&H55)> 
        //  ReadOnly Property BackgroundSavingStatus() As Integer
        //  <DispId(&H56)> 
        //  ReadOnly Property BackgroundPrintingStatus() As Integer
        //  <DispId(&H57)> 
        //  Property Left() As Integer
        //  <DispId(&H58)> 
        //  Property Top() As Integer
        //  <DispId(&H59)> 
        //  Property Width() As Integer
        //  <DispId(90)> 
        //  Property Height() As Integer
        //  <DispId(&H5B)> 
        //  Property WindowState() As WdWindowState
        //  <DispId(&H5C)> 
        //  Property DisplayAutoCompleteTips() As Boolean
        //  <DispId(&H5D)> 
        //  ReadOnly Property Options() As <MarshalAs(UnmanagedType.Interface)> Options
        //  <DispId(&H5E)> 
        //  Property DisplayAlerts() As WdAlertLevel
        //  <DispId(&H5F)> 
        //  ReadOnly Property CustomDictionaries() As <MarshalAs(UnmanagedType.Interface)> Dictionaries
        //  <DispId(&H60)> 
        //  ReadOnly Property PathSeparator() As <MarshalAs(UnmanagedType.BStr)> String
        //  <DispId(&H61)> 
        //  WriteOnly Property StatusBar() As String
        //  <DispId(&H62)> 
        //  ReadOnly Property MAPIAvailable() As Boolean
        //  <DispId(&H63)> 
        //  Property DisplayScreenTips() As Boolean
        //  <DispId(100)> 
        //  Property EnableCancelKey() As WdEnableCancelKey
        //  <DispId(&H65)> 
        //  ReadOnly Property UserControl() As Boolean
        //  <DispId(&H67)> 
        //  ReadOnly Property FileSearch() As <MarshalAs(UnmanagedType.Interface)> FileSearch
        //  <DispId(&H68)> 
        //  ReadOnly Property MailSystem() As WdMailSystem
        //  <DispId(&H69)> 
        //  Property DefaultTableSeparator() As <MarshalAs(UnmanagedType.BStr)> String
        //  <DispId(&H6A)> 
        //  Property ShowVisualBasicEditor() As Boolean
        //  <DispId(&H6C)> 
        //  Property BrowseExtraFileTypes() As <MarshalAs(UnmanagedType.BStr)> String
        //  <DispId(&H6D)> 
        //  ReadOnly Property IsObjectValid(ByVal [Object] As Object) As Boolean
        //  <DispId(110)> 
        //  ReadOnly Property HangulHanjaDictionaries() As <MarshalAs(UnmanagedType.Interface)> HangulHanjaConversionDictionaries
        //  <DispId(&H15C)> 
        //  ReadOnly Property MailMessage() As <MarshalAs(UnmanagedType.Interface)> MailMessage
        //  <DispId(&H182)> 
        //  ReadOnly Property FocusInMailHeader() As Boolean
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H451)> 
        //Sub Quit(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional SaveChanges As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional OriginalFormat As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional RouteDocument As Object)
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H12D)> 
        //  Sub ScreenRefresh()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), TypeLibFunc(CShort(&H40)), DispId(&H12E)> 
        //Sub PrintOutOld(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Background As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Append As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Range As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional OutputFileName As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional From As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional [To] As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Item As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Copies As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Pages As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PageType As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintToFile As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Collate As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional FileName As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ActivePrinterMacGX As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ManualDuplexPrint As Object)
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H12F)> 
        //  Sub LookupNameProperties(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal Name As String)
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H130)> 
        //  Sub SubstituteFont(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal UnavailableFont As String, <[In](), MarshalAs(UnmanagedType.BStr)> ByVal SubstituteFont As String)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H131)> 
        //Function Repeat(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Times As Object) As Boolean
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(310)> 
        //  Sub DDEExecute(<[In]()> ByVal Channel As Integer, <[In](), MarshalAs(UnmanagedType.BStr)> ByVal Command As String)
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H137)> 
        //  Function DDEInitiate(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal App As String, <[In](), MarshalAs(UnmanagedType.BStr)> ByVal Topic As String) As Integer
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H138)> 
        //  Sub DDEPoke(<[In]()> ByVal Channel As Integer, <[In](), MarshalAs(UnmanagedType.BStr)> ByVal Item As String, <[In](), MarshalAs(UnmanagedType.BStr)> ByVal Data As String)
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H139)> 
        //  Function DDERequest(<[In]()> ByVal Channel As Integer, <[In](), MarshalAs(UnmanagedType.BStr)> ByVal Item As String) As <MarshalAs(UnmanagedType.BStr)> String
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H13A)> 
        //  Sub DDETerminate(<[In]()> ByVal Channel As Integer)
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H13B)> 
        //  Sub DDETerminateAll()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H13C)> 
        //Function BuildKeyCode(<[In]> ByVal Arg1 As WdKey, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Arg2 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Arg3 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Arg4 As Object) As Integer
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H13D)> 
        //Function KeyString(<[In]> ByVal KeyCode As Integer, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional KeyCode2 As Object) As <MarshalAs(UnmanagedType.BStr)> String
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [PreserveSig]
        [DispId(dispId: 0x13E)]
        int OrganizerCopy([In][MarshalAs(UnmanagedType.BStr)] string Source,
                          [In][MarshalAs(UnmanagedType.BStr)] string Destination,
                          [In][MarshalAs(UnmanagedType.BStr)] string Name, [In] WdOrganizerObject Object);

        //Calc the gap between  OrganizerCopy and ShowWindowsInTaskbar 
        //ShowWindowsInTaskbar offset (201) -OrganizerCopy offset (134) = (201-(134+1)) = 66
        [MethodImpl(MethodCodeType = MethodCodeType.Runtime)]
        [EditorBrowsable(EditorBrowsableState.Never)]
        void _VtblGap_66();

        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H13F)> 
        //  Sub OrganizerDelete(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal Source As String, <[In](), MarshalAs(UnmanagedType.BStr)> ByVal Name As String, <[In]()> ByVal [Object] As WdOrganizerObject)
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(320)> 
        //  Sub OrganizerRename(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal Source As String, <[In](), MarshalAs(UnmanagedType.BStr)> ByVal Name As String, <[In](), MarshalAs(UnmanagedType.BStr)> ByVal NewName As String, <[In]()> ByVal [Object] As WdOrganizerObject)
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H141)> 
        //  Sub AddAddress(<[In](), MarshalAs(UnmanagedType.SafeArray, SafeArraySubType:=VarEnum.VT_BSTR)> ByRef TagID As Array, <[In](), MarshalAs(UnmanagedType.SafeArray, SafeArraySubType:=VarEnum.VT_BSTR)> ByRef Value As Array)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H142)> 
        //Function GetAddress(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Name As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional AddressProperties As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional UseAutoText As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional DisplaySelectDialog As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional SelectDialog As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CheckNamesDialog As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional RecentAddressesChoice As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional UpdateRecentAddresses As Object) As <MarshalAs(UnmanagedType.BStr)> String
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H143)> 
        //  Function CheckGrammar(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal [String] As String) As Boolean
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H144)> 
        //Function CheckSpelling(<[In], MarshalAs(UnmanagedType.BStr)> ByVal Word As String, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional IgnoreUppercase As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional MainDictionary As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary2 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary3 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary4 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary5 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary6 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary7 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary8 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary9 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary10 As Object) As Boolean
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H146)> 
        //  Sub ResetIgnoreAll()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H147)> 
        //Function GetSpellingSuggestions(<[In], MarshalAs(UnmanagedType.BStr)> ByVal Word As String, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional IgnoreUppercase As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional MainDictionary As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional SuggestionMode As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary2 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary3 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary4 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary5 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary6 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary7 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary8 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary9 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary10 As Object) As <MarshalAs(UnmanagedType.Interface)> SpellingSuggestions
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H148)> 
        //  Sub GoBack()
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H149)> 
        //  Sub Help(<[In](), MarshalAs(UnmanagedType.Struct)> ByRef HelpType As Object)
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(330)> 
        //  Sub AutomaticChange()
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H14B)> 
        //  Sub ShowMe()
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H14C)> 
        //  Sub HelpTool()
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H159)> 
        //  Function NewWindow() As <MarshalAs(UnmanagedType.Interface)> Window
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H15A)> 
        //  Sub ListCommands(<[In]()> ByVal ListAllCommands As Boolean)
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H15D)> 
        //  Sub ShowClipboard()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(350)> 
        //Sub OnTime(<[In], MarshalAs(UnmanagedType.Struct)> ByRef [When] As Object, <[In], MarshalAs(UnmanagedType.BStr)> ByVal Name As String, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Tolerance As Object)
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H15F)> 
        //  Sub NextLetter()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H161)> 
        //Function MountVolume(<[In], MarshalAs(UnmanagedType.BStr)> ByVal Zone As String, <[In], MarshalAs(UnmanagedType.BStr)> ByVal Server As String, <[In], MarshalAs(UnmanagedType.BStr)> ByVal Volume As String, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional User As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional UserPassword As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional VolumePassword As Object) As Short
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H162)> 
        //  Function CleanString(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal [String] As String) As <MarshalAs(UnmanagedType.BStr)> String
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H164)> 
        //  Sub SendFax()
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H165)> 
        //  Sub ChangeFileOpenDirectory(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal Path As String)
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), TypeLibFunc(CShort(&H40)), DispId(&H166)> 
        //  Sub RunOld(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal MacroName As String)
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H167)> 
        //  Sub GoForward()
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(360)> 
        //  Sub Move(<[In]()> ByVal Left As Integer, <[In]()> ByVal Top As Integer)
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H169)> 
        //  Sub Resize(<[In]()> ByVal Width As Integer, <[In]()> ByVal Height As Integer)
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(370)> 
        //  Function InchesToPoints(<[In]()> ByVal Inches As Single) As Single
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H173)> 
        //  Function CentimetersToPoints(<[In]()> ByVal Centimeters As Single) As Single
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H174)> 
        //  Function MillimetersToPoints(<[In]()> ByVal Millimeters As Single) As Single
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H175)> 
        //  Function PicasToPoints(<[In]()> ByVal Picas As Single) As Single
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H176)> 
        //  Function LinesToPoints(<[In]()> ByVal Lines As Single) As Single
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(380)> 
        //  Function PointsToInches(<[In]()> ByVal Points As Single) As Single
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H17D)> 
        //  Function PointsToCentimeters(<[In]()> ByVal Points As Single) As Single
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H17E)> 
        //  Function PointsToMillimeters(<[In]()> ByVal Points As Single) As Single
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H17F)> 
        //  Function PointsToPicas(<[In]()> ByVal Points As Single) As Single
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H180)> 
        //  Function PointsToLines(<[In]()> ByVal Points As Single) As Single
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H181)> 
        //  Sub Activate()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H183)> 
        //Function PointsToPixels(<[In]> ByVal Points As Single, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional fVertical As Object) As Single
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H184)> 
        //Function PixelsToPoints(<[In]> ByVal Pixels As Single, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional fVertical As Object) As Single
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(400)> 
        //  Sub KeyboardLatin()
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H191)> 
        //  Sub KeyboardBidi()
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H192)> 
        //  Sub ToggleKeyboard()
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H1BE)> 
        //  Function Keyboard(<[In]()> Optional ByVal LangId As Integer = 0) As Integer
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H194)> 
        //  Function ProductCode() As <MarshalAs(UnmanagedType.BStr)> String
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H195)> 
        //  Function DefaultWebOptions() As <MarshalAs(UnmanagedType.Interface)> DefaultWebOptions
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), TypeLibFunc(CShort(&H40)), DispId(&H197)> 
        //  Sub DiscussionSupport(<[In](), MarshalAs(UnmanagedType.Struct)> ByRef Range As Object, <[In](), MarshalAs(UnmanagedType.Struct)> ByRef cid As Object, <[In](), MarshalAs(UnmanagedType.Struct)> ByRef piCSE As Object)
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H19E)> 
        //  Sub SetDefaultTheme(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal Name As String, <[In]()> ByVal DocumentType As WdDocumentMedium)
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H1A0)> 
        //  Function GetDefaultTheme(<[In]()> ByVal DocumentType As WdDocumentMedium) As <MarshalAs(UnmanagedType.BStr)> String
        //  <DispId(&H185)> 
        //  ReadOnly Property EmailOptions() As <MarshalAs(UnmanagedType.Interface)> EmailOptions
        //  <DispId(&H187)> 
        //  ReadOnly Property Language() As MsoLanguageID
        //  <DispId(&H6F)> 
        //  ReadOnly Property COMAddIns() As <MarshalAs(UnmanagedType.Interface)> COMAddIns
        //  <DispId(&H70)> 
        //  Property CheckLanguage() As Boolean
        //  <DispId(&H193)> 
        //  ReadOnly Property LanguageSettings() As <MarshalAs(UnmanagedType.Interface)> LanguageSettings
        //  <DispId(&H196)> 
        //  ReadOnly Property Dummy1() As Boolean
        //  <DispId(&H199)> 
        //  ReadOnly Property AnswerWizard() As <MarshalAs(UnmanagedType.Interface)> AnswerWizard
        //  <DispId(&H1BF)> 
        //  Property FeatureInstall() As MsoFeatureInstall
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), TypeLibFunc(CShort(&H40)), DispId(&H1BC)> 
        //Sub PrintOut2000(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Background As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Append As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Range As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional OutputFileName As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional From As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional [To] As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Item As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Copies As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Pages As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PageType As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintToFile As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Collate As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional FileName As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ActivePrinterMacGX As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ManualDuplexPrint As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintZoomColumn As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintZoomRow As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintZoomPaperWidth As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintZoomPaperHeight As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H1BD)> 
        //Function Run(<[In], MarshalAs(UnmanagedType.BStr)> ByVal MacroName As String, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg1 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg2 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg3 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg4 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg5 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg6 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg7 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg8 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg9 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg10 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg11 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg12 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg13 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg14 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg15 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg16 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg17 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg18 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg19 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg20 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg21 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg22 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg23 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg24 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg25 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg26 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg27 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg28 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg29 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional varg30 As Object) As <MarshalAs(UnmanagedType.Struct)> Object
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H1C0)> 
        //Sub PrintOut(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Background As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Append As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Range As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional OutputFileName As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional From As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional [To] As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Item As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Copies As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Pages As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PageType As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintToFile As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Collate As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional FileName As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ActivePrinterMacGX As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ManualDuplexPrint As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintZoomColumn As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintZoomRow As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintZoomPaperWidth As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintZoomPaperHeight As Object)
        //  <DispId(&H1C1)> 
        //  Property AutomationSecurity() As MsoAutomationSecurity
        //  <DispId(450)> 
        //  ReadOnly Property FileDialog(ByVal FileDialogType As MsoFileDialogType) As <MarshalAs(UnmanagedType.Interface)> FileDialog
        //  <DispId(&H1C3)> 
        //  Property EmailTemplate() As <MarshalAs(UnmanagedType.BStr)> String

        //Changed from a property Get/Let to a Get method
        // ''<DispId(&H1C4)> 
        // ''Property ShowWindowsInTaskbar() As Boolean
        [DispId(dispId: 0x1C4)]
        [PreserveSig]
        int ShowWindowsInTaskbar(ref bool retVal);

        [MethodImpl(MethodCodeType = MethodCodeType.Runtime)]
        [EditorBrowsable(EditorBrowsableState.Never)]
        void _VtblGap_38();

        [DispId(dispId: 0x1E6)]
        UndoRecord UndoRecord { get; }
    }
}
