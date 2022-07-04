//**************************************************
//** © 2018-2021 Litera Corp. All Rights Reserved.
//***************************************************

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using JetBrains.Annotations;
using LevitJames.Core;
using LevitJames.MSOffice.Internal;
using LevitJames.TextServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;

namespace LevitJames.MSOffice.MSWord
{
    /// <summary>
    ///     Represents a set of compiler extensions that are added to the Word and Office object models.
    /// </summary>
    /// <remarks>
    ///     The naming convention of Extension of methods is as follows:
    ///     <para>
    ///         If the Extension method is designed to replace an existing Word member then the Extension members
    ///         Name will be the same as the original Word member name.
    ///         However, the Extension member will end with the LJ suffix.
    ///     </para>
    ///     <para>
    ///         Common .Net methods that help make using the Word object model easier to use.
    ///         Such methods include ExistsLJ and TryGetItemLJ.
    ///     </para>
    ///     <para>
    ///         If an extension method requires a minimum version of Word to work the
    ///         Extension method name is suffixed with the version number of
    ///         Word required. For example a method that is only available in version 12 of word will be suffixed with 12 like
    ///         so: ExistsLJ12.
    ///     </para>
    ///     <para>
    ///         If you try to call an Extension method for a newer version of word than the version of
    ///         Word in use the extension method will simply return the default value for the type, for example null or false.
    ///         No Exception is ever thrown.
    ///     </para>
    ///     <para>
    ///         Any extension methods that add completely new functionality are always named with a LJ *Prefix*,
    ///         such as LJDialogs.
    ///     </para>
    /// </remarks>
    public static partial class Extensions
    {
        // ReSharper disable InconsistentNaming
        // ReSharper disable SuspiciousTypeConversion.Global
        private static WordBasic _wordBasic;

        //Word.Application'


        //Word.Application
        /// <summary>
        ///     True displays opened documents in the task bar, the default Single Document Interface (SDI). False
        ///     lists opened documents only in the Window menu, providing the appearance of a Multiple Document
        ///     Interface (MDI). Unlike the ShowWindowsInTaskBar function provided my Word this method does not throw
        ///     an exception. If there is an error the method will return a state of Unknown.
        /// </summary>
        /// <param name="source">A Word.Application instance.</param>
        /// <returns>A WindowsInTaskbar Enum value</returns>
        /// <remarks>This method always returns true for Word.Versions 15 and after.</remarks>
        public static WordBoolean ShowWindowsInTaskBarLJ([NotNull] this Application source)
        {
            Check.NotNull(source, nameof(source));

            if (source.VersionLJ() >= OfficeVersion.Office2013)
                return WordBoolean.True;

            if (!(source is WordApplication11 wordApp11))
                return 0;

            var retVal = false;
            if (wordApp11.ShowWindowsInTaskbar(ref retVal) == 0)
            {
                return retVal ? WordBoolean.True : WordBoolean.False;
            }

            return WordBoolean.Unknown;
        }


        public static string VersionForDisplayLJ(this Application @this)
        {
            //User could install x86 version in x64 folders?? perhapse....but ignoring.

            var hMso = NativeMethods.GetModuleHandle("MSO.DLL");
            string msoPath = null;
            if (hMso != IntPtr.Zero)
            {
                var buffer = new StringBuilder(260);
                //Mso dll is loaded so we can find the required paths using the module handle.
                NativeMethods.GetModuleFileName(hMso, buffer, buffer.Capacity);

                msoPath = buffer.ToString();
            }

            var path = @this.Path;
            // For testing
            // path = @"C:\Program Files (x86)\Microsoft Office\Office15"
            // path = @"C:\Program Files (x64)\Microsoft Office\Office15"
            var isX64 = !path.StartsWith(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), StringComparison.OrdinalIgnoreCase);
            if (string.IsNullOrEmpty(msoPath))
            {
                //out of process
                if (int.TryParse(@this.Version.Substring(0, 2), out var officeVersion))
                {
                    if (isX64)
                    {
                        // Not easy to get C:\Program Files (x64) under a 32 bit process so we have to resort to using the Application.Path
                        var progFilesPath = path.Substring(0, path.IndexOf(Path.DirectorySeparatorChar, 3));
                        msoPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles);
                        msoPath = msoPath.Replace(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), progFilesPath);
                    }
                    else
                    {
                        msoPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFilesX86);
                    }

                    msoPath = Path.Combine(msoPath, $@"Microsoft Shared\Office{officeVersion}\Mso.dll");
                }
            }

            if (!File.Exists(msoPath))
                return string.Empty;

            if (msoPath == null)
                return string.Empty;

            var msoFileVersion = FileVersionInfo.GetVersionInfo(msoPath);
            var wordPath = Path.Combine(@this.Path, "winword.exe");
            var wordFileVersion = FileVersionInfo.GetVersionInfo(wordPath);

            var sp = @this.VersionServicePackLJ();
            if (string.IsNullOrEmpty(sp) == false)
            {
                sp += " ";
            }

            var prodName = wordFileVersion.ProductName.Replace("Office", "Word").Replace("system", string.Empty).Trim();
            var wordProdVer = wordFileVersion.ProductVersion;
            var msoProdVer = msoFileVersion.ProductVersion;

            return $"{prodName} ({wordProdVer}) {sp} MSO ({msoProdVer}) {(isX64 ? "64" : "32")}-bit";
        }


        /// <summary>
        ///     Returns a Word.Document object that represents the active document (the document with the focus).
        ///     Unlike the ActiveDocument function provided my Word this method does not throw
        ///     an exception. instead if there is an error the method will return a state of Unknown.
        /// </summary>
        /// <param name="source">A Word.Application instance.</param>
        /// <returns>A Word.Document if there is an active document, null otherwise</returns>
        
        public static Document ActiveDocumentLJ([NotNull] this Application source)
        {
            Check.NotNull(source, nameof(source));

            if (!(source is WordApplication11 wordApp11))
                return null;

            Document activeDocument = null;
            return wordApp11.ActiveDocument(ref activeDocument) == 0 ? activeDocument : null;
        }


        /// <summary>
        ///     Returns a Word.Document object that represents the active document (the document with the focus).
        ///     Unlike the ActiveDocument function provided my Word this method does not throw
        ///     an exception. instead if there is an error the method will return a state of Unknown.
        /// </summary>
        /// <param name="source">A Word.Application instance.</param>
        /// <returns>A Word.Document if there is an active document, null otherwise</returns>
        
        public static Window ActiveWindowLJ([NotNull] this Application source)
        {
            Check.NotNull(source, nameof(source));

            if (!(source is WordApplication11 wordApp11))
                return null;

            Window activeWindow = null;
            return wordApp11.ActiveWindow(ref activeWindow) == 0 ? activeWindow : null;
        }

        public static OfficeUITheme GetOfficeTheme([NotNull] this Application source)
        {
            var officeTheme = OfficeUITheme.None;
            var wordVersion = VersionLJ(source);

            if (wordVersion >= OfficeVersion.Office2007)
            {
                OfficeUITheme defultTheme;
                switch (wordVersion)
                {
                case OfficeVersion.Office2007:
                    defultTheme = OfficeUITheme.Blue;
                    break;
                case OfficeVersion.Office2010:
                    defultTheme = OfficeUITheme.Black;
                    break;
                default:
                    defultTheme = OfficeUITheme.Blue;
                    break;
                }

                var keyPath = string.Format(CultureInfo.InvariantCulture,
                                            "HKEY_CURRENT_USER\\Software\\Microsoft\\Office\\{0}.0\\Common",
                                            Convert.ToInt32(wordVersion));

                var keyName = wordVersion >= OfficeVersion.Office2013 ? "UI Theme" : "Theme";
                officeTheme = (OfficeUITheme)Registry.GetValue(keyPath, keyName, defultTheme);

                switch (wordVersion)
                {
                case OfficeVersion.Office2013:
                    officeTheme = (OfficeUITheme)((int)officeTheme + 4);
                    break;
                case OfficeVersion.Office2016:
                    //2016 
                    //Colorful =0
                    //DarkGray =3
                    //Black = 4
                    //White = 5
                    //SystemSetting = 6

                    if (officeTheme == 0)
                        officeTheme = OfficeUITheme.Colorful;
                    else if ((int)officeTheme == 3)
                        officeTheme = OfficeUITheme.DarkGrey;
                    else if ((int)officeTheme == 4)
                        officeTheme = OfficeUITheme.Black;
                    else if ((int)officeTheme == 5)
                        officeTheme = OfficeUITheme.White;
                    else if ((int) officeTheme == 6)
                    {
                            //Use System setting
                            var val = (int)Registry.GetValue("HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize", "AppsUseLightTheme", 1);
                            if (val == 1)
                                officeTheme = OfficeUITheme.White;
                            else
                                officeTheme = OfficeUITheme.Black;
                    }
                        


                    break;
                }
            }
           
            return officeTheme;
        }

        /// <summary>
        ///     Resolves the following word objects to a Word.Window. The objects resolved are Window,Document,Application &amp;
        ///     Pane. If the wordObject passed is null or Application then the Application.ActiveWindow is returned.
        /// </summary>
        /// <param name="wordObject">A Word instance to resolve.</param>
        /// <returns>A Window instance or null if the object passed cannot be resolved.</returns>
        public static Window WindowFromObject([NotNull] this Application source, object wordObject)
        {
            if ((wordObject is Window wordWindow))
                return wordWindow;

            if (wordObject == null || wordObject is Application)
                return source.ActiveWindowLJ();

            if (wordObject is Document wordDoc)
                return wordDoc.ActiveWindow;

            if (wordObject is Pane wordPane)
                return (Window) wordPane.Parent;

            return null;
        }


        /// <summary>
        ///     Provides a friendly name for the Word version by determining the service pack (SP) from the Build number.
        /// </summary>
        /// <param name="this"></param>
        /// <returns>
        ///     "none" if no service pack is applied, an empty string if the service pack cannot be determined, else returns a
        ///     string that representing the service pack applied, if an exact match is not found then the string is appended with
        ///     "(patched)"
        /// </returns>
        /// <remarks>This method will need updating when newer service packs and versions of Word are released.</remarks>
        public static string VersionServicePackLJ(this Application @this)
        {
            Check.NotNull(@this, "this");

            var buildNumber = int.Parse(@this.Build.Substring(@this.Build.LastIndexOf(".", StringComparison.Ordinal) + 1));

            var versionValue = Convert.ToInt32(@this.Version.Val());
            switch (versionValue)
            {
            case 11:
                // Word 2003 (11)
                // https://support.microsoft.com/kb/821549
                if (buildNumber < 5604)
                    return "RTM";
                if (buildNumber == 5604)
                    return string.Empty;
                if (buildNumber < 6359)
                    return "(patched)";
                if (buildNumber == 6359)
                    return "SP1";
                if (buildNumber < 7969)
                    return "SP1 (patched)";
                if (buildNumber == 7969)
                    return "SP2";
                if (buildNumber < 8173)
                    return "SP2 (patched)";
                if (buildNumber == 8173)
                    return "SP3";
                if (buildNumber > 8173)
                    return "SP3 (patched)";

                return string.Empty; // Never accessed

            case 12:
                // Word 2007 (12)
                // https://support.microsoft.com/kb/928116
                if (buildNumber < 4518)
                    return "(Beta)";
                if (buildNumber == 4518)
                    return "RTM";
                if (buildNumber < 6211)
                    return "(patched)";
                if (buildNumber == 6211)
                    return "SP1";
                if (buildNumber < 6425)
                    return "SP1 (patched)";
                if (buildNumber == 6425)
                    return "SP2";
                if (buildNumber < 6612)
                    return "SP2 (patched)";
                if (buildNumber == 6612)
                    return "SP3";
                if (buildNumber > 6612)
                    return "SP3 (patched)";

                return string.Empty; // Never accessed

            case 14:
                // Word 2010 (14)
                // https://support.microsoft.com/kb/2121559
                if (buildNumber < 4762)
                    return "(Beta)";
                if (buildNumber == 4762)
                    return "RTM";
                if (buildNumber < 6024)
                    return "(patched)";
                if (buildNumber == 6024)
                    return "SP1";
                if (buildNumber < 7012)
                    return "SP1 (patched)";
                if (buildNumber == 7012)
                    return "SP2";
                if (buildNumber > 7012)
                    return "SP2 (patched)";

                return string.Empty; // Never accessed

            case 15:
                // Word 2013 (15)
                // https://support.microsoft.com/en-us/kb/2951141
                if (buildNumber < 4420)
                    return "(Beta)";
                if (buildNumber == 4420)
                    return "RTM";
                if (buildNumber < 4569)
                    return "(patched)";
                if (buildNumber == 4569)
                    return "SP1";
                if (buildNumber > 4569)
                    return "SP1 (patched)";

                return string.Empty; // Never accessed

            case 16:
                // Word 2016 (16)
                // 4351 March 2016 security update (right mouse click fix)  https://support.microsoft.com/en-us/kb/3114855
                if (buildNumber < 4266)
                    return "(Beta)";
                if (buildNumber == 4266)
                    return "RTM";
                if (buildNumber < 4351)
                    return "(patched before March 2016)";
                if (buildNumber == 4351)
                    return "(patched March 2016)";
                if (buildNumber > 4351)
                    return "(patched after March 2016)";

                return string.Empty; // Never accessed

            default:
                return string.Empty;
            }
        }


        /// <summary>
        ///     Activates the Word.Window without throwing an exception.
        /// </summary>
        /// <param name="source">A Word.Application instance.</param>
        /// <returns>
        ///     true on success; otherwise false if the there is no active window or the window could not be activated for
        ///     some reason.
        /// </returns>
        
        public static bool ActivateLJ([NotNull] this Window source)
        {
            Check.NotNull(source, nameof(source));

            var wordWindow11 = source as Window11;
            return wordWindow11?.Activate() == 0;
        }


        /// <summary>
        ///     Activates the ActiveWindow without throwing an exception.
        /// </summary>
        /// <param name="source">A Word.Application instance.</param>
        /// <returns>
        ///     true on success; otherwise false if the there is no active window or the window could not be activated for
        ///     some reason.
        /// </returns>
        
        public static bool ActivateWindowLJ([NotNull] this Document source)
        {
            Check.NotNull(source, nameof(source));
            var wordWindow = source.ActiveWindow;
            return wordWindow != null && ActivateLJ(wordWindow);
        }


        /// <summary>
        ///     Returns a WordBasic class that defines some common WordBasic commands that are not available in full object model.
        /// </summary>
        /// <param name="source">A Word.Application instance.</param>
        
        public static WordBasic WordBasicLJ([NotNull] this Application source)
        {
            Check.NotNull(source, nameof(source));
            return _wordBasic ?? (_wordBasic = new WordBasic(source));
        }


        /// <summary>
        ///     Returns an UndoRecord instance, introduced in Word 2010 (14).
        ///     For earlier versions of Word, a null value is returned.
        /// </summary>
        /// <param name="source">A Word.Application instance.</param>
        
        public static UndoRecord UndoRecordLJ([NotNull] this Application source)
        {
            Check.NotNull(source, nameof(source));

            if (source is WordApplication11 wordApp14 && source.VersionLJ() >= OfficeVersion.Office2010)
            {
                return wordApp14.UndoRecord;
            }

            return null;
        }


        /// <summary>
        ///     Returns if Word in running out of process or not.
        /// </summary>
        /// <param name="source">A Word.Application instance.</param>
        /// <returns>true if the Word.Application is running out of process; false otherwise;</returns>
        /// <remarks>
        ///     Some Word properties cannot be set when Word is running out of Process. Specifically CommandBarButton.Picture
        ///     and CommandBarButton.Mask. Setting LJOutOfProcess to true will make the PictureLJ and MaskLJ ignore the set/get
        ///     property request.
        /// </remarks>
        [SuppressMessage("Microsoft.Usage", "CA1801:ReviewUnusedParameters", MessageId = "this")]
        public static bool LJOutOfProcess(Application source)
        {
            return OutOfProcess;
        }

        /// <summary>
        ///     Sets whether the Word.Application is running out of process or not.
        /// </summary>
        /// <param name="source">A Word.Application instance.</param>
        /// <param name="value">true if the Word.Application is running out of process; false otherwise;</param>
        /// <remarks>
        ///     Some Word properties cannot be set when Word is running out of Process. Specifically CommandBarButton.Picture
        ///     and CommandBarButton.Mask. Setting LJOutOfProcess to true will make the PictureLJ and MaskLJ ignore the set/get
        ///     property request.
        /// </remarks>
        public static void LJOutOfProcess(Application source, bool value)
        {
            OutOfProcess = value;
        }


        /// <summary>
        ///     Opens or gets an Word.Application instance
        /// </summary>
        /// <param name="source">A Word.Application instance.</param>
        /// <param name="visible">
        ///     true to set the visibility of the Word Application to true. This is only used if a new
        ///     Word.Application instance is created.
        /// </param>
        /// <param name="addDocument">
        ///     true to add a new document to the Word Application. This is only used if a new
        ///     Word.Application instance is created.
        /// </param>
        /// <returns>An existing, or new Word Application instance.</returns>
        public static Application GetOrStartNewWordApplication([CanBeNull] this Application source, bool visible = true, bool addDocument = true)
        {
            Application wordApp;
            try
            {
                wordApp = (Application) Marshal.GetActiveObject("Word.Application");
            }
            catch (Exception)
            {
                wordApp = (Application) Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));
                wordApp.Visible = visible;
                if (addDocument)
                    wordApp.Documents.Add();
            }

            return wordApp;
        }

        /// <summary>
        ///     Opens or gets an Word.Application instance and opens or gets the document using the file name provided
        /// </summary>
        /// <param name="source">A Word.Application instance.</param>
        /// <param name="fileName">The filename of a document to open.</param>
        /// <param name="visible">
        ///     true to set the visibility of the Word Application to true. This is only used if a new
        ///     Word.Application instance is created.
        /// </param>
        /// <param name="doc">The opened Word Document specified in the fileName.</param>
        /// <returns>A Word.Application instance.</returns>
        public static Application GetOrStartNewWordApplication([CanBeNull] this Application source, string fileName, out Document doc, bool visible = true)
        {
            var wordApp = GetOrStartNewWordApplication(null, visible, addDocument: false);
            doc = wordApp.Documents.Open(fileName);
            return wordApp;
        }


        /// <summary>
        ///     An extension helper for calling the IsObjectValid[] property
        /// </summary>
        /// <param name="source">A Word.Application instance.</param>
        /// <param name="obj">A Word object to check.</param>
        
        public static bool IsObjectValid(this Application source, object obj) => source.IsObjectValid[obj];

        public static FileDialog FileDialog(this Application source, MsoFileDialogType type) => source.FileDialog[type];


        internal static bool OutOfProcess { get; private set; }


        /// <summary>
        ///     Parses the major version number from the Application.Version string and returns the version as an integer.
        /// </summary>
        /// <param name="source">A Word.Application instance.</param>
        /// <returns>The version number as an integer value.</returns>
        public static OfficeVersion VersionLJ([NotNull] this Application source)
        {
            var version = source.Version.Substring(startIndex: 0, length: 2);
            return (OfficeVersion) int.Parse(version);
        }


        /// <summary>
        ///     Returns if the Word 2003 Compatibility Pack a conversion message if the document passed is an old .doc document.
        /// </summary>
        /// <param name="source">A Word.Application instance.</param>
        /// <param name="value">true to display a message;false otherwise</param>
        /// <returns>True if the value was changed;false if the value was not changed</returns>
        public static bool ShowDocXConversionMessage([NotNull] this Application source, bool value)
        {
            Check.NotNull(source, nameof(source));

            if (source.VersionLJ() != OfficeVersion.Office2007)
                return false;

            var curValue = ShowWord2003DocXConversionMessage(source);
            if (curValue == value)
                return false;

            var regKeyModified = false;
            const string subKey = "Software\\Microsoft\\Office\\12.0\\Word\\Options";
            const string valueName = "NoShowCnvMsg";
            if (value)
            {
                Registry.CurrentUser.DeleteValue(subKey, valueName);
            }
            else
            {
                using (var regKey = Registry.CurrentUser.CreateSubKey(subKey, RegistryKeyPermissionCheck.ReadWriteSubTree))
                {
                    regKey?.SetValue(valueName, 1, RegistryValueKind.DWord);
                }
            }

            return true;
        }


        /// <summary>
        ///     Returns if the Word 2003 Compatibility pack displays a conversion message if the document passed is an old .doc
        ///     document.
        /// </summary>
        /// <param name="source">A Word.Application instance.</param>
        /// <returns>True if a message will be displayed;false otherwise</returns>
        public static bool ShowWord2003DocXConversionMessage([NotNull] this Application source)
        {
            if (source.VersionLJ() != OfficeVersion.Office2007)
                return false;

            //Default is to display thus defaultValue = 1
            // NoShowCnvMsg = 1     : return false   
            // NoShowCnvMsg = 0     : return true
            // NoShowCnvMsg missing : return true
            return (int) Registry.CurrentUser.GetKeyValue(
                                                          subKey: "Software\\Microsoft\\Office\\12.0\\Word\\Options",
                                                          valueName: "NoShowCnvMsg",
                                                          defaultValue: 0) != 0;
        }


        /// <summary>
        ///     Returns if the Application.UserName value has not been set, or is set to a null or empty string.
        /// </summary>
        /// <param name="source">A Word.Application instance.</param>
        
        public static bool IsUserNameMissing([NotNull] this Application source)
        {
            Check.NotNull(source, nameof(source));

            var word11 = (WordApplication11) source;
            string userName = null;
            var retVal = word11.UserName(ref userName);
            return retVal != 0 || string.IsNullOrEmpty(userName);
        }


        /// <summary>
        ///     Closes a Word dialog.
        /// </summary>
        /// <param name="source">A Word.Application instance.</param>
        /// <param name="dialog">A WdWordDialog value representing the dialog to close</param>
        public static void CloseDialog([NotNull] this Application source, WdWordDialog dialog)
        {
            dynamic dlg = source.Dialogs[dialog];
            dlg.Close();

            //dlg.GetType().InvokeMember("Close", BindingFlags.Default | BindingFlags.InvokeMethod, null, dlg, null);
        }


        public static Document GetOpenDocument([NotNull] this Application source, string filePath)
        {
            foreach (Document doc in source.Documents)
            {
                if (string.Compare(doc.FullNameLJ(), filePath, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    return doc;
                }
            }

            return null;
        }

        // *********************************************
        // BA-806: TEMPORARY. 
        // Need to make sure we compare UNC to UNC. If file is on One Drive, it will not match as OneDrive files are in Url format
        // TO DO: Work the FullNameLJ to accommodate this problem and remove this code.
        public static Document GetOpenDocumentUnc([NotNull] this Application source, string filePath)
        {
            filePath = GetLocalPath(filePath); // Ensures it is in UNC form

            foreach (Document doc in source.Documents)
            {
                if (string.Compare(doc.FullNameLJ(), filePath, StringComparison.OrdinalIgnoreCase) == 0)
                    return doc;
                if (string.Compare(GetLocalPath(doc.FullNameLJ()), filePath, StringComparison.OrdinalIgnoreCase) == 0)
                    return doc;
            }

            return null;
        }

        private static string GetLocalPath(string sharePointUri)
        {
            try
            {
                var workbookUri = new Uri(sharePointUri);

                // If workbook path points to local file, return it as-is
                if (workbookUri.IsFile)
                    return sharePointUri;

                // Registry key names to loop

                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Environment", false))
                {
                    //Loop through all key names until the file is found
                    foreach (var keyName in new[] { "OneDriveCommercial", "OneDrive" })
                    {
                        var rootDirectory = key.GetValue(keyName).ToString();
                        if (string.IsNullOrEmpty(rootDirectory) == false && System.IO.Directory.Exists(rootDirectory))
                        {
                            // Create a queue consisting of path parts
                            var pathParts = new Queue<string>(workbookUri.LocalPath.Split('/'));
                            while (pathParts.Count != 0)
                            {
                                // Compose a local path by adding root directory and slashes in between
                                var localPath = string.Join(Path.DirectorySeparatorChar.ToString(), pathParts.Union(pathParts));
                                localPath = Path.Combine(rootDirectory, localPath);
                                if (File.Exists(localPath))
                                    return localPath;

                                // The file was not found - get rid of leftmost part of the path and try again
                                pathParts.Dequeue();
                            }
                        }
                    }
                }
            }
#pragma warning disable CA1031 // Do not catch general exception types
            catch
#pragma warning restore CA1031 // Do not catch general exception types
            {
                //If there's an issue with finding the local path, just return null
                return null;
            }

            
            return null;
        }
        // *********************************************



        public static bool IsDocumentOpen([NotNull] this Application source, string filePath)
        {
            return (source.GetOpenDocument(filePath) != null);
        }


        /// <summary>Sets connected state to on/off for a Word COMAddin.</summary>
        /// <param name="source">Word.Application object.</param>
        /// <param name="addinName">Name of the addin to set the loaded state.</param>
        /// <param name="isComAddin">ProgId of addin to disable.</param>
        /// <param name="connected">True/false. Sets Connected value for input COMAddin.</param>
        public static void SetComAddinConnectState([NotNull] this Application source, string addinName, bool isComAddin, bool loaded)
        {
            CheckNotNull(source);
            if (string.IsNullOrEmpty(addinName))
            {
                throw new ArgumentNullException(nameof(addinName));
            }

            if (isComAddin)
            {
                foreach (COMAddIn cai in source.COMAddIns)
                {
                    if (cai.ProgId.ToLower().Contains(addinName))
                    {
                        cai.Connect = loaded;
                    }
                }
            }
            else
            {
                // TODO: KDP - Untested code
                foreach (AddIn ai in source.AddIns)
                {
                    if (ai.Name.Equals(addinName, StringComparison.OrdinalIgnoreCase))
                    {
                        ai.Installed = loaded;
                    }
                }
            }
        }


        /// <summary>Determines if a Word Addin or COMAddin is loaded in the current Word session.</summary>
        /// <param name="source">Word.Application object.</param>
        /// <param name="addinName">
        ///     The Word Addin name. If a COMAddin, is the ProgId value. If a template Addin, is the Name
        ///     value.
        /// </param>
        /// <param name="isComAddin">If true, analyzing a COM Addin. If false, analyzing a Template Addin.</param>
        /// <returns>Returns true if Addin/COMAddin is loaded.</returns>
        public static bool IsAddinLoaded([NotNull] this Application source, string addinName, WordAddinType addinType = WordAddinType.Both)
        {
            Check.NotNull(source, nameof(source));
            Check.NotEmpty(addinName, nameof(addinName));

            var result = false;
            if (addinType == WordAddinType.ComAddin || addinType == WordAddinType.Both)
                result = source.COMAddIns.Cast<COMAddIn>()
                               .Any(cai =>
                                        string.Equals(cai.ProgId, addinName, StringComparison.OrdinalIgnoreCase) && cai.Connect);

            if (!result && (addinType == WordAddinType.DocumentAddin || addinType == WordAddinType.Both))
                result = source.AddIns.Cast<AddIn>()
                               .Any(ai =>
                                        string.Equals(ai.Name, addinName, StringComparison.OrdinalIgnoreCase) && ai.Installed);

            return result;
        }
    }
}