// © Copyright 2018 Levit & James, Inc.

using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using JetBrains.Annotations;
using LevitJames.Core;
using LevitJames.Core.Reflection;
using LevitJames.MSOffice.MSWord;
using LevitJames.Win32;
using Microsoft.Office.Core;

namespace LevitJames.MSOffice.Addin
{

    // ReSharper disable InconsistentNaming
    internal interface IGetRibbonUI
    {
        IRibbonUI GetRibbonUI();
    }
    // ReSharper restore InconsistentNaming

    //Note ClassInterface attribute must be set to AutoDispatch or IRibbonExtensibility will not work.

    /// <summary>
    ///     Add-ins inherit this class to provide a connection to Office applications.
    /// </summary>
    /// <remarks>
    ///     The inherited class must be ,marked with the ComVisible(true) attribute in order for IRibbonExtensibility to work.
    ///     The assembly does not however need to be registered with reg-asm
    /// </remarks>
    [ComVisible(visibility: true)]
    [Guid("6B6CB094-8726-4493-B40A-F3C733FCA7BB")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public abstract class OfficeAddinConnection : MarshalByRefObject,
                                                  IDTExtensibility2, 
                                                  IRibbonExtensibility, 
                                                  IAddinCommand, 
                                                  ICustomTaskPaneConsumer,
                                                  IDisposable,
                                                  IRibbonCallbacks, IAddinApp

    {
        private ICTPFactory _ctpFactoryInst;
        private bool _doStartupComplete;
 
        private IGetRibbonUI _getRibbonUi;
        private IRibbonUI _ribbonUi;
 
        private string _assemblyResolveLocation;

        public static void Create<T>(string name, OfficeAddinAdapter officeAddinAdapter, string rootAddinLocation) where T : OfficeAddinConnection
        {
            var info = AppDomain.CurrentDomain.SetupInformation;
            var configurationFile = info.ConfigurationFile;
            info.ApplicationBase = Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBaseFullPath());
            info.ConfigurationFile = configurationFile;

            var appDomain = AppDomain.CreateDomain(name, new System.Security.Policy.Evidence(), info);
            // ReSharper disable AssignNullToNotNullAttribute
            var connection = (T)appDomain.CreateInstanceAndUnwrap(typeof(T).Assembly.FullName, typeof(T).FullName);
            // ReSharper restore AssignNullToNotNullAttribute

            connection.Initialize(officeAddinAdapter.GetRibbonUI(), rootAddinLocation);

            officeAddinAdapter.Initialize(connection);

        }

        private void Initialize(IGetRibbonUI ribbonUi, string rootAddinLocation)
        {
            _getRibbonUi = ribbonUi;
            if (!Directory.Exists(rootAddinLocation))
                return;

            _assemblyResolveLocation = rootAddinLocation;
            AppDomain.CurrentDomain.AssemblyResolve += AppDomainAssemblyResolveHandler;
        }

        private Assembly AppDomainAssemblyResolveHandler(object sender, ResolveEventArgs args)
        {
            var name = new AssemblyName(args.Name);
            var path = Path.Combine(_assemblyResolveLocation, name.Name + ".dll");
            if (File.Exists(path))
                return Assembly.LoadFrom(path);

            //path = Path.Combine(this._assemblyResolveLocation, name.Name + ".exe");
            //if (File.Exists(path))
            //    return Assembly.LoadFrom(path);

            return null;
        }

 
        public override object InitializeLifetimeService()
        {
            return null;
        }

        public static OfficeAddinConnection Instance { get; private set; }

        protected OfficeAddinConnection()
        {
        }

        protected OfficeAddinConnection(object application) 
        {
            ((IDTExtensibility2) this).OnConnection(application, AddinConnectMode.Startup, addInInst: null, custom: null);
            ((IDTExtensibility2) this).OnStartupComplete(custom: null);
        }

        public string GetProgId()
        {
	        var addinInfo = Instance.GetType().GetCustomAttribute<AddinAttribute>() ?? new AddinAttribute();
	        addinInfo.EnsureValues(Instance.GetType());
	        return addinInfo.ProgId;
        }

        /// <summary>
        ///     Returns the host Application i.e. Application
        /// </summary>
        public object Application { get; private set; }

        public IOfficeAddin AddinApp { get; private set; }

        public object Tag { get; set; }


        public string StartCommand { get; private set; }

        public string StartCommandParameters { get; private set; }

        public void Execute(string commandName, object parameter)
        {
            ExecuteAddinCommand(commandName, parameter);
        }

#pragma warning disable CA1033 // Interface methods should be callable by child types
        void ICustomTaskPaneConsumer.CTPFactoryAvailable(ICTPFactory factoryInst)
#pragma warning restore CA1033 // Interface methods should be callable by child types
        {
            _ctpFactoryInst = factoryInst;
        }


        void IDTExtensibility2.OnAddInsUpdate(ref Array custom)
        {
            try
            {
                OnAddinUpdate(custom);

                // When an addin is manually enabled via Tools.ComAddins it does not call IDTExtensibility2_OnStartupComplete 
                // so we manually call IDTExtensibility2_OnStartupComplete if the connectMode in IDTExtensibility2_OnConnection was AfterStartup
                if (_doStartupComplete)
                {
                    _doStartupComplete = false;
                    ((IDTExtensibility2) this).OnStartupComplete(custom);
                }
            }
            catch (Exception ex)
            {
                ApplicationHelper.ShowMessage(ex.Message, "Error updating Levit & James, inc Addin.");
                throw;
            }
        }

        void IDTExtensibility2.OnBeginShutdown(ref Array custom)
        {
            try
            {
                OnBeginAddinShutdown(custom);
            }
            catch (Exception ex)
            {
                ApplicationHelper.ShowMessage(ex.Message, "Error shutting down Levit & James, inc Addin.");
                throw;
            }
        }

        void IDTExtensibility2.OnConnection(object application, AddinConnectMode connectMode, object addInInst,
                                            ref Array custom)
        {
            try
            {
                if (connectMode == AddinConnectMode.AfterStartup)
                    //because AddinConnectMode.AfterStartup = 0 we use a flag rather than store the connectMode
                {
                    _doStartupComplete = true;
                }

                Application = application;
                if (Marshal.IsTypeVisibleFromCom(GetType()))
                {
                    if (addInInst is COMAddIn officeAddin)
                    {
                        officeAddin.Object = this;
                    }
                }

                OnAddinConnected(application, connectMode, addInInst, ref custom);
            }
            catch (Exception ex)
            {
                ApplicationHelper.ShowMessage(ex.Message, "Error starting Levit & James, inc Addin.");
                throw;
            }
        }


        void IDTExtensibility2.OnDisconnection(AddinDisconnectMode removeMode, ref Array custom)
        {
            OnAddinDisconnected(removeMode, custom);
        }

        void IDTExtensibility2.OnStartupComplete(ref Array custom)
        {
            GetStartupArguments();
            AddinApp?.Initialize(custom);
            OnAddinStartupComplete(custom);
        }


#pragma warning disable CA1033 // Interface methods should be callable by child types
        string IRibbonExtensibility.GetCustomUI(string ribbonId)
#pragma warning restore CA1033 // Interface methods should be callable by child types
        {
            return OnGetRibbon(ribbonId);
        }

        /// <summary>
        /// Invalidates the ribbon ui.
        /// </summary>
        public void InvalidateRibbon()
        {
            if (_getRibbonUi != null)
                _getRibbonUi.GetRibbonUI()?.Invalidate();
            else
                _ribbonUi?.Invalidate();
        }

        /// <summary>
        /// Invalidates the ribbon control with the supplied id.
        /// </summary>
        public void InvalidateRibbonControl(string id)
        {
            if (_getRibbonUi != null)
                _getRibbonUi.GetRibbonUI()?.InvalidateControl(id);
            else
                _ribbonUi?.Invalidate();
        }

        // ReSharper disable once InconsistentNaming
        public IRibbonUI RibbonUI
        {
            get
            {
                if (_getRibbonUi != null)
                    return _getRibbonUi.GetRibbonUI();

                return _ribbonUi;
            }
        }

        public static AddinApplicationInfo AddinEntryAssembly { get; private set; }

        protected virtual IOfficeAddin CreateAddinApp() => null;

        [ComRegisterFunction]
        private static void RegisterClass([NotNull] Type registerType) =>
            new AddinRegistration().Register(registerType);

        [ComUnregisterFunction]
        private static void UnRegisterClass([NotNull] Type registerType) => new AddinRegistration().UnRegister(registerType);


        /// <summary>
        ///     Calls AddinApp.StartAddin if AddinApp is not null.
        /// </summary>
        public void StartAddin()
        {
            try
            {
                if (AddinApp?.Running == true)
                    return;

                AddinApp?.StartAddin();
            }
            catch (Exception ex)
            {
                ApplicationHelper.ShowMessage(ex.Message, "Error starting Levit & James, inc Addin.");
                throw;
            }
        }


        protected virtual void OnAddinStartupComplete(Array custom)
        {
            //Default Implementation does nothing
        }
 
        protected virtual void OnAddinConnected(object application, AddinConnectMode connectMode, object addinInstance, ref Array custom)
        {
            Instance = this;
            AddinEntryAssembly = new AddinApplicationInfo(this);
            CustomSettingsProviderOptions.SetDefault(new CustomSettingsProviderOptions(AddinEntryAssembly.EntryAssembly));

            AddinApp = CreateAddinApp();
    
            if (WordExtensions.Singleton == null)
	            WordExtensions.NewWordExtensions((Microsoft.Office.Interop.Word.Application)application);

            InitDpi();

        }

        
        private void InitDpi()
        {
	        var opusAppHandle = WordExtensions.GetWordApplicationWindow();
	        if (opusAppHandle == IntPtr.Zero)
		        return;

	        var dpiContext = DpiAwarenessContextHandle.FromWindow(opusAppHandle);
            var awareness = dpiContext.Resolve();

            if (dpiContext == DpiAwareness.Unaware)
	        {
		        if (DpiAwarenessContextHandle.IsProcessDPIAware())
			        awareness = DpiAwareness.SystemAware;
	        }

            if (awareness == DpiAwareness.PerMonitorAware)
            {
                AppContext.SetSwitch("Switch.System.Windows.DoNotScaleForDpiChanges", false);
                awareness = DpiAwareness.PerMonitorAware2;
            }

            DpiAwarenessContext.Awareness = awareness;
        }

        /// <summary>
        ///     Called when the addin has been disconnected from its host.
        /// </summary>
        /// <param name="removeMode"></param>
        /// <param name="custom">The AddinDisconnectedEventArgs arguments to pass.</param>
        protected virtual void OnAddinDisconnected(AddinDisconnectMode removeMode, Array custom)
        {
           Dispose();
        }


        protected virtual void OnAddinUpdate(Array custom) { }


        protected virtual void OnBeginAddinShutdown(Array custom) { }


        /// <summary>
        ///     Override this method to return an xml string representing the Ribbon UI for the Office application
        /// </summary>
        /// <param name="ribbonId">The id of the ribbon to return.</param>
        /// <returns>An xml string representing the ribbon controls to display.</returns>
        /// <remarks>
        ///     This member is only applicable to Office applications version 2007 and newer. If the RibbonUserActionAdapter
        ///     has been added to the UserActionManager.Adapters collection then this method will not be called.
        /// </remarks>
        protected virtual string OnGetRibbon(string ribbonId) => AddinApp?.GetCustomUI(ribbonId);


        /// <summary>
        ///     Override this method to return an
        /// </summary>
        /// <param name="ribbonControl">A ribbonControl instance identifying the ribbon control requesting the image.</param>
        /// <returns>An Drawing.Image to use for the ribbon</returns>
        /// <remarks>
        ///     This member is only applicable to Office applications version 2007 and newer. If the RibbonUserActionAdapter
        ///     has been added to the UserActionManager.Adapters collection then this method will not be called.
        /// </remarks>
        protected virtual Image OnRibbonGetButtonImage(IRibbonControl ribbonControl)
        {
            return null;
        }


        [ComVisible(visibility: true)]
        [EditorBrowsable(EditorBrowsableState.Never)]
        public string Version() => typeof(OfficeAddinConnection).Assembly.GetName().Version.ToString();


        private OfficeVersion _officeVersion= OfficeVersion.Unknown;
        public OfficeVersion OfficeVersion
        {
            get
            {
                if (_officeVersion != OfficeVersion.Unknown)
                    return _officeVersion;
                if (Application == null)
                    return OfficeVersion.Unknown;
                if (Application.GetValueByName("Version") is string version)
                {
                    version = version.Substring(startIndex: 0, length: 2);
                    _officeVersion = (OfficeVersion)int.Parse(version, CultureInfo.InvariantCulture);
                }

                return _officeVersion;
            }
        }
        //IRibbonUI


        /// <summary>
        ///     The callback that Office 2007 and above will call when the ribbon ui is loaded.
        ///     This member is only called when the xml returned in OnRibbonGetCustomUI contains "onLoad='RibbonCallbackOnLoad'"
        /// </summary>
        /// <param name="ribbon">The IRibbonUI interface to use to update the ribbon display.</param>
        
        [ComVisible(visibility: true)]
        [EditorBrowsable(EditorBrowsableState.Never)]
        public void RibbonCallbackOnLoad(IRibbonUI ribbon)
        {
            _ribbonUi = ribbon;
           

        }
        //[ComVisible(visibility: true)]
        //[EditorBrowsable(EditorBrowsableState.Never)]
        //public void RibbonCallbackOnLoad(Func<IRibbonUI> ribbon)
        //{
        //    _ribbonUi = ribbon;

        //}

        [ComVisible(visibility: true)]
        [EditorBrowsable(EditorBrowsableState.Never)]
        public void RunCommand(string commandName)
        {
            ExecuteAddinCommand(commandName, parameter: null);
        }

        [ComVisible(visibility: true)]
        [EditorBrowsable(EditorBrowsableState.Never)]
        public void RunCommand2(string commandName, string parameter)
        {
            if (string.IsNullOrWhiteSpace(parameter))
            {
                parameter = null;
            }

            ExecuteAddinCommand(commandName, parameter);
        }

        protected virtual void ExecuteAddinCommand(string commandName, object parameter) { }


        protected virtual void OnRibbonAction(IRibbonControl control)
        {
            if (control == null)
                return;

            //using (new DpiAwarenessContext())
	            OnRibbonActionCore(control);
        }

        private void OnRibbonActionCore(IRibbonControl control)
        {
            try
            {
                if (GetRibbonAdapter(control)?.RibbonCallbackOnAction(control) == false)
                {
                    UserActionManager.TryExecuteUserAction(control.Id, control.Tag);
                }
            }
            catch (Exception ex)
            {
                ApplicationHelper.ShowMessage(ex.Message, "Error in Office Ribbon callback: " + new StackTrace().GetFrame(0).GetMethod().Name);
                throw;
            }
        }
        //Ribbon Control callbacks

        [CLSCompliant(isCompliant: false)]
        [ComVisible(visibility: true)]
        [EditorBrowsable(EditorBrowsableState.Never)]
#pragma warning disable IDE0060 // Remove unused parameter
        public void RibbonCallbackBackStageOnShow(object control)
#pragma warning restore IDE0060 // Remove unused parameter
        {
            WordExtensions.InBackstage = true;
        }

        [CLSCompliant(isCompliant: false)]
        [ComVisible(visibility: true)]
        [EditorBrowsable(EditorBrowsableState.Never)]
#pragma warning disable IDE0060 // Remove unused parameter
        public void RibbonCallbackBackStageOnHide(object control)
#pragma warning restore IDE0060 // Remove unused parameter
        {
            WordExtensions.InBackstage = false;
        }


        ///// <summary>
        /////     Called by Office 2007 and above, to notify when a ribbon button was clicked.
        ///// </summary>
        ///// <param name="control">The Ribbon control that executed the callback.</param>
        
        //[CLSCompliant(isCompliant: false)]
        //[ComVisible(visibility: true)]
        //[EditorBrowsable(EditorBrowsableState.Never)]
        //public void RibbonCallbackOnAction(IRibbonControl control)
        //{
        //    OnRibbonAction(control);
        //}


        /// <summary>
        ///     Called by Office 2007 and above, to notify when a ribbon button was clicked.
        /// </summary>
        /// <param name="control">The Ribbon control that executed the callback.</param>
        /// <param name="cancelDefault"></param>
        
        [CLSCompliant(isCompliant: false)]
        [ComVisible(visibility: true)]
        [EditorBrowsable(EditorBrowsableState.Never)]
#pragma warning disable IDE0060 // Remove unused parameter
        public void RibbonCallbackOnAction(IRibbonControl control)
#pragma warning restore IDE0060 // Remove unused parameter
        {
            OnRibbonAction(control);
        }

        /// <summary>
        ///     Called by Office 2007 and above, to notify when a ribbon button was clicked.
        /// </summary>
        /// <param name="control">The Ribbon control that executed the callback.</param>
        /// <param name="pressed"></param>
        [ComVisible(visibility: true)]
        [EditorBrowsable(EditorBrowsableState.Never)]
        [CLSCompliant(isCompliant: false)]
#pragma warning disable IDE0060 // Remove unused parameter
        public void RibbonCallbackToggleButtonOnAction(IRibbonControl control, bool pressed)
#pragma warning restore IDE0060 // Remove unused parameter
        {
            OnRibbonAction(control);
        }


        private RibbonUserActionAdapter GetRibbonAdapter(IRibbonControl control)
        {
            try
            {
                if (AddinApp?.Running == false)
                {
                    //We do not start the application if the ribbon item is a context menu.
                    if (control.Id.EndsWith(RibbonUserActionAdapter.ContextMenuIdIdentifier, StringComparison.OrdinalIgnoreCase))
                        return null;

                    StartAddin();
                }
                   

                return UserActionManager.Adapters.GetItem<RibbonUserActionAdapter>();
            }
            catch (Exception ex)
            {
                ApplicationHelper.ShowMessage(ex.Message, "Error in Office Ribbon callback: " + new StackTrace().GetFrame(0).GetMethod().Name);
                throw;
            }
        }

        [ComVisible(true), EditorBrowsable(EditorBrowsableState.Never), CLSCompliant(false)]
        public void RibbonCallbackDropdownAction(IRibbonControl control, string selectedId, int selectedIndex)
            => DoRibbonCallback(control, (object)null, (ctl, adapter, param) => adapter.RibbonCallbackDropDownAction(ctl, selectedId, selectedIndex));
  
        [return: MarshalAs(UnmanagedType.IDispatch)]
        [CLSCompliant(false), ComVisible(true), EditorBrowsable(EditorBrowsableState.Never)]
        public object RibbonCallbackGetButtonImage(IRibbonControl control) =>
            DoRibbonCallbackFunc(control, (object)null, (ctl, adapter, _) => adapter.RibbonCallbackGetButtonImage(ctl));

        [ComVisible(true), EditorBrowsable(EditorBrowsableState.Never), CLSCompliant(false)]
        public string RibbonCallbackGetComboText(IRibbonControl control) =>
            DoRibbonCallbackFunc(control, (object)null, (ctl, adapter, _) => adapter.RibbonCallbackGetComboText(ctl));

        [CLSCompliant(false), ComVisible(true), EditorBrowsable(EditorBrowsableState.Never)]
        public string RibbonCallbackGetDescription(IRibbonControl control) =>
            DoRibbonCallbackFunc(control, (object)null, (ctl, adapter, _) => adapter.RibbonCallbackGetDescription(ctl));

        [CLSCompliant(false), ComVisible(true), EditorBrowsable(EditorBrowsableState.Never)]
        public string RibbonCallbackGetEditValue(IRibbonControl control) =>
            DoRibbonCallbackFunc(control, (object)null, (ctl, adapter, _) => adapter.RibbonCallbackGetEditValue(ctl));

        [CLSCompliant(false), ComVisible(true), EditorBrowsable(EditorBrowsableState.Never)]
        public bool RibbonCallbackGetEnabled(IRibbonControl control) =>
            DoRibbonCallbackFunc(control, (object)null, (ctl, adapter, _) => adapter.RibbonCallbackGetEnabled(ctl));

        [ComVisible(true), EditorBrowsable(EditorBrowsableState.Never), CLSCompliant(false)]
        public int RibbonCallbackGetItemCount(IRibbonControl control) =>
            DoRibbonCallbackFunc(control, (object)null, (ctl, adapter, _) => adapter.RibbonCallbackGetItemCount(ctl));

        [ComVisible(true), EditorBrowsable(EditorBrowsableState.Never), CLSCompliant(false)]
        public string RibbonCallbackGetItemId(IRibbonControl control, int index) =>
            DoRibbonCallbackFunc(control, index, (ctl, adapter, param) => adapter.RibbonCallbackGetItemId(ctl, param));

        [return: MarshalAs(UnmanagedType.IDispatch)]
        [ComVisible(true), EditorBrowsable(EditorBrowsableState.Never), CLSCompliant(false)]
        public object RibbonCallbackGetItemImage(IRibbonControl control, int index) =>
            DoRibbonCallbackFunc(control, index, (ctl, adapter, param) => adapter.RibbonCallbackGetItemImage(ctl, param));

        [ComVisible(true), EditorBrowsable(EditorBrowsableState.Never), CLSCompliant(false)]
        public string RibbonCallbackGetItemLabel(IRibbonControl control, int index) =>
            DoRibbonCallbackFunc(control, index, (ctl, adapter, param) => adapter.RibbonCallbackGetItemLabel(ctl, param));

        [ComVisible(true), EditorBrowsable(EditorBrowsableState.Never), CLSCompliant(false)]
        public string RibbonCallbackGetItemScreenTip(IRibbonControl control, int index) =>
            DoRibbonCallbackFunc(control, index, (ctl, adapter, param) => adapter.RibbonCallbackGetItemScreenTip(ctl, param));

        [CLSCompliant(false), ComVisible(true), EditorBrowsable(EditorBrowsableState.Never)]
        public string RibbonCallbackGetKeyTip(IRibbonControl control) =>
            DoRibbonCallbackFunc<object, string>(control, null, (ctl, adapter, _) => adapter.RibbonCallbackGetKeyTip(ctl));

        [CLSCompliant(false), ComVisible(true), EditorBrowsable(EditorBrowsableState.Never)]
        public string RibbonCallbackGetLabel(IRibbonControl control) =>
            DoRibbonCallbackFunc<object, string>(control, null, (ctl, adapter, _) => adapter.RibbonCallbackGetLabel(ctl));

        [CLSCompliant(false), ComVisible(true), EditorBrowsable(EditorBrowsableState.Never)]
        public bool RibbonCallbackGetPressed(IRibbonControl control) =>
            DoRibbonCallbackFunc<object, bool>(control, null, (ctl, adapter, _) => adapter.RibbonCallbackGetPressed(ctl));

        [CLSCompliant(false), ComVisible(true), EditorBrowsable(EditorBrowsableState.Never)]
        public string RibbonCallbackGetScreenTip(IRibbonControl control) =>
            DoRibbonCallbackFunc<object, string>(control, null, (ctl, adapter, _) => adapter.RibbonCallbackGetScreenTip(ctl));

        [ComVisible(true), EditorBrowsable(EditorBrowsableState.Never), CLSCompliant(false)]
        public int RibbonCallbackGetSelectedItemIndex(IRibbonControl control) =>
            DoRibbonCallbackFunc<object, int>(control, null, (ctl, adapter, _) => adapter.RibbonCallbackGetSelectedItemIndex(ctl));

        [CLSCompliant(false), ComVisible(true), EditorBrowsable(EditorBrowsableState.Never)]
        public int RibbonCallbackGetSize(IRibbonControl control) =>
            DoRibbonCallbackFunc<object, int>(control, null, (ctl, adapter, _) => adapter.RibbonCallbackGetSize(ctl));

        [CLSCompliant(false), ComVisible(true), EditorBrowsable(EditorBrowsableState.Never)]
        public string RibbonCallbackGetSuperTip(IRibbonControl control) =>
            DoRibbonCallbackFunc<object, string>(control, null, (ctl, adapter, _) => adapter.RibbonCallbackGetSuperTip(ctl));

        [CLSCompliant(false), ComVisible(true), EditorBrowsable(EditorBrowsableState.Never)]
        public string RibbonCallbackGetTabLabel(IRibbonControl control) =>
            DoRibbonCallbackFunc<object, string>(control, null, (ctl, adapter, _) => adapter.RibbonCallbackGetLabel(ctl));

        [CLSCompliant(false), ComVisible(true), EditorBrowsable(EditorBrowsableState.Never)]
        public bool RibbonCallbackGetVisible(IRibbonControl control) =>
            DoRibbonCallbackFunc<object, bool>(control, null, (ctl, adapter, _) => adapter.RibbonCallbackGetVisible(ctl));

        [CLSCompliant(false), ComVisible(true), EditorBrowsable(EditorBrowsableState.Never)]
        public bool RibbonCallbackGetVisibleTab(IRibbonControl control)
        {
            if (AddinApp?.Running == false && (control != null && control.Tag.Contains(OfficeRibbonHelper.HideTabUntilStartedRibbonTagArgument)))
                return false;

            return DoRibbonCallbackFunc<object, bool>(control, null, (ctl, adapter, _) => adapter.RibbonCallbackGetVisible(ctl));
        }
 
        [CLSCompliant(false), ComVisible(true), EditorBrowsable(EditorBrowsableState.Never)]
        public void RibbonCallbackOnChange(IRibbonControl control, string text)
        {
            DoRibbonCallback(control, text, (ctl, adapter, param) => adapter.RibbonCallbackOnChange(ctl, param));
        }

        private void DoRibbonCallback<TParam>(IRibbonControl control, TParam param, Action<IRibbonControl, RibbonUserActionAdapter, TParam> callback)
        {
 
            if (control == null)
                return;

            try
            {
                var ribbonAdapter = GetRibbonAdapter(control);

                if (ribbonAdapter != null)
                    callback(control, ribbonAdapter, param);
            }
            catch (Exception exception1)
            {
                ApplicationHelper.ShowMessage(exception1.Message, "Error in Office Ribbon callback: " + new StackTrace().GetFrame(1).GetMethod().Name, 0, IntPtr.Zero);
                throw;
            }

        }

        private TResult DoRibbonCallbackFunc<TParam, TResult>(IRibbonControl control, TParam param, Func<IRibbonControl, RibbonUserActionAdapter, TParam, TResult> callback)
        {
 
            if (control == null)
                return default;
 
            try
            {
                var ribbonAdapter = GetRibbonAdapter(control);
      
                if (ribbonAdapter != null)
                    return callback(control, ribbonAdapter, param);

                return default;
            }
            catch (Exception exception1)
            {
                ApplicationHelper.ShowMessage(exception1.Message, "Error in Office Ribbon callback: " + new StackTrace().GetFrame(1).GetMethod().Name, 0, IntPtr.Zero);
                throw;
            }
 
        }


        private void GetStartupArguments()
        {
            var commandLine = Environment.CommandLine;

            var nPos = commandLine.IndexOf(" \"//?", StringComparison.Ordinal);
            if (nPos > 0)
            {
                nPos += 5;
                var nPosEnd = commandLine.IndexOf(value: '\"', startIndex: nPos);
                if (nPosEnd == -1)
                {
                    nPosEnd = commandLine.Length;
                }

                commandLine = commandLine.Substring(nPos, nPosEnd - nPos);
                nPos = commandLine.IndexOf(value: '|');
                if (nPos > -1)
                {
                    if (nPos > 0)
                    {
                        StartCommand = commandLine.Substring(startIndex: 0, length: nPos).Trim();
                        if (!string.IsNullOrEmpty(StartCommand))
                        {
                            StartCommandParameters = commandLine.Substring(nPos + 1).Trim();
                        }
                    }
                }
                else
                {
                    StartCommand = commandLine.Trim();
                }
            }
        }


        internal CustomTaskPane CreateTaskPane(Type activeXControl, string title, object parentWindow)
        {
            if (_ctpFactoryInst == null)
                return null;

            var activeXControlId = activeXControl.GetCustomAttribute<ProgIdAttribute>();
            return _ctpFactoryInst.CreateCTP(activeXControlId.Value, title, parentWindow);
        }


 
        protected virtual void Dispose(bool disposing)
        {
 
            if (disposing)
            {
                if (WordExtensions.Singleton != null)
                {
                    WordExtensions.Singleton.OnAddinDisconnected();

                    WordExtensions.Singleton.Dispose();
                }

                AddinEntryAssembly = null;
                _ctpFactoryInst = null;
                _ribbonUi = null;
                _getRibbonUi = null;
                Application = null;

                (AddinApp as IDisposable)?.Dispose();
                AddinApp = null;
                Instance = null;
            }
        }
 
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~OfficeAddinConnection()
        {
            Dispose(false);
        }

        public void HideTaskPanes(object document)
        {
            AddinApp.HideTaskPanes(document);
        }


    }


}