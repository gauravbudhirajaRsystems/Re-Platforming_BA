// © Copyright 2018 Levit & James, Inc.

using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using LevitJames.Core;
using LevitJames.Core.Diagnostics;
using LevitJames.MSOffice.MSWord;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;

namespace LevitJames.AddinApplicationFramework
{
    public class AddinAppTracer : AddinAppBase
    {
        public TraceSourceLJ Tracer { get; private set; }

        protected TraceListener DefaultListener { get; set; }
        protected TraceListener TraceListener { get; set; }
        protected TraceListener ErrorListener { get; set; }
        protected SourceSwitch SourceSwitch { get; private set; }

        public TraceOptions DefaultTraceListenerOptions { get; set; } = TraceOptions.DateTime;
 
        protected override void Initialize(object data)
        {
            base.Initialize(data);

            ConfigureTraceListeners();
        }


        protected internal virtual void ConfigureTraceListeners()
        {
            SourceSwitch = new SourceSwitch("ShowAll") {
                                                           Level = App.UserSettings.TraceLogLevel == TraceLevel.Verbose
                                                                       ? SourceLevels.Verbose
                                                                       : SourceLevels.Information
                                                       };

            ConfigureDefaultListener();
            ConfigureTraceListener();
            ConfigureErrorListener();
            ConfigureTracer();

            // Attach Listeners to Trace and to TraceSources
            AddListeners(Trace.Listeners);
 
        }


        protected virtual void ConfigureDefaultListener()
        {
            if (DefaultListener == null)
            {
                var listener =  new CompactDefaultTraceListener(supportsWriteHeader:false) {
                                                                      TraceOutputOptions = DefaultTraceListenerOptions,
                                                                };
                DefaultListener = listener;
            }

            DefaultListener.Filter = null;
        }


        protected virtual void ConfigureTraceListener()
        {
            if (App.UserSettings.TraceToLogFile && TraceListener == null)
            {
                var fs = new FileStream(App.Paths.TraceLogFile, FileMode.Create, FileAccess.Write, FileShare.ReadWrite);
                var traceLogStreamWriter = new StreamWriter(fs) {AutoFlush = true};
                var tracer= new CompactTextWriterTraceListener(traceLogStreamWriter)
                {
                    TraceOutputOptions = DefaultTraceListenerOptions
                };

                TraceListener = tracer;

                tracer.WriteHeader += (s, e) => WriteSystemSummaryInfo((IWriteTraceHeader)s, Tracer.Switch.Level == SourceLevels.Verbose);
            }
            else
            {
                TraceListener?.Dispose();
            }
        }


        protected void ConfigureErrorListener()
        {
            if (ErrorListener != null)
                return;

            // extend any prior Error.log file
            var errFile = App.Paths.ErrorLogFile;

            //Delete the error log if it is greater than 500KB
            if (File.Exists(errFile))
            {
                var fileInfo = new FileInfo(errFile);
                if (fileInfo.Length / 1024 > 500)
                    FileOperations.Delete(errFile, "Cannot delete Error Log");
            }

            var fs = new FileStream(App.Paths.ErrorLogFile, FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
            var errorLogStreamWriter = new StreamWriter(fs) {AutoFlush = true};

            var errorListener = new ErrorListener(errorLogStreamWriter);

            errorListener.WriteHeader += (s, e) => WriteSystemSummaryInfo((IWriteTraceHeader)s, verbose:true);
      
            errorListener.Filter = new EventTypeFilter(SourceLevels.All);
            errorListener.TriggerFilter = new EventTypeFilter(SourceLevels.Warning);
            errorListener.TraceOutputOptions = TraceOptions.DateTime;
            
            ErrorListener = errorListener;
        }


        private void ConfigureTracer()
        {
            Tracer = (TraceSourceLJ) CreateTraceSource(App.Environment.ProductName);
            // Set the level of tracing to log to any attached listeners
            // from the MSOffice assembly. If the level is set to TraceLevel.Info, the app will log both
            // error- and info-level statements, while if it is set to TraceLevel.Verbose, the app will
            // log errors, info, and verbose trace statements.
            WordExtensions.TraceLevel = App.UserSettings.TraceLogLevel;

            if (Debugger.IsAttached || App.UserSettings.DisplayFailedAssertions && App.UserSettings.DisplayIncludesWarnings)
            {
                Tracer.ForwardToDebugFailSourceLevels = SourceLevels.Warning;
            }
            else
            {
                Tracer.ForwardToDebugFailSourceLevels = SourceLevels.Error;
            }
        }


        protected virtual void WriteSystemSummaryInfo(IWriteTraceHeader headerWriter, bool verbose)
        {

            var sb = new StringBuilder();
   
            //General
            AppendLine(sb, "Data Path:", App.Paths.ProgramDataPath);
            AppendLine(sb, "Version:", App.Version.ToString());
            AppendLine(sb, "Word Version:", WordExtensions.WordApplication.Build);
            AppendLine(sb, "DateTime (Utc):", DateTime.Now.ToUniversalTime());

            headerWriter.WriteHeaderSection(App.Environment.ProductName,sb.ToString());

            //Word Templates
            if (verbose)
                WriteWordTemplateAddinsHeaderSection(headerWriter);
            //ComAddins
            if (verbose)
                WriteWordComAddinsHeaderSection(headerWriter);

            //Environment
            sb.Clear();
            AppendLine(sb, "OSVersion:", Environment.OSVersion + $" \"{OSVersionHelper.FriendlyName()}\"");
            AppendLine(sb, "Is64BitProcess:", Environment.Is64BitProcess.ToString());
            AppendLine(sb, "Is64BitOperatingSystem:", Environment.Is64BitOperatingSystem);
            AppendLine(sb, "Language:", CultureInfo.CurrentCulture.DisplayName + ", LCID:" + CultureInfo.CurrentCulture.LCID);
            AppendLine(sb, "dotNet Version:", Environment.Version.ToString());

            headerWriter.WriteHeaderSection("Environment:", sb.ToString());
 
            if (verbose)
            {
                sb.Clear();
                AppendLine(sb, "Loaded Modules:");
 
                foreach (ProcessModule pm in Process.GetCurrentProcess().Modules)
                {
                    try
                    {
                        AppendLine(sb, pm.FileName + ", File Version:=" + pm.FileVersionInfo.FileVersion);
                    }
                    catch (Exception)
                    {
                        AppendLine(sb, pm.FileName + ", File Version COULD NOT BE DETERMINED");
                    }
                }

                headerWriter.WriteHeaderSection("Loaded Modules:", sb.ToString());
            }
 
        }

        private static void AppendLine(StringBuilder sb, string lineHeader, object value = null, int indent = 0)
        {
            if (value == null)
            {
                sb.AppendIndentedLine(lineHeader, indent);
            }
            else
            {
                sb.AppendIndented(lineHeader.PadRight(totalWidth: 30), indent);
                sb.AppendLine(value.ToString());
            }
        }


        private static void SetListenerPresence(TraceListenerCollection listeners, TraceListener listener, bool include)
        {
            if (listeners == null || listener == null)
                return;

            if (!include)
            {
                listeners.Remove(listener);
                return;
            }

            if (!listeners.Contains(listener))
                listeners.Add(listener);
        }


        private void AddListeners(TraceListenerCollection listeners)
        {
            listeners.Clear();
            SetListenerPresence(listeners, DefaultListener, Debugger.IsAttached && Debugger.IsLogging() || App.UserSettings.DisplayFailedAssertions);
            SetListenerPresence(listeners, ErrorListener, include: true);
            SetListenerPresence(listeners, TraceListener, App.UserSettings.TraceToLogFile);
        }


        public TraceSource CreateTraceSource(string name)
        {
            var source = new TraceSourceLJ(name) {Switch = SourceSwitch};

            AddListeners(source.Listeners);

            if (Debugger.IsAttached || App.UserSettings.DisplayFailedAssertions && App.UserSettings.DisplayIncludesWarnings)
                source.ForwardToDebugFailSourceLevels = SourceLevels.Warning;
            else
                source.ForwardToDebugFailSourceLevels = SourceLevels.Error;

            return source;
        }


        public void SaveTraceLogForError()
        {
            var traceFileName = App.Paths.ErrorLogFile;
            if (!File.Exists(traceFileName)) return;

            var newFileName = App.Paths.AppendDateTimeToFileName(traceFileName, "Trace_");
            try
            {
                File.Copy(traceFileName, newFileName);
            }
            catch (Exception)
            {
                // ignored
            }
        }


        private static void WriteWordTemplateAddinsHeaderSection(IWriteTraceHeader headerWriter)
        {
            var sb = new StringBuilder();

            var added = false;
            foreach (AddIn tAddin in WordExtensions.WordApplication.AddIns)
            {
                if (added == false)
                {
                    sb.AppendLine("Word Addins");
                    added = true;
                }

                string name = null;
                try
                {
                    name = tAddin.Name;
                }
                catch
                {
                    // ignored
                }

                if (string.IsNullOrEmpty(name))
                {
                    name = "Unknown";
                }

                sb.AppendLine(name + ", Installed:=" + tAddin.Installed + ", AutoLoad:=" + tAddin.Autoload);
                sb.AppendLine("  Path:" + tAddin.Path);
            }

            if (added)
                headerWriter.WriteHeaderSection("Word Template Addins", sb.ToString());
 
        }
 
        private void WriteWordComAddinsHeaderSection(IWriteTraceHeader headerWriter)
        {
            var sb = new StringBuilder();
            var added = false;
 
            foreach (COMAddIn cAddin in App.WordApplication.COMAddIns)
            {
 
                added = true;
      
                sb.AppendIndentedLine(cAddin.Description + ", Connect:=" + cAddin.Connect, 1);
                if (string.IsNullOrEmpty(cAddin.Guid))
                    continue;

                try
                {
                    using (var registryKey = Registry.ClassesRoot.OpenSubKey("CLSID\\" + cAddin.Guid + "\\InprocServer32",
                                                                             writable: false))
                    {
                        if (registryKey != null)
                        {
                            var value = Convert.ToString(registryKey.GetValue(name: null));
                            if (!string.IsNullOrEmpty(value))
                            {
                                sb.AppendIndentedLine("Path:" + value, 2);
                                if (File.Exists(value))
                                {
                                    var version = FileVersionInfo.GetVersionInfo(value);
                                    sb.AppendIndentedLine("Product Version:" + version.ProductVersion + ", File Version:" +
                                                  version.FileVersion + ", PR:" + version.IsPreRelease,2);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    sb.AppendIndentedLine("Error retrieving ComAddin Info: " + ex.Message, 2);
                }
            }
            if (added)
                headerWriter.WriteHeaderSection("Word Com Addins", sb.ToString());
 
        }
    }
}