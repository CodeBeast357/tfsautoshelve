using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using VsExt.AutoShelve.EventArgs;
using VsExt.AutoShelve.IO;
using VsExt.AutoShelve.Packaging;

namespace VsExt.AutoShelve
{
    /// <summary>
    /// <para>This is the class that implements the package exposed by this assembly.</para>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the 
    /// IVsPackage interface and uses the registration attributes defined in the framework to 
    /// register itself and its components with the shell.
    /// </para>
    /// </summary>
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [ProvideAutoLoad(UIContextGuids80.SolutionExists)] // https://msdn.microsoft.com/en-us/library/microsoft.visualstudio.shell.interop.uicontextguids80.aspx
    [ProvideMenuResource("Menus.ctmenu", 1)]
    // This attribute is used to include custom options in the Tools->Options dialog
    [ProvideOptionPage(typeof(OptionsPageGeneral), "TFS Auto Shelve", OptionsPageGeneral.GeneralCategory, 101, 106, true)]
    [ProvideService(typeof(TfsAutoShelve))]
    // This attribute is used to register the information needed to show this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "6.2", IconResourceID = 400)]
    [Guid(GuidList.GuidAutoShelvePkgString)]
    public class VsExtAutoShelvePackage : Package, IVsSolutionEvents, IDisposable
    {
        private IAutoShelve _autoShelve;
        private DTE2 _dte;
        private IVsActivityLog _log;
        private OleMenuCommand _menuRunState;
        private OptionsPageGeneral _options;
        private uint _solutionEventsCookie;
        private IVsSolution2 _solutionService;
        private bool _isPaused;

        public static string ExtensionName => Resources.ExtensionName;

        private string _menuTextRunning => string.Format(Resources.MenuTextRunning, ExtensionName);
        private string _menuTextStopped => string.Format(Resources.MenuTextStopped, ExtensionName);

        /// <summary>
        /// Default constructor of the package.
        /// Inside this method you can place any initialization code that does not require 
        /// any Visual Studio service because at this point the package object is created but 
        /// not sited yet inside Visual Studio environment. The place to do all the other 
        /// initialization is the Initialize method.
        /// </summary>
        public VsExtAutoShelvePackage()
        {
            Trace.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this));

            IServiceContainer serviceContainer = this as IServiceContainer;
            ServiceCreatorCallback callback = new ServiceCreatorCallback(CreateService);
            serviceContainer.AddService(typeof(ISAutoShelve), callback, true);
        }

        private object CreateService(IServiceContainer container, Type serviceType)
        {
            if (typeof(ISAutoShelve) == serviceType)
            {
                if (_autoShelve == null)
                    _autoShelve = new TfsAutoShelve(this);
                return _autoShelve;
            }
            return null;
        }

        /////////////////////////////////////////////////////////////////////////////
        // Overridden Package Implementation
        #region Package Members

        private void AutoShelve_OnShelvesetCreated(object sender, ShelvesetCreatedEventArgs e)
        {
            if (e.IsSuccess)
            {
                if (e.ShelvesetChangeCount != 0)
                {
                    var str = string.Format(Resources.ShelvesetSuccess, e.ShelvesetChangeCount, e.ShelvesetName);
                    if (e.ShelvesetsPurgeCount > 0)
                    {
                        str += string.Format(Resources.ShelvesetSucessPurged, _autoShelve.MaximumShelvesets, e.ShelvesetsPurgeCount);
                    }
                    WriteToStatusBar(str);
                    WriteLineToOutputWindow(str);
                }
            }
            else
            {
                WriteException(e.Error);
            }
        }

        private void AutoShelve_OnTfsConnectionError(object sender, TfsConnectionErrorEventArgs e)
        {
            WriteLineToOutputWindow(Resources.ErrorNotConnected);
            WriteException(e.Error);
        }

        private void AutoShelve_OnStart(object sender, System.EventArgs e)
        {
            DisplayRunState();
        }

        private void AutoShelve_OnStop(object sender, System.EventArgs e)
        {
            DisplayRunState();
        }

        private void DisplayRunState()
        {
            string str1;
            if (_isPaused)
            {
                str1 = string.Format(Resources.StatePaused, ExtensionName);
            }
            else if (_autoShelve.IsRunning)
            {
                str1 = string.Format(Resources.StateRunning, ExtensionName);
            }
            else
            {
                str1 = string.Format(Resources.StateStopped, ExtensionName);
            }
            WriteToStatusBar(str1);
            WriteLineToOutputWindow(str1);
            ToggleMenuCommandRunStateText(_menuRunState);
        }

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            try
            {
                Trace.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this));

                base.Initialize();

                // InitializePackageServices
                _dte = (DTE2)GetGlobalService(typeof(DTE));

                _log = GetService(typeof(SVsActivityLog)) as IVsActivityLog;

                // Initialize Tools->Options Page
                _options = (OptionsPageGeneral)GetDialogPage(typeof(OptionsPageGeneral));

                // Initialize Solution Service Events
                _solutionService = (IVsSolution2)GetGlobalService(typeof(SVsSolution));
                _solutionService?.AdviseSolutionEvents(this, out _solutionEventsCookie);

                var debuggerEvents = _dte.Events.DebuggerEvents;
                debuggerEvents.OnEnterRunMode += new _dispDebuggerEvents_OnEnterRunModeEventHandler(OnEnterRunMode);
                debuggerEvents.OnEnterDesignMode += new _dispDebuggerEvents_OnEnterDesignModeEventHandler(OnEnterDesignMode);

                //InitializeOutputWindowPane
                if (_dte.ToolWindows.OutputWindow.OutputWindowPanes.Cast<OutputWindowPane>().All(p => p.Name != _options.OutputPane))
                {
                    _dte.ToolWindows.OutputWindow.OutputWindowPanes.Add(_options.OutputPane);
                }

                InitializeMenus();
                InitializeAutoShelve();
            }
            catch (Exception ex)
            {
                WriteException(ex);
            }
        }

        /// <summary>
        /// <para>
        /// Register your own package service
        /// http://technet.microsoft.com/en-us/office/bb164693(v=vs.71).aspx
        /// http://blogs.msdn.com/b/aaronmar/archive/2004/03/12/88646.aspx
        /// http://social.msdn.microsoft.com/Forums/vstudio/en-US/be755076-6e07-4025-93e7-514cd4019dcb/register-own-service?forum=vsx
        /// IVsRunningDocumentTable rdt = Package.GetGlobalService(typeof(SVsRunningDocumentTable)) as IVsRunningDocumentTable;
        /// rdt.AdviseRunningDocTableEvents(new YourRunningDocTableEvents());
        /// rdt.GetDocumentInfo(docCookie, ...)
        /// One of the out params is RDT_ProjSlnDocument; this will be set for your solution file. Note this flag also covers projects. Once you have sufficiently determined it is your solution you're set.
        /// </para>
        /// <para>
        /// http://msdn.microsoft.com/en-us/library/microsoft.visualstudio.shell.interop.ivsrunningdoctableevents.onaftersave.aspx
        /// http://social.msdn.microsoft.com/Forums/vstudio/it-IT/fd513e71-bb23-4de0-b631-35bfbdfdd4f5/visual-studio-isolated-shell-onsolutionsaved-event?forum=vsx
        /// </para>
        /// </summary>
        private void InitializeAutoShelve()
        {
            _autoShelve = GetGlobalService(typeof(ISAutoShelve)) as TfsAutoShelve;
            if (_autoShelve != null)
            {
                // Property Initialization
                _autoShelve.MaximumShelvesets = _options.MaximumShelvesets;
                _autoShelve.ShelvesetName = _options.ShelvesetName;
                _autoShelve.Interval = _options.Interval;
            }
            AttachEvents();
        }

        private void InitializeMenus()
        {
            if (GetService(typeof(IMenuCommandService)) is OleMenuCommandService mcs)
            {
                var commandId = new CommandID(GuidList.GuidAutoShelveCmdSet, PkgCmdIdList.CmdidAutoShelve);
                var oleMenuCommand = new OleMenuCommand(MenuItemCallbackAutoShelveRunState, commandId)
                {
                    Text = _menuTextStopped
                };
                _menuRunState = oleMenuCommand;
                mcs.AddCommand(_menuRunState);

                var commandId1 = new CommandID(GuidList.GuidAutoShelveCmdSet, PkgCmdIdList.CmdidAutoShelveNow);
                var menuAutoShelveNow = new OleMenuCommand(MenuItemCallbackRunNow, commandId1);
                mcs.AddCommand(menuAutoShelveNow);
            }
        }

        #endregion

        #region Package Menu Commands

        private void MenuItemCallbackAutoShelveRunState(object sender, System.EventArgs e)
        {
            try
            {
                _isPaused = false; // this prevents un-pause following a manual start/stop
                if (_autoShelve.IsRunning)
                {
                    _autoShelve.Stop();
                }
                else
                {
                    _autoShelve.Start();
                }
            }
            catch
            {
                // swallow exceptions
            }
        }

        private void MenuItemCallbackRunNow(object sender, System.EventArgs e)
        {
            _autoShelve.CreateShelveset(true);
        }

        #endregion

        #region Local Methods

        private void AttachEvents()
        {
            if (_autoShelve != null)
            {
                _autoShelve.Stopped += AutoShelve_OnStop;
                _autoShelve.Started += AutoShelve_OnStart;
                _autoShelve.ShelvesetCreated += AutoShelve_OnShelvesetCreated;
                _autoShelve.TfsConnectionErrorReceived += AutoShelve_OnTfsConnectionError;
            }
            if (_options != null)
            {
                _options.OnOptionsChanged += Options_OnOptionsChanged;
            }
        }

        private void DetachEvents()
        {
            if (_autoShelve != null)
            {
                _autoShelve.Stopped -= AutoShelve_OnStop;
                _autoShelve.Started -= AutoShelve_OnStart;
                _autoShelve.ShelvesetCreated -= AutoShelve_OnShelvesetCreated;
                _autoShelve.TfsConnectionErrorReceived -= AutoShelve_OnTfsConnectionError;
            }
            if (_options != null)
            {
                _options.OnOptionsChanged -= Options_OnOptionsChanged;
            }
        }

        private void Options_OnOptionsChanged(object sender, OptionsChangedEventArgs e)
        {
            if (_autoShelve != null)
            {
                _autoShelve.MaximumShelvesets = e.MaximumShelvesets;
                _autoShelve.ShelvesetName = e.ShelvesetName;
                _autoShelve.Interval = e.Interval;
            }
        }

        private void ToggleMenuCommandRunStateText(object sender)
        {
            try
            {
                if (sender is OleMenuCommand menuCommand && menuCommand.CommandID.Guid == GuidList.GuidAutoShelveCmdSet)
                {
                    menuCommand.Text = _autoShelve.IsRunning ? _menuTextRunning : _menuTextStopped;
                }
            }
            catch
            {
                // swallow exceptions 
            }
        }

        public void WriteToActivityLog(string message, string stackTrace)
        {
            try
            {
                _log?.LogEntry(3, "VsExtAutoShelvePackage", string.Format(CultureInfo.CurrentCulture, "Message: {0} Stack Trace: {1}", message, stackTrace));
            }
            catch
            {
                // swallow exceptions
            }
        }

        private void WriteException(Exception ex)
        {
            WriteToStatusBar(string.Format(Resources.ErrorMessage, ExtensionName));
            WriteLineToOutputWindow(ex.Message);
            WriteLineToOutputWindow(ex.StackTrace);
            WriteToActivityLog(ex.Message, ex.StackTrace);
        }

        private void WriteLineToOutputWindow(string outputText)
        {
            WriteToOutputWindow(outputText, true);
        }

        private void WriteToOutputWindow(string outputText, bool newLine = false)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(_options.OutputPane))
                {
                    var oWindow = _dte.ToolWindows.OutputWindow.OutputWindowPanes.Item(_options.OutputPane);
                    oWindow.OutputString(outputText);
                    if (newLine)
                        oWindow.OutputString(Environment.NewLine);
                }
            }
            catch
            {
                // swallow exceptions
            }
        }

        private void WriteToStatusBar(string text)
        {
            try
            {
                _dte.StatusBar.Text = text;
            }
            catch
            {
                // swallow exceptions
            }
        }

        #endregion

        #region DebuggerEvents

        private void OnEnterRunMode(dbgEventReason Reason)
        {
            if (_options.PauseWhileDebugging && !_isPaused)
            {
                _isPaused = true;
                _autoShelve.Stop();
            }
        }

        private void OnEnterDesignMode(dbgEventReason Reason)
        {
            if (_isPaused)
            {
                _autoShelve.CreateShelveset();
                _autoShelve.Start();
                _isPaused = false;
            }
        }

        #endregion

        #region IVsSolutionEvents

        public int OnAfterCloseSolution(object pUnkReserved)
        {
            if (_autoShelve?.IsRunning == true)
            {
                _autoShelve.CreateShelveset();
            }
            return 0;
        }

        public int OnAfterLoadProject(IVsHierarchy pStubHierarchy, IVsHierarchy pRealHierarchy) { return 0; }

        public int OnAfterOpenProject(IVsHierarchy pHierarchy, int fAdded) { return 0; }

        public int OnAfterOpenSolution(object pUnkReserved, int fNewSolution)
        {
            if (_autoShelve?.IsRunning == false)
            {
                _autoShelve.Start();
            }
            return 0;
        }

        public int OnBeforeCloseProject(IVsHierarchy pHierarchy, int fRemoved) { return 0; }

        public int OnBeforeCloseSolution(object pUnkReserved) { return 0; }

        public int OnBeforeUnloadProject(IVsHierarchy pRealHierarchy, IVsHierarchy pStubHierarchy) { return 0; }

        public int OnQueryCloseProject(IVsHierarchy pHierarchy, int fRemoving, ref int pfCancel) { return 0; }

        public int OnQueryCloseSolution(object pUnkReserved, ref int pfCancel) { return 0; }

        public int OnQueryUnloadProject(IVsHierarchy pRealHierarchy, ref int pfCancel) { return 0; }

        #endregion

        #region IDisposable

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~VsExtAutoShelvePackage()
        {
            // Finalizer calls Dispose(false)
            Dispose(false);
        }

        // The bulk of the clean-up code is implemented in Dispose(bool)
        protected override void Dispose(bool disposing)
        {
            try
            {
                DetachEvents();

                if (_solutionService != null && _solutionEventsCookie != 0)
                {
                    _solutionService.UnadviseSolutionEvents(_solutionEventsCookie);
                    _solutionEventsCookie = 0;
                    _solutionService = null;
                }
            }
            catch
            {
                // swallow exceptions
            }
            base.Dispose(disposing);
        }

        #endregion

    }
}