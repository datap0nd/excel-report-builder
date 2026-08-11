using System;
using System.Runtime.InteropServices;
using ExcelReportBuilder.AddIn.Host;
using ExcelReportBuilder.AddIn.Hosting;
using ExcelReportBuilder.AddIn.Interop;
using ExcelReportBuilder.AddIn.Ribbon;
using Microsoft.Win32;

namespace ExcelReportBuilder.AddIn.Com
{
    [ComVisible(true)]
    [Guid(ClassId)]
    [ProgId(ProgramId)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IExcelReportBuilderRibbonCallbacks))]
    public sealed class ExcelReportBuilderAddIn :
        IDTExtensibility2,
        IRibbonExtensibility,
        ICustomTaskPaneConsumer,
        IExcelReportBuilderRibbonCallbacks
    {
        public const string ProgramId = "ExcelReportBuilder.AddIn";
        public const string ClassId = "F953480C-A73C-4121-9E21-18676EC34CE8";

        private const string AddInRegistryPath = @"Software\Microsoft\Office\Excel\Addins\" + ProgramId;
        private const string TaskPaneTitle = "Excel Report Builder";
        private const int DefaultTaskPaneWidth = 420;

        private object? _application;
        private object? _addInInstance;
        private ICTPFactory? _taskPaneFactory;
        private object? _taskPane;
        private object? _ribbonUi;
        private bool _showWhenFactoryIsReady;

        public string GetCustomUI(string ribbonId)
        {
            return RibbonMarkup.CustomUi;
        }

        public void OnConnection(
            object application,
            ExtConnectMode connectMode,
            object addInInstance,
            ref Array custom)
        {
            _application = application;
            _addInInstance = addInInstance;
            TaskPaneBootstrapper.HostServiceFactory =
                () => new ExcelReportBuilderHostService(application);
        }

        public void OnDisconnection(ExtDisconnectMode removeMode, ref Array custom)
        {
            TaskPaneBootstrapper.HostServiceFactory = null;
            TearDownTaskPane();

            _ribbonUi = null;
            _taskPaneFactory = null;
            _addInInstance = null;
            _application = null;
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
            InvalidateRibbon();
        }

        public void OnBeginShutdown(ref Array custom)
        {
            TaskPaneBootstrapper.HostServiceFactory = null;
            TearDownTaskPane();
        }

        public void CTPFactoryAvailable(ICTPFactory taskPaneFactory)
        {
            _taskPaneFactory = taskPaneFactory;

            if (_showWhenFactoryIsReady)
            {
                SetTaskPaneVisible(true);
            }

            InvalidateRibbon();
        }

        // Office resolves ribbon callbacks through IDispatch on this COM-visible class.
        public void OnRibbonLoad(object ribbonUi)
        {
            _ribbonUi = ribbonUi;
            InvalidateRibbon();
        }

        public void OnToggleTaskPane(object control, bool isPressed)
        {
            _showWhenFactoryIsReady = isPressed;
            SetTaskPaneVisible(isPressed);
            InvalidateRibbon();
        }

        public bool GetTaskPanePressed(object control)
        {
            return IsTaskPaneVisible();
        }

        public bool GetTaskPaneEnabled(object control)
        {
            return _taskPaneFactory != null;
        }

        [ComRegisterFunction]
        public static void Register(Type registeredType)
        {
            using (RegistryKey? key = Registry.CurrentUser.CreateSubKey(AddInRegistryPath))
            {
                if (key == null)
                {
                    throw new InvalidOperationException("Could not create the per-user Excel add-in registration key.");
                }

                key.SetValue("Description", "Build checked dense management reports from workbook data.");
                key.SetValue("FriendlyName", "Excel Report Builder");
                key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
            }
        }

        [ComUnregisterFunction]
        public static void Unregister(Type registeredType)
        {
            Registry.CurrentUser.DeleteSubKeyTree(AddInRegistryPath, throwOnMissingSubKey: false);
        }

        private void SetTaskPaneVisible(bool visible)
        {
            if (visible && _taskPane == null)
            {
                TryCreateTaskPane();
            }

            if (_taskPane != null)
            {
                LateBoundCom.TrySetProperty(_taskPane, "Visible", visible);
            }
        }

        private void TryCreateTaskPane()
        {
            if (_taskPaneFactory == null || _taskPane != null)
            {
                return;
            }

            object parentWindow = Type.Missing;
            if (LateBoundCom.TryGetProperty(_application, "ActiveWindow", out object? activeWindow)
                && activeWindow != null)
            {
                parentWindow = activeWindow;
            }

            try
            {
                _taskPane = _taskPaneFactory.CreateCTP(
                    TaskPaneHost.ProgramId,
                    TaskPaneTitle,
                    parentWindow);

                LateBoundCom.TrySetProperty(_taskPane, "Width", DefaultTaskPaneWidth);
            }
            catch (Exception exception) when (
                exception is COMException
                || exception is InvalidOperationException
                || exception is MissingMemberException)
            {
                _taskPane = null;
                _showWhenFactoryIsReady = false;
            }
        }

        private bool IsTaskPaneVisible()
        {
            if (LateBoundCom.TryGetProperty(_taskPane, "Visible", out object? value)
                && value is bool isVisible)
            {
                return isVisible;
            }

            return false;
        }

        private void InvalidateRibbon()
        {
            LateBoundCom.TryInvoke(_ribbonUi, "Invalidate");
        }

        private void TearDownTaskPane()
        {
            if (_taskPane == null)
            {
                return;
            }

            LateBoundCom.TrySetProperty(_taskPane, "Visible", false);
            LateBoundCom.TryInvoke(_taskPane, "Delete");
            LateBoundCom.FinalRelease(_taskPane);
            _taskPane = null;
        }
    }
}
