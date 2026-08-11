using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using ExcelReportBuilder.AddIn.Views;
using Microsoft.Win32;

namespace ExcelReportBuilder.AddIn.Hosting
{
    [ComVisible(true)]
    [Guid(ClassId)]
    [ProgId(ProgramId)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class TaskPaneHost : UserControl
    {
        public const string ProgramId = "ExcelReportBuilder.TaskPaneHost";
        public const string ClassId = "A3F4E10D-0DD1-420E-8B6F-E0A654BBEA16";

        private readonly ElementHost _elementHost;
        private readonly ReportBuilderView _reportBuilderView;

        public TaskPaneHost()
        {
            AutoScaleMode = AutoScaleMode.Dpi;
            BackColor = Color.White;
            AccessibleName = "Excel Report Builder task pane";
            AccessibleDescription = "Choose data, build a report, use chat, and review checks.";
            AccessibleRole = AccessibleRole.Client;
            MinimumSize = new Size(320, 480);
            TabStop = true;

            _reportBuilderView = new ReportBuilderView(TaskPaneBootstrapper.CreateHostService());
            _elementHost = new ElementHost
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Child = _reportBuilderView,
                TabStop = true
            };

            Controls.Add(_elementHost);
        }

        [ComRegisterFunction]
        public static void Register(Type registeredType)
        {
            string controlKeyPath = $@"CLSID\{{{registeredType.GUID:D}}}\Control";
            using (RegistryKey? controlKey = Registry.ClassesRoot.CreateSubKey(controlKeyPath))
            {
                if (controlKey == null)
                {
                    throw new InvalidOperationException("Could not mark the task-pane host as an ActiveX control.");
                }
            }
        }

        [ComUnregisterFunction]
        public static void Unregister(Type registeredType)
        {
            string controlKeyPath = $@"CLSID\{{{registeredType.GUID:D}}}\Control";
            Registry.ClassesRoot.DeleteSubKey(controlKeyPath, throwOnMissingSubKey: false);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _reportBuilderView.Dispose();
                _elementHost.Child = null;
                _elementHost.Dispose();
            }

            base.Dispose(disposing);
        }
    }
}
