using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using System.Windows.Interop;
using WpfRenderOptions = System.Windows.Media.RenderOptions;
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
        private readonly PivotPlusView _pivotPlusView;

        public TaskPaneHost()
        {
            // Office custom task panes host WPF through a WinForms ActiveX control.
            // Some Office/graphics-driver combinations expose the ElementHost as a
            // black surface when WPF uses hardware composition. The pane is small
            // and interaction-heavy, so software composition is the reliable choice.
            WpfRenderOptions.ProcessRenderMode = RenderMode.SoftwareOnly;

            AutoScaleMode = AutoScaleMode.Dpi;
            BackColor = Color.White;
            AccessibleName = "PivotTable Plus task pane";
            AccessibleDescription = "Edit the selected native PivotTable and add validated extras.";
            AccessibleRole = AccessibleRole.Client;
            MinimumSize = new Size(320, 480);
            TabStop = true;

            _pivotPlusView = new PivotPlusView(TaskPaneBootstrapper.CreateHostService());
            _elementHost = new ElementHost
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Child = _pivotPlusView,
                TabStop = true
            };

            Controls.Add(_elementHost);
        }

        [ComRegisterFunction]
        public static void Register(Type registeredType)
        {
            string controlKeyPath = @"CLSID\" + registeredType.GUID.ToString("B") + @"\Control";
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
            string controlKeyPath = @"CLSID\" + registeredType.GUID.ToString("B") + @"\Control";
            Registry.ClassesRoot.DeleteSubKey(controlKeyPath, throwOnMissingSubKey: false);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _pivotPlusView.Dispose();
                _elementHost.Child = null;
                _elementHost.Dispose();
            }

            base.Dispose(disposing);
        }
    }
}
