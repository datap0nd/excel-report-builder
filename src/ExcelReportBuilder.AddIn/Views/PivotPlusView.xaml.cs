using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using ExcelReportBuilder.AddIn.Host;
using ExcelReportBuilder.AddIn.Presentation;

namespace ExcelReportBuilder.AddIn.Views
{
    public partial class PivotPlusView : UserControl, IDisposable
    {
        private readonly PivotPlusViewModel viewModel;
        private bool initialized;
        private bool disposed;

        public PivotPlusView()
            : this(new SyntheticPivotPlusHostService())
        {
        }

        public PivotPlusView(IPivotPlusHostService hostService)
        {
            InitializeComponent();
            viewModel = new PivotPlusViewModel(hostService);
            DataContext = viewModel;
            Loaded += OnLoaded;
        }

        public PivotPlusViewModel ViewModel => viewModel;

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            Loaded -= OnLoaded;
            viewModel.Dispose();
        }

        private async void OnLoaded(object sender, RoutedEventArgs eventArgs)
        {
            if (initialized || disposed) return;
            initialized = true;

            // Office raises Loaded while it is still servicing the Ribbon callback
            // that creates and shows the custom task pane. Inspecting Excel's COM
            // object model synchronously at that point can block the host before WPF
            // receives its first render pass. Resume only after the callback has
            // unwound and higher-priority layout/render work has completed.
            await Dispatcher.Yield(DispatcherPriority.ContextIdle);
            if (disposed) return;

            await viewModel.InitializeAsync();
        }
    }
}
