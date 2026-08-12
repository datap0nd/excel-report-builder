using System;
using System.Windows;
using System.Windows.Controls;
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
            await viewModel.InitializeAsync();
        }
    }
}
