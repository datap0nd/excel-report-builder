using System;
using System.Collections.Specialized;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using ExcelReportBuilder.AddIn.Host;
using ExcelReportBuilder.AddIn.Presentation;

namespace ExcelReportBuilder.AddIn.Views
{
    public partial class ReportBuilderView : UserControl, IDisposable
    {
        private readonly ShellViewModel _viewModel;
        private bool _disposed;

        public ReportBuilderView()
            : this(new SyntheticReportBuilderHostService())
        {
        }

        public ReportBuilderView(IReportBuilderHostService hostService)
        {
            InitializeComponent();
            _viewModel = new ShellViewModel(hostService);
            DataContext = _viewModel;
            _viewModel.Activity.Entries.CollectionChanged += OnActivityEntriesChanged;
            _viewModel.ApiKeyClearRequested += OnApiKeyClearRequested;
        }

        public ShellViewModel ViewModel => _viewModel;

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            _viewModel.Activity.Entries.CollectionChanged -= OnActivityEntriesChanged;
            _viewModel.ApiKeyClearRequested -= OnApiKeyClearRequested;
            _viewModel.Dispose();
        }

        private void OnActivityEntriesChanged(object? sender, NotifyCollectionChangedEventArgs eventArgs)
        {
            if (eventArgs.NewItems == null || eventArgs.NewItems.Count == 0)
            {
                return;
            }

            object newestEntry = eventArgs.NewItems[eventArgs.NewItems.Count - 1]!;
            Dispatcher.BeginInvoke(
                new Action(() => ActivityList.ScrollIntoView(newestEntry)),
                DispatcherPriority.Background);
        }

        private void OnApiKeyChanged(object sender, RoutedEventArgs eventArgs)
        {
            if (_disposed || !(sender is PasswordBox passwordBox))
            {
                return;
            }

            _viewModel.SetApiKey(passwordBox.SecurePassword);
        }

        private void OnApiKeyClearRequested(object? sender, EventArgs eventArgs)
        {
            if (!_disposed && ApiKeyBox.Password.Length != 0)
            {
                ApiKeyBox.Password = string.Empty;
            }
        }
    }
}
