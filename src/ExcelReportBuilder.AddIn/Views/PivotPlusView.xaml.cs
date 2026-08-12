using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using ExcelReportBuilder.AddIn.Host;
using ExcelReportBuilder.AddIn.Presentation;
using ExcelReportBuilder.Core.PivotPlus;

namespace ExcelReportBuilder.AddIn.Views
{
    public partial class PivotPlusView : UserControl, IDisposable
    {
        private readonly PivotPlusViewModel viewModel;
        private readonly DispatcherTimer syncTimer;
        private Point dragStart;
        private PivotPlusFieldRow? pendingFieldDrag;
        private PivotPlusPlacementRow? pendingPlacementDrag;
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
            syncTimer = new DispatcherTimer(DispatcherPriority.Background)
            {
                Interval = TimeSpan.FromSeconds(2)
            };
            syncTimer.Tick += OnSyncTimerTick;
            Loaded += OnLoaded;
        }

        public PivotPlusViewModel ViewModel => viewModel;

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            Loaded -= OnLoaded;
            syncTimer.Stop();
            syncTimer.Tick -= OnSyncTimerTick;
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
            if (!disposed) syncTimer.Start();
        }

        private async void OnSyncTimerTick(object? sender, EventArgs eventArgs)
        {
            if (disposed) return;
            await viewModel.SyncAsync();
        }

        private void OnDragOriginMouseDown(object sender, MouseButtonEventArgs eventArgs)
        {
            dragStart = eventArgs.GetPosition(this);
            DependencyObject? source = eventArgs.OriginalSource as DependencyObject;
            pendingFieldDrag = FindItem<PivotPlusFieldRow>(source);
            pendingPlacementDrag = FindItem<PivotPlusPlacementRow>(source);
        }

        private void OnFieldListMouseMove(object sender, MouseEventArgs eventArgs)
        {
            if (eventArgs.LeftButton != MouseButtonState.Pressed ||
                !ExceededDragThreshold(eventArgs.GetPosition(this)) ||
                sender is not ListBox listBox)
            {
                return;
            }

            PivotPlusFieldRow? hit = FindItem<PivotPlusFieldRow>(
                eventArgs.OriginalSource as DependencyObject) ?? pendingFieldDrag;
            if (hit == null) return;
            IReadOnlyList<PivotPlusFieldRow> fields = listBox.SelectedItems
                .OfType<PivotPlusFieldRow>()
                .Where(item => ReferenceEquals(item, hit) || listBox.SelectedItems.Contains(hit))
                .ToList();
            if (fields.Count == 0 || !fields.Contains(hit)) fields = new[] { hit };

            pendingFieldDrag = null;
            DragDrop.DoDragDrop(listBox, new PivotPlusFieldDragData(fields), DragDropEffects.Copy);
        }

        private void OnPlacementListMouseMove(object sender, MouseEventArgs eventArgs)
        {
            if (eventArgs.LeftButton != MouseButtonState.Pressed ||
                !ExceededDragThreshold(eventArgs.GetPosition(this)) ||
                sender is not ListBox listBox)
            {
                return;
            }

            PivotPlusPlacementRow? placement = FindItem<PivotPlusPlacementRow>(
                eventArgs.OriginalSource as DependencyObject) ?? pendingPlacementDrag;
            if (placement == null) return;
            pendingPlacementDrag = null;
            DragDrop.DoDragDrop(
                listBox,
                new PivotPlusPlacementDragData(placement),
                DragDropEffects.Move);
        }

        private void OnAreaDragOver(object sender, DragEventArgs eventArgs)
        {
            eventArgs.Effects = eventArgs.Data.GetDataPresent(typeof(PivotPlusFieldDragData))
                ? DragDropEffects.Copy
                : eventArgs.Data.GetDataPresent(typeof(PivotPlusPlacementDragData))
                    ? DragDropEffects.Move
                    : DragDropEffects.None;
            eventArgs.Handled = true;
        }

        private void OnAreaDrop(object sender, DragEventArgs eventArgs)
        {
            if (sender is not ListBox listBox ||
                !Enum.TryParse(Convert.ToString(listBox.Tag), true, out PivotFieldArea area))
            {
                return;
            }

            int insertionIndex = FindInsertionIndex(listBox, eventArgs.GetPosition(listBox));
            if (eventArgs.Data.GetData(typeof(PivotPlusFieldDragData)) is PivotPlusFieldDragData fields)
            {
                viewModel.DropFields(fields.Fields, area, insertionIndex);
                eventArgs.Effects = DragDropEffects.Copy;
            }
            else if (eventArgs.Data.GetData(typeof(PivotPlusPlacementDragData)) is PivotPlusPlacementDragData placement)
            {
                viewModel.MovePlacement(placement.Placement, area, insertionIndex);
                eventArgs.Effects = DragDropEffects.Move;
            }

            eventArgs.Handled = true;
        }

        private void OnPlacementKeyDown(object sender, KeyEventArgs eventArgs)
        {
            if (eventArgs.Key == Key.Delete && sender is ListBox { SelectedItem: PivotPlusPlacementRow placement })
            {
                viewModel.RemovePlacement(placement);
                eventArgs.Handled = true;
            }
        }

        private bool ExceededDragThreshold(Point current)
        {
            return Math.Abs(current.X - dragStart.X) >= SystemParameters.MinimumHorizontalDragDistance ||
                   Math.Abs(current.Y - dragStart.Y) >= SystemParameters.MinimumVerticalDragDistance;
        }

        private static T? FindItem<T>(DependencyObject? source) where T : class
        {
            DependencyObject? current = source;
            while (current != null)
            {
                if (current is FrameworkElement { DataContext: T item }) return item;
                current = VisualTreeHelper.GetParent(current);
            }

            return null;
        }

        private static int FindInsertionIndex(ListBox listBox, Point point)
        {
            for (int index = 0; index < listBox.Items.Count; index++)
            {
                if (listBox.ItemContainerGenerator.ContainerFromIndex(index) is ListBoxItem item)
                {
                    Point relative = item.TranslatePoint(new Point(0, item.ActualHeight / 2), listBox);
                    if (point.Y < relative.Y) return index;
                }
            }

            return listBox.Items.Count;
        }

        private sealed class PivotPlusFieldDragData
        {
            public PivotPlusFieldDragData(IReadOnlyList<PivotPlusFieldRow> fields)
            {
                Fields = fields;
            }

            public IReadOnlyList<PivotPlusFieldRow> Fields { get; }
        }

        private sealed class PivotPlusPlacementDragData
        {
            public PivotPlusPlacementDragData(PivotPlusPlacementRow placement)
            {
                Placement = placement;
            }

            public PivotPlusPlacementRow Placement { get; }
        }
    }
}
