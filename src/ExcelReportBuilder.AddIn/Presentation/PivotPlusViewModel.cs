using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Input;
using ExcelReportBuilder.AddIn.Host;
using ExcelReportBuilder.Core.PivotPlus;

namespace ExcelReportBuilder.AddIn.Presentation
{
    public sealed class PivotPlusFieldRow
    {
        public PivotPlusFieldRow(PivotPlusFieldSnapshot source)
        {
            Source = source ?? throw new ArgumentNullException(nameof(source));
        }

        internal PivotPlusFieldSnapshot Source { get; }
        public string Name => Source.Name;
        public string Caption => Source.Caption;
        public bool IsMeasure => Source.IsMeasure;
        public string KindLabel => Source.IsMeasure ? "Measure" : "Field";

        public bool IsPeriodField => PivotPlusPeriodFieldClassifier.IsPeriodField(Caption);

        public int PeriodOrder => PivotPlusPeriodFieldClassifier.PeriodOrder(Caption);
    }

    public sealed class PivotPlusPlacementRow
    {
        public PivotPlusPlacementRow(PivotPlusPlacementSnapshot source)
        {
            FieldName = source.FieldName;
            Caption = source.Caption;
            Area = source.Area;
            Aggregation = source.Aggregation;
            NumberFormatCode = source.NumberFormatCode;
        }

        public PivotPlusPlacementRow(PivotPlusFieldRow field, PivotFieldArea area)
        {
            FieldName = field.Name;
            Caption = field.Caption;
            Area = area;
            Aggregation = area == PivotFieldArea.Values
                ? PivotAggregationFunction.Sum
                : (PivotAggregationFunction?)null;
            NumberFormatCode = area == PivotFieldArea.Values ? "#,##0" : string.Empty;
        }

        public string FieldName { get; }
        public string Caption { get; }
        public PivotFieldArea Area { get; }
        public PivotAggregationFunction? Aggregation { get; }
        public string NumberFormatCode { get; }
    }

    public sealed class PivotPlusViewModel : ObservableObject, IDisposable
    {
        private readonly IPivotPlusHostService hostService;
        private readonly CancellationTokenSource lifetime = new CancellationTokenSource();
        private readonly List<PivotPlusFieldRow> allFields = new List<PivotPlusFieldRow>();
        private PivotPlusPaneSnapshot? snapshot;
        private PivotPlusFieldRow? selectedField;
        private PivotPlusPlacementRow? selectedPlacement;
        private PivotPlusFieldRow? selectedPortionValueField;
        private PivotPlusFieldRow? selectedPortionDetailField;
        private PivotPlusFieldRow? selectedPeriodField;
        private string portionCaption = "Portion %";
        private string searchText = string.Empty;
        private string targetLabel = "Select a cell inside a PivotTable";
        private string sourceKindLabel = "No PivotTable selected";
        private string statusMessage = "Select a PivotTable, then choose Refresh.";
        private bool isBusy;
        private bool hasPendingChanges;
        private bool disposed;
        private string snapshotSignature = string.Empty;

        public PivotPlusViewModel(IPivotPlusHostService hostService)
        {
            this.hostService = hostService ?? throw new ArgumentNullException(nameof(hostService));
            RefreshCommand = new AsyncRelayCommand(RefreshAsync, () => !IsBusy);
            AddFieldCommand = new RelayCommand(AddSelectedField, CanAddSelectedField);
            RemoveFieldCommand = new RelayCommand(RemoveSelectedPlacement, () => SelectedPlacement != null && !IsBusy);
            MoveUpCommand = new RelayCommand(() => MoveSelectedPlacement(-1), () => CanMoveSelected(-1));
            MoveDownCommand = new RelayCommand(() => MoveSelectedPlacement(1), () => CanMoveSelected(1));
            ApplyCommand = new AsyncRelayCommand(ApplyAsync, () => HasPendingChanges && !IsBusy);
            EnableDataModelCommand = new AsyncRelayCommand(EnableDataModelAsync, () => CanEnableDataModel && !IsBusy);
            AddPortionCommand = new AsyncRelayCommand(AddPortionAsync, CanAddPortion);
            UndoCommand = new AsyncRelayCommand(UndoAsync, () => snapshot != null && !IsBusy);
            GroupDateCommand = new AsyncRelayCommand(GroupDateAsync, CanGroupDate);
            OpenExcelFieldListCommand = new RelayCommand(OpenExcelFieldList, () => !IsBusy);
        }

        public ObservableCollection<PivotPlusFieldRow> Fields { get; } = new ObservableCollection<PivotPlusFieldRow>();
        public ObservableCollection<PivotPlusFieldRow> PortionValueFields { get; } = new ObservableCollection<PivotPlusFieldRow>();
        public ObservableCollection<PivotPlusFieldRow> PortionDetailFields { get; } = new ObservableCollection<PivotPlusFieldRow>();
        public ObservableCollection<PivotPlusFieldRow> PeriodFields { get; } = new ObservableCollection<PivotPlusFieldRow>();
        public ObservableCollection<PivotPlusPlacementRow> Filters { get; } = new ObservableCollection<PivotPlusPlacementRow>();
        public ObservableCollection<PivotPlusPlacementRow> Columns { get; } = new ObservableCollection<PivotPlusPlacementRow>();
        public ObservableCollection<PivotPlusPlacementRow> Rows { get; } = new ObservableCollection<PivotPlusPlacementRow>();
        public ObservableCollection<PivotPlusPlacementRow> Values { get; } = new ObservableCollection<PivotPlusPlacementRow>();

        public ICommand RefreshCommand { get; }
        public ICommand AddFieldCommand { get; }
        public ICommand RemoveFieldCommand { get; }
        public ICommand MoveUpCommand { get; }
        public ICommand MoveDownCommand { get; }
        public ICommand ApplyCommand { get; }
        public ICommand EnableDataModelCommand { get; }
        public ICommand AddPortionCommand { get; }
        public ICommand UndoCommand { get; }
        public ICommand GroupDateCommand { get; }
        public ICommand OpenExcelFieldListCommand { get; }

        public string SearchText
        {
            get => searchText;
            set
            {
                if (SetProperty(ref searchText, value ?? string.Empty))
                {
                    RefreshFieldFilter();
                }
            }
        }

        public PivotPlusFieldRow? SelectedField
        {
            get => selectedField;
            set
            {
                if (SetProperty(ref selectedField, value)) RaiseCommandStates();
            }
        }

        public PivotPlusPlacementRow? SelectedPlacement
        {
            get => selectedPlacement;
            set
            {
                if (SetProperty(ref selectedPlacement, value)) RaiseCommandStates();
            }
        }

        public PivotPlusFieldRow? SelectedPortionValueField
        {
            get => selectedPortionValueField;
            set
            {
                if (SetProperty(ref selectedPortionValueField, value)) RaiseCommandStates();
            }
        }

        public PivotPlusFieldRow? SelectedPortionDetailField
        {
            get => selectedPortionDetailField;
            set
            {
                if (SetProperty(ref selectedPortionDetailField, value)) RaiseCommandStates();
            }
        }

        public string PortionCaption
        {
            get => portionCaption;
            set => SetProperty(ref portionCaption, value ?? string.Empty);
        }

        public string TargetLabel
        {
            get => targetLabel;
            private set => SetProperty(ref targetLabel, value);
        }

        public string SourceKindLabel
        {
            get => sourceKindLabel;
            private set => SetProperty(ref sourceKindLabel, value);
        }

        public string StatusMessage
        {
            get => statusMessage;
            private set => SetProperty(ref statusMessage, value);
        }

        public bool IsBusy
        {
            get => isBusy;
            private set
            {
                if (SetProperty(ref isBusy, value)) RaiseCommandStates();
            }
        }

        public bool HasPendingChanges
        {
            get => hasPendingChanges;
            private set
            {
                if (SetProperty(ref hasPendingChanges, value)) RaiseCommandStates();
            }
        }

        public bool SupportsExtras => snapshot?.SupportsExtras == true;
        public bool CanEnableDataModel => snapshot?.CanEnableDataModel == true;

        public bool HasPeriodFields => PeriodFields.Count > 0;

        public string PeriodSummary => HasPeriodFields
            ? PeriodFields.Count + " date/period fields detected"
            : "No date/period fields detected";

        public Task InitializeAsync()
        {
            return RefreshAsync();
        }

        public PivotPlusFieldRow? SelectedPeriodField
        {
            get => selectedPeriodField;
            set
            {
                if (SetProperty(ref selectedPeriodField, value)) RaiseCommandStates();
            }
        }

        public async Task SyncAsync()
        {
            if (disposed || IsBusy || HasPendingChanges) return;

            try
            {
                PivotPlusPaneSnapshot next = await hostService
                    .InspectAsync(lifetime.Token)
                    .ConfigureAwait(true);
                string signature = SnapshotSignature(next);
                if (!string.Equals(signature, snapshotSignature, StringComparison.Ordinal))
                {
                    LoadSnapshot(next);
                    StatusMessage = "Synced with Excel.";
                }
            }
            catch (OperationCanceledException) when (lifetime.IsCancellationRequested)
            {
            }
            catch (Exception)
            {
                // Auto-sync is opportunistic. Explicit commands continue to surface
                // actionable errors in the status region.
            }
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            lifetime.Cancel();
            lifetime.Dispose();
        }

        private async Task RefreshAsync()
        {
            await RunHostOperationAsync(
                "Reading the selected PivotTable…",
                token => hostService.InspectAsync(token),
                "Ready. Standard field changes remain a preview until Apply.").ConfigureAwait(true);
        }

        private async Task ApplyAsync()
        {
            IReadOnlyList<PivotPlusPlacementRequest> requests = BuildPlacementRequests();
            await RunHostOperationAsync(
                "Applying the preview to the native PivotTable…",
                token => hostService.ApplyLayoutAsync(requests, token),
                "Applied. Excel's PivotTable and Field List remain native and refreshable.").ConfigureAwait(true);
        }

        private async Task AddPortionAsync()
        {
            PivotPlusFieldRow valueField = SelectedPortionValueField ??
                throw new InvalidOperationException("Choose a numeric value field.");
            PivotPlusFieldRow detailField = SelectedPortionDetailField ??
                throw new InvalidOperationException("Choose the detail row field.");
            await RunHostOperationAsync(
                "Creating a validated parent-portion measure…",
                token => hostService.AddParentPortionAsync(
                    valueField.Name,
                    detailField.Name,
                    PortionCaption,
                    token),
                "Portion added inside the native PivotTable. Use Undo to reverse this session's extra.").ConfigureAwait(true);
        }

        private async Task EnableDataModelAsync()
        {
            await RunHostOperationAsync(
                "Creating a verified Data Model replacement…",
                token => hostService.EnableDataModelAsync(token),
                "PivotTable+ is enabled. The selected object is still one native PivotTable.").ConfigureAwait(true);
        }

        private async Task UndoAsync()
        {
            await RunHostOperationAsync(
                "Restoring the state before the last extra…",
                token => hostService.UndoLastExtraAsync(token),
                "The last PivotTable+ extra was undone.").ConfigureAwait(true);
        }

        private async Task GroupDateAsync(object? parameter)
        {
            PivotPlusFieldRow field = SelectedPeriodField ??
                throw new InvalidOperationException("Choose a date or period field.");
            if (!Enum.TryParse(Convert.ToString(parameter), true, out PivotDateGrouping grouping))
            {
                throw new ArgumentException("Choose Month, Quarter, Year, or Ungrouped.");
            }

            await RunHostOperationAsync(
                "Updating native Excel date grouping…",
                token => hostService.GroupDateAsync(field.Name, grouping, token),
                grouping == PivotDateGrouping.Ungrouped
                    ? "Date grouping removed."
                    : "Dates grouped by " + grouping.ToString().ToLowerInvariant() + ".")
                .ConfigureAwait(true);
        }

        private async Task RunHostOperationAsync(
            string progress,
            Func<CancellationToken, Task<PivotPlusPaneSnapshot>> operation,
            string success)
        {
            if (disposed) return;
            IsBusy = true;
            StatusMessage = progress;
            try
            {
                PivotPlusPaneSnapshot next = await operation(lifetime.Token).ConfigureAwait(true);
                LoadSnapshot(next);
                StatusMessage = success;
            }
            catch (OperationCanceledException) when (lifetime.IsCancellationRequested)
            {
                StatusMessage = "Operation cancelled.";
            }
            catch (Exception exception)
            {
                StatusMessage = FormatOperationError(exception);
            }
            finally
            {
                IsBusy = false;
            }
        }

        private void LoadSnapshot(PivotPlusPaneSnapshot next)
        {
            snapshot = next;
            TargetLabel = next.WorksheetName + " · " + next.PivotTableName;
            SourceKindLabel = GetSourceKindLabel(next.SourceKind);
            allFields.Clear();
            allFields.AddRange(next.Fields.Select(field => new PivotPlusFieldRow(field)));
            RefreshFieldFilter();
            Replace(PortionValueFields, allFields.Where(field => !field.IsMeasure));
            Replace(PortionDetailFields, allFields.Where(field => !field.IsMeasure));
            Replace(
                PeriodFields,
                allFields.Where(field => field.IsPeriodField)
                    .OrderBy(field => field.PeriodOrder)
                    .ThenBy(field => field.Caption, StringComparer.CurrentCultureIgnoreCase));
            Replace(Filters, CreateAreaRows(next, PivotFieldArea.Filter));
            Replace(Columns, CreateAreaRows(next, PivotFieldArea.Column));
            Replace(Rows, CreateAreaRows(next, PivotFieldArea.Row));
            Replace(Values, CreateAreaRows(next, PivotFieldArea.Values));
            SelectedField = Fields.FirstOrDefault();
            SelectedPlacement = null;
            SelectedPortionValueField = PortionValueFields.FirstOrDefault(field =>
                Values.Any(value => string.Equals(value.FieldName, field.Name, StringComparison.OrdinalIgnoreCase)))
                ?? PortionValueFields.FirstOrDefault();
            SelectedPortionDetailField = PortionDetailFields.FirstOrDefault(field =>
                Rows.Any(row => string.Equals(row.FieldName, field.Name, StringComparison.OrdinalIgnoreCase)))
                ?? PortionDetailFields.FirstOrDefault();
            SelectedPeriodField = PeriodFields.FirstOrDefault(field =>
                Columns.Concat(Rows).Any(item =>
                    string.Equals(item.FieldName, field.Name, StringComparison.OrdinalIgnoreCase)))
                ?? PeriodFields.FirstOrDefault();
            HasPendingChanges = false;
            snapshotSignature = SnapshotSignature(next);
            RaisePropertyChanged(nameof(SupportsExtras));
            RaisePropertyChanged(nameof(CanEnableDataModel));
            RaisePropertyChanged(nameof(HasPeriodFields));
            RaisePropertyChanged(nameof(PeriodSummary));
        }

        private static string FormatOperationError(Exception exception)
        {
            if (exception is ExcelReportBuilder.Excel.PivotPlus.PivotMutationException mutation)
            {
                Exception cause = mutation.InnerException ?? mutation;
                while (cause.InnerException != null && cause is not AggregateException)
                {
                    cause = cause.InnerException;
                }

                string detail = cause is AggregateException aggregate
                    ? string.Join(" | ", aggregate.Flatten().InnerExceptions.Select(item => item.Message))
                    : cause.Message;
                return mutation.Message + " Stage: " + mutation.FailedStep + ". " + detail;
            }

            return exception.Message;
        }

        public void DropFields(
            IEnumerable<PivotPlusFieldRow> fields,
            PivotFieldArea area,
            int insertionIndex)
        {
            if (fields == null) return;
            ObservableCollection<PivotPlusPlacementRow> target = AreaCollection(area);
            int nextIndex = Math.Max(0, Math.Min(insertionIndex, target.Count));
            bool changed = false;
            foreach (PivotPlusFieldRow field in fields.Distinct())
            {
                if (!SupportsArea(field.Source.SupportedAreas, area))
                {
                    StatusMessage = field.Caption + " cannot be placed in " + AreaLabel(area) + ".";
                    continue;
                }

                if (area != PivotFieldArea.Values && target.Any(item =>
                    string.Equals(item.FieldName, field.Name, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                target.Insert(nextIndex++, new PivotPlusPlacementRow(field, area));
                changed = true;
            }

            if (changed)
            {
                SelectedPlacement = target[Math.Max(0, nextIndex - 1)];
                MarkPreviewChanged();
            }
        }

        public void MovePlacement(
            PivotPlusPlacementRow placement,
            PivotFieldArea targetArea,
            int insertionIndex)
        {
            if (placement == null) return;
            PivotPlusFieldRow? field = allFields.FirstOrDefault(item =>
                string.Equals(item.Name, placement.FieldName, StringComparison.OrdinalIgnoreCase));
            if (field == null || !SupportsArea(field.Source.SupportedAreas, targetArea))
            {
                StatusMessage = placement.Caption + " cannot be placed in " + AreaLabel(targetArea) + ".";
                return;
            }

            ObservableCollection<PivotPlusPlacementRow> source = AreaCollection(placement.Area);
            ObservableCollection<PivotPlusPlacementRow> target = AreaCollection(targetArea);
            int sourceIndex = source.IndexOf(placement);
            if (sourceIndex < 0) return;
            if (targetArea != PivotFieldArea.Values && target.Any(item =>
                !ReferenceEquals(item, placement) &&
                string.Equals(item.FieldName, placement.FieldName, StringComparison.OrdinalIgnoreCase)))
            {
                StatusMessage = placement.Caption + " is already in " + AreaLabel(targetArea) + ".";
                return;
            }

            source.RemoveAt(sourceIndex);
            int targetIndex = Math.Max(0, Math.Min(insertionIndex, target.Count));
            if (ReferenceEquals(source, target) && sourceIndex < insertionIndex)
            {
                targetIndex = Math.Max(0, targetIndex - 1);
            }

            PivotPlusPlacementRow moved = ReferenceEquals(source, target)
                ? placement
                : new PivotPlusPlacementRow(field, targetArea);
            target.Insert(targetIndex, moved);
            SelectedPlacement = moved;
            MarkPreviewChanged();
        }

        public void RemovePlacement(PivotPlusPlacementRow placement)
        {
            if (placement == null) return;
            if (AreaCollection(placement.Area).Remove(placement))
            {
                SelectedPlacement = null;
                MarkPreviewChanged();
            }
        }

        private void RefreshFieldFilter()
        {
            string query = SearchText.Trim();
            IEnumerable<PivotPlusFieldRow> matches = allFields.Where(field =>
                query.Length == 0 ||
                field.Caption.IndexOf(query, StringComparison.CurrentCultureIgnoreCase) >= 0 ||
                field.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0);
            Replace(Fields, matches);
            if (SelectedField != null && !Fields.Contains(SelectedField)) SelectedField = null;
        }

        private void AddSelectedField(object? parameter)
        {
            if (SelectedField == null || parameter == null) return;
            if (!Enum.TryParse(Convert.ToString(parameter), true, out PivotFieldArea area)) return;
            if (!SupportsArea(SelectedField.Source.SupportedAreas, area))
            {
                StatusMessage = SelectedField.Caption + " cannot be placed in " + AreaLabel(area) + ".";
                return;
            }

            ObservableCollection<PivotPlusPlacementRow> target = AreaCollection(area);
            if (area != PivotFieldArea.Values && target.Any(item =>
                string.Equals(item.FieldName, SelectedField.Name, StringComparison.OrdinalIgnoreCase)))
            {
                StatusMessage = SelectedField.Caption + " is already in " + AreaLabel(area) + ".";
                return;
            }

            target.Add(new PivotPlusPlacementRow(SelectedField, area));
            SelectedPlacement = target[target.Count - 1];
            MarkPreviewChanged();
        }

        private void RemoveSelectedPlacement()
        {
            if (SelectedPlacement == null) return;
            AreaCollection(SelectedPlacement.Area).Remove(SelectedPlacement);
            SelectedPlacement = null;
            MarkPreviewChanged();
        }

        private void MoveSelectedPlacement(int delta)
        {
            if (SelectedPlacement == null) return;
            ObservableCollection<PivotPlusPlacementRow> collection = AreaCollection(SelectedPlacement.Area);
            int current = collection.IndexOf(SelectedPlacement);
            int target = current + delta;
            if (current < 0 || target < 0 || target >= collection.Count) return;
            collection.Move(current, target);
            MarkPreviewChanged();
        }

        private bool CanMoveSelected(int delta)
        {
            if (SelectedPlacement == null || IsBusy) return false;
            ObservableCollection<PivotPlusPlacementRow> collection = AreaCollection(SelectedPlacement.Area);
            int current = collection.IndexOf(SelectedPlacement);
            int target = current + delta;
            return current >= 0 && target >= 0 && target < collection.Count;
        }

        private void MarkPreviewChanged()
        {
            HasPendingChanges = true;
            StatusMessage = "Preview changed. Choose Apply to update Excel.";
            RaiseCommandStates();
        }

        private IReadOnlyList<PivotPlusPlacementRequest> BuildPlacementRequests()
        {
            var result = new List<PivotPlusPlacementRequest>();
            AppendArea(result, Filters, PivotFieldArea.Filter);
            AppendArea(result, Columns, PivotFieldArea.Column);
            AppendArea(result, Rows, PivotFieldArea.Row);
            AppendArea(result, Values, PivotFieldArea.Values);
            return result;
        }

        private static void AppendArea(
            ICollection<PivotPlusPlacementRequest> result,
            IEnumerable<PivotPlusPlacementRow> placements,
            PivotFieldArea area)
        {
            int position = 0;
            foreach (PivotPlusPlacementRow placement in placements)
            {
                result.Add(new PivotPlusPlacementRequest(
                    placement.FieldName,
                    area,
                    ++position,
                    placement.Caption,
                    placement.Aggregation,
                    placement.NumberFormatCode));
            }
        }

        private void OpenExcelFieldList()
        {
            try
            {
                hostService.OpenExcelFieldList();
                StatusMessage = "Excel's native PivotTable Fields pane was toggled.";
            }
            catch (Exception exception)
            {
                StatusMessage = exception.Message;
            }
        }

        private bool CanAddSelectedField(object? parameter)
        {
            return SelectedField != null && parameter != null && !IsBusy;
        }

        private bool CanAddPortion()
        {
            return SupportsExtras && SelectedPortionValueField != null &&
                   SelectedPortionDetailField != null && !IsBusy;
        }

        private bool CanGroupDate(object? parameter)
        {
            return SelectedPeriodField != null && parameter != null && !IsBusy;
        }

        private void RaiseCommandStates()
        {
            (RefreshCommand as AsyncRelayCommand)?.RaiseCanExecuteChanged();
            (AddFieldCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (RemoveFieldCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (MoveUpCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (MoveDownCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (ApplyCommand as AsyncRelayCommand)?.RaiseCanExecuteChanged();
            (EnableDataModelCommand as AsyncRelayCommand)?.RaiseCanExecuteChanged();
            (AddPortionCommand as AsyncRelayCommand)?.RaiseCanExecuteChanged();
            (UndoCommand as AsyncRelayCommand)?.RaiseCanExecuteChanged();
            (GroupDateCommand as AsyncRelayCommand)?.RaiseCanExecuteChanged();
            (OpenExcelFieldListCommand as RelayCommand)?.RaiseCanExecuteChanged();
        }

        private ObservableCollection<PivotPlusPlacementRow> AreaCollection(PivotFieldArea area)
        {
            switch (area)
            {
                case PivotFieldArea.Filter: return Filters;
                case PivotFieldArea.Column: return Columns;
                case PivotFieldArea.Row: return Rows;
                case PivotFieldArea.Values: return Values;
                default: throw new ArgumentOutOfRangeException(nameof(area));
            }
        }

        private static IEnumerable<PivotPlusPlacementRow> CreateAreaRows(
            PivotPlusPaneSnapshot source,
            PivotFieldArea area)
        {
            return source.Placements.Where(item => item.Area == area)
                .OrderBy(item => item.Position)
                .Select(item => new PivotPlusPlacementRow(item));
        }

        private static bool SupportsArea(PivotFieldAreaSupport support, PivotFieldArea area)
        {
            PivotFieldAreaSupport required;
            switch (area)
            {
                case PivotFieldArea.Row: required = PivotFieldAreaSupport.Row; break;
                case PivotFieldArea.Column: required = PivotFieldAreaSupport.Column; break;
                case PivotFieldArea.Filter: required = PivotFieldAreaSupport.Filter; break;
                case PivotFieldArea.Values: required = PivotFieldAreaSupport.Values; break;
                default: return false;
            }

            return (support & required) == required;
        }

        private static string AreaLabel(PivotFieldArea area)
        {
            return area == PivotFieldArea.Values ? "Values" : area + "s";
        }

        private static string GetSourceKindLabel(PivotSourceKind sourceKind)
        {
            switch (sourceKind)
            {
                case PivotSourceKind.DataModel: return "Data Model PivotTable";
                case PivotSourceKind.ExternalOlap: return "External OLAP PivotTable";
                case PivotSourceKind.WorksheetTable: return "Worksheet table PivotTable";
                case PivotSourceKind.WorksheetRange: return "Worksheet range PivotTable";
                default: return sourceKind.ToString();
            }
        }

        private static void Replace<T>(ObservableCollection<T> target, IEnumerable<T> values)
        {
            target.Clear();
            foreach (T value in values) target.Add(value);
        }

        private static string SnapshotSignature(PivotPlusPaneSnapshot value)
        {
            return string.Join("\u001f", new[]
            {
                value.WorksheetName,
                value.PivotTableName,
                value.SourceKind.ToString(),
                string.Join("\u001e", value.Fields.Select(field =>
                    field.Name + "\u001d" + field.Caption + "\u001d" + field.SupportedAreas)),
                string.Join("\u001e", value.Placements
                    .OrderBy(item => item.Area)
                    .ThenBy(item => item.Position)
                    .Select(item => item.Area + "\u001d" + item.Position + "\u001d" +
                        item.FieldName + "\u001d" + item.Caption + "\u001d" +
                        item.Aggregation + "\u001d" + item.NumberFormatCode))
            });
        }
    }

    internal static class PivotPlusPeriodFieldClassifier
    {
        private static readonly string[] Months =
        {
            "jan", "feb", "mar", "apr", "may", "jun",
            "jul", "aug", "sep", "oct", "nov", "dec"
        };

        public static bool IsPeriodField(string caption)
        {
            string normalized = Normalize(caption);
            return PeriodOrder(normalized) < 100 ||
                   normalized.Contains("date") ||
                   normalized.Contains("month") ||
                   normalized.Contains("quarter") ||
                   normalized.Contains("year") ||
                   normalized.Contains("period");
        }

        public static int PeriodOrder(string caption)
        {
            string normalized = Normalize(caption);
            for (int index = 0; index < Months.Length; index++)
            {
                if (normalized == Months[index] || normalized.StartsWith(Months[index] + " "))
                {
                    return index + 1;
                }
            }

            return 100;
        }

        private static string Normalize(string value)
        {
            return (value ?? string.Empty).Trim().ToLowerInvariant();
        }
    }
}
