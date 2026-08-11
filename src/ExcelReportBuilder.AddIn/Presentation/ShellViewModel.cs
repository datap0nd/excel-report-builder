using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Security;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Windows.Media;
using ExcelReportBuilder.AddIn.Activity;
using ExcelReportBuilder.AddIn.Host;
using ExcelReportBuilder.Agent.Configuration;

namespace ExcelReportBuilder.AddIn.Presentation
{
    public sealed class ShellViewModel : ObservableObject, IDisposable
    {
        private readonly IReportBuilderHostService _hostService;
        private readonly HashSet<FieldPlacement> _subscribedPlacements = new HashSet<FieldPlacement>();
        private readonly HashSet<ObservableObject> _subscribedManualRules = new HashSet<ObservableObject>();
        private ShellSurface _currentSurface;
        private SourceSnapshot? _source;
        private ReportSpecificationSnapshot? _agentAppliedSpecification;
        private SourceColumnSnapshot? _selectedField;
        private FieldPlacement? _selectedPlacement;
        private ManualTransformRule? _selectedTransformRule;
        private ManualCalculatedMetricRule? _selectedCalculatedMetric;
        private ManualReportBlockRule? _selectedReportBlock;
        private ManualCheckRule? _selectedCheckRule;
        private string _selectedPeriodMode = "Date column";
        private string _selectedOutputStyle = "Dense management block";
        private string _selectedPeriodColumn = "Date";
        private string _headerPattern = string.Empty;
        private string _reportingYearText = string.Empty;
        private string _chatDraft = string.Empty;
        private string _endpointBaseUrl = "http://127.0.0.1:11434/v1";
        private string _modelId = AgentDefaults.Model;
        private bool _allowRemoteHttp;
        private bool _allowRemoteWorkbookData;
        private SecureString? _apiKey;
        private bool _savedApiKeyAvailable;
        private string? _savedCredentialBaseUrl;
        private bool _savedCredentialHasProtectedKey;
        private long _endpointConfigurationVersion;
        private string _endpointStateLabel = "Not checked";
        private string _endpointValidationMessage = "Check the endpoint before sending a request.";
        private bool _endpointCheckPassed;
        private WideHeaderMappingPreview? _wideHeaderPreview;
        private bool _reportingYearRequired;
        private bool _periodMappingConfirmed;
        private bool _hasBuiltDraft;
        private bool _hasRunChecks;
        private bool _checksPassed;
        private bool _isPublished;
        private string _lastDraftLabel = "No managed draft built";
        private bool _repeatRowLabels;
        private bool _insertBlankRows;
        private bool _freezeHeaders = true;
        private bool _showRowGrandTotals = true;
        private bool _showColumnGrandTotals = true;
        private int _rowIndent = 1;
        private string _rowGrandTotalLabel = "Grand Total";
        private string _columnGrandTotalLabel = "Grand Total";
        private bool _manualProjectionComplete = true;
        private bool _manualRestrictionMessageShown;
        private bool _applyingManualEditingState;
        private bool _disposed;

        public event EventHandler? ApiKeyClearRequested;

        public ShellViewModel()
            : this(new SyntheticReportBuilderHostService())
        {
        }

        public ShellViewModel(IReportBuilderHostService hostService)
        {
            _hostService = hostService ?? throw new ArgumentNullException(nameof(hostService));
            _hostService.ActivityReported += OnHostActivityReported;

            Activity = new OperationActivityController();
            Activity.PropertyChanged += OnActivityPropertyChanged;
            Columns = new ObservableCollection<SourceColumnSnapshot>();
            AvailableFields = new ObservableCollection<SourceColumnSnapshot>();
            Placements = new ObservableCollection<FieldPlacement>();
            TransformRules = new ObservableCollection<ManualTransformRule>();
            CalculatedMetrics = new ObservableCollection<ManualCalculatedMetricRule>();
            ReportBlocks = new ObservableCollection<ManualReportBlockRule>();
            RequiredChecks = new ObservableCollection<ManualCheckRule>();
            ChatLines = new ObservableCollection<ChatLine>();
            CheckLines = new ObservableCollection<CheckLine>();
            WideHeaderMappings = new ObservableCollection<WideHeaderMappingRowSnapshot>();
            NormalizedSampleRows = new ObservableCollection<NormalizedSampleRowSnapshot>();
            AvailableModels = new ObservableCollection<string>();
            PeriodModes = new ReadOnlyCollection<string>(new[]
            {
                "Date column",
                "Wide period headers",
                "No period columns"
            });
            OutputStyles = new ReadOnlyCollection<string>(new[]
            {
                "Standard matrix",
                "Metric stack",
                "Dense management block"
            });
            PlacementSortOptions = new ReadOnlyCollection<string>(new[]
            {
                "Default order",
                "Ascending",
                "Descending"
            });
            SubtotalPlacementOptions = new ReadOnlyCollection<string>(new[]
            {
                "After members",
                "Before members"
            });
            ValueAggregationOptions = new ReadOnlyCollection<string>(new[]
            {
                "Sum",
                "Count",
                "Distinct count",
                "Average",
                "Minimum",
                "Maximum"
            });
            NumberFormatOptions = new ReadOnlyCollection<string>(new[]
            {
                "General",
                "#,##0",
                "#,##0.00",
                "0.0%",
                "0.00%"
            });
            TransformOperations = new ReadOnlyCollection<string>(new[]
            {
                "Keep columns",
                "Remove columns",
                "Reorder columns",
                "Rename column",
                "Convert type",
                "Trim text",
                "Replace value",
                "Normalize blanks",
                "Normalize errors",
                "Fill down",
                "Map values",
                "Filter rows",
                "Exclude total rows",
                "Derive period parts",
                "Arithmetic"
            });
            CalculatedMetricKinds = new ReadOnlyCollection<string>(new[]
            {
                "Add",
                "Subtract",
                "Multiply",
                "Safe divide",
                "Ratio",
                "Difference",
                "Percentage change",
                "Percentage-point difference",
                "Share of parent",
                "Share of report total",
                "Weighted average",
                "Filtered aggregate"
            });
            RequiredCheckKinds = new ReadOnlyCollection<string>(new[]
            {
                "Total preservation",
                "Required values",
                "Non-negative",
                "Balance"
            });

            Placements.CollectionChanged += OnPlacementsChanged;
            TransformRules.CollectionChanged += OnManualRulesChanged;
            CalculatedMetrics.CollectionChanged += OnManualRulesChanged;
            ReportBlocks.CollectionChanged += OnManualRulesChanged;
            RequiredChecks.CollectionChanged += OnManualRulesChanged;
            SelectSurfaceCommand = new RelayCommand(SelectSurface);
            SelectSourceCommand = new AsyncRelayCommand(SelectSourceAsync, CanStartOperation);
            ConfirmPeriodMappingCommand = new AsyncRelayCommand(ConfirmPeriodMappingAsync, CanConfirmPeriodMapping);
            PreviewWideHeaderMappingCommand = new AsyncRelayCommand(
                PreviewWideHeaderMappingAsync,
                CanPreviewWideHeaderMapping);
            AddPlacementCommand = new RelayCommand(AddPlacement, CanAddPlacement);
            RemovePlacementCommand = new RelayCommand(RemoveSelectedPlacement, () => SelectedPlacement != null && CanEditManualSpecification && !Activity.IsOperationActive);
            MovePlacementUpCommand = new RelayCommand(() => MoveSelectedPlacement(-1), () => CanMoveSelectedPlacement(-1));
            MovePlacementDownCommand = new RelayCommand(() => MoveSelectedPlacement(1), () => CanMoveSelectedPlacement(1));
            AddTransformCommand = new RelayCommand(AddTransformRule, () => CanEditManualSpecification && !Activity.IsOperationActive && _source != null);
            RemoveTransformCommand = new RelayCommand(RemoveSelectedTransformRule, () => CanEditManualSpecification && SelectedTransformRule != null && !Activity.IsOperationActive);
            AddCalculatedMetricCommand = new RelayCommand(AddCalculatedMetric, () => CanEditManualSpecification && !Activity.IsOperationActive && _source != null);
            RemoveCalculatedMetricCommand = new RelayCommand(RemoveSelectedCalculatedMetric, () => CanEditManualSpecification && SelectedCalculatedMetric != null && !Activity.IsOperationActive);
            AddReportBlockCommand = new RelayCommand(AddReportBlock, () => CanEditManualSpecification && !Activity.IsOperationActive && _source != null && ReportBlocks.Count < 8);
            RemoveReportBlockCommand = new RelayCommand(RemoveSelectedReportBlock, () => CanEditManualSpecification && SelectedReportBlock != null && ReportBlocks.Count > 1 && !Activity.IsOperationActive);
            AddRequiredCheckCommand = new RelayCommand(AddRequiredCheck, () => CanEditManualSpecification && !Activity.IsOperationActive && _source != null);
            RemoveRequiredCheckCommand = new RelayCommand(RemoveSelectedRequiredCheck, () => CanEditManualSpecification && SelectedCheckRule != null && !Activity.IsOperationActive);
            BuildDraftCommand = new AsyncRelayCommand(BuildDraftAsync, CanBuildDraft);
            TogglePauseCommand = new RelayCommand(TogglePause, () => Activity.CanPause);
            CancelCommand = new RelayCommand(CancelOperation, () => Activity.CanCancel);
            SendChatCommand = new AsyncRelayCommand(SendChatAsync, CanSendChat);
            DiscoverModelsCommand = new AsyncRelayCommand(DiscoverModelsAsync, CanDiscoverModels);
            CheckEndpointCommand = new AsyncRelayCommand(CheckEndpointAsync, CanCheckEndpoint);
            RunChecksCommand = new AsyncRelayCommand(RunChecksAsync, CanRunChecks);
            PublishCommand = new AsyncRelayCommand(PublishAsync, () => CanPublish);

            ChatLines.Add(new ChatLine(
                "Builder",
                "Describe the report change you want. I can propose only bounded changes to Rows, Columns, Values, Filters, formatting, and checks."));
            SavedEndpointSettingsSnapshot? savedEndpoint = _hostService.SavedEndpointSettings;
            if (savedEndpoint != null)
            {
                _endpointBaseUrl = savedEndpoint.BaseUrl;
                _modelId = savedEndpoint.ModelId;
                _allowRemoteHttp = savedEndpoint.AllowRemoteHttp;
                _allowRemoteWorkbookData = savedEndpoint.AllowRemoteWorkbookData;
                _savedApiKeyAvailable = savedEndpoint.HasProtectedApiKey;
                _savedCredentialBaseUrl = savedEndpoint.BaseUrl;
                _savedCredentialHasProtectedKey = savedEndpoint.HasProtectedApiKey;
                EndpointValidationMessage = savedEndpoint.HasProtectedApiKey
                    ? "Saved endpoint settings loaded. The API key is protected for this Windows user."
                    : "Saved endpoint settings loaded.";
            }

            if (_hostService.IsSynthetic)
            {
                SeedSyntheticShellState();
                AvailableModels.Add("sample-balanced");
                ModelId = "sample-balanced";
            }
        }

        public OperationActivityController Activity { get; }

        public ObservableCollection<SourceColumnSnapshot> Columns { get; }

        public ObservableCollection<SourceColumnSnapshot> AvailableFields { get; }

        public ObservableCollection<FieldPlacement> Placements { get; }

        public ObservableCollection<ManualTransformRule> TransformRules { get; }

        public ObservableCollection<ManualCalculatedMetricRule> CalculatedMetrics { get; }

        public ObservableCollection<ManualReportBlockRule> ReportBlocks { get; }

        public ObservableCollection<ManualCheckRule> RequiredChecks { get; }

        public ObservableCollection<ChatLine> ChatLines { get; }

        public ObservableCollection<CheckLine> CheckLines { get; }

        public ObservableCollection<WideHeaderMappingRowSnapshot> WideHeaderMappings { get; }

        public ObservableCollection<NormalizedSampleRowSnapshot> NormalizedSampleRows { get; }

        public ObservableCollection<string> AvailableModels { get; }

        public IReadOnlyList<string> PeriodModes { get; }

        public IReadOnlyList<string> OutputStyles { get; }

        public IReadOnlyList<string> PlacementSortOptions { get; }

        public IReadOnlyList<string> SubtotalPlacementOptions { get; }

        public IReadOnlyList<string> ValueAggregationOptions { get; }

        public IReadOnlyList<string> NumberFormatOptions { get; }

        public IReadOnlyList<string> TransformOperations { get; }

        public IReadOnlyList<string> CalculatedMetricKinds { get; }

        public IReadOnlyList<string> RequiredCheckKinds { get; }

        public ICommand SelectSurfaceCommand { get; }

        public AsyncRelayCommand SelectSourceCommand { get; }

        public AsyncRelayCommand ConfirmPeriodMappingCommand { get; }

        public AsyncRelayCommand PreviewWideHeaderMappingCommand { get; }

        public RelayCommand AddPlacementCommand { get; }

        public RelayCommand RemovePlacementCommand { get; }

        public RelayCommand MovePlacementUpCommand { get; }

        public RelayCommand MovePlacementDownCommand { get; }

        public RelayCommand AddTransformCommand { get; }

        public RelayCommand RemoveTransformCommand { get; }

        public RelayCommand AddCalculatedMetricCommand { get; }

        public RelayCommand RemoveCalculatedMetricCommand { get; }

        public RelayCommand AddReportBlockCommand { get; }

        public RelayCommand RemoveReportBlockCommand { get; }

        public RelayCommand AddRequiredCheckCommand { get; }

        public RelayCommand RemoveRequiredCheckCommand { get; }

        public AsyncRelayCommand BuildDraftCommand { get; }

        public RelayCommand TogglePauseCommand { get; }

        public RelayCommand CancelCommand { get; }

        public AsyncRelayCommand SendChatCommand { get; }

        public AsyncRelayCommand DiscoverModelsCommand { get; }

        public AsyncRelayCommand CheckEndpointCommand { get; }

        public AsyncRelayCommand RunChecksCommand { get; }

        public AsyncRelayCommand PublishCommand { get; }

        public ShellSurface CurrentSurface
        {
            get => _currentSurface;
            set
            {
                if (!SetProperty(ref _currentSurface, value))
                {
                    return;
                }

                RaisePropertyChanged(nameof(CurrentSurfaceTitle));
                RaisePropertyChanged(nameof(CurrentSurfaceSummary));
            }
        }

        public string CurrentSurfaceTitle => CurrentSurface.ToString();

        public string CurrentSurfaceSummary
        {
            get
            {
                switch (CurrentSurface)
                {
                    case ShellSurface.Build:
                        return "Place workbook columns into the report layout.";
                    case ShellSurface.Chat:
                        return "Ask for bounded changes to the saved report setup.";
                    case ShellSurface.Checks:
                        return "Review independent checks before publishing.";
                    default:
                        return "Choose and confirm the workbook data to use.";
                }
            }
        }

        public SourceColumnSnapshot? SelectedField
        {
            get => _selectedField;
            set
            {
                if (SetProperty(ref _selectedField, value))
                {
                    AddPlacementCommand.RaiseCanExecuteChanged();
                }
            }
        }

        public FieldPlacement? SelectedPlacement
        {
            get => _selectedPlacement;
            set
            {
                if (!SetProperty(ref _selectedPlacement, value))
                {
                    return;
                }

                RemovePlacementCommand.RaiseCanExecuteChanged();
                MovePlacementUpCommand.RaiseCanExecuteChanged();
                MovePlacementDownCommand.RaiseCanExecuteChanged();
            }
        }

        public ManualTransformRule? SelectedTransformRule
        {
            get => _selectedTransformRule;
            set
            {
                if (SetProperty(ref _selectedTransformRule, value))
                {
                    RemoveTransformCommand.RaiseCanExecuteChanged();
                }
            }
        }

        public ManualCalculatedMetricRule? SelectedCalculatedMetric
        {
            get => _selectedCalculatedMetric;
            set
            {
                if (SetProperty(ref _selectedCalculatedMetric, value))
                {
                    RemoveCalculatedMetricCommand.RaiseCanExecuteChanged();
                }
            }
        }

        public ManualReportBlockRule? SelectedReportBlock
        {
            get => _selectedReportBlock;
            set
            {
                if (SetProperty(ref _selectedReportBlock, value))
                {
                    RemoveReportBlockCommand.RaiseCanExecuteChanged();
                }
            }
        }

        public ManualCheckRule? SelectedCheckRule
        {
            get => _selectedCheckRule;
            set
            {
                if (SetProperty(ref _selectedCheckRule, value))
                {
                    RemoveRequiredCheckCommand.RaiseCanExecuteChanged();
                }
            }
        }

        public bool RepeatRowLabels
        {
            get => _repeatRowLabels;
            set
            {
                if (!DemandManualEditing(nameof(RepeatRowLabels))) return;
                if (SetProperty(ref _repeatRowLabels, value)) MarkSpecificationDirty();
            }
        }

        public bool InsertBlankRows
        {
            get => _insertBlankRows;
            set
            {
                if (!DemandManualEditing(nameof(InsertBlankRows))) return;
                if (SetProperty(ref _insertBlankRows, value)) MarkSpecificationDirty();
            }
        }

        public bool FreezeHeaders
        {
            get => _freezeHeaders;
            set
            {
                if (!DemandManualEditing(nameof(FreezeHeaders))) return;
                if (SetProperty(ref _freezeHeaders, value)) MarkSpecificationDirty();
            }
        }

        public bool ShowRowGrandTotals
        {
            get => _showRowGrandTotals;
            set
            {
                if (!DemandManualEditing(nameof(ShowRowGrandTotals))) return;
                if (SetProperty(ref _showRowGrandTotals, value)) MarkSpecificationDirty();
            }
        }

        public bool ShowColumnGrandTotals
        {
            get => _showColumnGrandTotals;
            set
            {
                if (!DemandManualEditing(nameof(ShowColumnGrandTotals))) return;
                if (SetProperty(ref _showColumnGrandTotals, value)) MarkSpecificationDirty();
            }
        }

        public int RowIndent
        {
            get => _rowIndent;
            set
            {
                if (!DemandManualEditing(nameof(RowIndent))) return;
                int bounded = Math.Max(0, Math.Min(15, value));
                if (SetProperty(ref _rowIndent, bounded)) MarkSpecificationDirty();
            }
        }

        public string RowGrandTotalLabel
        {
            get => _rowGrandTotalLabel;
            set
            {
                if (!DemandManualEditing(nameof(RowGrandTotalLabel))) return;
                if (SetProperty(ref _rowGrandTotalLabel, value ?? string.Empty)) MarkSpecificationDirty();
            }
        }

        public string ColumnGrandTotalLabel
        {
            get => _columnGrandTotalLabel;
            set
            {
                if (!DemandManualEditing(nameof(ColumnGrandTotalLabel))) return;
                if (SetProperty(ref _columnGrandTotalLabel, value ?? string.Empty)) MarkSpecificationDirty();
            }
        }

        public string SelectedPeriodMode
        {
            get => _selectedPeriodMode;
            set
            {
                if (!DemandManualEditing(nameof(SelectedPeriodMode))) return;
                if (SetProperty(ref _selectedPeriodMode, value ?? string.Empty))
                {
                    ClearWideHeaderPreview(keepReportingYearRequirement: false);
                    RaisePropertyChanged(nameof(IsWideHeaderMode));
                    InvalidatePeriodMapping();
                }
            }
        }

        public string SelectedOutputStyle
        {
            get => _selectedOutputStyle;
            set
            {
                if (!DemandManualEditing(nameof(SelectedOutputStyle))) return;
                if (SetProperty(ref _selectedOutputStyle, value ?? "Dense management block"))
                {
                    MarkSpecificationDirty();
                }
            }
        }

        public string SelectedPeriodColumn
        {
            get => _selectedPeriodColumn;
            set
            {
                if (!DemandManualEditing(nameof(SelectedPeriodColumn))) return;
                if (SetProperty(ref _selectedPeriodColumn, value ?? string.Empty))
                {
                    InvalidatePeriodMapping();
                }
            }
        }

        public string HeaderPattern
        {
            get => _headerPattern;
            set
            {
                if (!DemandManualEditing(nameof(HeaderPattern))) return;
                if (SetProperty(ref _headerPattern, value ?? string.Empty))
                {
                    ClearWideHeaderPreview(keepReportingYearRequirement: true);
                    InvalidatePeriodMapping();
                }
            }
        }

        public string ReportingYearText
        {
            get => _reportingYearText;
            set
            {
                if (!DemandManualEditing(nameof(ReportingYearText))) return;
                if (SetProperty(ref _reportingYearText, value ?? string.Empty))
                {
                    ClearWideHeaderPreview(keepReportingYearRequirement: true);
                    InvalidatePeriodMapping();
                }
            }
        }

        public string ChatDraft
        {
            get => _chatDraft;
            set
            {
                if (SetProperty(ref _chatDraft, value ?? string.Empty))
                {
                    SendChatCommand.RaiseCanExecuteChanged();
                }
            }
        }

        public string EndpointBaseUrl
        {
            get => _endpointBaseUrl;
            set
            {
                string next = value ?? string.Empty;
                string previous = _endpointBaseUrl;
                if (SetProperty(ref _endpointBaseUrl, next))
                {
                    _endpointConfigurationVersion++;
                    if (!AgentEndpointCredentialScope.Matches(previous, next))
                    {
                        ResetEndpointScopedSecurityState(next);
                    }

                    InvalidateEndpointCheck();
                }
            }
        }

        public string ModelId
        {
            get => _modelId;
            set
            {
                if (SetProperty(ref _modelId, value ?? string.Empty))
                {
                    _endpointConfigurationVersion++;
                    InvalidateEndpointCheck();
                }
            }
        }

        public bool AllowRemoteHttp
        {
            get => _allowRemoteHttp;
            set
            {
                if (SetProperty(ref _allowRemoteHttp, value))
                {
                    _endpointConfigurationVersion++;
                    InvalidateEndpointCheck();
                }
            }
        }

        public bool AllowRemoteWorkbookData
        {
            get => _allowRemoteWorkbookData;
            set
            {
                if (SetProperty(ref _allowRemoteWorkbookData, value))
                {
                    _endpointConfigurationVersion++;
                    InvalidateEndpointCheck();
                }
            }
        }

        public bool IsWideHeaderMode => string.Equals(
            SelectedPeriodMode,
            "Wide period headers",
            StringComparison.Ordinal);

        public bool ReportingYearRequired => _reportingYearRequired;

        public string ReportingYearRequirementLabel => ReportingYearRequired
            ? "Required because at least one month header has no year."
            : "Enter a year only when source headers omit it.";

        public string MappingPreviewState => _wideHeaderPreview == null
            ? "Not previewed"
            : ReportingYearRequired
                ? "Year required"
                : _wideHeaderPreview.TotalPreservation == TotalPreservationState.Pass
                    ? "Ready"
                    : "Needs attention";

        public string ProjectedNormalizedRowsLabel => _wideHeaderPreview == null
            ? "Projected rows unavailable"
            : $"{_wideHeaderPreview.ProjectedNormalizedRowCount:N0} projected normalized rows";

        public string TotalPreservationLabel
        {
            get
            {
                if (_wideHeaderPreview == null)
                {
                    return "Not checked";
                }

                switch (_wideHeaderPreview.TotalPreservation)
                {
                    case TotalPreservationState.Pass:
                        return "Totals preserved";
                    case TotalPreservationState.Fail:
                        return "Totals do not match";
                    default:
                        return "Waiting for required input";
                }
            }
        }

        public string TotalPreservationDetail => _wideHeaderPreview?.TotalPreservationDetail
            ?? "Preview the mapping to compare source and normalized totals.";

        public Brush TotalPreservationBrush => _wideHeaderPreview?.TotalPreservation == TotalPreservationState.Pass
            ? new SolidColorBrush(Color.FromRgb(26, 102, 62))
            : _wideHeaderPreview?.TotalPreservation == TotalPreservationState.Fail
                ? new SolidColorBrush(Color.FromRgb(137, 35, 38))
                : new SolidColorBrush(Color.FromRgb(105, 68, 0));

        public string EndpointStateLabel
        {
            get => _endpointStateLabel;
            private set => SetProperty(ref _endpointStateLabel, value);
        }

        public string EndpointValidationMessage
        {
            get => _endpointValidationMessage;
            private set => SetProperty(ref _endpointValidationMessage, value);
        }

        public Brush EndpointStateBrush => _endpointCheckPassed
            ? new SolidColorBrush(Color.FromRgb(26, 102, 62))
            : new SolidColorBrush(Color.FromRgb(105, 68, 0));

        public string ApiKeyStateLabel => _apiKey != null && _apiKey.Length > 0
            ? "API key held in memory for this Excel session"
            : _savedApiKeyAvailable
                ? "Protected API key available for this Windows user"
                : "No API key entered";

        public string SourceName => _source?.DisplayName ?? "No workbook data selected";

        public string SourceLocation => _source?.Location ?? "Choose a table or rectangular range.";

        public string SourceSummary => _source == null
            ? "No columns available"
            : $"{_source.RowCount:N0} rows · {_source.Columns.Count} columns";

        public string SourceKindLabel => _hostService.IsSynthetic
            ? "Synthetic smoke mode"
            : _source == null
                ? "Waiting for data"
                : "Workbook data";

        public string PeriodStatusLabel => _periodMappingConfirmed ? "Confirmed" : "Needs confirmation";

        public Brush PeriodStatusBrush => _periodMappingConfirmed
            ? new SolidColorBrush(Color.FromRgb(26, 102, 62))
            : new SolidColorBrush(Color.FromRgb(105, 68, 0));

        public string LastDraftLabel
        {
            get => _lastDraftLabel;
            private set => SetProperty(ref _lastDraftLabel, value);
        }

        public string PublishGateState => _isPublished ? "Published" : CanPublish ? "Ready" : "Blocked";

        public string PublishGateLabel
        {
            get
            {
                if (_isPublished)
                {
                    return _hostService.IsSynthetic
                        ? "The synthetic managed draft is marked as published."
                        : "The checked managed draft is published.";
                }

                if (!_hasBuiltDraft)
                {
                    return "Build a managed draft before publishing.";
                }

                if (!_hasRunChecks)
                {
                    return "Run all checks before publishing.";
                }

                if (!_checksPassed)
                {
                    return "Resolve failed checks before publishing.";
                }

                return "All checks passed. The managed draft can be published.";
            }
        }

        public Brush PublishGateBrush
        {
            get
            {
                if (_isPublished || CanPublish)
                {
                    return new SolidColorBrush(Color.FromRgb(26, 102, 62));
                }

                return new SolidColorBrush(Color.FromRgb(105, 68, 0));
            }
        }

        public bool CanPublish => _hasBuiltDraft
            && _hasRunChecks
            && _checksPassed
            && !_isPublished
            && !Activity.IsOperationActive;

        public void SetApiKey(SecureString? apiKey)
        {
            _apiKey?.Dispose();
            _apiKey = apiKey != null && apiKey.Length > 0 ? apiKey.Copy() : null;
            _endpointConfigurationVersion++;
            RaisePropertyChanged(nameof(ApiKeyStateLabel));
            InvalidateEndpointCheck();
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            _hostService.ActivityReported -= OnHostActivityReported;
            Activity.PropertyChanged -= OnActivityPropertyChanged;
            Placements.CollectionChanged -= OnPlacementsChanged;
            TransformRules.CollectionChanged -= OnManualRulesChanged;
            CalculatedMetrics.CollectionChanged -= OnManualRulesChanged;
            ReportBlocks.CollectionChanged -= OnManualRulesChanged;
            RequiredChecks.CollectionChanged -= OnManualRulesChanged;
            foreach (FieldPlacement placement in _subscribedPlacements)
            {
                placement.PropertyChanged -= OnPlacementPropertyChanged;
            }

            _subscribedPlacements.Clear();
            foreach (ObservableObject rule in _subscribedManualRules)
            {
                rule.PropertyChanged -= OnManualRulePropertyChanged;
            }

            _subscribedManualRules.Clear();

            Activity.Dispose();
            _apiKey?.Dispose();
            _apiKey = null;
            if (_hostService is IDisposable disposableHostService)
            {
                disposableHostService.Dispose();
            }
        }

        private void SeedSyntheticShellState()
        {
            ApplySource(SyntheticReportBuilderHostService.CreateSampleSource());
            Placements.Add(new FieldPlacement(PlacementBucket.Rows, "Category", "Default order"));
            Placements.Add(new FieldPlacement(PlacementBucket.Columns, "Date", "Ascending"));
            Placements.Add(new FieldPlacement(PlacementBucket.Values, "Amount", "Sum", numberFormat: "#,##0.00"));
            Placements.Add(new FieldPlacement(PlacementBucket.Filters, "Region", "All"));
            SelectedPlacement = Placements[0];
            _periodMappingConfirmed = false;
            RaisePeriodProperties();
            MarkSpecificationDirty();
        }

        private void SelectSurface(object? parameter)
        {
            if (parameter is ShellSurface surface)
            {
                CurrentSurface = surface;
            }
        }

        private async Task SelectSourceAsync()
        {
            Activity.Begin(
                "Select workbook data",
                ActivityStage.Inspecting,
                "Starting data selection.",
                "The host service owns all Excel reads.");

            try
            {
                SourceSnapshot source = await _hostService.SelectCurrentDataAsync(Activity.CancellationToken);
                ApplySource(source);
                Activity.Complete(
                    "Workbook data is ready.",
                    $"{source.RowCount:N0} rows and {source.Columns.Count} columns were returned by the host.");
            }
            catch (OperationCanceledException)
            {
                EnsureCancelled();
            }
            catch (Exception)
            {
                Activity.Fail(
                    "Workbook data could not be selected.",
                    "Review the active selection and try again.");
            }
            finally
            {
                RefreshCommandStates();
            }
        }

        private async Task ConfirmPeriodMappingAsync()
        {
            PeriodMappingSnapshot mapping = CreatePeriodMappingSnapshot();
            Activity.Begin(
                "Confirm period layout",
                ActivityStage.Planning,
                "Checking the selected period layout.",
                mapping.Mode);

            try
            {
                await _hostService.ConfirmPeriodMappingAsync(mapping, Activity.CancellationToken);
                _periodMappingConfirmed = true;
                RaisePeriodProperties();
                MarkSpecificationDirty();
                Activity.Complete(
                    "Period layout confirmed.",
                    "The saved report setup can now use this period mapping.");
            }
            catch (OperationCanceledException)
            {
                EnsureCancelled();
            }
            catch (Exception)
            {
                Activity.Fail(
                    "The period layout could not be confirmed.",
                    "Review the period mode and column, then try again.");
            }
            finally
            {
                RefreshCommandStates();
            }
        }

        private async Task PreviewWideHeaderMappingAsync()
        {
            int? reportingYear = TryGetReportingYear(out int parsedYear) ? parsedYear : (int?)null;
            Activity.Begin(
                "Preview wide period mapping",
                ActivityStage.Inspecting,
                "Classifying wide period headers.",
                "The preview shows Period, Metric, projected rows, samples, and total preservation.");

            try
            {
                WideHeaderMappingPreview preview = await _hostService.PreviewWideHeaderMappingAsync(
                    HeaderPattern,
                    reportingYear,
                    Activity.CancellationToken);
                ApplyWideHeaderPreview(preview);
                Activity.Complete(
                    preview.RequiresReportingYear
                        ? "Mapping preview needs a reporting year."
                        : "Wide period mapping preview is ready.",
                    preview.TotalPreservationDetail);
            }
            catch (OperationCanceledException)
            {
                EnsureCancelled();
            }
            catch (Exception)
            {
                Activity.Fail(
                    "Wide period headers could not be previewed.",
                    "Review the source headers and reporting year, then try again.");
            }
            finally
            {
                RefreshCommandStates();
            }
        }

        private void AddPlacement(object? parameter)
        {
            if (!(parameter is PlacementBucket bucket) || SelectedField == null)
            {
                return;
            }

            FieldPlacement? existing = Placements.FirstOrDefault(
                placement => placement.Bucket == bucket
                    && string.Equals(placement.ColumnName, SelectedField.Name, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                SelectedPlacement = existing;
                return;
            }

            string setting = bucket == PlacementBucket.Values
                ? "Sum"
                : bucket == PlacementBucket.Filters
                    ? "All"
                    : "Default order";
            FieldPlacement placement = new FieldPlacement(
                bucket,
                SelectedField.Name,
                setting,
                numberFormat: bucket == PlacementBucket.Values ? "#,##0.00" : "General");
            Placements.Add(placement);
            SelectedPlacement = placement;
            MarkSpecificationDirty();
        }

        private bool CanAddPlacement(object? parameter)
        {
            return parameter is PlacementBucket
                && SelectedField != null
                && CanEditManualSpecification
                && !Activity.IsOperationActive;
        }

        private void AddTransformRule()
        {
            var rule = new ManualTransformRule
            {
                Operation = "Trim text",
                Column = SelectedField?.Name ?? AvailableFields.FirstOrDefault()?.Name ?? string.Empty
            };
            TransformRules.Add(rule);
            SelectedTransformRule = rule;
            MarkSpecificationDirty();
        }

        private void RemoveSelectedTransformRule()
        {
            if (SelectedTransformRule == null) return;
            int index = TransformRules.IndexOf(SelectedTransformRule);
            TransformRules.Remove(SelectedTransformRule);
            SelectedTransformRule = TransformRules.Count == 0
                ? null
                : TransformRules[Math.Min(index, TransformRules.Count - 1)];
            MarkSpecificationDirty();
        }

        private void AddCalculatedMetric()
        {
            string[] valueNames = Placements
                .Where(placement => placement.Bucket == PlacementBucket.Values)
                .Select(placement => placement.Setting + " of " + placement.ColumnName)
                .ToArray();
            var rule = new ManualCalculatedMetricRule
            {
                Label = "Calculated metric " + (CalculatedMetrics.Count + 1).ToString(),
                Kind = "Ratio",
                Primary = valueNames.FirstOrDefault() ?? string.Empty,
                Secondary = valueNames.Skip(1).FirstOrDefault() ?? valueNames.FirstOrDefault() ?? string.Empty,
                NumberFormat = "0.0%"
            };
            CalculatedMetrics.Add(rule);
            SelectedCalculatedMetric = rule;
            MarkSpecificationDirty();
        }

        private void RemoveSelectedCalculatedMetric()
        {
            if (SelectedCalculatedMetric == null) return;
            int index = CalculatedMetrics.IndexOf(SelectedCalculatedMetric);
            CalculatedMetrics.Remove(SelectedCalculatedMetric);
            SelectedCalculatedMetric = CalculatedMetrics.Count == 0
                ? null
                : CalculatedMetrics[Math.Min(index, CalculatedMetrics.Count - 1)];
            MarkSpecificationDirty();
        }

        private void AddReportBlock()
        {
            int number = ReportBlocks.Count + 1;
            var block = new ManualReportBlockRule
            {
                Title = number == 1 ? "Management report" : "Management report " + number.ToString(),
                WorksheetName = number == 1 ? "Report" : "Report " + number.ToString(),
                AnchorCell = "A1",
                OutputStyle = SelectedOutputStyle,
                OwnedRows = 500,
                OwnedColumns = 64
            };
            ReportBlocks.Add(block);
            SelectedReportBlock = block;
            MarkSpecificationDirty();
        }

        private void RemoveSelectedReportBlock()
        {
            if (SelectedReportBlock == null || ReportBlocks.Count <= 1) return;
            int index = ReportBlocks.IndexOf(SelectedReportBlock);
            ReportBlocks.Remove(SelectedReportBlock);
            SelectedReportBlock = ReportBlocks[Math.Min(index, ReportBlocks.Count - 1)];
            MarkSpecificationDirty();
        }

        private void AddRequiredCheck()
        {
            FieldPlacement? firstValue = Placements.FirstOrDefault(
                placement => placement.Bucket == PlacementBucket.Values);
            var rule = new ManualCheckRule
            {
                Kind = "Total preservation",
                Metric = firstValue == null
                    ? string.Empty
                    : firstValue.Setting + " of " + firstValue.ColumnName,
                ToleranceText = "0"
            };
            RequiredChecks.Add(rule);
            SelectedCheckRule = rule;
            MarkSpecificationDirty();
        }

        private void RemoveSelectedRequiredCheck()
        {
            if (SelectedCheckRule == null) return;
            int index = RequiredChecks.IndexOf(SelectedCheckRule);
            RequiredChecks.Remove(SelectedCheckRule);
            SelectedCheckRule = RequiredChecks.Count == 0
                ? null
                : RequiredChecks[Math.Min(index, RequiredChecks.Count - 1)];
            MarkSpecificationDirty();
        }

        private void RemoveSelectedPlacement()
        {
            if (SelectedPlacement == null)
            {
                return;
            }

            int currentIndex = Placements.IndexOf(SelectedPlacement);
            Placements.Remove(SelectedPlacement);
            SelectedPlacement = Placements.Count == 0
                ? null
                : Placements[Math.Min(currentIndex, Placements.Count - 1)];
            MarkSpecificationDirty();
        }

        private void MoveSelectedPlacement(int direction)
        {
            if (SelectedPlacement == null)
            {
                return;
            }

            int currentIndex = Placements.IndexOf(SelectedPlacement);
            int candidateIndex = currentIndex + direction;
            while (candidateIndex >= 0 && candidateIndex < Placements.Count)
            {
                if (Placements[candidateIndex].Bucket == SelectedPlacement.Bucket)
                {
                    Placements.Move(currentIndex, candidateIndex);
                    MarkSpecificationDirty();
                    return;
                }

                candidateIndex += direction;
            }
        }

        private bool CanMoveSelectedPlacement(int direction)
        {
            if (SelectedPlacement == null || !CanEditManualSpecification || Activity.IsOperationActive)
            {
                return false;
            }

            int currentIndex = Placements.IndexOf(SelectedPlacement);
            int candidateIndex = currentIndex + direction;
            while (candidateIndex >= 0 && candidateIndex < Placements.Count)
            {
                if (Placements[candidateIndex].Bucket == SelectedPlacement.Bucket)
                {
                    return true;
                }

                candidateIndex += direction;
            }

            return false;
        }

        private async Task BuildDraftAsync()
        {
            Activity.Begin(
                "Build managed draft",
                ActivityStage.Inspecting,
                "Preparing the saved report setup.",
                "The source remains unchanged.");

            try
            {
                BuildDraftResult result = await _hostService.BuildManagedDraftAsync(
                    CreateSpecificationSnapshot(),
                    Activity.CancellationToken);
                _hasBuiltDraft = true;
                _hasRunChecks = false;
                _checksPassed = false;
                _isPublished = false;
                CheckLines.Clear();
                LastDraftLabel = $"{result.DraftName} · {result.OutputRows:N0} output rows";
                RaisePublishProperties();
                Activity.Complete(
                    "Managed draft is ready for checks.",
                    LastDraftLabel);
            }
            catch (OperationCanceledException)
            {
                EnsureCancelled();
            }
            catch (ReportSetupValidationException exception)
            {
                Activity.Fail(
                    "The report setup needs attention.",
                    exception.Message);
            }
            catch (Exception)
            {
                Activity.Fail(
                    "The managed draft could not be built.",
                    "Review the report setup and the latest activity, then try again.");
            }
            finally
            {
                RefreshCommandStates();
            }
        }

        private void TogglePause()
        {
            if (Activity.State == OperationState.Paused)
            {
                _hostService.RequestResume();
            }
            else
            {
                _hostService.RequestPause();
            }

            Activity.TogglePause();
            RefreshCommandStates();
        }

        private void CancelOperation()
        {
            _hostService.RequestCancel();
            Activity.Cancel();
            RefreshCommandStates();
        }

        private async Task SendChatAsync()
        {
            string request = ChatDraft.Trim();
            ChatLines.Add(new ChatLine("You", request));
            ChatDraft = string.Empty;
            Activity.Begin(
                "Review chat request",
                ActivityStage.Planning,
                "Reading the request against the saved report setup.",
                "The assistant has bounded report-specification tools only.");

            try
            {
                ChatRunResult result;
                using (SecureString? apiKey = CopyApiKey())
                {
                    result = await _hostService.RunChatAsync(
                        request,
                        CreateSpecificationSnapshot(),
                        CreateEndpointSettingsSnapshot(),
                        apiKey,
                        Activity.CancellationToken);
                }

                ChatLines.Add(new ChatLine("Builder", result.Response));
                if (result.AppliedSpecification != null)
                {
                    if (result.Published)
                    {
                        throw new InvalidOperationException(
                            "The guarded agent returned an invalid published result.");
                    }

                    if (_source != null)
                    {
                        ApplySource(new SourceSnapshot(
                            _source.DisplayName,
                            _source.Location,
                            _source.RowCount,
                            _source.Columns,
                            _source.IsSynthetic,
                            result.AppliedSpecification,
                            "Chat-applied report setup. The managed draft remains unpublished."));
                    }

                    _agentAppliedSpecification = result.AppliedSpecification;
                    _hasBuiltDraft = result.HasManagedDraft;
                    _hasRunChecks = result.Checks.Count > 0;
                    _checksPassed = result.AllChecksPassed &&
                        result.Checks.Count > 0 &&
                        result.Checks.All(check => check.Passed);
                    _isPublished = false;
                    CheckLines.Clear();
                    foreach (HostCheckResult check in result.Checks)
                    {
                        CheckLines.Add(new CheckLine(
                            check.Name,
                            check.Passed ? "Pass" : "Fail",
                            check.Detail,
                            check.Passed));
                    }

                    foreach (ChatChangeSnapshot change in result.Changes)
                    {
                        ChatLines.Add(new ChatLine(
                            "Builder",
                            change.Category + ": " + change.Description));
                    }

                    LastDraftLabel = result.ManagedDraftName + " · " +
                        result.OutputRows.ToString("N0") + " normalized rows · unpublished";
                    ApplyManualEditingState(result.AppliedSpecification);
                    RaisePublishProperties();
                    if (!_checksPassed)
                    {
                        throw new InvalidOperationException(
                            "The guarded agent result did not contain a complete passing check set.");
                    }

                    Activity.Complete(
                        "Managed draft built and checked.",
                        "Review the applied change summary. Publishing still requires your click.");
                }
                else
                {
                    Activity.Complete(
                        "Chat proposal is ready for review.",
                        "No workbook changes were published.");
                }
            }
            catch (OperationCanceledException)
            {
                EnsureCancelled();
            }
            catch (Exception)
            {
                ChatLines.Add(new ChatLine(
                    "Builder",
                    "I could not finish that request. Review the latest activity and try again."));
                Activity.Fail(
                    "The chat request could not be completed.",
                    "The saved report setup was not changed.");
            }
            finally
            {
                RefreshCommandStates();
            }
        }

        private async Task DiscoverModelsAsync()
        {
            if (!TryValidateEndpoint(requireModel: false, out string validationMessage))
            {
                SetEndpointState(false, "Needs attention", validationMessage);
                return;
            }

            Activity.Begin(
                "Discover models",
                ActivityStage.Ready,
                "Requesting the model list.",
                "The configured host service owns the network request.");

            long requestedConfigurationVersion = _endpointConfigurationVersion;
            ModelEndpointSettingsSnapshot requestedSettings = CreateEndpointSettingsSnapshot();
            bool requestHasApiKey = HasEndpointCredentialFor(requestedSettings.BaseUrl);
            try
            {
                IReadOnlyList<string> models;
                using (SecureString? apiKey = CopyApiKey())
                {
                    models = await _hostService.DiscoverModelsAsync(
                        requestedSettings,
                        apiKey,
                        Activity.CancellationToken);
                }

                if (!TryPersistCurrentEndpointSettings(
                        requestedConfigurationVersion,
                        requestedSettings,
                        requestHasApiKey))
                {
                    InvalidateEndpointCheck();
                    Activity.Complete(
                        "Endpoint settings changed.",
                        "The model list response was discarded. Run discovery again for the current settings.");
                    return;
                }

                AvailableModels.Clear();
                foreach (string model in models.Where(model => !string.IsNullOrWhiteSpace(model)).Distinct())
                {
                    AvailableModels.Add(model);
                }

                if (AvailableModels.Count > 0
                    && !AvailableModels.Any(model => string.Equals(model, ModelId, StringComparison.Ordinal)))
                {
                    ModelId = AvailableModels[0];
                }

                EndpointStateLabel = "Models found";
                EndpointValidationMessage = AvailableModels.Count == 0
                    ? "The endpoint returned no models. Enter a model ID manually."
                    : $"{AvailableModels.Count} models available. Check the endpoint before sending.";
                RaisePropertyChanged(nameof(EndpointStateBrush));
                Activity.Complete(
                    "Model discovery completed.",
                    EndpointValidationMessage);
            }
            catch (OperationCanceledException)
            {
                EnsureCancelled();
            }
            catch (Exception)
            {
                SetEndpointState(
                    false,
                    "Needs attention",
                    "Models could not be discovered. Check the base URL, API key, and network permission.");
                Activity.Fail("Model discovery failed.", EndpointValidationMessage);
            }
            finally
            {
                RefreshCommandStates();
            }
        }

        private async Task CheckEndpointAsync()
        {
            if (!TryValidateEndpoint(requireModel: true, out string validationMessage))
            {
                SetEndpointState(false, "Needs attention", validationMessage);
                return;
            }

            Activity.Begin(
                "Check model endpoint",
                ActivityStage.Ready,
                "Checking the configured endpoint and model.",
                "The API key is passed in memory and is never displayed.");

            long requestedConfigurationVersion = _endpointConfigurationVersion;
            ModelEndpointSettingsSnapshot requestedSettings = CreateEndpointSettingsSnapshot();
            bool requestHasApiKey = HasEndpointCredentialFor(requestedSettings.BaseUrl);
            try
            {
                EndpointCheckResult result;
                using (SecureString? apiKey = CopyApiKey())
                {
                    result = await _hostService.CheckEndpointAsync(
                        requestedSettings,
                        apiKey,
                        Activity.CancellationToken);
                }

                if (!TryPersistCurrentEndpointSettings(
                        requestedConfigurationVersion,
                        requestedSettings,
                        requestHasApiKey))
                {
                    InvalidateEndpointCheck();
                    Activity.Complete(
                        "Endpoint settings changed.",
                        "The check result was discarded. Check the current endpoint again before Chat can send Data.");
                    return;
                }

                SetEndpointState(
                    result.Succeeded,
                    result.Succeeded ? "Ready" : "Needs attention",
                    result.Message);
                if (result.Succeeded)
                {
                    Activity.Complete("Model endpoint is ready.", result.Message);
                }
                else
                {
                    Activity.Fail("Model endpoint check failed.", result.Message);
                }
            }
            catch (OperationCanceledException)
            {
                EnsureCancelled();
            }
            catch (Exception)
            {
                SetEndpointState(
                    false,
                    "Needs attention",
                    "The endpoint could not be reached. Check the base URL, model ID, API key, and network permission.");
                Activity.Fail("Model endpoint check failed.", EndpointValidationMessage);
            }
            finally
            {
                RefreshCommandStates();
            }
        }

        private async Task RunChecksAsync()
        {
            Activity.Begin(
                "Run checks",
                ActivityStage.Checking,
                "Starting independent checks.",
                "Publishing remains blocked until every required check passes.");

            try
            {
                IReadOnlyList<HostCheckResult> results = await _hostService.RunChecksAsync(
                    Activity.CancellationToken);
                CheckLines.Clear();
                foreach (HostCheckResult result in results)
                {
                    CheckLines.Add(new CheckLine(
                        result.Name,
                        result.Passed ? "Pass" : "Fail",
                        result.Detail,
                        result.Passed));
                }

                _hasRunChecks = true;
                _checksPassed = results.Count > 0 && results.All(result => result.Passed);
                _isPublished = false;
                RaisePublishProperties();
                if (_checksPassed)
                {
                    Activity.Complete(
                        "All required checks passed.",
                        "The managed draft is ready to publish.");
                }
                else
                {
                    Activity.Fail(
                        "One or more checks failed.",
                        "Publishing remains blocked until the failed checks are resolved.");
                }
            }
            catch (OperationCanceledException)
            {
                EnsureCancelled();
            }
            catch (Exception)
            {
                Activity.Fail(
                    "Checks could not be completed.",
                    "Publishing remains blocked. Run checks again after reviewing the host connection.");
            }
            finally
            {
                RefreshCommandStates();
            }
        }

        private async Task PublishAsync()
        {
            Activity.Begin(
                "Publish managed draft",
                ActivityStage.Checking,
                "Confirming the publish gate.",
                "Every required check has passed.");

            try
            {
                PublishResult result = await _hostService.PublishManagedDraftAsync(
                    Activity.CancellationToken);
                _isPublished = true;
                RaisePublishProperties();
                Activity.Complete("Managed draft published.", result.Message);
            }
            catch (OperationCanceledException)
            {
                EnsureCancelled();
            }
            catch (Exception)
            {
                Activity.Fail(
                    "The managed draft could not be published.",
                    "The draft remains managed and unpublished.");
            }
            finally
            {
                RefreshCommandStates();
            }
        }

        private bool CanStartOperation()
        {
            return !Activity.IsOperationActive;
        }

        private bool CanConfirmPeriodMapping()
        {
            if (_source == null
                || string.IsNullOrWhiteSpace(SelectedPeriodMode)
                || !CanEditManualSpecification
                || Activity.IsOperationActive)
            {
                return false;
            }

            if (!IsWideHeaderMode)
            {
                return true;
            }

            return _wideHeaderPreview != null
                && _wideHeaderPreview.TotalPreservation == TotalPreservationState.Pass
                && (!ReportingYearRequired || TryGetReportingYear(out _));
        }

        private bool CanPreviewWideHeaderMapping()
        {
            bool yearIsBlank = string.IsNullOrWhiteSpace(ReportingYearText);
            return _source != null
                && IsWideHeaderMode
                && CanEditManualSpecification
                && (yearIsBlank || TryGetReportingYear(out _))
                && !Activity.IsOperationActive;
        }

        private bool CanBuildDraft()
        {
            bool hasExactCanonicalSetup =
                _agentAppliedSpecification != null &&
                _agentAppliedSpecification.HasCanonicalReportSpec;
            return _source != null
                && _periodMappingConfirmed
                && (hasExactCanonicalSetup ||
                    (Placements.Any(placement => placement.Bucket == PlacementBucket.Rows) &&
                     Placements.Any(placement => placement.Bucket == PlacementBucket.Values)))
                && !Activity.IsOperationActive;
        }

        private bool CanSendChat()
        {
            return !string.IsNullOrWhiteSpace(ChatDraft)
                && _endpointCheckPassed
                && (!IsRemoteEndpoint() || AllowRemoteWorkbookData)
                && !Activity.IsOperationActive;
        }

        private bool CanDiscoverModels()
        {
            return !Activity.IsOperationActive
                && TryValidateEndpoint(requireModel: false, out _);
        }

        private bool CanCheckEndpoint()
        {
            return !Activity.IsOperationActive
                && TryValidateEndpoint(requireModel: true, out _);
        }

        private bool CanRunChecks()
        {
            return _hasBuiltDraft && !Activity.IsOperationActive;
        }

        private void ApplySource(SourceSnapshot source)
        {
            _source = source;
            _agentAppliedSpecification = null;
            _manualProjectionComplete = true;
            _manualRestrictionMessageShown = false;
            Columns.Clear();
            AvailableFields.Clear();
            Placements.Clear();
            TransformRules.Clear();
            CalculatedMetrics.Clear();
            ReportBlocks.Clear();
            RequiredChecks.Clear();
            WideHeaderMappings.Clear();
            NormalizedSampleRows.Clear();
            foreach (SourceColumnSnapshot column in source.Columns)
            {
                Columns.Add(column);
                AvailableFields.Add(column);
            }

            ReportSpecificationSnapshot? saved = source.SavedReportSetup;
            if (saved != null)
            {
                PeriodMappingSnapshot mapping = saved.PeriodMapping;
                _selectedPeriodMode = mapping.Mode;
                _selectedPeriodColumn = mapping.PeriodColumn;
                _headerPattern = mapping.HeaderPattern;
                _reportingYearText = mapping.ReportingYear?.ToString() ?? string.Empty;
                _selectedOutputStyle = saved.OutputStyle;
                foreach (WideHeaderMappingRowSnapshot row in mapping.WideHeaderMappings)
                {
                    WideHeaderMappings.Add(row);
                }

                foreach (FieldPlacementSnapshot placement in saved.Placements)
                {
                    EnsureSavedFieldIsAvailable(placement.ColumnName);
                    Placements.Add(new FieldPlacement(
                        placement.Bucket,
                        placement.ColumnName,
                        placement.Setting,
                        placement.ShowSubtotals,
                        placement.SubtotalPlacement,
                        string.Join("; ", placement.MemberOrder),
                        placement.NumberFormat));
                }

                foreach (ManualTransformSnapshot transform in saved.Transforms)
                {
                    TransformRules.Add(new ManualTransformRule
                    {
                        Operation = transform.Operation,
                        Column = transform.Column,
                        OutputColumn = transform.OutputColumn,
                        Details = transform.Details
                    });
                }

                foreach (ManualCalculatedMetricSnapshot metric in saved.CalculatedMetrics)
                {
                    CalculatedMetrics.Add(new ManualCalculatedMetricRule
                    {
                        Label = metric.Label,
                        Kind = metric.Kind,
                        Primary = metric.Primary,
                        Secondary = metric.Secondary,
                        Details = metric.Details,
                        NumberFormat = metric.NumberFormat
                    });
                }

                foreach (ManualReportBlockSnapshot block in saved.Blocks)
                {
                    ReportBlocks.Add(new ManualReportBlockRule(
                        block.StableId,
                        block.CanonicalBlockId,
                        block.CanonicalOwnershipId)
                    {
                        Title = block.Title,
                        WorksheetName = block.WorksheetName,
                        AnchorCell = block.AnchorCell,
                        OutputStyle = block.OutputStyle,
                        OwnedRows = block.OwnedRows,
                        OwnedColumns = block.OwnedColumns
                    });
                }

                foreach (ManualCheckSnapshot check in saved.Checks)
                {
                    RequiredChecks.Add(new ManualCheckRule
                    {
                        Kind = check.Kind,
                        Metric = check.Metric,
                        ComparedMetric = check.ComparedMetric,
                        ToleranceText = check.Tolerance.ToString(CultureInfo.InvariantCulture)
                    });
                }

                _repeatRowLabels = saved.Layout.RepeatRowLabels;
                _insertBlankRows = saved.Layout.InsertBlankRows;
                _freezeHeaders = saved.Layout.FreezeHeaders;
                _showRowGrandTotals = saved.Layout.ShowRowGrandTotals;
                _showColumnGrandTotals = saved.Layout.ShowColumnGrandTotals;
                _rowIndent = saved.Layout.RowIndent;
                _rowGrandTotalLabel = saved.Layout.RowGrandTotalLabel;
                _columnGrandTotalLabel = saved.Layout.ColumnGrandTotalLabel;

                if (ReportBlocks.Count == 0 && saved.ManualProjectionComplete)
                {
                    ReportBlocks.Add(new ManualReportBlockRule
                    {
                        Title = "Management report",
                        WorksheetName = "Report",
                        AnchorCell = "A1",
                        OutputStyle = saved.OutputStyle,
                        OwnedRows = 500,
                        OwnedColumns = 64
                    });
                }

                SelectedPlacement = Placements.FirstOrDefault();
                SelectedTransformRule = TransformRules.FirstOrDefault();
                SelectedCalculatedMetric = CalculatedMetrics.FirstOrDefault();
                SelectedReportBlock = ReportBlocks.FirstOrDefault();
                SelectedCheckRule = RequiredChecks.FirstOrDefault();
                _periodMappingConfirmed = true;
                ChatLines.Add(new ChatLine(
                    "Builder",
                    "A compatible saved report setup was restored for this exact workbook Data object."));
            }
            else
            {
                SourceColumnSnapshot? dateColumn = AvailableFields.FirstOrDefault(
                    column => string.Equals(column.Name, "Date", StringComparison.OrdinalIgnoreCase));
                _selectedPeriodMode = "Date column";
                _selectedPeriodColumn = dateColumn?.Name ?? AvailableFields.FirstOrDefault()?.Name ?? string.Empty;
                _headerPattern = string.Empty;
                _reportingYearText = string.Empty;
                _selectedOutputStyle = "Dense management block";
                _repeatRowLabels = false;
                _insertBlankRows = false;
                _freezeHeaders = true;
                _showRowGrandTotals = true;
                _showColumnGrandTotals = true;
                _rowIndent = 1;
                _rowGrandTotalLabel = "Grand Total";
                _columnGrandTotalLabel = "Grand Total";
                ReportBlocks.Add(new ManualReportBlockRule
                {
                    Title = "Management report",
                    WorksheetName = "Report",
                    AnchorCell = "A1",
                    OutputStyle = _selectedOutputStyle,
                    OwnedRows = 500,
                    OwnedColumns = 64
                });
                SelectedReportBlock = ReportBlocks[0];
                SelectedPlacement = null;
                _periodMappingConfirmed = false;
            }

            SelectedField = AvailableFields.FirstOrDefault();
            RaiseSourceProperties();
            _wideHeaderPreview = null;
            _reportingYearRequired = false;
            NormalizedSampleRows.Clear();
            RaisePropertyChanged(nameof(SelectedPeriodMode));
            RaisePropertyChanged(nameof(IsWideHeaderMode));
            RaisePropertyChanged(nameof(SelectedPeriodColumn));
            RaisePropertyChanged(nameof(HeaderPattern));
            RaisePropertyChanged(nameof(ReportingYearText));
            RaisePropertyChanged(nameof(SelectedOutputStyle));
            RaisePropertyChanged(nameof(RepeatRowLabels));
            RaisePropertyChanged(nameof(InsertBlankRows));
            RaisePropertyChanged(nameof(FreezeHeaders));
            RaisePropertyChanged(nameof(ShowRowGrandTotals));
            RaisePropertyChanged(nameof(ShowColumnGrandTotals));
            RaisePropertyChanged(nameof(RowIndent));
            RaisePropertyChanged(nameof(RowGrandTotalLabel));
            RaisePropertyChanged(nameof(ColumnGrandTotalLabel));
            RaiseWideHeaderProperties();
            RaisePeriodProperties();
            MarkSpecificationDirty();
            if (saved != null && saved.HasCanonicalReportSpec)
            {
                _agentAppliedSpecification = saved;
            }
            _periodMappingConfirmed = saved != null;
            RaisePeriodProperties();
            LastDraftLabel = string.IsNullOrWhiteSpace(source.SavedReportSetupStatus)
                ? saved != null
                    ? "Compatible saved report setup restored"
                    : "No managed draft built"
                : source.SavedReportSetupStatus;
            ApplyManualEditingState(saved);
        }

        private void EnsureSavedFieldIsAvailable(string fieldName)
        {
            if (AvailableFields.Any(field => string.Equals(
                    field.Name,
                    fieldName,
                    StringComparison.OrdinalIgnoreCase)))
            {
                return;
            }

            string dataType = string.Equals(fieldName, "Period", StringComparison.OrdinalIgnoreCase)
                ? "Date"
                : string.Equals(fieldName, "Value", StringComparison.OrdinalIgnoreCase)
                    ? "Number"
                    : "Text";
            var derived = new SourceColumnSnapshot(
                fieldName,
                dataType,
                "Created by period normalization");
            AvailableFields.Add(derived);
        }

        private PeriodMappingSnapshot CreatePeriodMappingSnapshot()
        {
            int? reportingYear = TryGetReportingYear(out int parsedYear) ? parsedYear : (int?)null;
            return new PeriodMappingSnapshot(
                SelectedPeriodMode,
                SelectedPeriodColumn,
                HeaderPattern,
                reportingYear,
                WideHeaderMappings.ToArray());
        }

        private ModelEndpointSettingsSnapshot CreateEndpointSettingsSnapshot()
        {
            return new ModelEndpointSettingsSnapshot(
                EndpointBaseUrl.Trim(),
                ModelId.Trim(),
                AllowRemoteHttp,
                AllowRemoteWorkbookData);
        }

        private bool IsRemoteEndpoint()
        {
            return Uri.TryCreate(EndpointBaseUrl.Trim(), UriKind.Absolute, out var endpoint) &&
                !endpoint.IsLoopback;
        }

        private ReportSpecificationSnapshot CreateSpecificationSnapshot()
        {
            return _agentAppliedSpecification ?? new ReportSpecificationSnapshot(
                CreatePeriodMappingSnapshot(),
                Placements.Select(placement => placement.ToSnapshot()).ToArray(),
                SelectedOutputStyle,
                transforms: TransformRules.Select(rule => rule.ToSnapshot()).ToArray(),
                calculatedMetrics: CalculatedMetrics.Select(rule => rule.ToSnapshot()).ToArray(),
                blocks: ReportBlocks.Select(rule => rule.ToSnapshot()).ToArray(),
                layout: new ManualLayoutSnapshot
                {
                    RepeatRowLabels = RepeatRowLabels,
                    InsertBlankRows = InsertBlankRows,
                    FreezeHeaders = FreezeHeaders,
                    ShowRowGrandTotals = ShowRowGrandTotals,
                    ShowColumnGrandTotals = ShowColumnGrandTotals,
                    RowIndent = RowIndent,
                    RowGrandTotalLabel = RowGrandTotalLabel,
                    ColumnGrandTotalLabel = ColumnGrandTotalLabel
                },
                checks: RequiredChecks.Select(rule => rule.ToSnapshot()).ToArray());
        }

        public bool CanEditManualSpecification => _manualProjectionComplete;

        public string ManualEditingRestrictionMessage =>
            "This setup contains settings the manual editor cannot represent exactly. " +
            "The preview is read-only; rebuild it unchanged or ask Chat for a bounded change.";

        private bool DemandManualEditing(string propertyName)
        {
            if (CanEditManualSpecification)
            {
                return true;
            }

            ShowManualEditingRestriction();
            RaisePropertyChanged(propertyName);
            return false;
        }

        private void ShowManualEditingRestriction()
        {
            string message = ManualEditingRestrictionMessage;
            LastDraftLabel = message;
            if (!_manualRestrictionMessageShown)
            {
                ChatLines.Add(new ChatLine("Builder", message));
                _manualRestrictionMessageShown = true;
            }

            RefreshCommandStates();
        }

        private void ApplyManualEditingState(ReportSpecificationSnapshot? snapshot)
        {
            _applyingManualEditingState = true;
            try
            {
                _manualProjectionComplete = snapshot == null || snapshot.ManualProjectionComplete;
                RaisePropertyChanged(nameof(CanEditManualSpecification));
                RaisePropertyChanged(nameof(ManualEditingRestrictionMessage));
                _manualRestrictionMessageShown = false;
                bool isReadOnly = !_manualProjectionComplete;
                foreach (FieldPlacement placement in Placements)
                {
                    placement.IsReadOnly = isReadOnly;
                }

                foreach (ManualEditableObject rule in TransformRules.Cast<ManualEditableObject>()
                             .Concat(CalculatedMetrics)
                             .Concat(ReportBlocks)
                             .Concat(RequiredChecks))
                {
                    rule.IsReadOnly = isReadOnly;
                }

                if (isReadOnly)
                {
                    ShowManualEditingRestriction();
                }
                else
                {
                    RefreshCommandStates();
                }
            }
            finally
            {
                _applyingManualEditingState = false;
            }
        }

        private void InvalidatePeriodMapping()
        {
            _periodMappingConfirmed = false;
            RaisePeriodProperties();
            MarkSpecificationDirty();
        }

        private void ApplyWideHeaderPreview(WideHeaderMappingPreview preview)
        {
            _wideHeaderPreview = preview;
            _reportingYearRequired = preview.RequiresReportingYear;
            WideHeaderMappings.Clear();
            foreach (WideHeaderMappingRowSnapshot mapping in preview.HeaderMappings)
            {
                WideHeaderMappings.Add(mapping);
            }

            NormalizedSampleRows.Clear();
            foreach (NormalizedSampleRowSnapshot sampleRow in preview.SampleRows)
            {
                NormalizedSampleRows.Add(sampleRow);
            }

            RaiseWideHeaderProperties();
            ConfirmPeriodMappingCommand.RaiseCanExecuteChanged();
        }

        private void ClearWideHeaderPreview(bool keepReportingYearRequirement)
        {
            bool wasYearRequired = ReportingYearRequired;
            _wideHeaderPreview = null;
            _reportingYearRequired = keepReportingYearRequirement && wasYearRequired;
            WideHeaderMappings.Clear();
            NormalizedSampleRows.Clear();
            RaiseWideHeaderProperties();
        }

        private void RaiseWideHeaderProperties()
        {
            RaisePropertyChanged(nameof(ReportingYearRequired));
            RaisePropertyChanged(nameof(ReportingYearRequirementLabel));
            RaisePropertyChanged(nameof(MappingPreviewState));
            RaisePropertyChanged(nameof(ProjectedNormalizedRowsLabel));
            RaisePropertyChanged(nameof(TotalPreservationLabel));
            RaisePropertyChanged(nameof(TotalPreservationDetail));
            RaisePropertyChanged(nameof(TotalPreservationBrush));
        }

        private bool TryGetReportingYear(out int reportingYear)
        {
            return int.TryParse(ReportingYearText, out reportingYear)
                && reportingYear >= 1900
                && reportingYear <= 9999;
        }

        private bool TryValidateEndpoint(bool requireModel, out string validationMessage)
        {
            if (!Uri.TryCreate(EndpointBaseUrl.Trim(), UriKind.Absolute, out Uri? endpointUri)
                || (endpointUri.Scheme != Uri.UriSchemeHttp && endpointUri.Scheme != Uri.UriSchemeHttps))
            {
                validationMessage = "Enter an absolute HTTP or HTTPS base URL.";
                return false;
            }

            if (endpointUri.Scheme == Uri.UriSchemeHttp
                && !endpointUri.IsLoopback
                && !AllowRemoteHttp)
            {
                validationMessage = "Remote plain HTTP is blocked. Use HTTPS or explicitly allow remote HTTP.";
                return false;
            }

            if (requireModel && string.IsNullOrWhiteSpace(ModelId))
            {
                validationMessage = "Enter or discover a model ID.";
                return false;
            }

            validationMessage = endpointUri.Scheme == Uri.UriSchemeHttp && !endpointUri.IsLoopback
                ? "Remote plain HTTP is explicitly allowed. Credentials and requests are not encrypted."
                : "Endpoint settings are valid. Run Check endpoint to verify connectivity.";
            return true;
        }

        private void ResetEndpointScopedSecurityState(string nextBaseUrl)
        {
            _apiKey?.Dispose();
            _apiKey = null;
            _savedApiKeyAvailable = _savedCredentialHasProtectedKey &&
                _savedCredentialBaseUrl != null &&
                AgentEndpointCredentialScope.Matches(
                    _savedCredentialBaseUrl,
                    nextBaseUrl);

            if (_allowRemoteHttp)
            {
                _allowRemoteHttp = false;
                RaisePropertyChanged(nameof(AllowRemoteHttp));
            }

            if (_allowRemoteWorkbookData)
            {
                _allowRemoteWorkbookData = false;
                RaisePropertyChanged(nameof(AllowRemoteWorkbookData));
            }

            RaisePropertyChanged(nameof(ApiKeyStateLabel));
            ApiKeyClearRequested?.Invoke(this, EventArgs.Empty);
        }

        private bool HasEndpointCredentialFor(string baseUrl)
        {
            return (_apiKey != null && _apiKey.Length > 0) ||
                (_savedCredentialHasProtectedKey &&
                 _savedCredentialBaseUrl != null &&
                 AgentEndpointCredentialScope.Matches(_savedCredentialBaseUrl, baseUrl));
        }

        private void RecordPersistedEndpoint(string baseUrl, bool hasProtectedKey)
        {
            _savedCredentialBaseUrl = baseUrl;
            _savedCredentialHasProtectedKey = hasProtectedKey;
            _savedApiKeyAvailable = hasProtectedKey &&
                AgentEndpointCredentialScope.Matches(baseUrl, EndpointBaseUrl);
            RaisePropertyChanged(nameof(ApiKeyStateLabel));
        }

        private bool TryPersistCurrentEndpointSettings(
            long requestedConfigurationVersion,
            ModelEndpointSettingsSnapshot requestedSettings,
            bool requestHasApiKey)
        {
            if (requestedConfigurationVersion != _endpointConfigurationVersion)
            {
                return false;
            }

            // This is a bounded local DPAPI and file operation. Complete it without
            // yielding so UI-bound endpoint edits cannot interleave after the version gate.
            using (SecureString? apiKey = CopyApiKey())
            {
                _hostService.PersistEndpointSettingsAsync(
                        requestedSettings,
                        apiKey,
                        Activity.CancellationToken)
                    .GetAwaiter()
                    .GetResult();
            }

            if (requestedConfigurationVersion != _endpointConfigurationVersion)
            {
                return false;
            }

            RecordPersistedEndpoint(requestedSettings.BaseUrl, requestHasApiKey);
            return true;
        }

        private void InvalidateEndpointCheck()
        {
            _endpointCheckPassed = false;
            EndpointStateLabel = "Not checked";
            TryValidateEndpoint(requireModel: true, out string validationMessage);
            EndpointValidationMessage = validationMessage;
            RaisePropertyChanged(nameof(EndpointStateBrush));
            DiscoverModelsCommand.RaiseCanExecuteChanged();
            CheckEndpointCommand.RaiseCanExecuteChanged();
            SendChatCommand.RaiseCanExecuteChanged();
        }

        private void SetEndpointState(bool succeeded, string stateLabel, string validationMessage)
        {
            _endpointCheckPassed = succeeded;
            EndpointStateLabel = stateLabel;
            EndpointValidationMessage = validationMessage;
            RaisePropertyChanged(nameof(EndpointStateBrush));
            SendChatCommand.RaiseCanExecuteChanged();
        }

        private SecureString? CopyApiKey()
        {
            return _apiKey?.Copy();
        }

        private void MarkSpecificationDirty()
        {
            if (_applyingManualEditingState)
            {
                return;
            }

            if (!CanEditManualSpecification && _agentAppliedSpecification != null)
            {
                ShowManualEditingRestriction();
                return;
            }

            _agentAppliedSpecification = null;
            _hasBuiltDraft = false;
            _hasRunChecks = false;
            _checksPassed = false;
            _isPublished = false;
            LastDraftLabel = "Report setup changed · build a new managed draft";
            CheckLines.Clear();
            RaisePublishProperties();
            RefreshCommandStates();
        }

        private void RaiseSourceProperties()
        {
            RaisePropertyChanged(nameof(SourceName));
            RaisePropertyChanged(nameof(SourceLocation));
            RaisePropertyChanged(nameof(SourceSummary));
            RaisePropertyChanged(nameof(SourceKindLabel));
        }

        private void RaisePeriodProperties()
        {
            RaisePropertyChanged(nameof(PeriodStatusLabel));
            RaisePropertyChanged(nameof(PeriodStatusBrush));
            PreviewWideHeaderMappingCommand.RaiseCanExecuteChanged();
            ConfirmPeriodMappingCommand.RaiseCanExecuteChanged();
        }

        private void RaisePublishProperties()
        {
            RaisePropertyChanged(nameof(CanPublish));
            RaisePropertyChanged(nameof(PublishGateState));
            RaisePropertyChanged(nameof(PublishGateLabel));
            RaisePropertyChanged(nameof(PublishGateBrush));
        }

        private void RefreshCommandStates()
        {
            SelectSourceCommand.RaiseCanExecuteChanged();
            ConfirmPeriodMappingCommand.RaiseCanExecuteChanged();
            PreviewWideHeaderMappingCommand.RaiseCanExecuteChanged();
            AddPlacementCommand.RaiseCanExecuteChanged();
            RemovePlacementCommand.RaiseCanExecuteChanged();
            MovePlacementUpCommand.RaiseCanExecuteChanged();
            MovePlacementDownCommand.RaiseCanExecuteChanged();
            AddTransformCommand.RaiseCanExecuteChanged();
            RemoveTransformCommand.RaiseCanExecuteChanged();
            AddCalculatedMetricCommand.RaiseCanExecuteChanged();
            RemoveCalculatedMetricCommand.RaiseCanExecuteChanged();
            AddReportBlockCommand.RaiseCanExecuteChanged();
            RemoveReportBlockCommand.RaiseCanExecuteChanged();
            AddRequiredCheckCommand.RaiseCanExecuteChanged();
            RemoveRequiredCheckCommand.RaiseCanExecuteChanged();
            BuildDraftCommand.RaiseCanExecuteChanged();
            TogglePauseCommand.RaiseCanExecuteChanged();
            CancelCommand.RaiseCanExecuteChanged();
            SendChatCommand.RaiseCanExecuteChanged();
            DiscoverModelsCommand.RaiseCanExecuteChanged();
            CheckEndpointCommand.RaiseCanExecuteChanged();
            RunChecksCommand.RaiseCanExecuteChanged();
            PublishCommand.RaiseCanExecuteChanged();
            RaisePublishProperties();
        }

        private void EnsureCancelled()
        {
            if (Activity.State != OperationState.Cancelled)
            {
                Activity.Cancel();
            }
        }

        private void OnPlacementsChanged(object? sender, NotifyCollectionChangedEventArgs eventArgs)
        {
            foreach (FieldPlacement removed in _subscribedPlacements
                         .Where(placement => !Placements.Contains(placement))
                         .ToArray())
            {
                removed.PropertyChanged -= OnPlacementPropertyChanged;
                _subscribedPlacements.Remove(removed);
            }

            foreach (FieldPlacement placement in Placements)
            {
                if (_subscribedPlacements.Add(placement))
                {
                    placement.PropertyChanged += OnPlacementPropertyChanged;
                }
            }

            MovePlacementUpCommand.RaiseCanExecuteChanged();
            MovePlacementDownCommand.RaiseCanExecuteChanged();
            BuildDraftCommand.RaiseCanExecuteChanged();
        }

        private void OnPlacementPropertyChanged(object? sender, PropertyChangedEventArgs eventArgs)
        {
            MarkSpecificationDirty();
        }

        private void OnManualRulesChanged(object? sender, NotifyCollectionChangedEventArgs eventArgs)
        {
            var activeRules = TransformRules.Cast<ObservableObject>()
                .Concat(CalculatedMetrics)
                .Concat(ReportBlocks)
                .Concat(RequiredChecks)
                .ToArray();
            foreach (ObservableObject removed in _subscribedManualRules
                         .Where(rule => !activeRules.Contains(rule))
                         .ToArray())
            {
                removed.PropertyChanged -= OnManualRulePropertyChanged;
                _subscribedManualRules.Remove(removed);
            }

            foreach (ObservableObject rule in activeRules)
            {
                if (_subscribedManualRules.Add(rule))
                {
                    rule.PropertyChanged += OnManualRulePropertyChanged;
                }
            }

            AddReportBlockCommand.RaiseCanExecuteChanged();
            RemoveReportBlockCommand.RaiseCanExecuteChanged();
            RemoveTransformCommand.RaiseCanExecuteChanged();
            RemoveCalculatedMetricCommand.RaiseCanExecuteChanged();
            RemoveRequiredCheckCommand.RaiseCanExecuteChanged();
        }

        private void OnManualRulePropertyChanged(object? sender, PropertyChangedEventArgs eventArgs)
        {
            MarkSpecificationDirty();
        }

        private void OnActivityPropertyChanged(object? sender, PropertyChangedEventArgs eventArgs)
        {
            if (eventArgs.PropertyName == nameof(OperationActivityController.State)
                || eventArgs.PropertyName == nameof(OperationActivityController.IsOperationActive))
            {
                RefreshCommandStates();
            }
        }

        private void OnHostActivityReported(object? sender, HostActivityEventArgs eventArgs)
        {
            if (eventArgs.Kind == ActivityKind.Heartbeat)
            {
                Activity.Heartbeat(eventArgs.Detail);
                return;
            }

            Activity.Report(eventArgs.Stage, eventArgs.Message, eventArgs.Detail, eventArgs.Kind);
        }
    }
}
