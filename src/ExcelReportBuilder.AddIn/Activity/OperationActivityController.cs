using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Threading;
using System.Windows.Media;
using System.Windows.Threading;
using ExcelReportBuilder.AddIn.Presentation;

namespace ExcelReportBuilder.AddIn.Activity
{
    public sealed class OperationActivityController : ObservableObject, IDisposable
    {
        public static readonly TimeSpan HeartbeatInterval = TimeSpan.FromSeconds(15);

        private const int MaximumTimelineEntries = 200;

        private readonly Dispatcher _dispatcher;
        private readonly DispatcherTimer _clock;
        private readonly Stopwatch _elapsed = new Stopwatch();
        private CancellationTokenSource? _cancellation;
        private DateTimeOffset _lastHeartbeat;
        private OperationState _state;
        private ActivityStage _currentStage;
        private string _operationName = "No active operation";
        private string _elapsedLabel = "00:00";
        private bool _disposed;

        public OperationActivityController()
        {
            _dispatcher = Dispatcher.CurrentDispatcher;
            Entries = new ObservableCollection<ActivityEntry>();
            _state = OperationState.Idle;
            _currentStage = ActivityStage.Ready;
            _clock = new DispatcherTimer(DispatcherPriority.Background, _dispatcher)
            {
                Interval = TimeSpan.FromSeconds(1)
            };
            _clock.Tick += OnClockTick;
            _clock.Start();

            AddEntry(
                ActivityStage.Ready,
                ActivityKind.Result,
                "Ready for a report setup.",
                "Choose workbook data or open a saved report setup.");
        }

        public ObservableCollection<ActivityEntry> Entries { get; }

        public OperationState State
        {
            get => _state;
            private set
            {
                if (!SetProperty(ref _state, value))
                {
                    return;
                }

                RaisePropertyChanged(nameof(StateLabel));
                RaisePropertyChanged(nameof(StateBrush));
                RaisePropertyChanged(nameof(StateForeground));
                RaisePropertyChanged(nameof(CanPause));
                RaisePropertyChanged(nameof(CanCancel));
                RaisePropertyChanged(nameof(PauseLabel));
                RaisePropertyChanged(nameof(IsOperationActive));
            }
        }

        public ActivityStage CurrentStage
        {
            get => _currentStage;
            private set
            {
                if (!SetProperty(ref _currentStage, value))
                {
                    return;
                }

                RaisePropertyChanged(nameof(CurrentStageLabel));
                RaisePropertyChanged(nameof(StagePositionLabel));
                RaisePropertyChanged(nameof(StageProgress));
            }
        }

        public string OperationName
        {
            get => _operationName;
            private set => SetProperty(ref _operationName, value);
        }

        public string ElapsedLabel
        {
            get => _elapsedLabel;
            private set => SetProperty(ref _elapsedLabel, value);
        }

        public string StateLabel => ActivityLabels.State(State);

        public string CurrentStageLabel => ActivityLabels.Stage(CurrentStage);

        public string StagePositionLabel => CurrentStage == ActivityStage.Ready
            ? "Not started"
            : CurrentStage == ActivityStage.Complete
                ? "9 of 9"
                : $"{(int)CurrentStage} of 9";

        public double StageProgress => Math.Max(0, Math.Min(100, (int)CurrentStage * (100d / 9d)));

        public bool CanPause => State == OperationState.Running || State == OperationState.Paused;

        public bool CanCancel => State == OperationState.Running || State == OperationState.Paused;

        public bool IsOperationActive => CanCancel;

        public string PauseLabel => State == OperationState.Paused ? "_Resume" : "_Pause";

        public Brush StateBrush
        {
            get
            {
                switch (State)
                {
                    case OperationState.Running:
                        return Brushes.SeaGreen;
                    case OperationState.Paused:
                        return Brushes.DarkGoldenrod;
                    case OperationState.Cancelled:
                    case OperationState.Failed:
                        return Brushes.Firebrick;
                    case OperationState.Completed:
                        return Brushes.DarkGreen;
                    default:
                        return Brushes.SlateGray;
                }
            }
        }

        public Brush StateForeground
        {
            get
            {
                switch (State)
                {
                    case OperationState.Paused:
                        return new SolidColorBrush(Color.FromRgb(105, 68, 0));
                    case OperationState.Cancelled:
                    case OperationState.Failed:
                        return new SolidColorBrush(Color.FromRgb(137, 35, 38));
                    case OperationState.Running:
                    case OperationState.Completed:
                        return new SolidColorBrush(Color.FromRgb(26, 102, 62));
                    default:
                        return new SolidColorBrush(Color.FromRgb(75, 89, 82));
                }
            }
        }

        public CancellationToken CancellationToken => _cancellation?.Token ?? CancellationToken.None;

        public void Begin(
            string operationName,
            ActivityStage firstStage,
            string message,
            string detail = "")
        {
            Dispatch(() =>
            {
                ThrowIfDisposed();
                _cancellation?.Dispose();
                _cancellation = new CancellationTokenSource();
                _elapsed.Restart();
                _lastHeartbeat = DateTimeOffset.UtcNow;
                OperationName = operationName;
                CurrentStage = firstStage;
                State = OperationState.Running;
                UpdateElapsedLabel();
                AddEntry(firstStage, ActivityKind.Progress, message, detail);
            });
        }

        public void Report(
            ActivityStage stage,
            string message,
            string detail = "",
            ActivityKind kind = ActivityKind.Progress)
        {
            Dispatch(() =>
            {
                ThrowIfDisposed();
                CurrentStage = stage;
                if (State == OperationState.Idle
                    || State == OperationState.Cancelled
                    || State == OperationState.Completed
                    || State == OperationState.Failed)
                {
                    State = OperationState.Running;
                    _elapsed.Start();
                }

                _lastHeartbeat = DateTimeOffset.UtcNow;
                AddEntry(stage, kind, message, detail);
            });
        }

        public void Heartbeat(string detail = "Waiting for the next stage update.")
        {
            Dispatch(() =>
            {
                ThrowIfDisposed();
                _lastHeartbeat = DateTimeOffset.UtcNow;
                AddEntry(CurrentStage, ActivityKind.Heartbeat, "Activity heartbeat.", detail);
            });
        }

        public void TogglePause()
        {
            Dispatch(() =>
            {
                ThrowIfDisposed();
                if (State == OperationState.Running)
                {
                    _elapsed.Stop();
                    State = OperationState.Paused;
                    _lastHeartbeat = DateTimeOffset.UtcNow;
                    AddEntry(
                        CurrentStage,
                        ActivityKind.Control,
                        "Operation paused.",
                        "The current stage is retained until you resume or cancel.");
                }
                else if (State == OperationState.Paused)
                {
                    _elapsed.Start();
                    State = OperationState.Running;
                    _lastHeartbeat = DateTimeOffset.UtcNow;
                    AddEntry(CurrentStage, ActivityKind.Control, "Operation resumed.", CurrentStageLabel);
                }
            });
        }

        public void Cancel()
        {
            Dispatch(() =>
            {
                ThrowIfDisposed();
                if (!CanCancel)
                {
                    return;
                }

                _cancellation?.Cancel();
                _elapsed.Stop();
                State = OperationState.Cancelled;
                AddEntry(
                    CurrentStage,
                    ActivityKind.Control,
                    "Cancellation requested.",
                    "Managed draft work remains unpublished.");
            });
        }

        public void Complete(string message, string detail = "")
        {
            Dispatch(() =>
            {
                ThrowIfDisposed();
                _elapsed.Stop();
                CurrentStage = ActivityStage.Complete;
                State = OperationState.Completed;
                AddEntry(ActivityStage.Complete, ActivityKind.Result, message, detail);
                UpdateElapsedLabel();
            });
        }

        public void Fail(string message, string detail)
        {
            Dispatch(() =>
            {
                ThrowIfDisposed();
                _elapsed.Stop();
                State = OperationState.Failed;
                AddEntry(CurrentStage, ActivityKind.Error, message, detail);
                UpdateElapsedLabel();
            });
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            _clock.Stop();
            _clock.Tick -= OnClockTick;
            _cancellation?.Dispose();
            _cancellation = null;
        }

        private void OnClockTick(object? sender, EventArgs eventArgs)
        {
            if (_disposed)
            {
                return;
            }

            UpdateElapsedLabel();

            if (!IsOperationActive || DateTimeOffset.UtcNow - _lastHeartbeat < HeartbeatInterval)
            {
                return;
            }

            _lastHeartbeat = DateTimeOffset.UtcNow;
            if (State == OperationState.Paused)
            {
                AddEntry(
                    CurrentStage,
                    ActivityKind.Heartbeat,
                    "Operation remains paused.",
                    "Resume or cancel when ready.");
            }
            else
            {
                AddEntry(
                    CurrentStage,
                    ActivityKind.Heartbeat,
                    "Status heartbeat.",
                    "No new stage update in the last 15 seconds. Pause and cancel remain available.");
            }
        }

        private void UpdateElapsedLabel()
        {
            TimeSpan elapsed = _elapsed.Elapsed;
            ElapsedLabel = elapsed.TotalHours >= 1
                ? elapsed.ToString(@"hh\:mm\:ss")
                : elapsed.ToString(@"mm\:ss");
        }

        private void AddEntry(
            ActivityStage stage,
            ActivityKind kind,
            string message,
            string detail)
        {
            Entries.Add(new ActivityEntry(DateTimeOffset.UtcNow, stage, kind, message, detail));
            while (Entries.Count > MaximumTimelineEntries)
            {
                Entries.RemoveAt(0);
            }
        }

        private void Dispatch(Action action)
        {
            if (_dispatcher.CheckAccess())
            {
                action();
                return;
            }

            _dispatcher.BeginInvoke(action, DispatcherPriority.Background);
        }

        private void ThrowIfDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(OperationActivityController));
            }
        }
    }
}
