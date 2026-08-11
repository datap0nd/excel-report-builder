namespace ExcelReportBuilder.AddIn.Activity
{
    public enum ActivityStage
    {
        Ready = 0,
        Inspecting = 1,
        Normalizing = 2,
        Planning = 3,
        BuildingPivots = 4,
        Rendering = 5,
        Calculating = 6,
        Checking = 7,
        Repairing = 8,
        Complete = 9
    }

    public enum ActivityKind
    {
        Progress,
        Heartbeat,
        Control,
        Check,
        Result,
        Error
    }

    public enum OperationState
    {
        Idle,
        Running,
        Paused,
        Cancelled,
        Completed,
        Failed
    }

    internal static class ActivityLabels
    {
        public static string Stage(ActivityStage stage)
        {
            switch (stage)
            {
                case ActivityStage.Inspecting:
                    return "Inspecting";
                case ActivityStage.Normalizing:
                    return "Normalizing";
                case ActivityStage.Planning:
                    return "Planning";
                case ActivityStage.BuildingPivots:
                    return "Building pivots";
                case ActivityStage.Rendering:
                    return "Rendering";
                case ActivityStage.Calculating:
                    return "Calculating";
                case ActivityStage.Checking:
                    return "Checking";
                case ActivityStage.Repairing:
                    return "Repairing";
                case ActivityStage.Complete:
                    return "Complete";
                default:
                    return "Ready";
            }
        }

        public static string Kind(ActivityKind kind)
        {
            switch (kind)
            {
                case ActivityKind.Heartbeat:
                    return "Heartbeat";
                case ActivityKind.Control:
                    return "Control";
                case ActivityKind.Check:
                    return "Check";
                case ActivityKind.Result:
                    return "Result";
                case ActivityKind.Error:
                    return "Error";
                default:
                    return "Progress";
            }
        }

        public static string State(OperationState state)
        {
            switch (state)
            {
                case OperationState.Running:
                    return "Running";
                case OperationState.Paused:
                    return "Paused";
                case OperationState.Cancelled:
                    return "Cancelled";
                case OperationState.Completed:
                    return "Complete";
                case OperationState.Failed:
                    return "Needs attention";
                default:
                    return "Ready";
            }
        }
    }
}
