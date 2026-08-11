using System;
using System.Collections.Generic;
using System.Globalization;

namespace ExcelReportBuilder.Excel.PivotPlus
{
    public sealed class PivotMutationStep
    {
        public PivotMutationStep(string name, Action apply, Action rollback)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentException("A mutation step name is required.", nameof(name));
            }

            Name = name;
            Apply = apply ?? throw new ArgumentNullException(nameof(apply));
            Rollback = rollback ?? throw new ArgumentNullException(nameof(rollback));
        }

        public string Name { get; }

        internal Action Apply { get; }

        internal Action Rollback { get; }
    }

    public sealed class PivotMutationException : Exception
    {
        internal PivotMutationException(
            string message,
            string failedStep,
            bool rollbackCompleted,
            Exception innerException)
            : base(message, innerException)
        {
            FailedStep = failedStep;
            RollbackCompleted = rollbackCompleted;
        }

        public string FailedStep { get; }

        public bool RollbackCompleted { get; }
    }

    /// <summary>
    /// Runs one bounded PivotTable mutation with update batching, a
    /// reentrancy guard, verification, and reverse-order rollback.
    /// </summary>
    public sealed class PivotMutationCoordinator
    {
        private readonly object synchronization = new object();
        private bool mutationActive;

        public void Execute(
            object pivotTable,
            IReadOnlyList<PivotMutationStep> steps,
            Action refresh,
            Action verify)
        {
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            if (steps == null) throw new ArgumentNullException(nameof(steps));
            if (refresh == null) throw new ArgumentNullException(nameof(refresh));
            if (verify == null) throw new ArgumentNullException(nameof(verify));

            Enter();
            dynamic pivot = pivotTable;
            bool originalManualUpdate;
            try
            {
                originalManualUpdate = Convert.ToBoolean(
                    pivot.ManualUpdate,
                    CultureInfo.InvariantCulture);
                pivot.ManualUpdate = true;
            }
            catch (Exception exception)
            {
                Exit();
                throw new InvalidOperationException(
                    "Excel did not expose a writable ManualUpdate state for the selected PivotTable.",
                    exception);
            }

            var applied = new List<PivotMutationStep>();
            string failedStep = "prepare";
            try
            {
                foreach (PivotMutationStep step in steps)
                {
                    if (step == null)
                    {
                        throw new ArgumentException("A PivotTable mutation step cannot be null.", nameof(steps));
                    }

                    failedStep = step.Name;
                    applied.Add(step);
                    step.Apply();
                }

                failedStep = "refresh";
                pivot.ManualUpdate = originalManualUpdate;
                refresh();

                failedStep = "verify";
                verify();
            }
            catch (Exception mutationFailure)
            {
                bool rollbackCompleted = RollBack(
                    pivot,
                    originalManualUpdate,
                    applied,
                    refresh,
                    out Exception? rollbackFailure);
                Exception cause = rollbackFailure == null
                    ? mutationFailure
                    : new AggregateException(mutationFailure, rollbackFailure);
                throw new PivotMutationException(
                    rollbackCompleted
                        ? "The PivotTable+ change failed and the prior PivotTable state was restored."
                        : "The PivotTable+ change failed and rollback did not complete.",
                    failedStep,
                    rollbackCompleted,
                    cause);
            }
            finally
            {
                try
                {
                    pivot.ManualUpdate = originalManualUpdate;
                }
                finally
                {
                    Exit();
                }
            }
        }

        private static bool RollBack(
            dynamic pivot,
            bool originalManualUpdate,
            IReadOnlyList<PivotMutationStep> applied,
            Action refresh,
            out Exception? failure)
        {
            var failures = new List<Exception>();
            try
            {
                pivot.ManualUpdate = true;
            }
            catch (Exception exception)
            {
                failures.Add(exception);
            }

            for (var index = applied.Count - 1; index >= 0; index--)
            {
                try
                {
                    applied[index].Rollback();
                }
                catch (Exception exception)
                {
                    failures.Add(new InvalidOperationException(
                        "Rollback failed for step '" + applied[index].Name + "'.",
                        exception));
                }
            }

            try
            {
                pivot.ManualUpdate = originalManualUpdate;
                refresh();
            }
            catch (Exception exception)
            {
                failures.Add(exception);
            }

            failure = failures.Count == 0 ? null : new AggregateException(failures);
            return failures.Count == 0;
        }

        private void Enter()
        {
            lock (synchronization)
            {
                if (mutationActive)
                {
                    throw new InvalidOperationException(
                        "A PivotTable+ mutation is already active. Recursive PivotTable updates are not allowed.");
                }

                mutationActive = true;
            }
        }

        private void Exit()
        {
            lock (synchronization)
            {
                mutationActive = false;
            }
        }
    }
}
