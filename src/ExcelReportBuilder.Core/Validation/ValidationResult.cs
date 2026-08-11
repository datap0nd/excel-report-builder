using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelReportBuilder.Core.Validation
{
    public enum ValidationSeverity
    {
        Warning,
        Error
    }

    public sealed class ValidationIssue
    {
        public string Code { get; set; } = string.Empty;

        public string Path { get; set; } = string.Empty;

        public string Message { get; set; } = string.Empty;

        public ValidationSeverity Severity { get; set; }
    }

    public sealed class ValidationResult
    {
        private readonly List<ValidationIssue> _issues = new List<ValidationIssue>();

        public IReadOnlyList<ValidationIssue> Issues => _issues;

        public bool IsValid => !_issues.Any(issue => issue.Severity == ValidationSeverity.Error);

        public void AddError(string code, string path, string message)
        {
            _issues.Add(new ValidationIssue
            {
                Code = code,
                Path = path,
                Message = message,
                Severity = ValidationSeverity.Error
            });
        }

        public void AddWarning(string code, string path, string message)
        {
            _issues.Add(new ValidationIssue
            {
                Code = code,
                Path = path,
                Message = message,
                Severity = ValidationSeverity.Warning
            });
        }
    }

    public sealed class InvalidReportSpecException : Exception
    {
        public InvalidReportSpecException(ValidationResult validation)
            : base("The report specification is invalid.")
        {
            Validation = validation ?? throw new ArgumentNullException(nameof(validation));
        }

        public ValidationResult Validation { get; }
    }
}
