# Contributing

Contributions are welcome when they preserve the deterministic and public-safe
product boundary.

## Before opening a pull request

1. Use only generated synthetic fixtures and generic field names.
2. Do not attach workbooks, exports, screenshots, logs, prompts, or transcripts.
3. Keep model capabilities inside the existing typed tool boundary.
4. Add tests for every calculation, transformation, or permission change.
5. Run the build, tests, and `scripts/Test-PublicSafety.ps1`.

Pull requests that include confidential data or broaden the model into general
Excel or operating-system automation will be closed.
