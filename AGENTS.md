# Repository Instructions

## Public Safety

- This is a public repository. Use only generated synthetic data and generic
  names.
- Never commit real workbooks, exports, screenshots, prompts, transcripts,
  endpoints, credentials, user paths, employer details, internal product names,
  colleague names, network locations, or copied report formulas.
- Do not copy source or history from private repositories. Reimplement generic
  behavior from public requirements and synthetic tests.
- Refer to the output style as a "dense management report".

## Delivery

- Work from `plan.md`, one numbered task at a time.
- Keep `docs/agent_handoff.md` current before stopping.
- Run relevant tests and public-safety checks before delivery.
- Completed scoped work must be committed, pushed, and present on `main`.
- Preserve unrelated work and generated local artifacts.

## Product Boundaries

- The model may propose only validated PivotTable+ operations.
- No arbitrary DAX, MDX, worksheet formulas, VBA, COM calls, shell commands,
  filesystem tools, workbook saves, publishing, or deletion may be exposed to
  the model.
- The add-in may change the explicitly selected PivotTable only after preview
  and confirmation. It owns only the measures, named sets, queries,
  connections, and metadata that it creates; it never claims ownership of a
  user's PivotTable or source data.
- Supported features must remain inside one real Excel PivotTable. Do not
  silently substitute a formula-backed companion report.
- Never truncate data silently or infer a missing reporting year.
- Source period grain and member coverage must be validated before generating
  a period calculation or asymmetric set.
- Continuous, specific progress is required for every long-running operation.
