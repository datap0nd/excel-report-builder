# Public Repository Safety

This repository must be independently understandable without real organizational
material.

## Allowed

- Generated synthetic data with labels such as Region North, Channel Retail,
  Product Alpha, Actual, Plan, and Prior Year.
- Schema fragments and expected numeric results created specifically for tests.
- Generic diagrams and hand-authored vector icons.
- Release workbooks generated in CI from source code and attached as artifacts.

## Prohibited

- Real or lightly anonymized workbooks, exports, formulas, screenshots, and
  report layouts.
- Company, project, market, colleague, customer, server, network, or local-user
  identifiers.
- Credentials, API keys, endpoints, model transcripts, prompts, and logs.
- Absolute user paths and internal file names.
- Source or history copied from private repositories.

When uncertain, create a new synthetic example rather than redacting a real
one. Redaction is easy to reverse or miss; generation has no original secret.
