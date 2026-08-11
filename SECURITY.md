# Security Policy

## Supported versions

Security fixes are applied to the latest tagged release and `main` during the
prototype phase.

## Reporting a vulnerability

Use GitHub's private vulnerability reporting for this repository. Do not open
a public issue containing workbook data, credentials, endpoints, or exploit
details.

## Security boundaries

- The model has no Excel, COM, VBA, formula, filesystem, shell, save, publish,
  delete, email, or external-action capability.
- Model tools accept strict versioned arguments and operate only on the active
  job's managed draft scope.
- Workbook cells, headers, comments, formulas, and model output cannot grant
  new capabilities.
- The add-in rejects unknown tools, schema versions, formula strings, code, and
  unmanaged object identities.
- Endpoint credentials are encrypted for the current Windows user and excluded
  from diagnostics.
- Every agent job launches its own worker on a cryptographically random pipe.
  A one-time HMAC proof is verified before the add-in sends endpoint credentials,
  prompts, column names, or workbook samples. The worker accepts one connection
  and the legitimate pipe is restricted to the current Windows user.
- Session API keys and remote-data consent are scoped to the exact endpoint.
  Changing its origin or case-sensitive path clears both and invalidates any
  in-flight endpoint result.
- HTTP is accepted automatically only for loopback endpoints. Non-loopback
  plain HTTP requires a persisted explicit warning and opt-in.
- Query generation is limited to the selected in-workbook source and cannot
  introduce files, URLs, databases, native queries, or credentials.
- Builds occur on managed drafts. Publishing requires a user action, freezes
  generated formulas to values, and commits all managed outputs as one
  compensating transaction. The add-in never saves the workbook automatically.

## Diagnostic policy

Logs may contain timestamps, stages, error categories, counts, object IDs,
duration, HRESULT values, and bounded provider diagnostic codes. They must not
contain cell values, prompts, model output, API keys, endpoints, workbook paths,
sheet contents, or generated formulas.

## Known limits

This project cannot protect against a compromised Windows account, maliciously
modified binaries, vulnerabilities in Excel or the configured model server, or
an administrator replacing installed files. Unsigned prototype installers may
trigger Windows trust warnings until release signing is configured.
