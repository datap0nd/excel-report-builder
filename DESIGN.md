---
name: "Excel Report Builder"
description: "A precise, spreadsheet-native operations workbench for building and checking managed report drafts."
colors:
  canvas: "#F7F9F7"
  surface: "#FFFFFF"
  subtle-surface: "#EFF3F0"
  rule: "#D3DAD5"
  rule-strong: "#AEBBB3"
  text: "#1F2924"
  text-muted: "#52615A"
  accent: "#1F6B45"
  accent-hover: "#185537"
  accent-soft: "#E4F0E9"
  focus: "#005FCC"
  danger: "#982F32"
  danger-soft: "#F7E9E9"
  warning: "#765100"
  info: "#245A77"
  check: "#5A4B8A"
  control-hover: "#F2F5F3"
  control-hover-border: "#7E9186"
  control-pressed: "#E4EAE6"
  progress-track: "#DDE4DF"
  splitter: "#C6D0C9"
typography:
  surface-title:
    fontFamily: "Segoe UI"
    fontSize: "20px"
    fontWeight: 600
  product-title:
    fontFamily: "Segoe UI"
    fontSize: "18px"
    fontWeight: 600
  section-title:
    fontFamily: "Segoe UI"
    fontSize: "14px"
    fontWeight: 600
  body:
    fontFamily: "Segoe UI"
    fontSize: "13px"
    fontWeight: 400
  detail:
    fontFamily: "Segoe UI"
    fontSize: "11px"
    fontWeight: 400
  timing:
    fontFamily: "Consolas"
    fontSize: "11px"
    fontWeight: 400
rounded:
  square: "0px"
  control: "2px"
spacing:
  micro: "3px"
  compact: "4px"
  control-gap: "6px"
  row: "8px"
  field: "10px"
  section-gap: "12px"
  surface-inset: "16px"
  section-break: "18px"
components:
  button-standard:
    backgroundColor: "{colors.surface}"
    textColor: "{colors.text}"
    typography: "{typography.body}"
    rounded: "{rounded.control}"
    padding: "5px 12px"
    height: "32px"
  button-primary:
    backgroundColor: "{colors.accent}"
    textColor: "{colors.surface}"
    typography: "{typography.body}"
    rounded: "{rounded.control}"
    padding: "5px 12px"
    height: "32px"
  input:
    backgroundColor: "{colors.surface}"
    textColor: "{colors.text}"
    typography: "{typography.body}"
    rounded: "{rounded.square}"
    padding: "5px 8px"
    height: "32px"
  navigation-active:
    backgroundColor: "{colors.accent-soft}"
    textColor: "{colors.accent}"
    typography: "{typography.body}"
    rounded: "{rounded.square}"
    padding: "4px 6px"
    height: "40px"
---

# Design System: Excel Report Builder

## Overview

**Creative North Star: "The Visible Operations Ledger"**

Excel Report Builder is an operations-strip workbench, not a dashboard. It borrows the useful visual qualities of a working spreadsheet: white working surfaces, cool rules, compact rows, explicit labels, and a persistent record of what changed. The interface should feel calm under long-running work because state, ownership, and the next safe action remain visible.

The visual hierarchy follows the report-building sequence: choose Data, configure Build, request bounded changes in Chat, then review Checks before publishing. One active work surface occupies the center. The current operation stays above it and the activity timeline stays below it. Accent color is reserved for selection, progress, successful results, and primary actions so it remains meaningful.

**Key Characteristics:**

- Spreadsheet-native density without visual clutter.
- Flat, ruled surfaces rather than nested cards.
- Ink-green actions against neutral whites and cool gray-greens.
- Persistent operational feedback with tabular timing.
- Square status fields and nearly square controls.
- Explicit safety language around drafts, checks, cancellation, and publishing.

### Token reference

The YAML frontmatter is the normative token source. This table records the intended use of each token family without redefining its values.

| Token family | Role | Application rule |
| --- | --- | --- |
| `colors.canvas`, `colors.surface`, `colors.subtle-surface` | Tonal layers | Separate the task pane, working surface, and status bands without shadows. |
| `colors.rule`, `colors.rule-strong` | Structure | Divide rows, sections, panes, and persistent operational regions. |
| `colors.text`, `colors.text-muted` | Reading hierarchy | Use full ink for actions and conclusions; use muted ink for context, labels, and details. |
| `colors.accent`, `colors.accent-hover`, `colors.accent-soft` | Primary intent | Mark the active destination, primary actions, progress, selected timeline rows, and successful results. |
| `colors.focus` | Keyboard focus | Draw a visible focus boundary independent of selection or status color. |
| `colors.danger`, `colors.warning`, `colors.info`, `colors.check` | Semantic events | Pair every semantic color with a readable state, kind, or message label. |
| `typography.*` | Compact hierarchy | Keep the scale narrow and use weight, spacing, and alignment before adding new sizes. |
| `spacing.*` | Working rhythm | Use compact control gaps within a task and larger section breaks between decisions. |
| `rounded.*` | Form language | Keep structural regions square and limit softening to the slight control radius. |

## Colors

The palette is a cool spreadsheet neutral system with one restrained ink-green accent and explicit semantic colors for operational feedback.

### Primary

- **Ink Green** (`colors.accent`): primary actions, active navigation, progress, selected activity rows, and successful result labels.
- **Deep Ink Green** (`colors.accent-hover`): hover feedback for primary actions.
- **Washed Green** (`colors.accent-soft`): active navigation and selected timeline backgrounds where a quiet, full-row state is needed.

### Neutral

- **Sheet Canvas** (`colors.canvas`): the outer task-pane background and alternating data rows.
- **Working White** (`colors.surface`): the active work surface, inputs, lists, and timeline.
- **Cool Status Wash** (`colors.subtle-surface`): operation strips, data-mode labels, saved setup summaries, connection states, and publish gates.
- **Cool Rule** (`colors.rule`): routine row and section separation.
- **Strong Cool Rule** (`colors.rule-strong`): boundaries between persistent regions and standard control borders.
- **Deep Worksheet Ink** (`colors.text`): primary copy and control text.
- **Muted Worksheet Ink** (`colors.text-muted`): field labels, supporting copy, timestamps, locations, and secondary state detail.

### Semantic

- **Keyboard Blue** (`colors.focus`): the only focus-ring color. It must remain visually distinct from active green states.
- **Error Red** (`colors.danger`): destructive action labels, unsafe transport warnings, failed operations, and error events.
- **Control Amber** (`colors.warning`): paused work and control events.
- **Information Blue** (`colors.info`): progress and heartbeat events.
- **Check Violet** (`colors.check`): independent check events in the timeline.

**The Paired Signal Rule.** Color never carries operational meaning alone. Every dot, status brush, or gate color is paired with a text label and, when needed, explanatory detail.

**The One Accent Rule.** Green means current intent, forward progress, or a successful managed result. It is not decorative fill.

## Typography

**Interface Font:** Segoe UI

**Timing Font:** Consolas

Segoe UI keeps the add-in consistent with the Windows and Excel environment. Consolas is used narrowly for elapsed time and timestamps so changing numeric widths remain stable and easy to scan.

### Hierarchy

| Role | Token | Use |
| --- | --- | --- |
| Surface title | `typography.surface-title` | The Data, Build, Chat, and Checks headings. |
| Product title | `typography.product-title` | The fixed task-pane identity at the top. |
| Section title | `typography.section-title` | Section headings and the Activity header. |
| Body | `typography.body` | Controls, values, messages, and ordinary explanatory copy. |
| Detail | `typography.detail` | Requirements, warnings, samples, state details, and compact metadata. |
| Timing | `typography.timing` | Timeline timestamps and compact elapsed-time readouts. |

Semibold is the primary emphasis mechanism. It identifies titles, state labels, result kinds, and important totals without introducing display typography. Sentence case is the default. Access-key underscores remain available in control labels.

**The Narrow Scale Rule.** Do not add decorative type sizes. Use the existing title, section, body, and detail roles, then solve hierarchy with spacing, rules, and weight.

**The Stable Time Rule.** Durations and timestamps use tabular-looking Consolas numerals and remain right-aligned or column-aligned where practical.

## Layout

The task pane has a minimum working size of 320 by 480 device-independent pixels. Its vertical structure is fixed in meaning and flexible in height:

1. Product identity and data-mode indicator.
2. Four equal-width navigation destinations.
3. Persistent current-operation strip.
4. One active, vertically scrolling work surface.
5. A five-pixel resizable divider.
6. Persistent activity timeline.

The work surface and timeline share remaining height at a 3:2 ratio before the user adjusts the splitter. The work surface preserves at least 120 pixels and the timeline preserves at least 96 pixels. Main work surfaces use a 16-pixel horizontal inset, compact 12-pixel top rhythm, and no horizontal scrolling.

### Narrow-pane behavior

- Keep the four destinations in one equal-width row. Labels stay short and retain keyboard access.
- Stack fields and decision groups vertically. Never introduce a desktop-style side rail inside the docked pane.
- Use wrapping action rows so secondary buttons flow to the next line instead of creating horizontal scrolling.
- Allow the current operation name to trim before pause or cancel controls disappear.
- Keep tables inside bounded vertical regions. Give descriptive columns the flexible width and keep small status columns fixed.
- Wrap explanatory text, warnings, state details, and source summaries.
- Preserve both the operation strip and activity timeline at narrow widths. They are functional safety surfaces, not optional chrome.

**The One Work Surface Rule.** Data, Build, Chat, and Checks occupy the same central region, and only the selected destination is visible.

**The No Horizontal Scroll Rule.** At the minimum pane width, content reflows, wraps, or trims. The core workflow never depends on panning sideways.

**The Persistent Feedback Rule.** Resizing may change the relative height of the work surface and timeline, but it must not remove the current-operation strip or collapse the timeline below its minimum.

## Elevation & Depth

The interface is flat. It uses tonal layering and rules, not shadows, to show containment and depth. Working white sits on the sheet canvas; subtle washes identify operational and guarded states; stronger rules separate the current operation, work surface, and timeline.

The activity timeline uses a one-pixel vertical spine and row dividers to convey sequence. Tables use horizontal grid lines and alternating neutral rows, with vertical rules limited to column headers. This preserves spreadsheet structure without creating a boxed grid around every cell.

**The Flat Workbench Rule.** Do not add drop shadows, floating cards, glass effects, or layered panels. Use the established surface tones and rule weights.

**The Boundary Has Meaning Rule.** A strong rule separates persistent regions. A routine rule separates content within a region. Do not interchange them as decoration.

## Shapes

The form language is square and workmanlike. Structural surfaces, status fields, grids, lists, inputs, and navigation remain square. Buttons use only a two-pixel corner radius, enough to distinguish a control from a ruled table cell without making the interface feel soft.

Status markers are small circles on otherwise rectilinear surfaces. Their circular form makes live state easy to locate, but the adjacent label supplies the meaning. Progress is a thin three-pixel strip rather than a large decorative meter.

**The Almost-Square Rule.** Rounded containers, pills, and capsule buttons do not belong in this workbench. Only interactive buttons receive the slight control radius.

## Components

### Product header

The product header is a compact identity row with the product name, the "Managed draft workbench" qualifier, and a bordered data-mode field. The mode field is informational, not a promotional badge. It uses the subtle surface, a routine rule, detail typography, and wrapping text.

### Navigation

Four flat, equal-width destinations form the main navigation. The selected destination combines ink-green text, a green bottom rule, a washed-green background, and semibold weight. Unselected destinations remain transparent with muted text. Navigation is reachable through both tab order and Ctrl+1 through Ctrl+4.

### Buttons

- **Standard:** working white, deep text, strong rule, 32-pixel minimum height, and slight two-pixel corners.
- **Primary:** ink-green fill and border, white semibold text, and deep-green hover feedback.
- **Danger:** standard white surface with red text and a restrained red border. It signals a destructive or interrupting action without becoming a solid alarm block.
- **Operational compact:** pause and cancel controls may use a 28-pixel minimum height inside the persistent status strip.
- **States:** standard controls shift to the neutral hover and pressed tones. Disabled controls use 48 percent opacity and return the cursor to the standard arrow.

Buttons use access keys, visible focus treatment, automation names, and concise tooltips for keyboard shortcuts where relevant.

### Inputs and fields

Text, password, and selection fields use white backgrounds, strong cool borders, compact internal padding, and a 32-pixel minimum height. Keyboard focus changes the field border to Keyboard Blue and adds the shared two-pixel visible focus boundary. Field labels use muted body text and sit four pixels above their controls.

Multiline requests, mapping explanations, transport warnings, and validation details wrap. Password values are never rendered back into the interface. Error and disabled states must remain understandable through text, not border color alone.

### Tables and lists

Tables are read-oriented, flat, and dense. Column headers use the subtle surface, muted semibold labels, and a bottom rule. Rows use horizontal rules, alternating working white and sheet canvas, full-row selection, and no row header. Fixed-width columns hold compact types or statuses; the descriptive column absorbs remaining width.

The chat transcript and activity timeline use ruled list rows instead of message bubbles or cards. Speaker, event kind, stage, message, and detail remain textually explicit.

### Current-operation strip

The operation strip is always visible above the active work surface. It contains:

- Text state and a paired status marker.
- The current operation name, trimmed when horizontal space is constrained.
- Pause or resume and cancel controls when safe.
- Stage position and stage label.
- Monospaced elapsed time.
- A thin stage-progress strip.

The implemented sequence is Inspecting, Normalizing, Planning, Building pivots, Rendering, Calculating, Checking, Repairing, and Complete. Ready is the pre-operation state. Host operations may make the current operation text more specific, but they must not hide the visible stage.

### Activity timeline

The timeline is the signature component. Each row aligns a timestamp, a vertical sequence spine with a semantic marker, and a kind-stage-message-detail stack. Kinds include Progress, Heartbeat, Control, Check, Result, and Error. Selected rows use the washed-green surface. Rows are virtualized and recycled so the visible audit remains responsive, and the controller retains the most recent 200 entries.

The timeline is an assistive live region with polite announcements. Each row exposes a concise accessible summary, not merely its visual marker.

### Continuous-feedback contract

Every manual build and agent job uses the same activity system.

- Start by naming the operation and first stage before work begins.
- Report meaningful stage changes, decisions, affected managed objects, validation results, and failures as typed entries.
- If no new event arrives during active work, emit a heartbeat at least every 15 seconds. A heartbeat explains that work continues and that pause and cancel remain available.
- Show elapsed time throughout the operation.
- Announce an unavoidable blocking workbook operation before it begins.
- Keep pause and cancel enabled only while the host can honor them safely. Log pause, resume, and cancellation as Control entries.
- Preserve managed draft work as unpublished after cancellation.
- End with a Result or Error entry that states what changed, what was checked, and what remains unresolved.
- Never replace the operation strip and timeline with an unexplained spinner.

### Publishing gate

Checks presents a text-labeled publish gate, independent check results, and separate Run checks and Publish managed draft actions. A failed or incomplete gate is explained in text. Publishing remains an explicit primary action after checks; the interface never implies that building a draft saves or publishes the workbook.

### Interaction and accessibility rules

- Preserve the visible two-pixel Keyboard Blue focus style on every interactive control, table cell, list item, and splitter.
- Maintain logical tab order through the shell and local navigation group.
- Keep Ctrl+1 through Ctrl+4 for surface navigation, Ctrl+P for pause or resume, Ctrl+Shift+C for cancel, Ctrl+Shift+B for build, and Ctrl+Enter for chat submission.
- Retain automation names on work surfaces, controls, tables, the splitter, and operational regions. Use help text when a field has a safety or interpretation constraint.
- Pair colors, markers, progress, and gate states with labels. Do not communicate success, warning, or failure by hue alone.
- Keep action labels concrete and workbook-oriented. Avoid database, scripting, or model-internals jargon in the default workflow.
- Do not animate layout changes. State transitions are immediate so the task pane remains stable during workbook work.

### Implementation notes

- The implemented source of visual truth is `ReportBuilderView.xaml`. Reusable brushes and control styles live in the user-control resources and should be promoted to a shared resource dictionary only when another production surface needs them.
- The view is WPF hosted in a Windows task pane. Measurements are device-independent pixels, and native keyboard, automation, focus, and high-DPI behavior take precedence over web conventions.
- The view model supplies semantic operation-state brushes as well as text labels. Any future token consolidation must preserve the paired-label behavior and sufficient text contrast.
- The activity controller owns the 15-second heartbeat, one-second elapsed-time refresh, safe pause and cancel state, and bounded timeline history. Do not reproduce those behaviors as view-only timers.
- Data, Build, Chat, and Checks use one shared specification and one shared activity feed. New commands must enter that same state and feedback path.

## Do's and Don'ts

### Do

- **Do** keep one primary action visually dominant within the current decision group.
- **Do** use rules, row alignment, and compact labels to make complex report operations scannable.
- **Do** keep the current operation and activity history visible during long work.
- **Do** state exactly what a build, check, cancel, or publish action will affect.
- **Do** preserve keyboard access, visible focus, automation names, text wrapping, and minimum pane dimensions.
- **Do** use synthetic, generic examples in screenshots, fixtures, and public documentation.

### Don't

- **Don't** turn the task pane into a card dashboard, chat-first shell, or collection of floating status tiles.
- **Don't** hide long work behind a spinner, indeterminate message, or silent disabled interface.
- **Don't** use green, red, amber, blue, or violet without a nearby text label that explains the state.
- **Don't** add pills, large corner radii, shadows, gradients, decorative illustrations, or marketing copy.
- **Don't** collapse the timeline to make more room for forms. Let the user resize it within the established minimums.
- **Don't** auto-publish, auto-save, or visually imply that a managed draft is final.
- **Don't** expose raw formulas, scripts, data-engine terminology, or arbitrary agent capabilities in the default experience.
