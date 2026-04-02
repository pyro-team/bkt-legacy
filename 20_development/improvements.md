# Improvements Roadmap For `20_development/src`

## Summary

This document reviews the exported VBA modules in `20_development/src` for redundancy, performance, maintainability, and reliability. The codebase is small enough to improve substantially without changing its functional scope, but two modules dominate the risk surface:

- `ToolboxRibbon.bas` is the main complexity hotspot. It combines ribbon lifecycle, UI dispatch, property access, enablement logic, unit conversion, and shape mutation in one large callback module.
- `ToolboxActions.bas` mixes many unrelated behaviors such as shape formatting, text manipulation, slide/export operations, and presentation-wide cleanup.

The overall recommendation is a phased safe-first refactor:

1. remove dead code and duplication that do not change behavior,
2. reduce repeated COM access and repeated selection/ribbon work,
3. centralize control/property handling behind shared helpers,
4. split modules by responsibility while keeping ribbon callback names stable.

This roadmap treats any macros referenced by `20_development/ribbonUI.xml` as public interfaces. Early phases should preserve those names and signatures, even if the implementation moves behind adapters.

## Executive Summary

Below are the highest-value findings, ordered by correctness and operational risk first, then performance, then duplication and structure.

| # | Finding | Severity | Payoff | Effort |
|---|---|---|---|---|
| 1 | Ribbon/property logic is duplicated across large `Select Case` blocks in `ToolboxRibbon.bas` (`GetEditBoxValue`, `GetShapeSettingSingle`, `SetShapeSettingSingle`, `isEnabled`, increment/reset paths). | High | High | Medium |
| 2 | `ActiveWindow.selection` and nested `ShapeRange` access are repeated heavily across callbacks and actions, increasing COM overhead and making logic inconsistent. | High | High | Medium |
| 3 | Broad `On Error Resume Next` and empty handlers hide real failures in ribbon invalidation, shape operations, and helper functions. | High | Medium | Low |
| 4 | The previous `MinMax copy.bas` duplicate issue is resolved; keep `MinMax.bas` as the only implementation and avoid reintroducing parallel copies. | Resolved | Medium | Low |
| 5 | `ToolboxActions.bas` combines unrelated domains and should be split into smaller modules to reduce regression risk and improve testability. | Medium | High | Medium |
| 6 | Ribbon invalidation is eager and global (`myRibbon.Invalidate` in many callbacks plus invalidation on multiple application events), which can make the UI feel slower than necessary. | Medium | Medium | Medium |
| 7 | Selection sorting helpers rebuild arrays from live COM objects every time and are called repeatedly for the same selection. | Medium | Medium | Low |
| 8 | Several actions assume valid shape/text state without a shared precondition layer, leading to inconsistent guard logic and error handling. | Medium | Medium | Medium |
| 9 | `ribbonUI.xml` still references old controls such as `ebParIndent` and `ebRectCorner2` that are partly commented out in VBA, indicating interface drift. | Medium | Medium | Low |
| 10 | User prompts (`MsgBox`, `InputBox`) are mixed into reusable action logic, which makes future automation and testing harder. | Low | Medium | Medium |
| 11 | Presentation-wide scans are used in several operations where tags or cached lookups could cut work on large decks. | Low | Medium | Medium |
| 12 | Mac/Windows divergence is handled inline in business modules, which is workable now but should be isolated further before larger refactors. | Low | Medium | Medium |

## Quick Wins

These are low-risk changes that should be implemented first because they reduce noise and improve safety without changing public behavior.

## Current Status

### Already completed or largely in place

- `SelectionContext.bas` now exists and centralizes common selection access (`GetActiveShapeRange`, `GetActiveSlideRange`, `SelectionContainsTextFrame`, sorted shape helpers).
- `ToolboxRibbon.bas` already uses the shared selection helpers in key callback paths.
- `ToolboxActions.bas` already uses `GetActiveShapeRange` in many actions.
- `MinMax copy.bas` is no longer present, so the duplicate-file cleanup is effectively complete.
- Legacy selection-sorting wrappers have been removed from `ToolboxSelections.bas`; active callers now use the `SelectionContext`-based sort helpers.
- The broad `On Error Resume Next` usages targeted in `Helpers.bas`, `TriggerInvalidate.cls`, `ToolboxActions.bas`, and `ToolboxSelections.bas` have been replaced with guards or narrower helper-based handling.
- A small shared helper layer now exists in `ToolboxActions.bas` for repeated shape-range write/fallback operations.

### Still open and worth doing next

- remove or quarantine any remaining obsolete debug/test code in `ToolboxRibbon.bas`,
- keep trimming interface drift and stale comments where they no longer describe supported behavior,
- continue with small shared-helper extraction only when it reduces duplication without changing public callback contracts,
- defer larger module splits until new pain points appear during maintenance.

### 1. Remove dead and duplicate code

- Keep `MinMax.bas` as the single implementation. `MinMax copy.bas` no longer exists.
- Remove or quarantine unused test/debug procedures such as `test()` in `ToolboxRibbon.bas`.
- Remove stale commented-out branches that are no longer part of the supported feature set, especially:
  - old `ebParIndent` code paths,
  - old `ebRectCorner2` code paths,
  - commented test helpers and abandoned ribbon experiments.
- Review duplicate ribbon controls and IDs in `ribbonUI.xml`, especially repeated margin controls and controls that still point at partially removed handlers.

### 2. Replace broad silent error swallowing with narrow guards

- `On Error Resume Next` was removed from the utility conversions in `Helpers.bas`.
- In ribbon callbacks, prefer precondition checks such as:
  - `If myRibbon Is Nothing Then Exit Sub`
  - `If ActiveWindow Is Nothing Then Exit Sub`
  - `If sel Is Nothing Or sel.Type <> ppSelectionShapes Then Exit Sub`
- Keep silent failure only where Office event timing makes it unavoidable, and document those cases.
- Introduce one shared error-reporting helper for user-facing failures instead of hand-built `MsgBox "Fehler ..."` blocks across modules.

### 3. Reduce obviously repeated work

- Cache `ActiveWindow.Selection` and `ShapeRange` once per callback/action instead of repeatedly dereferencing them.
- Cache `ShapeRange.Count` and first shape once where the same values are used several times.
- In sorting helpers, avoid multiple `ActiveWindow.selection.ShapeRange.Count` calls inside the same function.
- In presentation-wide loops, cache `activePresentation` and slide counts locally.

### 4. Make ribbon callback compatibility explicit

- Build a callback inventory from `ribbonUI.xml` and annotate which VBA procedures are externally referenced.
- Mark those procedures as compatibility entrypoints in comments or documentation.
- For any future rename, keep the original callback as a thin wrapper.

## Medium Refactors

These changes bring most of the maintainability and performance gains while preserving current behavior and ribbon contracts.

### 1. Introduce a `SelectionContext` abstraction

This is already implemented enough to serve as the shared selection access layer for the current codebase. Continue using it as the single place for shared selection access and extend it only when a concrete new caller benefits.

Target capabilities:

- current `Selection`
- selection type
- `ShapeRange` if available
- selected shape count
- first shape
- sorted-by-top shape array
- sorted-by-left shape array
- active slide / active presentation when relevant

Recommended behavior:

- Build the context lazily so sorted arrays are only created if needed.
- Return `Nothing` or explicit flags when the current selection is not a shape selection.
- Use the helper in both `ToolboxRibbon.bas` and `ToolboxActions.bas`.
- Keep `ToolboxSelections.bas` focused on non-overlapping selection utilities.

Benefits:

- fewer COM roundtrips,
- less repeated guard logic,
- more consistent behavior between enablement, read, and write paths.

### 2. Replace property-specific `Select Case` sprawl with a property handler map

The main structural improvement is to centralize ribbon property metadata. Today the same control families are mapped repeatedly in:

- `GetEditBoxValue`
- `GetShapeSettingSingle`
- `SetShapeSettingSingle`
- `isEnabled`
- `ebPixelValue_onChange`
- `ChangeValueBy`
- `ResetPixelValue`

Target design:

- define a single metadata table or grouped helper functions for each property family:
  - margins,
  - geometry,
  - paragraph indents/spacings,
  - transparency,
  - line weight,
  - rounded corners,
  - spacing between shapes.
- expose shared operations:
  - `CanReadProperty`
  - `ReadProperty`
  - `WriteProperty`
  - `ResetProperty`
  - `StepProperty`
  - `ParseControlFamily`

Implementation guidance:

- VBA does not offer rich dictionaries by default in every environment, so a safe-first approach is grouped helper functions and normalized control IDs, not a complex dynamic registry.
- Normalize `incX`, `decX`, and `resX` to `ebX` early and operate on one property identity.
- Keep special cases such as `HSep`, `VSep`, `RectCorner`, and split/multiply controls as explicit branches, but move the common property set behind shared helpers.

### 3. Split `ToolboxActions.bas` by responsibility

Recommended module split:

- `ShapeOps.bas`
  - geometry/size/position helpers,
  - same height/width,
  - swap/arrange/split/multiply,
  - transparency and line weight setters.
- `TextOps.bas`
  - text move in/out of shapes,
  - split/join text shapes,
  - replace/remove text,
  - margin helpers,
  - language changes.
- `PresentationOps.bas`
  - clean author,
  - clean slide masters,
  - slide numbering,
  - create presentation from slide selection,
  - send slide selection by email.
- `RibbonActions.bas` or keep `ToolboxActions.bas` as a compatibility facade
  - existing `btnAction`-targeted public entrypoints that call the new modules.

This keeps early XML compatibility while allowing the internal code to become much easier to reason about.

### 4. Centralize user-facing prompts and guard messages

- Move repeated validation messages into small helpers.
- Keep domain logic separate from the prompt/interaction layer.
- This is especially useful for:
  - `MoveTextIntoShape`
  - `ReplaceAllText`
  - `SendEmailFromSlideSelection`
  - `CreatePresentationFromSlideSelection`
  - agenda actions in `ToolboxAgenda.cls`

## Deep Refactors

These are worthwhile, but they should come after the safer cleanup and shared-helper phases.

### 1. Redesign ribbon dispatch around explicit control families

Instead of a monolithic callback module, move toward:

- a small `ToolboxRibbon.bas` facade containing the Office callback entrypoints,
- one or more internal modules for:
  - ribbon state evaluation,
  - control/property access,
  - command dispatch,
  - invalidation policy.

This keeps the callback surface stable while making the internals modular.

### 2. Throttle or narrow ribbon invalidation

Current pattern:

- many action callbacks call `myRibbon.Invalidate`,
- `TriggerInvalidate.cls` invalidates on `SlideSelectionChanged`,
- `WindowActivate`,
- `WindowSelectionChange`.

Recommended redesign:

- invalidate only when state-relevant controls changed,
- consider `InvalidateControl` for hot controls if the Office host behaves reliably,
- add a simple re-entrancy or debounce guard so repeated event storms do not trigger redundant refreshes.

This needs care because Office ribbon behavior can be host-specific. Treat it as a later optimization after behavior is covered by manual regression checks.

### 3. Extract platform-specific concerns

Move platform divergence behind explicit helpers:

- `KeyState.bas` remains the platform boundary for modifier keys,
- template loading path logic in `Slides.bas` should be wrapped behind a path helper,
- any future file dialog or clipboard differences should not leak into general action code.

This reduces the risk that Windows-focused cleanups break Mac behavior.

## Performance Notes

### Main hotspots

#### Repeated COM access to `ActiveWindow.selection`

`ToolboxRibbon.bas` and `ToolboxActions.bas` repeatedly call chains like:

- `ActiveWindow.selection.Type`
- `ActiveWindow.selection.ShapeRange.Count`
- `ActiveWindow.selection.ShapeRange(1)`
- `ActiveWindow.View.Slide.Shapes`

Each call crosses the Office COM boundary. The current code does this many times inside the same callback, which is avoidable.

Recommendation:

- resolve selection once,
- hold local references,
- pass those references down to helpers.

#### Repeated sort/rebuild of selected shapes

`ActiveWindowSelectionSortedByTop` and `ActiveWindowSelectionSortedByLeft` rebuild a 2D array and sort it every time they are called. For spacing and join/connector operations, the same sorted order may be needed multiple times during one action.

Recommendation:

- make sorting part of `SelectionContext`,
- build each sorted view once per action.

#### Repeated full-slide or full-presentation scans

Examples:

- shape visibility restore in `ShowShapes`,
- line/fill/type matching in selection helpers,
- slide numbering add/remove checks,
- language assignment and cleanup operations.

Recommendation:

- keep scans where necessary, but avoid chaining several scans for one user action,
- use tags or a known naming convention where appropriate,
- separate “detect existing state” from “apply state” so repeated loops can be minimized.

#### Copy/paste/select churn

Some actions repeatedly select shapes or rely on clipboard-style operations:

- `MoveTextOutOfShapes`
- `MoveTextIntoShape`
- `SplitShapeByParagraphs`
- `JoinShapesWithText`
- `PasteAndReplace`

Recommendation:

- minimize `Select`, `Unselect`, and `DoEvents` usage where possible,
- prefer direct object operations unless Office requires selection-bound behavior.

## Reliability Notes

### Silent failures

Several modules swallow errors without logging or fallback behavior:

- ribbon invalidation,
- helper conversions,
- z-order adjustments,
- text and shape operations,
- selection range identification.

Risk:

- bugs become intermittent and hard to reproduce,
- real failures are misread as UI sluggishness or random Office behavior.

Recommendation:

- use narrow `On Error Resume Next` only around the one statement that genuinely needs it,
- clear or inspect `Err` immediately after,
- prefer `On Error GoTo` with a shared exit pattern elsewhere.

### Inconsistent precondition handling

Some actions validate selection type carefully, others assume `ShapeRange` or `SlideRange` exists, and some let the error handler absorb invalid states.

Recommendation:

- standardize guard helpers such as:
  - `TryGetSelectedShapes`
  - `TryGetSelectedSlides`
  - `HasSingleShapeSelection`
  - `HasExactShapeCount`

### Public interface drift between XML and VBA

The ribbon XML still references controls that have partially removed or commented-out backing logic, such as:

- `ebParIndent`
- `decParIndent`
- `incParIndent`
- `ebRectCorner2`
- `decRectCorner2`
- `incRectCorner2`

Risk:

- dead UI surface remains visible,
- maintenance becomes harder because it is unclear whether a branch is intentionally dormant or accidentally unfinished.

Recommendation:

- reconcile `ribbonUI.xml` and VBA together,
- either remove unsupported controls from XML or fully restore their backing logic.

### Mixed UI and business logic

Many procedures both decide what to do and manage prompts or message boxes. This makes reuse difficult and causes more edge-case branching inside the action code.

Recommendation:

- split “validate and ask” from “apply operation”,
- keep reusable action helpers side-effect-light.

## Suggested Target Architecture

The safest medium-term target is not a framework rewrite. It is a modest internal structure that keeps all existing ribbon callback names available.

### Recommended modules

- `ToolboxRibbon.bas`
  - Office ribbon callbacks only,
  - delegates to internal helpers,
  - preserves callback compatibility with `ribbonUI.xml`.
- `SelectionContext.bas` or `SelectionContext.cls`
  - cached selection state,
  - sorted shape lists,
  - shared precondition accessors.
- `RibbonPropertyAccess.bas`
  - normalized control ID parsing,
  - read/write/reset/step logic for shape properties.
- `ShapeOps.bas`
  - shape geometry, spacing, arranging, duplication, connector behavior.
- `TextOps.bas`
  - text-specific move/split/join/replace/margin/language logic.
- `PresentationOps.bas`
  - deck-wide and slide-wide operations.
- `ErrorHandling.bas`
  - shared user-facing error formatting and guarded execution helpers.

### Compatibility strategy

- keep current callback procedure names as public wrappers,
- move implementations behind internal helpers,
- update `ribbonUI.xml` only when there is a deliberate public contract change.

## Migration Guidance

### Phase 1: Safe cleanup

- Remove duplicate and dead modules.
- Reconcile ribbon XML with the implemented control surface.
- Normalize error handling in the most obvious silent-failure cases.
- Add documentation comments that identify public callback procedures.

### Phase 2: Shared selection and property helpers

- Introduce `SelectionContext`.
- Introduce normalized control-ID parsing.
- Collapse repeated property read/write logic behind shared helpers.
- Keep `ToolboxRibbon.bas` signatures unchanged.

### Phase 3: Module split

- Extract `ToolboxActions.bas` into focused internal modules.
- Keep existing public entrypoints or wrapper procedures.
- Move repeated prompt and validation logic into helpers.

### Phase 4: Performance-focused polish

- Revisit ribbon invalidation strategy.
- Revisit repeated presentation scans.
- Reduce selection churn and clipboard dependence where practical.

### Phase 5: Optional redesign

- If maintenance cost is still high, redesign dispatch internals more deeply while preserving the public callback layer.

## Test And Validation Scenarios

Manual validation is critical because this is Office/VBA code with strong UI coupling.

### Ribbon and callback integrity

- All callbacks referenced in `20_development/ribbonUI.xml` still resolve after module reorganization.
- Controls with duplicated IDs or stale XML references are either removed or fully backed by VBA code.
- `GetEnabled`, `getPressed`, `getText`, `onChange`, and `onAction` behavior remains unchanged for existing supported controls.

### Selection state behavior

- No selection: controls disable cleanly without errors.
- Single shape selection: geometry, margin, transparency, paragraph, and rounded-corner controls read/write correctly.
- Multi-shape selection: spacing, arrange, split, replace, and same-size actions behave consistently.
- Text selection vs shape selection: text insertion helpers and shape property callbacks fail gracefully.

### Geometry correctness

- Width/height/left/top changes still respect `ScaleFrom`.
- Rotated shapes still produce expected results for same-size and arrange operations.
- `HSep` and `VSep` edits behave the same after moving to cached sorted selections.

### Presentation-wide operations

- Slide numbering add/remove works across larger presentations.
- `setLanguage` still updates grouped shapes, tables, SmartArt, and charts.
- slide-selection export/email flows still work for saved and unsaved presentations with correct guard behavior.

### Platform compatibility

- Windows key modifiers still affect increment behavior.
- Mac builds still compile with conditional branches in `KeyState.bas` and `Slides.bas`.
- Mac-specific script availability handling remains isolated and does not break other ribbon callbacks.

## Assumptions And Defaults

- Language of this roadmap: English.
- Refactor strategy: phased safe-first.
- Scope: `20_development/src`, with `20_development/ribbonUI.xml` referenced only where callback compatibility and control drift matter.
- This document intentionally does not change code; it is a decision-ready implementation roadmap.
- Existing callback names referenced by ribbon XML are treated as public API until explicitly retired.
- The worktree may already contain unrelated local changes; implementation should avoid overwriting those without review.
