# In-App Updater Design (UPS Status Widget)

## Goal

Allow users to update from inside the app without manually opening GitHub and downloading assets.

## Constraints

- Windows desktop app (WinForms, .NET 8, single-file publish).
- Existing release source is GitHub Releases.
- Current app already checks latest version and can open release page.
- Update path must be safe and explicit (no silent hidden binary execution).

## Approaches Considered

1. Browser-only flow (status quo+)
- Behavior: app detects update and opens release page.
- Pros: simplest, lowest risk.
- Cons: still manual download/install, does not meet goal.

2. In-app download of portable EXE and self-replace
- Behavior: app downloads next EXE and swaps current binary.
- Pros: no external installer dependency.
- Cons: hard on Windows while running, requires robust rollback/locking logic.

3. In-app download of installer and handoff to installer (recommended)
- Behavior: app checks latest release, selects matching installer asset (x64/x86), downloads it, verifies hash when available, launches installer, exits app.
- Pros: reliable Windows update model, reuses existing installer pipeline, clear user consent.
- Cons: requires installer asset consistency in releases.

## Selected Design

Use approach 3.

### Components

- `UpdateCore` module:
  - parse latest release payload
  - parse semantic version from tag (`vX.Y.Z`)
  - select architecture-specific installer asset
  - download asset and optionally verify SHA-256 digest
- `Program/UpsWidget` integration:
  - tray action: `Install latest update`
  - existing `Check for updates` flow upgraded with install option
  - safe user prompt before install
  - launch installer and exit app

### UX Flow

1. User clicks `Check for updates` or `Install latest update`.
2. App fetches latest GitHub release metadata.
3. If no update: message `You are up to date`.
4. If update exists:
   - prompt with options to install now or open release notes.
5. On install:
   - choose installer asset for current architecture
   - download to `%TEMP%\\UpsStatusWidget\\updates\\`
   - verify SHA-256 if release digest provided
   - run installer via shell and exit app

### Error Handling

- Network/API failures: non-blocking message (`Unable to check/download update now`).
- Asset not found: fallback to release page.
- Hash mismatch: block execution and show error.
- Installer launch failure: show error and keep app running.

### Testing Strategy

- Unit tests for:
  - tag/version parsing
  - installer asset selection for x64/x86
  - digest parsing/validation helper behavior
- Manual validation:
  - simulate latest version > current
  - download + launch installer
  - ensure app exits only after successful handoff

## Execution Plan (Subtasks)

1. Create issues for updater sub-work.
2. Implement `UpdateCore` module and tests.
3. Integrate install-update UX in tray and runtime flow.
4. Update docs and close all updater issues.
5. Tag and trigger release build.
