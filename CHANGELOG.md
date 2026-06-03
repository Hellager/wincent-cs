# Changelog

## [v0.2.4] - 2026-06-03

### Added
- Add synchronous `QuickAccessManager` facade with option/result types, structured Wincent exceptions, and retry policy support.
- Add native-first Quick Access query, recent-file removal, frequent-folder pin/unpin, and Explorer refresh paths with PowerShell fallbacks.
- Add Windows infrastructure helpers for STA execution, COM initialization, native backing-file handles, Windows Recent folder resolution, PowerShell error classification, and Windows path comparison.
- Add Recent `.lnk` deep cleanup for remove operations through `RemoveOptions.DeepCleanRecentLinks`.
- Add best-effort batch add/remove APIs with per-item `BatchResult` and `BatchFailure` reporting.
- Add `QuickAccessLock` APIs for locking Recent Files, Frequent Folders, or both backing files, including Windows Recent shortcut snapshots and optional cleanup on unlock.
- Add Quick Access visibility APIs backed by Explorer registry settings: `IsVisible`, `SetVisible`, `ShowSection`, and `HideSection`.
- Add DestList metadata parsing for Recent Files and Frequent Folders, including public `AutomaticDestinations`, `DestList`, and `DestListEntry` metadata types.
- Add experimental rebuild-based DestList removal with deleted-shortcut reporting and rebuild status reporting.
- Expand `TestConsole` into an interactive API exercise CLI with list, add/remove, batch, lock, visibility, and DestList commands.
- Add unit coverage for new infrastructure, native query/mutation paths, visibility, locks, Recent link cleanup, DestList parsing/removal, and public API baselines.

### Changed
- Update project version to 0.2.4.
- Migrate `TestWincent` to an SDK-style `net48` test project using PackageReference.
- Split large phase-0 data structures into focused source files for options, batch results, lock types, DestList types, retry policy, and exceptions.
- Harden script storage and execution by removing global state, fixing dynamic script path handling, improving quoting, and cleaning scripts by assembly version.
- Simplify and rewrite English and Chinese README files around the current 0.2.4 API surface.

### Fixed
- Fix native COM resource cleanup and timeout handling for native query, mutation, Explorer refresh, and Recent link cleanup paths.
- Fix `QuickAccessLock.Unlock` to dispose handles safely and report deleted shortcuts correctly.
- Fix Quick Access backing-file cache behavior when Recent Files or Frequent Folders data files are missing.
- Fix drive and UNC root path normalization for Windows path equality.
- Fix DestList parser validation, unsupported-version reporting, `LastEntryId` parsing, and experimental removal failure reporting.
- Fix frequent-folder removal behavior for unpinned folders and Windows 10/Windows 11 Shell verb differences.

### Removed
- Remove public async manager APIs, `IQuickAccessManager`, `ExecutionFeasibilityStatus`, and the public `ClearCache` facade.
- Remove legacy `packages.config` test dependency management.

## [v0.1.4] - 2025-04-22

### Feat
- Add cache for ScriptExecutor
- More strategies support
- Cleanup older scripts

### Fixed
- Fixed COM initialization issues in `AddFileToRecentDocs` method
  - Improved COM initialization and cleanup process
  - Added `COINIT_DISABLE_OLE1DDE` flag for better compatibility
  - Optimized error handling and resource cleanup
  - Fixed `RPC_E_CHANGED_MODE` (0x80010106) error
- Unpin folder issue
- Check strategies' timeout issue

### Changed
- Improved code comments for better clarity and internationalization
- Optimized resource management to ensure proper unmanaged resource release
- Enhanced error message readability with hexadecimal error code display
- Remove unnecessary dependies
- New implemention of `ScriptExecutor` and `QuickAccessManager`
