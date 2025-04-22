# Changelog

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
