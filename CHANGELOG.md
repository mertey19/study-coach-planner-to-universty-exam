# Changelog

All notable changes to this project will be documented in this file.

## v1.0.0 - 2026-03-21

- Refactored application into clearer modules (`storage`, `services`, `config`, `data`).
- Added robust Python 3.8/3.9 typing compatibility updates.
- Added UI theming with multiple panel color presets and live switching.
- Added right-side analytics panel with:
  - Weekly plan vs done chart
  - Weekly completion chart
  - Exam score chart
  - Subject net trend chart
  - TYT/AYT and time-range filters
- Added study workflow improvements:
  - Week copy feature
  - Settings window
  - Keyboard shortcuts
- Added data safety and portability:
  - Auto backup
  - Timestamped backups in `backups/`
  - JSON export/import
  - Restore from backup
  - Full data reset with backup
- Fixed mouse wheel interaction issues around combobox/dropdown controls.
- Standardized time labels to `HH:MM` format and expanded to `00:00-23:00`.
- Added executable packaging support:
  - `requirements.txt`
  - `EXE_KURULUM.md`
  - Onefile and onedir Windows builds
  - Release ZIP packages
