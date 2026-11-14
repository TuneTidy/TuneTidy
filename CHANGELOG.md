# Changelog

All notable changes to TuneTidy are listed here. Versions follow [Semantic Versioning](https://semver.org).

## [1.0.0] - 2026-11-14

First public release, based on the earlier private music organizer script.

### Added
- Settings menu: folders, folder layout, genres, lookup services and API keys, on/off switches, logs, reset.
- First-run setup; no personal paths in the script.
- New default layout `Genre\BPM + key` on one level, plus a choice of four layouts.
- Key groups (1-6 neighbouring Camelot numbers per folder) and key notation (Camelot, Open Key, musical).
- BPM groups per genre family (DnB: 0-169, 170-175, 176+) and a half-time BPM fix.
- Editable genre list (`GenreRules.json`) with about 40 genres, families and BPM ranges.
- Detail level: subgenre folders or one folder per family.
- Automatic discovery of new genres from your tags and online lookups, with a guessed family.
- MusicBrainz genre list used as a dictionary to filter out junk tags.
- M3U playlists pointing at the original files, so Rekordbox keeps cue points.
- Duplicate warnings, free-space check, long-path handling, log cleanup and a log summary.
- Several new-tracks folders.
- Unattended mode (`-Unattended`) and preview mode (`-WhatIf`).
- API keys encrypted with Windows DPAPI; environment variables as an alternative.
- Pester tests and a GitHub workflow that runs them.

### Changed
- Settings, keys and caches moved to `%APPDATA%\TuneTidy`; decisions, history and logs to `<output folder>\_TuneTidy`.
- The lookup cache stores raw tags, so changes to the genre list also apply to earlier lookups.
- Requests identify the app with the project URL.
- Function names follow PowerShell's approved verbs.

### Migration
- Data from the earlier script (`Organized Music\_Organizer`) is imported on the first run after you confirm: decisions, copy history, cached lookups and API keys (encrypted on import; the old plain-text key file is deleted).
