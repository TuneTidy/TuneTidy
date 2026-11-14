# Changelog

All notable changes to TuneTidy are listed here. Versions follow [Semantic Versioning](https://semver.org).

## [1.0.0] - unreleased

First public release.

### Added
- Pick the Rekordbox export at the start of every run: recent exports (with date and track count), a Browse window, or a pasted path.
- Playlist mode: pick one or more playlists from the export's playlist tree (numbers, ranges or search); the choice is remembered.
- `-Playlist` and `-XmlPath` parameters for unattended runs; unattended runs repeat the last kind of run.
- Settings menu: folders, folder layout, genres, lookup services and API keys, on/off switches, logs, reset.
- First-run setup; no personal paths in the script.
- Default layout `Genre\BPM + key` on one level, plus a choice of four layouts.
- Key groups (1-6 neighbouring Camelot numbers per folder) and key notation (Camelot, Open Key, musical).
- BPM bands, optional BPM groups per genre family, and a half-time BPM fix.
- Editable genre list (`GenreRules.json`) with about 40 genres, families and BPM ranges.
- Detail level: subgenre folders or one folder per family.
- Automatic discovery of new genres from your tags and online lookups, with a guessed family.
- MusicBrainz genre list used as a dictionary to filter out junk tags.
- M3U playlists pointing at the original files, so Rekordbox keeps cue points.
- Duplicate warnings, free-space check, long-path handling, log cleanup and a log summary.
- Several new-tracks folders.
- Unattended mode (`-Unattended`) and preview mode (`-WhatIf`).
- API keys encrypted with Windows DPAPI; environment variables as an alternative.
- Pester tests and a GitHub workflow that runs them on Windows PowerShell 5.1 and PowerShell 7.
