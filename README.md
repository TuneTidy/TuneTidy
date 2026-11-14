# TuneTidy

Sort a DJ music library into **genre / BPM / key** folders, straight from a Rekordbox export.

```
Organized Music\
  Liquid DnB\
    170-175 BPM - Keys 07-09\
    176+ BPM - Keys 10-12\
  Neurofunk\
    170-175 BPM - Keys 01-03\
  Tech House\
    125-129 BPM - Keys 04-06\
  _Review\          (genre unknown - check these by hand)
  _Playlists\       (the same structure as Rekordbox playlists)
```

Neighbouring Camelot keys mix well, so each key folder holds a few of them (8A and 8B are relative keys, 7 and 9 are next door). Everything in one folder works together harmonically and is in the same tempo range.

## Your files are safe

TuneTidy **never moves, renames or deletes your original files**. It only *copies* them into its own output folder. The only files it ever moves are its own earlier copies inside that folder, and only after you say yes (for example when you fix a genre or change the layout). Run it with `-WhatIf` to see what would happen without touching any music files.

## What it does

1. Reads your Rekordbox collection (an XML export).
2. Works out a genre for every track, in this order:
   - genres you decided yourself in earlier runs,
   - the Genre tag in Rekordbox (tidied up, so "Drum & Bass", "Drum and Bass" and "DnB" all become **DnB**),
   - the name of the folder the track sits in,
   - online lookups on MusicBrainz, Discogs and Last.fm (optional; you approve the suggestions).
3. Shows you a plan, which you can edit in Excel if you like.
4. Copies the tracks into genre / BPM / key folders.
5. Writes M3U playlists with the same structure that point at your **original** files, so you can import them into Rekordbox and keep all your cue points, hot cues and beatgrids.

It remembers your answers, so the next run only asks about new tracks.

## Requirements

- Windows 10 or 11
- Windows PowerShell 5.1 (built into Windows) or PowerShell 7
- Rekordbox (any version that can export XML)

## Getting started

1. Download `TuneTidy.ps1` from this repository (or the latest release) into a folder of your choice.
2. Windows blocks scripts downloaded from the internet. Right-click the file > **Properties** > tick **Unblock** > OK.
3. In Rekordbox: **File > Export Collection in xml format**, and save it, for example, to your Desktop as `export.xml`.
4. Right-click `TuneTidy.ps1` > **Run with PowerShell**.

   Or open PowerShell in that folder and run:

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\TuneTidy.ps1
   ```

The first run asks a few questions: your music folder, where the sorted copies should go, your "new tracks" folder (where downloads land), and optional API keys. Everything can be changed later under **Settings**.

## The start menu

```
  1) Whole collection (first-time organize, or re-check everything)
  2) Only new tracks (your new-tracks folders)
  3) Settings
  4) Quit
```

For new tracks: import them into Rekordbox first, export a fresh XML, then choose 2.

## Folder layout

Choose under **Settings > Folder layout**:

| Layout | Example |
|---|---|
| Genre\BPM + key, one level (default) | `Liquid DnB\170-175 BPM - Keys 07-09` |
| Genre\BPM\Key, two levels | `Liquid DnB\170-175 BPM\Keys 07-09` |
| Genre\Key | `Liquid DnB\Keys 07-09` |
| Genre only | `Liquid DnB` |

You can also set:

- **Keys per folder**: 1 (a folder per Camelot number, `08 - 8A 8B`), 2, 3 (default, four folders), 4 or 6.
- **Key notation**: Camelot (`8A`), Open Key (`1m`) or musical (`Am`).
- **BPM groups per genre family**: DnB defaults to `170,176`, which gives `0-169`, `170-175` and `176+`. Other families use 5-BPM bands (`120-124 BPM`), which you can change.
- **Half-time fix**: Rekordbox sometimes detects DnB at 87 instead of 174. TuneTidy doubles such values for the BPM folders.

## Genres

TuneTidy ships with about 40 genres (DnB styles, house, techno, trance, dubstep, garage, breaks, hip hop and more), each with a **family** and a typical BPM range. Manage them under **Settings > Genres**: add, rename, merge, delete, change families and BPM ranges, or switch the **detail level**:

- **Detailed**: every subgenre gets its own folder (Neurofunk, Liquid DnB, ...).
- **Broad**: one folder per family (DnB, House, ...).

**New genres are found automatically.** When your Rekordbox tags or the online services mention a genre that is not in your list yet ("Halftime", "Minimal Drum & Bass"), TuneTidy shows them at the end of the genre step with their track counts and a guessed family. One answer adds them all.

To tell real genres from junk tags ("seen live", "favourites", artist names), TuneTidy uses MusicBrainz's public genre list, refreshed about once a month. That list is only a dictionary: folders are only made for genres that are actually in your collection.

The genre list is a plain file (`%APPDATA%\TuneTidy\GenreRules.json`), so advanced users can edit it directly. `GenreRules.default.json` in this repository shows the default list.

## Online lookups and API keys

| Service | Account | How to get a key |
|---|---|---|
| MusicBrainz | not needed | - |
| Discogs | free | discogs.com > Settings > Developers > Generate new token |
| Last.fm | free | last.fm/api/account/create (you only need the "API key") |

MusicBrainz alone finds only some tracks; Discogs and Last.fm find many more. Add, change, test or remove keys under **Settings > Lookup services and API keys**.

**How keys are kept safe:**

- Saved keys are **encrypted with Windows DPAPI**: only your Windows user on your PC can read them. A copy of the file on another computer or in a backup is useless.
- The key file is also locked so that other Windows accounts cannot open it.
- Keys are never shown on screen, never written to logs, and never included in error messages.
- Alternatively, set the environment variables `TUNETIDY_DISCOGS_TOKEN` and `TUNETIDY_LASTFM_KEY`. They win over saved keys, and nothing is stored on disk.

Lookups are cached, so each track is looked up only once. MusicBrainz is asked at most once per second, as its rules require.

## Rekordbox playlists (keep your cue points)

Cue points, hot cues and beatgrids belong to your *original* files in Rekordbox. If you load the copies from `Organized Music` onto a USB stick, they arrive without cues.

So TuneTidy also writes one `.m3u8` playlist per folder into `Organized Music\_Playlists`, pointing at your originals and sorted by key and BPM. Import them into Rekordbox (**File > Import > Import Playlist**), then export to USB as usual: same structure, all cues kept.

## Settings

Everything is under **3) Settings** in the start menu:

- **Folders**: music folder, output folder, new-tracks folders (several allowed), Rekordbox export.
- **Folder layout**: see above.
- **Genres**: see above.
- **Lookup services and API keys**: see above.
- **On/off switches**: every feature can be switched off. That covers each lookup service, folder-name hints, same-artist guessing, web-rip file name parsing, the BPM check, the half-time fix, the genre dictionary, new-genre suggestions, playlists, duplicate warnings, the free-space check, the genre log, the offers to move copies and edit the plan, and network paths.
- **Logs**: a summary of the last run, open the latest log, open the log folder, how many logs to keep.
- **Reset**: clear the lookup cache, forget your decisions (a backup is kept), restore default settings or the default genre list.

## Where TuneTidy keeps its files

| What | Where |
|---|---|
| Settings, API keys, lookup cache, genre list | `%APPDATA%\TuneTidy` |
| Your decisions, copy history, plan, logs | `<output folder>\_TuneTidy` |
| Playlists | `<output folder>\_Playlists` |

Nothing is ever stored next to the script, so the script folder can be shared or synced safely.

## Automatic runs (advanced)

TuneTidy can run without questions, for example every Sunday evening:

```powershell
.\TuneTidy.ps1 -Mode New -Unattended
```

In unattended mode, new tracks are copied using your saved settings and the saved Rekordbox export. Online genre suggestions are **not** applied: those tracks wait in `_Review` until your next normal run, where the suggestions appear instantly from the cache. Earlier copies are never moved.

To create a weekly Windows task (adjust the path):

```
schtasks /Create /TN "TuneTidy new tracks" /SC WEEKLY /D SUN /ST 20:00 /TR "powershell.exe -NoProfile -ExecutionPolicy Bypass -File \"C:\Tools\TuneTidy\TuneTidy.ps1\" -Mode New -Unattended"
```

Remember that Rekordbox only knows new tracks after you import them and export a fresh XML.

## Privacy

- Only **artist and title** are sent online, and only to the services you switch on.
- Requests identify the app as `TuneTidy/<version> ( https://github.com/Tunetidy/Tunetidy )`, as MusicBrainz and Discogs require. An email address is optional.
- No telemetry, no analytics, no accounts.

## Troubleshooting

**"Running scripts is disabled on this system"**: use **Run with PowerShell** from the right-click menu, or start it with `powershell -ExecutionPolicy Bypass -File .\TuneTidy.ps1`.

**Excel shows the plan in one column**: Excel in some languages expects `;`. TuneTidy reads both, so it's fine to save with either. Use **Save As > CSV UTF-8**.

**A track is in the wrong genre**: run again and open the plan when asked, change its Genre, save and close. TuneTidy remembers it and moves the copy.

**Something went wrong**: see the summary under **Settings > Logs**, or open an issue with the error message.

## Development

Tests use [Pester 5](https://pester.dev):

```powershell
Install-Module Pester -MinimumVersion 5.5.0 -Scope CurrentUser
Invoke-Pester -Path .\tests
```

They run automatically on GitHub for Windows PowerShell 5.1 and PowerShell 7.

The script is organised in steps (STEP 0 to STEP 7). Input is read in one place (`Get-TracksFromRekordboxXml`), so other inputs can be added as separate functions that return the same fields.

## Roadmap

- A .NET version with a proper Windows installer, published as an exe
- Writing found genres into the copies' file tags
- Serato and Traktor input
- A window-based interface

## License

MIT - see [LICENSE](LICENSE).
