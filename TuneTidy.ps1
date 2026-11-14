<#
.SYNOPSIS
TuneTidy - sorts a DJ music library into genre / BPM / key folders, using a Rekordbox XML export.

.DESCRIPTION
Run it with no arguments and answer the prompts. The first run asks a few setup questions;
everything can be changed later under "Settings" in the start menu.

TuneTidy never moves, renames or deletes your ORIGINAL files. It only copies them into your
output folder ('Organized Music' by default). The only files it ever moves are its own earlier
copies inside that folder (after you confirm), so they follow corrected genres or a new layout.

Genres come from, in this order: your earlier decisions, the Rekordbox Genre tag, the name of
the folder a track sits in, and optional online lookups (MusicBrainz, Discogs, Last.fm).
Online, only artist and title are sent.

It also writes M3U playlists that point at your ORIGINAL files, so you can import the same
structure into Rekordbox and keep all cue points, hot cues and beatgrids.

Input today: a Rekordbox XML export. Everything after STEP 3 works on a plain track list
(Artist, Title, Genre, BPM, Key, SourcePath), so other inputs (M3U, Serato, Traktor) can be
added later as extra input functions without touching the rest.

.PARAMETER XmlPath
Path to a Rekordbox XML export (File > Export Collection in xml format). If omitted, the last
used export is offered, or you are asked.

.PARAMETER Mode
'Initial' = whole collection, 'New' = only tracks inside your new-tracks folder(s).
If omitted, a menu is shown.

.PARAMETER Unattended
Runs with saved settings and asks nothing (for a scheduled task). Defaults to Mode 'New'.
Genre suggestions found online are NOT applied - they wait for your next normal run.
Earlier copies are not moved.

.PARAMETER WhatIf
Preview only: shows what would be copied, moved and written, but changes no music files.

.EXAMPLE
.\TuneTidy.ps1
Interactive run with menus.

.EXAMPLE
.\TuneTidy.ps1 -Mode New -Unattended
Copies new tracks from your new-tracks folder(s) without asking anything.

.EXAMPLE
.\TuneTidy.ps1 -WhatIf
Shows what would happen without copying or moving any music files.

.NOTES
Name:     TuneTidy
Version:  1.0.0
License:  MIT (see LICENSE)
Project:  https://github.com/Tunetidy/Tunetidy
Requires: Windows, Windows PowerShell 5.1 or PowerShell 7+

.LINK
https://github.com/Tunetidy/Tunetidy
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$XmlPath,
    [ValidateSet('Initial', 'New')]
    [string]$Mode,
    [switch]$Unattended
)

$ErrorActionPreference = 'Stop'
# Add TLS 1.2 without switching off newer protocols Windows already allows.
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

$TuneTidyVersion = '1.0.0'
$ProjectUrl      = 'https://github.com/Tunetidy/Tunetidy'

# -WhatIf is handled by TuneTidy itself: music files are not touched, but settings, the lookup
# cache, your decisions and the log are still saved (they are TuneTidy's own files).
$script:Preview = [bool]$WhatIfPreference
$WhatIfPreference = $false
$script:Unattended = [bool]$Unattended

# ---- Where TuneTidy keeps its files -----------------------------------------
# Settings, API keys and caches: %APPDATA%\TuneTidy (standard Windows place, never in your music).
# Decisions, copy history, plan and logs: '<output folder>\_TuneTidy' (they belong to that library).

$AppDataRoot    = Join-Path ([Environment]::GetFolderPath('ApplicationData')) 'TuneTidy'
$SettingsPath   = Join-Path $AppDataRoot 'Settings.json'
$KeysPath       = Join-Path $AppDataRoot 'ApiKeys.json'
$CachePath      = Join-Path $AppDataRoot 'TagCache.json'
$RulesPath      = Join-Path $AppDataRoot 'GenreRules.json'
$DictionaryPath = Join-Path $AppDataRoot 'GenreDictionary.json'

$DataFolderName     = '_TuneTidy'
$PlaylistFolderName = '_Playlists'
$ReviewFolder       = '_Review'     # genre unknown
$SamplesFolder      = '_Samples'
$SamplesGenre       = 'Samples'
$MaxPathLength      = 250           # Windows' classic limit is 260 characters

# Environment variables win over saved keys (nothing is stored on disk then).
$EnvDiscogs = 'TUNETIDY_DISCOGS_TOKEN'
$EnvLastFm  = 'TUNETIDY_LASTFM_KEY'

# YouTube / label channels that appear in the Artist field. They are never treated as artists.
# Your own list lives in Settings.json (ChannelNames) and starts as a copy of this one.
$DefaultChannelNames = @(
    'Liquicity', 'UKF', 'UKF Drum & Bass', 'UKF Dubstep', 'DnB Allstars', 'Drum&BassArena',
    'DrumAndBassArena', 'On Scope', 'DnB Ranger', 'Mentalo Music', 'Monstercat', 'Trap Nation',
    'Bass Nation', 'Chill Nation', 'Proximity', 'MrSuicideSheep', 'Majestic Casual', 'Selected',
    'NoCopyrightSounds', 'NCS', 'Arctic Empire', 'Lowly Palace', 'Mixmag', 'Boiler Room',
    'Hospital Records', 'RAM Records', 'Shogun Audio', 'Viper Recordings', 'Critical Music',
    'Metalheadz', 'Blackout Music', 'Elevate Records', 'Lifted Music', 'WeRave You'
)
$ChannelPattern = '(?i)\b(records|recordings|music|tv|official|channel|vevo|nation|network|allstars|arena|uploads?|promotions?|repost|archive)\s*$'

# ---- Genre matching ---------------------------------------------------------

$MinWinnerShare = 0.6      # winning genre family needs >= 60% of the combined evidence
$SourceWeight   = @{ 'MusicBrainz' = 1.0; 'Discogs' = 1.0; 'Last.fm' = 0.7 }

# Bracket contents that say nothing about which song it is.
$JunkInBrackets = '(?i)^(?:(?:original|extended|radio|club|album|instrumental|dub|vip|short|full|main)(?:\s+(?:mix|version|edit))?|mixed|clean|explicit|dirty|bonus track)$|official|video|audio|lyric|visuali[sz]er|\b4k\b|\bhd\b|out now|free download|premiere|kbps|remaster'

# ---- Folder layout ----------------------------------------------------------

$LayoutNames = [ordered]@{
    'GenreBpmKeyCombined' = 'Genre\BPM + key (one level)'
    'GenreBpmKey'         = 'Genre\BPM\Key (two levels)'
    'GenreKey'            = 'Genre\Key'
    'Genre'               = 'Genre only'
}
$NotationNames = [ordered]@{ 'Camelot' = 'Camelot (8A)'; 'OpenKey' = 'Open Key (1m)'; 'Musical' = 'Musical (Am)' }
# Index = Camelot number (1-12)
$MinorNames = @('', 'Abm', 'Ebm', 'Bbm', 'Fm', 'Cm', 'Gm', 'Dm', 'Am', 'Em', 'Bm', 'F#m', 'C#m')
$MajorNames = @('', 'B', 'F#', 'Db', 'Ab', 'Eb', 'Bb', 'F', 'C', 'G', 'D', 'A', 'E')

# ---- On/off switches (shown in Settings > On/off switches) -------------------

$SwitchInfo = [ordered]@{
    'MusicBrainz'     = 'Look up genres on MusicBrainz (free, no key)'
    'Discogs'         = 'Look up genres on Discogs (needs a free token)'
    'LastFm'          = 'Look up genres on Last.fm (needs a free API key)'
    'FolderHints'     = 'Use your own folder names as genre suggestions'
    'SameArtistGuess' = "Guess a genre from the same artist's other tracks"
    'SlugSplitting'   = 'Work out artist and title from web-rip file names'
    'ArtistFallback'  = "Use an artist's overall tags if a track is not found (less accurate)"
    'BpmCheck'        = 'Reject suggested genres whose BPM does not fit'
    'HalfTimeFix'     = 'Double half-time BPM for BPM folders (87 -> 174)'
    'GenreDictionary' = 'Use the MusicBrainz genre list to filter out junk tags'
    'AskNewGenres'    = 'Offer to add new genres found in your tags or online'
    'Playlists'       = 'Write M3U playlists (pointing at your originals) for Rekordbox'
    'Duplicates'      = 'Warn about tracks that exist as several files'
    'FreeSpaceCheck'  = 'Check free disk space before copying'
    'GenreLog'        = 'Write a detailed genre matching log'
    'OfferMoveCopies' = 'Offer to move earlier copies that are now in the wrong folder'
    'OfferPlanEdit'   = 'Offer to edit the plan in your spreadsheet app'
    'NetworkPaths'    = 'Allow tracks on network shares (\\server\...)'
}

$Inv = [Globalization.CultureInfo]::InvariantCulture
$script:RunStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$script:UserAgent = "TuneTidy/$TuneTidyVersion"
$script:ArtistCache = @{}
$script:ApiFailures = @{}
$script:LastApiError = @{}
$script:DisabledSources = @{}
$script:Keys = @{ Discogs = ''; LastFm = ''; DiscogsFrom = 'not set'; LastFmFrom = 'not set' }
$script:GenreLogPath = ''
$script:GenreLogStarted = $false
$script:DecisionsChanged = $false
$script:Decisions = @{}
$script:Discovered = @{}
$script:GenreDictionary = @{}

# =============================================================================
# Defaults
# =============================================================================

function Get-DefaultSettings {
    return [ordered]@{
        SettingsVersion  = 1
        MusicRoot        = ''
        OutputFolder     = ''
        NewTracksFolders = @()
        XmlPath          = ''
        MbContact        = ''
        FolderLayout     = 'GenreBpmKeyCombined'
        KeyGroupSize     = 3            # Camelot numbers per key folder (1-6)
        KeyNotation      = 'Camelot'    # Camelot / OpenKey / Musical
        DetailLevel      = 'Detailed'   # Detailed = subgenre folders, Broad = one folder per family
        BpmBandWidth     = 5            # used for families without their own BPM groups
        BpmGroups        = [ordered]@{ 'DnB' = '170,176' }   # boundaries -> 0-169 | 170-175 | 176+
        KeepLogs         = 10
        ChannelNames     = @($DefaultChannelNames)
        Switches         = [ordered]@{
            MusicBrainz     = $true
            Discogs         = $true
            LastFm          = $true
            FolderHints     = $true
            SameArtistGuess = $true
            SlugSplitting   = $true
            ArtistFallback  = $false
            BpmCheck        = $true
            HalfTimeFix     = $true
            GenreDictionary = $true
            AskNewGenres    = $true
            Playlists       = $true
            Duplicates      = $true
            FreeSpaceCheck  = $true
            GenreLog        = $true
            OfferMoveCopies = $true
            OfferPlanEdit   = $true
            NetworkPaths    = $false
        }
    }
}

function Get-DefaultGenreRules {
    # Order matters: the first genre whose pattern fits a tag wins, so specific genres come first.
    # Patterns are regular expressions checked against the tag in lower case, with '&' written as 'and'.
    $genres = @(
        [ordered]@{ Name = 'Neurofunk'; Family = 'DnB'; Patterns = @('neuro ?funk', 'techstep'); Aliases = @('neuro', 'neuro dnb', 'neurofunk drum and bass'); Bpm = '' }
        [ordered]@{ Name = 'Jump Up'; Family = 'DnB'; Patterns = @('jump ?up'); Aliases = @('jumpup', 'jump up dnb', 'jump up drum and bass'); Bpm = '' }
        [ordered]@{ Name = 'Liquid DnB'; Family = 'DnB'; Patterns = @('\bliquid\b'); Aliases = @('liquid', 'liquid funk', 'liquid dnb', 'liquid drum and bass', 'liquid d and b'); Bpm = '' }
        [ordered]@{ Name = 'Jungle'; Family = 'DnB'; Patterns = @('\bjungle\b'); Aliases = @('ragga jungle'); Bpm = '' }
        [ordered]@{ Name = 'Halftime'; Family = 'DnB'; Patterns = @('half ?time'); Aliases = @('halftime dnb'); Bpm = '75-95,150-180' }
        [ordered]@{ Name = 'DnB'; Family = 'DnB'; Patterns = @('drum ?(and|n) ?bass', '\bdnb\b', '\bd (and|n) b\b'); Aliases = @('drum and bass', 'drum n bass', 'drumandbass', 'drumnbass', 'dnb', 'd and b', 'd n b'); Bpm = '155-190,77-95' }
        [ordered]@{ Name = 'Riddim'; Family = 'Dubstep'; Patterns = @('\briddim\b'); Aliases = @(); Bpm = '' }
        [ordered]@{ Name = 'Dubstep'; Family = 'Dubstep'; Patterns = @('dubstep', 'brostep'); Aliases = @('brostep'); Bpm = '135-152,67-76' }
        [ordered]@{ Name = 'Hardcore'; Family = 'Hardcore'; Patterns = @('happy hardcore', 'uk hardcore', 'gabber', 'hardcore techno', 'frenchcore'); Aliases = @(); Bpm = '160-220' }
        [ordered]@{ Name = 'Hardstyle'; Family = 'Hardstyle'; Patterns = @('hardstyle', 'rawstyle', 'hard dance'); Aliases = @(); Bpm = '140-165' }
        [ordered]@{ Name = 'Bass House'; Family = 'House'; Patterns = @('bass house'); Aliases = @(); Bpm = '' }
        [ordered]@{ Name = 'Tech House'; Family = 'House'; Patterns = @('tech house'); Aliases = @(); Bpm = '' }
        [ordered]@{ Name = 'Deep House'; Family = 'House'; Patterns = @('deep house'); Aliases = @(); Bpm = '' }
        [ordered]@{ Name = 'Progressive House'; Family = 'House'; Patterns = @('progressive house'); Aliases = @(); Bpm = '' }
        [ordered]@{ Name = 'Afro House'; Family = 'House'; Patterns = @('afro house'); Aliases = @(); Bpm = '' }
        [ordered]@{ Name = 'House'; Family = 'House'; Patterns = @('\bhouse\b'); Aliases = @(); Bpm = '110-135' }
        [ordered]@{ Name = 'Hard Techno'; Family = 'Techno'; Patterns = @('hard techno'); Aliases = @(); Bpm = '' }
        [ordered]@{ Name = 'Melodic Techno'; Family = 'Techno'; Patterns = @('melodic techno'); Aliases = @(); Bpm = '' }
        [ordered]@{ Name = 'Techno'; Family = 'Techno'; Patterns = @('techno'); Aliases = @(); Bpm = '115-160' }
        [ordered]@{ Name = 'Psytrance'; Family = 'Trance'; Patterns = @('psy ?trance', '\bgoa\b'); Aliases = @('psy', 'goa trance'); Bpm = '' }
        [ordered]@{ Name = 'Trance'; Family = 'Trance'; Patterns = @('trance'); Aliases = @(); Bpm = '120-150' }
        [ordered]@{ Name = 'UK Garage'; Family = 'UK Garage'; Patterns = @('uk garage', '\bukg\b', '\b2 ?step\b', 'speed garage', '\bgarage\b(?! rock)'); Aliases = @('garage', 'ukg', '2 step', '2step'); Bpm = '125-142' }
        [ordered]@{ Name = 'Grime'; Family = 'Grime'; Patterns = @('\bgrime\b'); Aliases = @(); Bpm = '135-145' }
        [ordered]@{ Name = 'Breaks'; Family = 'Breaks'; Patterns = @('breakbeat', '\bbreaks\b', 'big beat'); Aliases = @('breakbeat'); Bpm = '95-145' }
        [ordered]@{ Name = 'Trap'; Family = 'Trap'; Patterns = @('\btrap\b'); Aliases = @(); Bpm = '' }
        [ordered]@{ Name = 'Hip Hop'; Family = 'Hip Hop'; Patterns = @('hip ?hop', '\brap\b', 'boom ?bap'); Aliases = @('hiphop', 'rap'); Bpm = '' }
        [ordered]@{ Name = 'R&B'; Family = 'R&B'; Patterns = @('\br (and|n) b\b', '\brnb\b', 'rhythm and blues'); Aliases = @('rnb', 'r and b'); Bpm = '' }
        [ordered]@{ Name = 'Downtempo'; Family = 'Downtempo'; Patterns = @('downtempo', 'chill ?out', 'trip ?hop', '\blo ?fi\b', 'chillhop'); Aliases = @(); Bpm = '' }
        [ordered]@{ Name = 'Ambient'; Family = 'Ambient'; Patterns = @('ambient'); Aliases = @(); Bpm = '' }
        [ordered]@{ Name = 'Nu Disco'; Family = 'Disco'; Patterns = @('nu ?disco'); Aliases = @(); Bpm = '' }
        [ordered]@{ Name = 'Disco'; Family = 'Disco'; Patterns = @('disco'); Aliases = @(); Bpm = '' }
        [ordered]@{ Name = 'Funk & Soul'; Family = 'Funk & Soul'; Patterns = @('\bfunk\b', '\bsoul\b', 'motown'); Aliases = @('funk', 'soul'); Bpm = '' }
        [ordered]@{ Name = 'Reggae'; Family = 'Reggae'; Patterns = @('reggae', 'dancehall', '\bdub\b'); Aliases = @('dub'); Bpm = '' }
        [ordered]@{ Name = 'Latin'; Family = 'Latin'; Patterns = @('latin', 'reggaeton', 'salsa', 'bachata', 'cumbia'); Aliases = @(); Bpm = '' }
        [ordered]@{ Name = 'Pop'; Family = 'Pop'; Patterns = @('\bpop\b', 'synth ?pop', 'electropop', 'k ?pop'); Aliases = @(); Bpm = '' }
        [ordered]@{ Name = 'Metal'; Family = 'Metal'; Patterns = @('metal'); Aliases = @(); Bpm = '' }
        [ordered]@{ Name = 'Rock'; Family = 'Rock'; Patterns = @('\brock\b', 'punk', '\bindie\b', 'grunge'); Aliases = @('indie'); Bpm = '' }
        [ordered]@{ Name = 'Jazz'; Family = 'Jazz'; Patterns = @('jazz'); Aliases = @(); Bpm = '' }
        [ordered]@{ Name = 'Classical'; Family = 'Classical'; Patterns = @('classical', 'orchestral', 'baroque'); Aliases = @(); Bpm = '' }
    )
    return [ordered]@{
        RulesVersion = 1
        Genres       = $genres
        Ignored      = @('electronic', 'electronica', 'dance', 'edm', 'club', 'instrumental', 'experimental', 'bass', 'bass music', 'seen live', 'favorites', 'favourites', 'vocal', 'remix')
        Declined     = @()
    }
}

# =============================================================================
# Safe file helpers
# =============================================================================

# Text from YouTube titles / web services can start with "=", "+", "-" or "@". Excel would
# run such a cell as a formula, so those cells are written with a leading apostrophe.
function Export-SafeCsv($Objects, [string]$Path, [switch]$Append) {
    $safe = foreach ($o in @($Objects)) {
        if ($null -eq $o) { continue }
        $copy = [ordered]@{}
        foreach ($p in $o.PSObject.Properties) {
            $v = $p.Value
            if ($v -is [string] -and $v -match '^[=+\-@\t\r]') { $v = "'" + $v }
            $copy[$p.Name] = $v
        }
        [pscustomobject]$copy
    }
    if ($Append) { $safe | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8 -Append }
    else { $safe | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8 }
}

# Reads a CSV written by TuneTidy or saved by Excel (Excel in some languages uses ';').
function Import-SafeCsv([string]$Path) {
    $first = [string](Get-Content -LiteralPath $Path -TotalCount 1 -Encoding UTF8)
    $delim = if ($first -match ';' -and $first -notmatch ',') { ';' } else { ',' }
    $rows = @(Import-Csv -LiteralPath $Path -Delimiter $delim -Encoding UTF8)
    foreach ($r in $rows) {
        foreach ($p in $r.PSObject.Properties) {
            if ($p.Value -is [string] -and $p.Value -match "^'[=+\-@\t\r]") { $p.Value = $p.Value.Substring(1) }
        }
    }
    return $rows
}

# Write to a temp file first, so a crash never leaves a half-written file behind.
function Save-CsvAtomic($Objects, [string]$Path) {
    $tmp = "$Path.tmp"
    Export-SafeCsv $Objects $tmp
    Move-Item -LiteralPath $tmp -Destination $Path -Force
}

function Save-JsonAtomic($Object, [string]$Path, [switch]$Compress) {
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    $tmp = "$Path.tmp"
    if ($Compress) { $json = $Object | ConvertTo-Json -Depth 6 -Compress }
    else { $json = $Object | ConvertTo-Json -Depth 6 }
    Set-Content -LiteralPath $tmp -Value $json -Encoding UTF8
    Move-Item -LiteralPath $tmp -Destination $Path -Force
}

# Returns the parsed JSON, or $null (with a warning) if the file is missing or unreadable.
function Read-JsonFile([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }
    try { return (Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json) }
    catch { Write-Warning "Could not read ${Path}: $($_.Exception.Message)"; return $null }
}

# Only the current Windows user may read the file (on top of the encryption of the keys).
function Protect-PrivateFile([string]$Path) {
    try {
        $acl = Get-Acl -LiteralPath $Path
        $acl.SetAccessRuleProtection($true, $false)
        foreach ($rule in @($acl.Access)) { [void]$acl.RemoveAccessRule($rule) }
        $me = [Security.Principal.WindowsIdentity]::GetCurrent().Name
        $acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule($me, 'FullControl', 'Allow')))
        Set-Acl -LiteralPath $Path -AclObject $acl
    }
    catch { Write-Warning "Could not restrict access to ${Path}: $($_.Exception.Message)" }
}

function Get-FullPathOrNull([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    try { return ([IO.Path]::GetFullPath($Path)).TrimEnd('\') } catch { return $null }
}

function Test-InsideFolder([string]$Path, [string]$Root) {
    $full = Get-FullPathOrNull $Path
    $r = Get-FullPathOrNull $Root
    if (-not $full -or -not $r) { return $false }
    return $full.StartsWith($r + '\', [StringComparison]::OrdinalIgnoreCase)
}

# Only local drive paths (C:\...) are read unless network paths are switched on. Touching a
# \\server\ path makes Windows try to log in to that server.
function Test-AllowedSourcePath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    if (Test-Switch 'NetworkPaths') { return $true }
    return ($Path -match '^[A-Za-z]:\\')
}

function Get-FreePath([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return $Path }
    $dir = Split-Path -Parent $Path
    $base = [IO.Path]::GetFileNameWithoutExtension($Path)
    $ext = [IO.Path]::GetExtension($Path)
    $n = 2
    do {
        $candidate = Join-Path $dir ('{0} ({1}){2}' -f $base, $n, $ext)
        $n++
    } while (Test-Path -LiteralPath $candidate)
    return $candidate
}

# Keeps full paths under Windows' length limit by shortening the file name (never the folders).
# Returns $null if even a short name would not fit.
function Get-ShortPath([string]$Dir, [string]$FileName) {
    $full = Join-Path $Dir $FileName
    if ($full.Length -le $MaxPathLength) { return $full }
    $ext = [IO.Path]::GetExtension($FileName)
    $base = [IO.Path]::GetFileNameWithoutExtension($FileName)
    $room = $MaxPathLength - $Dir.Length - 1 - $ext.Length - 4    # 4 = room for " (2)"
    if ($room -lt 15) { return $null }
    $base = $base.Substring(0, [math]::Min($base.Length, $room)).TrimEnd(' ', '.')
    return (Join-Path $Dir ($base + $ext))
}

# Removes folders that became empty, walking up but never leaving $Root or entering $Keep.
function Remove-EmptyFolders([string[]]$Folders, [string]$Root, [string[]]$Keep = @()) {
    $rootFull = Get-FullPathOrNull $Root
    foreach ($start in @($Folders | Where-Object { $_ } | Select-Object -Unique)) {
        $d = Get-FullPathOrNull $start
        while ($d -and (Test-InsideFolder $d $Root) -and ($d -ine $rootFull)) {
            $protected = $false
            foreach ($k in @($Keep | Where-Object { $_ })) {
                if (((Get-FullPathOrNull $k) -ieq $d) -or (Test-InsideFolder $d $k)) { $protected = $true }
            }
            if ($protected) { break }
            if (-not (Test-Path -LiteralPath $d -PathType Container)) { break }
            if (Get-ChildItem -LiteralPath $d -Force | Select-Object -First 1) { break }
            Remove-Item -LiteralPath $d
            $d = Get-FullPathOrNull (Split-Path -Parent $d)
        }
    }
}

function Limit-Text([string]$Text, [int]$Max) {
    if (-not $Text -or $Text.Length -le $Max) { return $Text }
    return $Text.Substring(0, $Max) + '...'
}

function ConvertTo-IntOrDefault($Value, [int]$Default, [int]$Min, [int]$Max) {
    $n = 0
    if (-not [int]::TryParse(([string]$Value).Trim(), [ref]$n)) { return $Default }
    return [math]::Min($Max, [math]::Max($Min, $n))
}

# =============================================================================
# Settings
# =============================================================================

function Test-Switch([string]$Name) { return [bool]$script:Settings.Switches[$Name] }

# Missing values are filled from the defaults, so older settings files keep working.
function Read-Settings {
    $s = Get-DefaultSettings
    $j = Read-JsonFile $SettingsPath
    if (-not $j) { return $s }
    foreach ($key in @($s.Keys)) {
        $p = $j.PSObject.Properties[$key]
        if (-not $p -or $null -eq $p.Value) { continue }
        if ($key -eq 'Switches') {
            foreach ($q in $p.Value.PSObject.Properties) { $s.Switches[$q.Name] = [bool]$q.Value }
        }
        elseif ($key -eq 'BpmGroups') {
            $s.BpmGroups = [ordered]@{}
            foreach ($q in $p.Value.PSObject.Properties) { $s.BpmGroups[$q.Name] = [string]$q.Value }
        }
        elseif ($key -in @('NewTracksFolders', 'ChannelNames')) {
            $s[$key] = @($p.Value | Where-Object { $_ } | ForEach-Object { [string]$_ })
        }
        else { $s[$key] = $p.Value }
    }
    $s.KeyGroupSize = ConvertTo-IntOrDefault $s.KeyGroupSize 3 1 6
    $s.BpmBandWidth = ConvertTo-IntOrDefault $s.BpmBandWidth 5 1 50
    $s.KeepLogs = ConvertTo-IntOrDefault $s.KeepLogs 10 1 1000
    if ([string]$s.FolderLayout -notin @($LayoutNames.Keys)) { $s.FolderLayout = 'GenreBpmKeyCombined' }
    if ([string]$s.KeyNotation -notin @($NotationNames.Keys)) { $s.KeyNotation = 'Camelot' }
    if ([string]$s.DetailLevel -notin @('Detailed', 'Broad')) { $s.DetailLevel = 'Detailed' }
    foreach ($k in @('MusicRoot', 'OutputFolder', 'XmlPath', 'MbContact')) { $s[$k] = [string]$s[$k] }
    return $s
}

function Save-Settings { Save-JsonAtomic $script:Settings $SettingsPath }

function Update-UserAgent {
    # MusicBrainz and Discogs ask every app to name itself and give a way to contact its maker.
    $contact = $ProjectUrl
    if ($script:Settings.MbContact) { $contact = $contact + ' ; ' + [string]$script:Settings.MbContact }
    $script:UserAgent = 'TuneTidy/{0} ( {1} )' -f $TuneTidyVersion, $contact
}

# Paths that depend on the output folder.
function Initialize-Paths {
    $script:OutputRoot     = Get-FullPathOrNull ([string]$script:Settings.OutputFolder)
    $script:DataRoot       = Join-Path $script:OutputRoot $DataFolderName
    $script:PlaylistRoot   = Join-Path $script:OutputRoot $PlaylistFolderName
    $script:LogFolder      = Join-Path $script:DataRoot 'Logs'
    $script:PlanPath       = Join-Path $script:DataRoot 'Plan.csv'
    $script:HistoryPath    = Join-Path $script:DataRoot 'Processed.csv'
    $script:DecisionsPath  = Join-Path $script:DataRoot 'Decisions.csv'
    $script:DuplicatesPath = Join-Path $script:DataRoot 'Duplicates.csv'
    $newLog = Join-Path $script:LogFolder ('Genre-Log-{0}.csv' -f $script:RunStamp)
    if ($newLog -ne $script:GenreLogPath) { $script:GenreLogPath = $newLog; $script:GenreLogStarted = $false }
    $script:ChannelNorms = @($script:Settings.ChannelNames | ForEach-Object { Get-Norm $_ } | Where-Object { $_ })
    Update-UserAgent
}

# The output folder must be its own folder, never the music folder itself or a new-tracks folder.
# Returns '' when fine, otherwise the reason.
function Test-OutputFolderOk([string]$Path) {
    $p = Get-FullPathOrNull $Path
    if (-not $p) { return 'Please enter a full folder path.' }
    $m = Get-FullPathOrNull ([string]$script:Settings.MusicRoot)
    if ($m -and (($p -ieq $m) -or (Test-InsideFolder $m $p))) {
        return 'The output folder must be a separate folder (for example inside your music folder), not the music folder itself.'
    }
    foreach ($nf in @($script:Settings.NewTracksFolders)) {
        $f = Get-FullPathOrNull $nf
        if ($f -and (($p -ieq $f) -or (Test-InsideFolder $p $f) -or (Test-InsideFolder $f $p))) {
            return 'The output folder cannot be inside a new-tracks folder (or the other way round).'
        }
    }
    return ''
}

function Add-NewTracksFolder([string]$Path) {
    $f = Get-FullPathOrNull $Path
    if (-not $f -or -not (Test-Path -LiteralPath $f -PathType Container)) { Write-Warning "Folder not found: $Path"; return $false }
    if ($script:OutputRoot -and (($f -ieq $script:OutputRoot) -or (Test-InsideFolder $f $script:OutputRoot) -or (Test-InsideFolder $script:OutputRoot $f))) {
        Write-Warning 'A new-tracks folder cannot be inside the output folder (or the other way round).'
        return $false
    }
    if (@($script:Settings.NewTracksFolders) -contains $f) { return $true }
    $script:Settings.NewTracksFolders = @($script:Settings.NewTracksFolders) + $f
    return $true
}

# =============================================================================
# Genre rules (GenreRules.json)
# =============================================================================

# "Minimal Drum & Bass" -> '^minimal ?drum ?(and|n) ?bass$' (matches the whole tag only).
function New-GenrePattern([string]$Name) {
    $words = @((Get-Norm $Name) -split ' ' | Where-Object { $_ } | ForEach-Object {
        if ($_ -eq 'and') { '(and|n)' } else { [regex]::Escape($_) }
    })
    if ($words.Count -eq 0) { return '^$' }
    return ('^' + ($words -join ' ?') + '$')
}

function Set-GenreRules($Genres, $Ignored, $Declined) {
    $list = New-Object System.Collections.ArrayList
    $byName = @{}
    foreach ($g in @($Genres)) {
        if (-not $g -or -not $g.Name) { continue }
        $name = ([string]$g.Name).Trim()
        if (-not $name -or $byName.ContainsKey($name)) { continue }
        $patterns = @()
        foreach ($p in @($g.Patterns | Where-Object { $_ })) {
            try { [void][regex]::new([string]$p); $patterns += [string]$p }
            catch { Write-Warning "Genre '$name': pattern '$p' is not valid and is ignored." }
        }
        if ($patterns.Count -eq 0) { $patterns = @(New-GenrePattern $name) }
        $aliases = @($g.Aliases | Where-Object { $_ } | ForEach-Object { [string]$_ })
        $family = if ($g.Family) { ([string]$g.Family).Trim() } else { $name }
        $rule = [pscustomobject]@{
            Name        = $name
            Family      = $family
            Patterns    = $patterns
            Aliases     = $aliases
            NormAliases = @($aliases | ForEach-Object { Get-Norm $_ })
            Bpm         = ([string]$g.Bpm).Trim()
        }
        [void]$list.Add($rule)
        $byName[$name] = $rule
    }
    $script:GenreRules = @($list)
    $script:RuleByName = $byName
    $script:IgnoredRaw = @($Ignored | Where-Object { $_ } | ForEach-Object { [string]$_ } | Select-Object -Unique)
    $script:DeclinedRaw = @($Declined | Where-Object { $_ } | ForEach-Object { [string]$_ } | Select-Object -Unique)
    $script:IgnoredGenres = @{}
    foreach ($i in $script:IgnoredRaw) { $script:IgnoredGenres[(Get-Norm $i)] = $true }
    $script:DeclinedGenres = @{}
    foreach ($i in $script:DeclinedRaw) { $script:DeclinedGenres[(Get-Norm $i)] = $true }
}

function Update-GenreRuleSets { Set-GenreRules $script:GenreRules $script:IgnoredRaw $script:DeclinedRaw }

function Read-GenreRules {
    $data = Read-JsonFile $RulesPath
    if ($data -and $data.Genres) { Set-GenreRules $data.Genres $data.Ignored $data.Declined; return }
    $defaults = Get-DefaultGenreRules
    Set-GenreRules $defaults.Genres $defaults.Ignored $defaults.Declined
    if (Test-Path -LiteralPath $RulesPath) {
        Write-Warning "Your genre list ($RulesPath) could not be read, so the default list is used for this run. Fix or delete the file."
    }
    else { Save-GenreRules }
}

function Save-GenreRules {
    $genres = @(foreach ($r in $script:GenreRules) {
        [ordered]@{ Name = $r.Name; Family = $r.Family; Patterns = @($r.Patterns); Aliases = @($r.Aliases); Bpm = $r.Bpm }
    })
    $data = [ordered]@{
        RulesVersion = 1
        Help         = 'Checked from top to bottom; the first pattern that fits a tag wins. Patterns are regular expressions on the tag in lower case, with & written as and. Family groups subgenres. Bpm: accepted ranges like 155-190,77-95 (the first range is the main tempo). Ignored: junk tags. Declined: genres not to suggest again.'
        Genres       = $genres
        Ignored      = @($script:IgnoredRaw)
        Declined     = @($script:DeclinedRaw)
    }
    Save-JsonAtomic $data $RulesPath
}

# =============================================================================
# API keys (encrypted with Windows DPAPI: only this Windows user on this PC can read them)
# =============================================================================

function Protect-Secret([string]$Plain) {
    if (-not $Plain) { return '' }
    # -AsPlainText is needed to turn the pasted key into a SecureString; it is encrypted right away.
    $ss = ConvertTo-SecureString -String $Plain -AsPlainText -Force
    return (ConvertFrom-SecureString -SecureString $ss)
}

function ConvertFrom-SecureStringToPlain([Security.SecureString]$Secure) {
    if (-not $Secure -or $Secure.Length -eq 0) { return '' }
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secure)
    try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}

function Unprotect-Secret([string]$Cipher) {
    if (-not $Cipher) { return '' }
    try { return (ConvertFrom-SecureStringToPlain (ConvertTo-SecureString -String $Cipher)) }
    catch {
        Write-Warning 'A saved API key could not be unlocked (saved by another Windows user or on another PC?). Enter it again under Settings > Lookup services.'
        return ''
    }
}

function Get-SavedKeys {
    $k = @{ Discogs = ''; LastFm = '' }
    $j = Read-JsonFile $KeysPath
    if ($j) {
        foreach ($n in @('Discogs', 'LastFm')) {
            $p = $j.PSObject.Properties[$n]
            if ($p -and $p.Value) { $k[$n] = Unprotect-Secret ([string]$p.Value) }
        }
    }
    return $k
}

function Save-SavedKeys($Keys) {
    $data = [ordered]@{
        KeysVersion = 1
        Note        = 'Encrypted with Windows DPAPI: only this Windows user on this PC can read these values.'
        Discogs     = (Protect-Secret ([string]$Keys.Discogs))
        LastFm      = (Protect-Secret ([string]$Keys.LastFm))
    }
    Save-JsonAtomic $data $KeysPath
    Protect-PrivateFile $KeysPath
}

# Environment variables win over saved keys.
function Update-ApiKeys {
    $saved = Get-SavedKeys
    $script:Keys = @{ Discogs = ''; LastFm = ''; DiscogsFrom = 'not set'; LastFmFrom = 'not set' }
    foreach ($pair in @(@('Discogs', $EnvDiscogs), @('LastFm', $EnvLastFm))) {
        $name = $pair[0]
        $fromEnv = [Environment]::GetEnvironmentVariable($pair[1])
        if ($fromEnv) {
            $script:Keys[$name] = $fromEnv.Trim()
            $script:Keys[$name + 'From'] = 'environment variable ' + $pair[1]
        }
        elseif ($saved[$name]) {
            $script:Keys[$name] = $saved[$name]
            $script:Keys[$name + 'From'] = 'saved (encrypted)'
        }
    }
}

# $true = works, $false = rejected, $null = could not check (no internet?).
function Test-ApiKey([string]$Name, [string]$Key) {
    if (-not $Key) { return $false }
    try {
        if ($Name -eq 'Discogs') {
            $headers = @{ 'User-Agent' = $script:UserAgent; 'Authorization' = "Discogs token=$Key" }
            $null = Invoke-RestMethod -Uri 'https://api.discogs.com/oauth/identity' -Headers $headers -TimeoutSec 20 -UseBasicParsing
            return $true
        }
        $url = 'https://ws.audioscrobbler.com/2.0/?method=chart.gettopartists&limit=1&format=json&api_key=' + [uri]::EscapeDataString($Key)
        $r = Invoke-RestMethod -Uri $url -Headers @{ 'User-Agent' = $script:UserAgent } -TimeoutSec 20 -UseBasicParsing
        return (-not $r.error)
    }
    catch {
        $status = 0
        try { $status = [int]$_.Exception.Response.StatusCode } catch { }
        if ($status -in @(400, 401, 403)) { return $false }
        return $null
    }
}

# =============================================================================
# Logging
# =============================================================================

function Write-GenreLog($Row, [string]$Stage, [string]$Source = '', [string]$SuggestedGenre = '', [string]$Decision = '', [string]$Evidence = '', [string]$Rule = '') {
    if (-not (Test-Switch 'GenreLog') -or -not $script:GenreLogPath) { return }
    if (-not $script:GenreLogStarted -and -not (Test-Path -LiteralPath $script:LogFolder)) {
        New-Item -ItemType Directory -Force -Path $script:LogFolder | Out-Null
    }
    $entry = [pscustomobject]@{
        Time                   = (Get-Date).ToString('s')
        Stage                  = $Stage
        TrackID                = [string]$Row.TrackID
        Artist                 = [string]$Row.Artist
        Title                  = [string]$Row.Title
        OriginalRekordboxGenre = [string]$Row.OriginalGenre
        SearchArtist           = [string]$Row.ParsedArtist
        SearchTitle            = [string]$Row.ParsedTitle
        BPM                    = [string]$Row.BPM
        Source                 = $Source
        SuggestedGenre         = $SuggestedGenre
        MatchedEvidence        = $Evidence
        MatchedRule            = $Rule
        Decision               = $Decision
        FinalGenre             = [string]$Row.Genre
        SourcePath             = [string]$Row.SourcePath
    }
    Export-SafeCsv $entry $script:GenreLogPath -Append:$script:GenreLogStarted
    $script:GenreLogStarted = $true
}

function Get-LogFiles {
    if (-not $script:LogFolder -or -not (Test-Path -LiteralPath $script:LogFolder)) { return @() }
    return @(Get-ChildItem -LiteralPath $script:LogFolder -Filter 'Genre-Log-*.csv' -File | Sort-Object Name -Descending)
}

# Keeps the newest logs (this run's log counts as one of them).
function Remove-OldLogs {
    $keep = [math]::Max(1, [int]$script:Settings.KeepLogs)
    foreach ($f in @(Get-LogFiles | Select-Object -Skip ($keep - 1))) {
        try { Remove-Item -LiteralPath $f.FullName } catch { }
    }
}

# =============================================================================
# Prompts
# =============================================================================

# In unattended mode nothing is asked; $WhenUnattended is the answer used instead.
function Read-YesNo([string]$Question, [bool]$Default = $true, [bool]$WhenUnattended = $false) {
    if ($script:Unattended) { return $WhenUnattended }
    $suffix = if ($Default) { '[Y/n]' } else { '[y/N]' }
    while ($true) {
        $answer = (Read-Host "$Question $suffix").Trim().ToLowerInvariant()
        if ($answer -eq '') { return $Default }
        if ($answer -in @('y', 'yes')) { return $true }
        if ($answer -in @('n', 'no')) { return $false }
        Write-Host 'Please answer y or n.' -ForegroundColor Yellow
    }
}

function Read-Choice([string]$Prompt, [string[]]$Valid, [string]$Default = '') {
    while ($true) {
        $a = (Read-Host $Prompt).Trim().ToLowerInvariant()
        if ($a -eq '' -and $Default) { return $Default }
        if ($Valid -contains $a) { return $a }
        Write-Host ('Please type one of: ' + ($Valid -join ', ')) -ForegroundColor Yellow
    }
}

function Read-FolderPath([string]$Prompt, [string]$Default = '', [bool]$MustExist = $true, [bool]$AllowEmpty = $false) {
    while ($true) {
        $q = if ($Default) { "$Prompt [$Default]" } else { $Prompt }
        $v = (Read-Host $q).Trim().Trim('"').Trim()
        if (-not $v) {
            if ($Default) { $v = $Default }
            elseif ($AllowEmpty) { return '' }
            else { continue }
        }
        $full = Get-FullPathOrNull ([Environment]::ExpandEnvironmentVariables($v))
        $isNetwork = $full -and $full.StartsWith('\\')
        if (-not $full -or (-not ($full -match '^[A-Za-z]:') -and -not ($isNetwork -and (Test-Switch 'NetworkPaths')))) {
            Write-Host 'Please enter a full folder path, for example C:\Users\you\Music' -ForegroundColor Yellow
            continue
        }
        if ($MustExist -and -not (Test-Path -LiteralPath $full -PathType Container)) {
            Write-Host "Folder not found: $full" -ForegroundColor Yellow
            continue
        }
        return $full
    }
}

function Read-FilePath([string]$Prompt, [bool]$AllowEmpty = $true) {
    while ($true) {
        $v = (Read-Host $Prompt).Trim().Trim('"').Trim()
        if (-not $v -and $AllowEmpty) { return '' }
        if ($v -and (Test-Path -LiteralPath $v -PathType Leaf)) { return (Get-FullPathOrNull $v) }
        Write-Host 'File not found. Please paste the full path (tip: Shift + right-click the file > Copy as path).' -ForegroundColor Yellow
    }
}

# =============================================================================
# General helpers
# =============================================================================

function Convert-FileUriToPath([string]$Location) {
    if ([string]::IsNullOrWhiteSpace($Location)) { return $null }
    try {
        if ($Location -match '^file:') {
            # Rekordbox writes file://localhost/C:/path/file.mp3 - strip scheme/host manually,
            # because [Uri].LocalPath would turn "localhost" into a UNC share.
            $p = $Location -replace '^file://(localhost)?/', ''
            $p = [System.Uri]::UnescapeDataString($p)
            return ($p -replace '/', '\')
        }
        return [Environment]::ExpandEnvironmentVariables($Location)
    }
    catch { return $null }
}

function Format-FolderName([string]$Name) {
    if ([string]::IsNullOrWhiteSpace($Name)) { return 'Review' }
    $clean = $Name.Trim()
    foreach ($char in [IO.Path]::GetInvalidFileNameChars()) { $clean = $clean.Replace([string]$char, '_') }
    $clean = $clean.Trim().TrimEnd('.').Trim()
    if ([string]::IsNullOrWhiteSpace($clean)) { return 'Review' }
    # Windows device names cannot be folders
    if ($clean -match '^(?i)(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(\..*)?$') { $clean = '_' + $clean }
    if ($clean.Length -gt 60) { $clean = $clean.Substring(0, 60).Trim() }
    return $clean
}

function Get-Camelot([string]$Key) {
    if ([string]::IsNullOrWhiteSpace($Key)) { return '' }
    $k = $Key.Trim() -replace '\u266F', '#' -replace '\u266D', 'b'
    # already Camelot: "8A", "08a"
    if ($k -match '^0?([1-9]|1[0-2])\s*([AB])$') { return ('{0}{1}' -f [int]$Matches[1], $Matches[2].ToUpperInvariant()) }
    # Open Key: "1m" = 8A, "1d" = 8B
    if ($k -match '^([1-9]|1[0-2])\s*([md])$') {
        $n = (([int]$Matches[1] + 6) % 12) + 1
        $side = if ($Matches[2] -eq 'm') { 'A' } else { 'B' }
        return "$n$side"
    }
    # "A minor", "Amin", "C major", "Cmaj" -> "Am", "C"
    $k = $k -replace '(?i)\s*(minor|min)$', 'm' -replace '(?i)\s*(major|maj)$', '' -replace '\s+', ''
    $map = @{
        'B'='1B'; 'F#'='2B'; 'Gb'='2B'; 'Db'='3B'; 'C#'='3B'; 'Ab'='4B'; 'G#'='4B'; 'Eb'='5B'; 'D#'='5B'
        'Bb'='6B'; 'A#'='6B'; 'F'='7B'; 'C'='8B'; 'G'='9B'; 'D'='10B'; 'A'='11B'; 'E'='12B'
        'Abm'='1A'; 'G#m'='1A'; 'Ebm'='2A'; 'D#m'='2A'; 'Bbm'='3A'; 'A#m'='3A'; 'Fm'='4A'; 'Cm'='5A'
        'Gm'='6A'; 'Dm'='7A'; 'Am'='8A'; 'Em'='9A'; 'Bm'='10A'; 'F#m'='11A'; 'Gbm'='11A'; 'C#m'='12A'; 'Dbm'='12A'
    }
    if ($map.ContainsKey($k)) { return $map[$k] }
    return ''
}

function Get-RoundedBpm([string]$BpmText) {
    $bpmValue = 0.0
    if (-not [double]::TryParse(([string]$BpmText).Trim(), [Globalization.NumberStyles]::Float, $Inv, [ref]$bpmValue)) { return $null }
    if ($bpmValue -le 0 -or [double]::IsNaN($bpmValue) -or [double]::IsInfinity($bpmValue)) { return $null }
    return [int][math]::Round($bpmValue, 0, [MidpointRounding]::AwayFromZero)
}

function Get-TrackKey([string]$Source, [string]$Id) {
    $full = Get-FullPathOrNull $Source
    if ($full) { return $full.ToLowerInvariant() }
    return ('id:' + $Id)
}

# =============================================================================
# Folder layout
# =============================================================================

# Open Key number for a Camelot number: 8 (Am) -> 1 (1m).
function Get-OpenKeyNumber([int]$Camelot) { return (($Camelot + 4) % 12) + 1 }

# One key folder holds neighbouring numbers (A and B together, they mix). With 3 per folder:
# "Keys 01-03", "Keys 04-06", ... With 1 per folder: "08 - 8A 8B".
function Get-KeyLabel([string]$Camelot) {
    if ($Camelot -notmatch '^(\d{1,2})([AB])$') { return 'No key' }
    $c = [int]$Matches[1]
    $size = ConvertTo-IntOrDefault $script:Settings.KeyGroupSize 3 1 6
    $notation = [string]$script:Settings.KeyNotation
    $num = if ($notation -eq 'OpenKey') { Get-OpenKeyNumber $c } else { $c }
    if ($size -le 1) {
        switch ($notation) {
            'OpenKey' { return ('{0:D2} - {0}m {0}d' -f $num) }
            'Musical' { return ('{0:D2} - {1} {2}' -f $c, $MinorNames[$c], $MajorNames[$c]) }
            default   { return ('{0:D2} - {0}A {0}B' -f $c) }
        }
    }
    $lo = [int]([math]::Floor(($num - 1) / $size) * $size) + 1
    $hi = [math]::Min($lo + $size - 1, 12)
    if ($notation -eq 'Musical') { return ('Keys {0:D2}-{1:D2} {2}-{3}' -f $lo, $hi, $MinorNames[$lo], $MinorNames[$hi]) }
    return ('Keys {0:D2}-{1:D2}' -f $lo, $hi)
}

function Test-InRange([int]$Value, [string]$Range) {
    $parts = @($Range.Trim() -split '-')
    if ($parts.Count -ne 2) { return $false }
    $lo = 0; $hi = 0
    if (-not [int]::TryParse($parts[0], [ref]$lo) -or -not [int]::TryParse($parts[1], [ref]$hi)) { return $false }
    return ($Value -ge $lo -and $Value -le $hi)
}

# Rekordbox sometimes detects DnB at half speed (87 instead of 174). If the doubled value fits
# the genre's main tempo and the detected one does not, the doubled value is used for folders.
function Get-EffectiveBpm($Bpm, [string]$Genre) {
    $b = Get-RoundedBpm ([string]$Bpm)
    if ($null -eq $b) { return $null }
    if (-not (Test-Switch 'HalfTimeFix')) { return $b }
    $ranges = @((Get-GenreBpmRanges $Genre) -split ',' | Where-Object { $_.Trim() })
    if ($ranges.Count -eq 0) { return $b }
    if (-not (Test-InRange $b $ranges[0]) -and (Test-InRange ($b * 2) $ranges[0])) { return ($b * 2) }
    return $b
}

# DnB with groups '170,176' -> "0-169 BPM", "170-175 BPM", "176+ BPM".
# Families without groups use bands of BpmBandWidth: "120-124 BPM".
function Get-BpmLabel($Bpm, [string]$Genre) {
    $b = Get-EffectiveBpm $Bpm $Genre
    if ($null -eq $b) { return 'No BPM' }
    $family = Get-GenreFamily $Genre
    $groupText = ''
    foreach ($k in @($script:Settings.BpmGroups.Keys)) {
        if ([string]$k -ieq $family) { $groupText = [string]$script:Settings.BpmGroups[$k] }
    }
    $bounds = @($groupText -split '[,;\s]+' | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ } | Sort-Object -Unique)
    if ($bounds.Count -gt 0) {
        if ($b -lt $bounds[0]) { return ('0-{0} BPM' -f ($bounds[0] - 1)) }
        for ($i = 0; $i -lt $bounds.Count - 1; $i++) {
            if ($b -lt $bounds[$i + 1]) { return ('{0}-{1} BPM' -f $bounds[$i], ($bounds[$i + 1] - 1)) }
        }
        return ('{0}+ BPM' -f $bounds[-1])
    }
    $w = ConvertTo-IntOrDefault $script:Settings.BpmBandWidth 5 1 50
    $lo = [int]([math]::Floor($b / $w) * $w)
    return ('{0}-{1} BPM' -f $lo, ($lo + $w - 1))
}

function Get-GenreFolderName([string]$Genre) {
    if ($script:Settings.DetailLevel -eq 'Broad') { return (Format-FolderName (Get-GenreFamily $Genre)) }
    return (Format-FolderName $Genre)
}

function Get-TargetFolder($Item) {
    $g = [string]$Item.Genre
    if (-not $g -or $g -in @('Review', $ReviewFolder)) { return $ReviewFolder }
    if ($g -in @($SamplesGenre, $SamplesFolder)) { return $SamplesFolder }
    $gf = Get-GenreFolderName $g
    switch ([string]$script:Settings.FolderLayout) {
        'Genre'       { return $gf }
        'GenreKey'    { return (Join-Path $gf (Get-KeyLabel ([string]$Item.Camelot))) }
        'GenreBpmKey' { return (Join-Path $gf (Join-Path (Get-BpmLabel $Item.BPM $g) (Get-KeyLabel ([string]$Item.Camelot)))) }
        default       { return (Join-Path $gf ('{0} - {1}' -f (Get-BpmLabel $Item.BPM $g), (Get-KeyLabel ([string]$Item.Camelot)))) }
    }
}

function Set-PlanDestination($Item) {
    $Item.BPMRounded = Get-RoundedBpm ([string]$Item.BPM)
    $Item.DestinationFolder = Get-TargetFolder $Item
    $Item.DestinationPath = ''
    $Item.PathTooLong = $false
    $fileName = if ($Item.SourcePath) { [IO.Path]::GetFileName([string]$Item.SourcePath) } else { '' }
    if (-not $fileName -or -not $script:OutputRoot) { return }
    $dest = Get-ShortPath (Join-Path $script:OutputRoot $Item.DestinationFolder) $fileName
    if ($dest) { $Item.DestinationPath = $dest } else { $Item.PathTooLong = $true }
}

function Show-FolderCounts($Items) {
    $Items | Group-Object { ([string]$_.DestinationFolder -split '\\')[0] } | Sort-Object Name | ForEach-Object {
        $sub = @($_.Group | Select-Object -ExpandProperty DestinationFolder -Unique).Count
        Write-Host ('  {0,-28} {1,5} tracks in {2,3} folder(s)' -f $_.Name, $_.Count, $sub)
    }
}

# "Genre\BPM + key (one level), 3 keys per folder, Camelot - e.g. Liquid DnB\170-175 BPM - Keys 07-09"
function Get-LayoutSummary {
    $s = $script:Settings
    $example = Get-TargetFolder ([pscustomobject]@{ Genre = 'Liquid DnB'; BPM = '174'; Camelot = '8A' })
    $keys = if ([int]$s.KeyGroupSize -le 1) { '1 key per folder' } else { "$($s.KeyGroupSize) keys per folder" }
    return ('{0}, {1}, {2} - e.g. {3}' -f $LayoutNames[[string]$s.FolderLayout], $keys, $s.KeyNotation, $example)
}

# =============================================================================
# Decisions (your answers, remembered between runs)
# =============================================================================

function Get-Decisions {
    $d = @{}
    if ($script:DecisionsPath -and (Test-Path -LiteralPath $script:DecisionsPath)) {
        foreach ($r in (Import-SafeCsv $script:DecisionsPath)) {
            if (-not $r.TrackKey) { continue }
            $d[[string]$r.TrackKey] = [pscustomobject]@{
                TrackKey = [string]$r.TrackKey; Artist = [string]$r.Artist; Title = [string]$r.Title
                Genre = [string]$r.Genre; RejectedGenres = [string]$r.RejectedGenres
                Source = [string]$r.Source; DecidedAt = [string]$r.DecidedAt
            }
        }
    }
    return $d
}

function Set-Decision {
    param([hashtable]$Decisions, $Row, [string]$Genre = '', [string]$Rejected = '', [string]$Source = '')
    $key = [string]$Row.TrackKey
    if (-not $key) { return }
    $d = $Decisions[$key]
    if (-not $d) {
        $d = [pscustomobject]@{ TrackKey = $key; Artist = [string]$Row.Artist; Title = [string]$Row.Title; Genre = ''; RejectedGenres = ''; Source = ''; DecidedAt = '' }
        $Decisions[$key] = $d
    }
    if ($Genre) { $d.Genre = $Genre; $d.Source = $Source }
    if ($Rejected) {
        $list = @(([string]$d.RejectedGenres -split ';') | Where-Object { $_ })
        if ($list -notcontains $Rejected) { $list += $Rejected }
        $d.RejectedGenres = $list -join ';'
        if (-not $d.Genre) { $d.Source = $Source }
    }
    $d.DecidedAt = (Get-Date).ToString('s')
    $script:DecisionsChanged = $true
}

function Test-Rejected([hashtable]$Decisions, $Row, [string]$Genre) {
    $d = $Decisions[[string]$Row.TrackKey]
    if (-not $d) { return $false }
    return (@(([string]$d.RejectedGenres -split ';')) -contains $Genre)
}

function Save-Decisions([hashtable]$Decisions) {
    if (-not $script:DecisionsChanged) { return }
    if (-not (Test-Path -LiteralPath $script:DataRoot)) { New-Item -ItemType Directory -Force -Path $script:DataRoot | Out-Null }
    Save-CsvAtomic @($Decisions.Values | Sort-Object Artist, Title) $script:DecisionsPath
    $script:DecisionsChanged = $false
}

# Used when a genre is renamed or merged in the Genres menu.
function Update-DecisionGenre([string]$Old, [string]$New) {
    if (-not $script:DecisionsPath -or -not (Test-Path -LiteralPath $script:DecisionsPath)) { return 0 }
    $all = Get-Decisions
    $n = 0
    foreach ($d in $all.Values) { if ($d.Genre -ieq $Old) { $d.Genre = $New; $n++ } }
    if ($n -gt 0) { Save-CsvAtomic @($all.Values | Sort-Object Artist, Title) $script:DecisionsPath }
    return $n
}

# =============================================================================
# History (what was copied where)
# =============================================================================

function Get-History {
    $h = @{}
    if (Test-Path -LiteralPath $script:HistoryPath) {
        foreach ($r in (Import-SafeCsv $script:HistoryPath)) {
            $key = if ($r.PSObject.Properties['TrackKey'] -and $r.TrackKey) { [string]$r.TrackKey }
                   else { Get-TrackKey ([string]$r.SourcePath) ([string]$r.TrackID) }
            $h[$key] = [pscustomobject]@{
                TrackKey = $key; TrackID = [string]$r.TrackID; Artist = [string]$r.Artist; Title = [string]$r.Title
                SourcePath = [string]$r.SourcePath; CopiedPath = [string]$r.CopiedPath; CopiedAt = [string]$r.CopiedAt
            }
        }
    }
    return $h
}

# =============================================================================
# Input: Rekordbox XML  (an M3U / Serato / Traktor reader would return the same fields)
# =============================================================================

function Get-TracksFromRekordboxXml([string]$Path) {
    $settings = New-Object System.Xml.XmlReaderSettings
    $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit   # no external entities
    $settings.XmlResolver = $null
    $stream = [IO.File]::OpenRead($Path)
    try {
        $reader = [System.Xml.XmlReader]::Create($stream, $settings)
        $xml = New-Object System.Xml.XmlDocument
        $xml.XmlResolver = $null
        $xml.Load($reader)
        $reader.Close()
    }
    finally { $stream.Dispose() }

    $nodes = @($xml.SelectNodes('/DJ_PLAYLISTS/COLLECTION/TRACK'))
    if ($nodes.Count -eq 0) { throw 'No tracks found at /DJ_PLAYLISTS/COLLECTION/TRACK. Please use a Rekordbox XML export.' }
    foreach ($t in $nodes) {
        [pscustomobject]@{
            TrackID    = $t.GetAttribute('TrackID')
            Artist     = $t.GetAttribute('Artist')
            Title      = $t.GetAttribute('Name')
            Genre      = $t.GetAttribute('Genre')
            BPM        = $t.GetAttribute('AverageBpm')
            Key        = $t.GetAttribute('Tonality')
            SourcePath = Convert-FileUriToPath $t.GetAttribute('Location')
        }
    }
}

# =============================================================================
# Text helpers for matching
# =============================================================================

function Get-Norm([string]$s) {
    if (-not $s) { return '' }
    $s = $s.Normalize([Text.NormalizationForm]::FormD) -replace '\p{Mn}', ''
    $s = $s.ToLowerInvariant() -replace '&', ' and '
    $s = $s -replace '[^a-z0-9]+', ' '
    return ($s -replace '\s+', ' ').Trim()
}

function Get-FirstArtist([string]$Artist) {
    return ($Artist -split '\s+x\s+|\s*,\s*|\s*/\s*|\s+(?:feat\.?|ft\.?|vs\.?)\s+')[0].Trim()
}

function Get-ArtistNames([string]$Artist) {
    if (-not $Artist) { return @() }
    $names = @($Artist)
    $names += @($Artist -split '\s*(?:,|&|/|\band\b|\bx\b|\bvs\.?|\bfeat\.?|\bft\.?)\s*' | Where-Object { $_ })
    return @($names | ForEach-Object { Get-Norm $_ } | Where-Object { $_ } | Select-Object -Unique)
}

function Test-IsChannel([string]$Name) {
    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    if ($script:ChannelNorms -contains (Get-Norm $Name)) { return $true }
    return ($Name.Trim() -match $ChannelPattern)
}

# "DnB Allstars - Gino - Du Du Du" -> "Gino - Du Du Du"
function Remove-ChannelPrefix([string]$Text) {
    for ($i = 0; $i -lt 3; $i++) {
        if ($Text -notmatch '^(.+?)\s+-\s+(.+)$') { break }
        $head = $Matches[1]
        $tail = $Matches[2].Trim()
        if (-not (Test-IsChannel $head)) { break }
        $Text = $tail
    }
    return $Text
}

function ConvertTo-LuceneText([string]$s) {
    return ($s -replace '([+\-&|!(){}\[\]^"~*?:\\/])', '\$1')
}

# =============================================================================
# Genre rules: tag -> genre, genre -> family / BPM
# =============================================================================

# Rekordbox tags, folder names and online tags all go through the same rules, so
# "Drum & Bass", "Drum and Bass" and "DnB" all become "DnB". Junk tags return nothing.
function Get-RuleForTag([string]$Text) {
    $n = Get-Norm $Text
    if (-not $n -or $script:IgnoredGenres.ContainsKey($n)) { return $null }
    foreach ($rule in $script:GenreRules) {
        if (@($rule.NormAliases) -contains $n) { return $rule }    # exact other spellings ("neuro", "psy")
        foreach ($p in $rule.Patterns) { if ($n -match $p) { return $rule } }
    }
    return $null
}

# $true when the tag is just another way to write the genre's name ("Drum & Bass" for DnB),
# $false when it says more ("Minimal Drum & Bass" - a possible new subgenre).
function Test-ExactGenreName($Rule, [string]$Text) {
    $n = Get-Norm $Text
    if ($n -eq (Get-Norm $Rule.Name)) { return $true }
    if (@($Rule.NormAliases) -contains $n) { return $true }
    foreach ($p in $Rule.Patterns) { if ($n -match ('^(?:' + $p + ')$')) { return $true } }
    return $false
}

function Get-StandardGenre([string]$Name) {
    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }
    $rule = Get-RuleForTag $Name
    if ($rule) { return $rule.Name }
    return (Format-FolderName $Name)
}

function Get-GenreRule([string]$Genre) {
    if (-not $Genre) { return $null }
    return $script:RuleByName[$Genre]
}

function Get-GenreFamily([string]$Genre) {
    $r = Get-GenreRule $Genre
    if ($r -and $r.Family) { return [string]$r.Family }
    return $Genre
}

# A subgenre without its own BPM ranges uses its family's ranges.
function Get-GenreBpmRanges([string]$Genre) {
    $r = Get-GenreRule $Genre
    if ($r -and $r.Bpm) { return [string]$r.Bpm }
    $fr = Get-GenreRule (Get-GenreFamily $Genre)
    if ($fr -and $fr.Bpm) { return [string]$fr.Bpm }
    return ''
}

function Add-GenreRule([string]$Name, [string]$Family = '', [string]$Bpm = '') {
    $Name = ([string]$Name).Trim()
    if (-not $Name) { return }
    if ($Name -ceq $Name.ToLowerInvariant()) { $Name = (Get-Culture).TextInfo.ToTitleCase($Name) }
    if (Get-GenreRule $Name) { return }
    if (-not $Family) { $Family = $Name }
    $new = [pscustomobject]@{ Name = $Name; Family = $Family.Trim(); Patterns = @(New-GenrePattern $Name); Aliases = @(); Bpm = $Bpm }
    # New genres go to the top: their pattern only matches the exact name, so nothing else is shadowed.
    Set-GenreRules (@($new) + @($script:GenreRules)) $script:IgnoredRaw $script:DeclinedRaw
}

# =============================================================================
# New genres seen in your tags or online
# =============================================================================

function Add-DiscoveredGenre([string]$Name, [string]$From, [string]$Family = '') {
    $n = Get-Norm $Name
    if (-not $n -or $script:IgnoredGenres.ContainsKey($n) -or $script:DeclinedGenres.ContainsKey($n)) { return }
    foreach ($r in $script:GenreRules) { if ((Get-Norm $r.Name) -eq $n) { return } }
    $d = $script:Discovered[$n]
    if (-not $d) {
        $d = [pscustomobject]@{ Name = $Name.Trim(); Tracks = 0; From = $From; Family = $Family }
        $script:Discovered[$n] = $d
    }
    $d.Tracks = $d.Tracks + 1
    if ($Family -and -not $d.Family) { $d.Family = $Family }
    if ($d.From -notlike "*$From*") { $d.From = $d.From + ' + ' + $From }
}

# A Rekordbox tag that is not in the genre list, or says more than a known genre
# ("Minimal Drum & Bass"), is offered as a new genre. Its family is guessed from the name.
function Register-TagDiscovery([string]$Raw) {
    $n = Get-Norm $Raw
    if (-not $n -or $n -eq 'music' -or $script:IgnoredGenres.ContainsKey($n)) { return }
    $rule = Get-RuleForTag $Raw
    if ($rule) {
        if (Test-ExactGenreName $rule $Raw) { return }
        Add-DiscoveredGenre $Raw 'your tags' $rule.Family
    }
    else { Add-DiscoveredGenre $Raw 'your tags' '' }
}

# Returns $true when at least one genre was added.
function Invoke-NewGenreReview {
    if ($script:Unattended -or -not (Test-Switch 'AskNewGenres')) { return $false }
    $items = @($script:Discovered.Values | Sort-Object Tracks -Descending)
    if ($items.Count -eq 0) { return $false }
    $show = @($items | Select-Object -First 20)
    Write-Host ''
    Write-Host 'New genres seen that are not in your genre list yet:' -ForegroundColor Cyan
    $i = 0
    foreach ($d in $show) {
        $i++
        $fam = if ($d.Family) { $d.Family } else { $d.Name }
        Write-Host ('  {0,2}) {1,-30} {2,5} tracks   family: {3,-14} from {4}' -f $i, (Limit-Text $d.Name 30), $d.Tracks, $fam, $d.From)
    }
    if ($items.Count -gt $show.Count) { Write-Host "  ...and $($items.Count - $show.Count) more with fewer tracks (shown next time)." -ForegroundColor DarkGray }
    Write-Host '  Added genres get their own folder; the family decides BPM folders and the Broad detail level.' -ForegroundColor DarkGray
    Write-Host '  A = add all   R = choose one by one   S = skip for now (asked again)   N = never suggest these'
    $c = Read-Choice 'Choice [A/R/S/N]' @('a', 'r', 's', 'n') 's'
    $added = 0
    if ($c -eq 'a') {
        foreach ($d in $show) { Add-GenreRule $d.Name $d.Family; $added++ }
    }
    elseif ($c -eq 'n') {
        $script:DeclinedRaw = @($script:DeclinedRaw) + @($show | ForEach-Object { $_.Name })
    }
    elseif ($c -eq 'r') {
        Write-Host '  y = add   f = add with another family   n = never suggest   s = skip for now' -ForegroundColor DarkGray
        foreach ($d in $show) {
            $fam = if ($d.Family) { $d.Family } else { $d.Name }
            $a = Read-Choice ("  Add '{0}' (family {1})? [Y/f/n/s]" -f $d.Name, $fam) @('y', 'f', 'n', 's') 'y'
            switch ($a) {
                'y' { Add-GenreRule $d.Name $d.Family; $added++ }
                'f' {
                    $f = (Read-Host '  Family (for example DnB, House, Techno, or the genre itself)').Trim()
                    Add-GenreRule $d.Name $f
                    $added++
                }
                'n' { $script:DeclinedRaw = @($script:DeclinedRaw) + $d.Name }
            }
        }
    }
    Update-GenreRuleSets
    Save-GenreRules
    $script:Discovered = @{}
    if ($added -gt 0) { Write-Host "$added genre(s) added. Change them any time under Settings > Genres." -ForegroundColor Green }
    return ($added -gt 0)
}

# =============================================================================
# Search terms
# =============================================================================

# "Artist - Title.mp3", "01 Artist - Title.mp3", or a web-rip slug such as
# "ytmp3free.cc_artist-track-name-free-download-youtubemp3free.org.mp3".
function Get-NameFromFile([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    $n = [IO.Path]::GetFileNameWithoutExtension($Path)
    $n = $n -replace '\s*\(\d+\)$', ''                                        # "Track (1)" duplicate suffix
    $n = $n -replace '(?i)^(ytmp3free\.cc|y2mate\.com|ytmp3\.cc)[_\s-]*', ''
    $n = $n -replace '(?i)[-_\s]*youtubemp3(free)?(\.org|\.com)?$', ''
    $n = $n -replace '[\u2013\u2014]', '-'
    $n = $n -replace '^0\d\s+', ''                                             # "01 Artist - Title"
    $n = $n -replace '^\d{1,3}\s*[-._)]\s+', ''                                # "1. Artist - Title"
    for ($i = 0; $i -lt 5; $i++) {                                             # trailing web-rip junk
        $before = $n
        $n = $n -replace '(?i)\s*[\(\[][^\)\]]*(official|video|audio|lyric|free download|out now|premiere|kbps)[^\)\]]*[\)\]]\s*$', ''
        $n = $n -replace '(?i)[-_\s]+(free[-_\s]?download|official[-_\s]*(music[-_\s]*)?(video|audio)|out[-_\s]now|lyric[-_\s]video|visuali[sz]er)$', ''
        if ($n -eq $before) { break }
    }
    $n = $n -replace '^[\s_-]+|[\s_-]+$', ''
    if ([string]::IsNullOrWhiteSpace($n)) { return $null }
    if ($n -match '^(.+?)\s+-\s+(.+)$') {
        return [pscustomobject]@{ Artist = $Matches[1].Trim(); Title = $Matches[2].Trim(); Slug = '' }
    }
    if ($n -match '^(.+?)_-_(.+)$') {
        return [pscustomobject]@{ Artist = ($Matches[1] -replace '_', ' ').Trim(); Title = ($Matches[2] -replace '_', ' ').Trim(); Slug = '' }
    }
    if (($n -notmatch '\s') -and ($n -match '^[^-]+(-[^-]+)+$')) {
        $words = @(($n -replace '_', '-') -split '-' | Where-Object { $_ })
        if ($words.Count -ge 2) { return [pscustomobject]@{ Artist = ''; Title = ''; Slug = (($words -join ' ').ToLowerInvariant()) } }
    }
    return $null
}

# Keeps meaningful brackets like "(Till I Come)", drops junk, and recognises remixes.
function New-Terms([string]$Artist, [string]$Title, [string]$From) {
    $a = ($Artist -replace '["\\]', '').Trim()
    $t = ([string]$Title -replace '["\\]', '').Trim()
    $t = $t -replace '(?i)\s+(?:4k\s+)?(?:visuali[sz]er|official\s+(?:music\s+)?(?:video|audio)|lyrics?\s+video|out\s+now|free\s+download)\s*$', ''
    $t = $t -replace '(?i)\s+(?:feat\.?|ft\.?|featuring)\s+[^\(\[]*', ' '

    $remixer = ''
    $keep = @()
    foreach ($m in [regex]::Matches($t, '[\(\[]([^\)\]]*)[\)\]]')) {
        $inner = $m.Groups[1].Value.Trim()
        if (-not $inner) { continue }
        if ($inner -match '(?i)^(?:feat\.?|ft\.?|featuring)\s') { continue }
        if ($inner -match $JunkInBrackets) { continue }
        if ($inner -match '(?i)^(?:remix|bootleg|edit|flip|clip|rework)$') { continue }
        if ($inner -match '(?i)^(.+?)\s+(?:remix|bootleg|flip|rework|refix|edit|mix)$') { $remixer = $Matches[1].Trim(); continue }
        $keep += $inner
    }
    $base = ($t -replace '\s*[\(\[][^\)\]]*[\)\]]', '').Trim()
    if ($keep.Count -gt 0) { $base = $base + ' (' + ($keep -join ') (') + ')' }
    if (-not $a -or -not $base) { return $null }
    return [pscustomobject]@{ Artist = $a; Title = $base; Slug = ''; From = $From; Remixer = $remixer }
}

# Channel names in the Artist field are ignored. "ATB - Topic" is YouTube's own channel
# for the artist ATB, so only the " - Topic" part is removed.
function Get-SearchTerms([string]$Artist, [string]$Title, [string]$Path = '') {
    $a = ([string]$Artist -replace '\s*-\s*Topic$', '').Trim()
    if (Test-IsChannel $a) { $a = '' }
    $t = (([string]$Title).Trim() -replace '[\u2013\u2014]', '-')

    # 1) the tags
    if ($t -and $t -notmatch '(?i)ytmp3|youtubemp3|y2mate') {
        $t = Remove-ChannelPrefix $t
        if ($t -match '^(.+?)\s+-\s+(.+)$') { $a = $Matches[1].Trim(); $t = $Matches[2].Trim() }
        if ($a) {
            $terms = New-Terms $a $t 'Tags'
            if ($terms) { return $terms }
        }
    }
    # 2) the file name
    $f = Get-NameFromFile $Path
    if ($f) {
        if ($f.Artist) {
            $fa = $f.Artist; $ft = $f.Title
            if (Test-IsChannel $fa) {
                $rest = Remove-ChannelPrefix ("$fa - $ft")
                if ($rest -match '^(.+?)\s+-\s+(.+)$') { $fa = $Matches[1].Trim(); $ft = $Matches[2].Trim() } else { $fa = '' }
            }
            if ($fa) {
                $terms = New-Terms $fa $ft 'Filename'
                if ($terms) { return $terms }
            }
        }
        elseif ($f.Slug -and (Test-Switch 'SlugSplitting')) {
            return [pscustomobject]@{ Artist = ''; Title = ''; Slug = $f.Slug; From = 'Filename'; Remixer = '' }
        }
    }
    return $null
}

function Get-CacheKey($Terms) {
    if ($Terms.Artist -and $Terms.Title) {
        $k = '{0}|{1}' -f $Terms.Artist, $Terms.Title
        if ($Terms.Remixer) { $k += '|remix:' + $Terms.Remixer }
        return $k.ToLowerInvariant()
    }
    return ('slug|' + $Terms.Slug).ToLowerInvariant()
}

# Tracks with no usable genre tag that sit in <music folder>\<Folder>\track.mp3.
function Get-FolderHint([string]$Source) {
    $musicRoot = Get-FullPathOrNull ([string]$script:Settings.MusicRoot)
    if (-not $musicRoot -or [string]::IsNullOrWhiteSpace($Source)) { return $null }
    $dir = Get-FullPathOrNull (Split-Path -Parent $Source)
    if (-not $dir) { return $null }
    $parent = Get-FullPathOrNull (Split-Path -Parent $dir)
    if (-not $parent) { return $null }
    foreach ($nf in @($script:Settings.NewTracksFolders)) { if ((Get-FullPathOrNull $nf) -ieq $dir) { return $null } }
    if ($script:OutputRoot -and (($dir -ieq $script:OutputRoot) -or (Test-InsideFolder $dir $script:OutputRoot))) { return $null }
    $leaf = Split-Path -Leaf $dir
    if ($leaf -match '^(?i)(music|rekordbox|downloads?|new|unsorted|misc|various|temp)$') { return $null }
    $roots = @($musicRoot, (Join-Path $musicRoot 'Music'))
    foreach ($r in $roots) { if ($parent -ieq $r) { return $leaf } }
    return $null
}

# =============================================================================
# Lookup cache
# =============================================================================

# Values look like "drum and bass=3;liquid funk=12 || evidence": the raw tags each service
# returned, so changes to your genre list apply to cached lookups too.
function Get-TagCache {
    $c = @{}
    $j = Read-JsonFile $CachePath
    if ($j) { foreach ($p in $j.PSObject.Properties) { $c[$p.Name] = [string]$p.Value } }
    return $c
}

function Save-TagCache([hashtable]$Cache) { Save-JsonAtomic $Cache $CachePath -Compress }

function ConvertTo-CacheText($Res) {
    $pairs = @(foreach ($t in @($Res.Tags)) {
        if (-not $t -or -not $t.Name) { continue }
        $w = 1.0
        if ($null -ne $t.Weight -and "$($t.Weight)" -ne '') { $w = [double]$t.Weight }
        (([string]$t.Name) -replace '[;=|]', ' ').Trim() + '=' + $w.ToString('0.##', $Inv)
    })
    return (($pairs -join ';') + ' || ' + ([string]$Res.Evidence -replace '\|\|', '|'))
}

function ConvertFrom-CacheText([string]$Text) {
    $tags = @()
    $parts = ([string]$Text) -split ' \|\| ', 2
    foreach ($p in ($parts[0] -split ';')) {
        if ($p -match '^(.+)=(-?[0-9.]+)$') { $tags += [pscustomobject]@{ Name = $Matches[1]; Weight = [double]::Parse($Matches[2], $Inv) } }
    }
    $ev = if ($parts.Count -gt 1) { $parts[1] } else { '' }
    return [pscustomobject]@{ Tags = $tags; Evidence = $ev }
}

# =============================================================================
# Online lookups
# =============================================================================

# Returns the parsed response, or $null if the service could not be reached after several tries.
# Error messages never include the request URL (the Last.fm key is part of it).
function Invoke-Api([string]$Source, [string]$Url, [hashtable]$Headers, [int]$DelayMs) {
    $backoff = 4
    for ($attempt = 1; $attempt -le 4; $attempt++) {
        Start-Sleep -Milliseconds $DelayMs
        try {
            $result = Invoke-RestMethod -Uri $Url -Headers $Headers -TimeoutSec 30 -UseBasicParsing -ErrorAction Stop
            $script:ApiFailures[$Source] = 0
            return $result
        }
        catch {
            $status = 0
            try { $status = [int]$_.Exception.Response.StatusCode } catch { }
            $script:LastApiError[$Source] = if ($status) { "HTTP $status" } else { $_.Exception.GetType().Name }
            if ($status -eq 404) { $script:ApiFailures[$Source] = 0; return [pscustomobject]@{ NotFound = $true } }
            if ($status -eq 401 -or $status -eq 403) { break }
            if ($attempt -lt 4) { Start-Sleep -Seconds $backoff; $backoff = $backoff * 2 }
        }
    }
    $script:ApiFailures[$Source] = 1 + [int]$script:ApiFailures[$Source]
    Write-Warning "$Source not answering ($($script:LastApiError[$Source])); skipping this track for now."
    return $null
}

function Invoke-MusicBrainz([string]$Query, [string]$Entity) {
    $q = [uri]::EscapeDataString($Query)
    $url = "https://musicbrainz.org/ws/2/${Entity}?query=${q}&fmt=json&limit=10"
    # MusicBrainz allows one request per second.
    return Invoke-Api 'MusicBrainz' $url @{ 'User-Agent' = $script:UserAgent; 'Accept' = 'application/json' } 1100
}

function Test-ArtistMatch([string]$Credit, [string]$Artist) {
    $c = ' ' + (Get-Norm $Credit) + ' '
    foreach ($a in @($Artist, (Get-FirstArtist $Artist))) {
        $n = Get-Norm $a
        if ($n -and $c.Contains(" $n ")) { return $true }
    }
    return $false
}

function Test-TitleMatch([string]$Found, $Terms) {
    $f = Get-Norm $Found
    $t = Get-Norm $Terms.Title
    if (-not $f -or -not $t) { return $false }
    $fs = $f -replace ' ', ''
    $ts = $t -replace ' ', ''
    $same = ($f -eq $t) -or ($fs -eq $ts) -or $f.StartsWith("$t ") -or
            ($fs.StartsWith($ts) -and $f -match '\b(mix|remix|edit|version|vip|bootleg|rework|flip|refix|remaster(ed)?)\b')
    if (-not $same) { return $false }
    $remixWords = '\b(remix|bootleg|rework|flip|refix|cover)\b'
    if ($Terms.Remixer) { return (" $f ").Contains(' ' + (Get-Norm $Terms.Remixer) + ' ') }
    if ($fs -ne $ts -and $f -match $remixWords -and $t -notmatch $remixWords) { return $false }
    return $true
}

# Turns raw tags into scores per genre. Tags that are real genres (MusicBrainz genre list) but
# not in your list yet are returned as Unknown, so they can be offered as new genres.
function Get-GenreScores($Tags, [string]$Artist) {
    $scores = @{}
    $unknown = @()
    $ignore = Get-ArtistNames $Artist
    foreach ($tg in @($Tags)) {
        if (-not $tg -or -not $tg.Name) { continue }
        $n = Get-Norm ([string]$tg.Name)
        if (-not $n -or ($ignore -contains $n) -or $script:IgnoredGenres.ContainsKey($n)) { continue }   # artist names, junk
        $w = 1.0
        if ($null -ne $tg.Weight -and "$($tg.Weight)" -ne '') { $w = [double]$tg.Weight }
        if ($w -le 0) { continue }
        $rule = Get-RuleForTag ([string]$tg.Name)
        if ($rule) { $scores[$rule.Name] = [double]$scores[$rule.Name] + $w }
        elseif ($script:GenreDictionary.ContainsKey($n)) { $unknown += [string]$script:GenreDictionary[$n] }
    }
    return [pscustomobject]@{ Scores = $scores; Unknown = $unknown }
}

function Format-Scores([hashtable]$Scores) {
    $total = 0.0
    foreach ($v in $Scores.Values) { $total += [double]$v }
    if ($total -le 0) { return '' }
    $parts = @($Scores.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
        $_.Key + ' ' + [string][int][math]::Round(100 * [double]$_.Value / $total) + '%'
    })
    return ($parts -join ', ')
}

# The winning FAMILY needs MinWinnerShare of all evidence. Inside it, a subgenre wins when it
# has at least half of the family's evidence; otherwise the family itself is used.
function Select-Genre([hashtable]$Scores) {
    $total = 0.0
    foreach ($v in $Scores.Values) { $total += [double]$v }
    if ($total -le 0) { return $null }
    $fam = @{}
    foreach ($k in @($Scores.Keys)) { $f = Get-GenreFamily $k; $fam[$f] = [double]$fam[$f] + [double]$Scores[$k] }
    $top = $fam.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1
    $share = [double]$top.Value / $total
    if ($share -lt $MinWinnerShare) { return $null }
    $best = $Scores.GetEnumerator() |
        Where-Object { $_.Key -ne $top.Key -and (Get-GenreFamily $_.Key) -eq $top.Key } |
        Sort-Object Value -Descending | Select-Object -First 1
    if ($best -and ([double]$best.Value / [double]$top.Value) -ge 0.5) { return [pscustomobject]@{ Genre = $best.Key; Share = $share } }
    return [pscustomobject]@{ Genre = $top.Key; Share = $share }
}

function Test-GenreBpm([string]$Genre, $Bpm) {
    if (-not (Test-Switch 'BpmCheck')) { return $true }
    $ranges = @((Get-GenreBpmRanges $Genre) -split ',' | Where-Object { $_.Trim() })
    if ($ranges.Count -eq 0) { return $true }
    $b = Get-RoundedBpm ([string]$Bpm)
    if ($null -eq $b) { return $true }
    foreach ($range in $ranges) { if (Test-InRange $b $range) { return $true } }
    return $false
}

function New-SourceResult($Tags, $Evidence) {
    return [pscustomobject]@{ Tags = @($Tags); Evidence = (@($Evidence) | Where-Object { $_ } | Select-Object -First 2) -join '; ' }
}

function Find-GenreMusicBrainz($Terms) {
    $words = $Terms.Title
    if ($Terms.Remixer) { $words += ' ' + $Terms.Remixer }
    $artistNames = @($Terms.Artist)
    $first = Get-FirstArtist $Terms.Artist
    if ($first -and ($first -ine $Terms.Artist)) { $artistNames += $first }

    $tags = @(); $ev = @()
    foreach ($nm in $artistNames) {
        $q = 'recording:({0}) AND artist:"{1}"' -f (ConvertTo-LuceneText $words), ($nm -replace '[\\"]', '')
        $r = Invoke-MusicBrainz $q 'recording'
        if ($null -eq $r) { return $null }
        foreach ($rec in @($r.recordings)) {
            if (-not $rec -or -not $rec.tags) { continue }
            $credit = (@($rec.'artist-credit') | ForEach-Object { [string]$_.name + [string]$_.joinphrase }) -join ''
            if (-not (Test-ArtistMatch $credit $Terms.Artist)) { continue }
            if (-not (Test-TitleMatch ([string]$rec.title) $Terms)) { continue }
            $tags += @($rec.tags | ForEach-Object { [pscustomobject]@{ Name = [string]$_.name; Weight = $_.count } })
            $ev += "recording '$($rec.title)' by $credit"
        }
        if ($tags.Count -gt 0) { break }
    }
    if ($tags.Count -eq 0 -and (Test-Switch 'ArtistFallback') -and -not $Terms.Remixer) {
        $r = Invoke-MusicBrainz ('artist:"{0}"' -f ($first -replace '[\\"]', '')) 'artist'
        if ($null -eq $r) { return $null }
        $hit = @($r.artists) | Where-Object { $_.name -ieq $first } | Select-Object -First 1
        if ($hit -and $hit.tags) {
            $tags = @($hit.tags | ForEach-Object { [pscustomobject]@{ Name = [string]$_.name; Weight = $_.count } })
            $ev += "artist '$($hit.name)' overall tags (guess)"
        }
    }
    return New-SourceResult $tags $ev
}

function Find-GenreDiscogs($Terms) {
    $artist = if ($Terms.Remixer) { $Terms.Remixer } else { $Terms.Artist }
    $a = [uri]::EscapeDataString($artist)
    $t = [uri]::EscapeDataString(($Terms.Title -replace '[\(\)\[\]]', ''))
    $url = "https://api.discogs.com/database/search?type=release&artist=${a}&track=${t}&per_page=10"
    $headers = @{ 'User-Agent' = $script:UserAgent; 'Authorization' = "Discogs token=$($script:Keys.Discogs)" }
    $r = Invoke-Api 'Discogs' $url $headers 1200
    if ($null -eq $r) { return $null }
    $allowed = @((Get-Norm $artist), (Get-Norm (Get-FirstArtist $artist))) | Where-Object { $_ } | Select-Object -Unique
    $tags = @(); $ev = @()
    foreach ($res in @($r.results)) {
        if (-not $res) { continue }
        $rt = Get-Norm ([string]$res.title)    # "Artist - Release"; skips "Various - ..." compilations
        $ok = $false
        foreach ($n in $allowed) { if ($rt.StartsWith("$n ")) { $ok = $true } }
        if (-not $ok) { continue }
        foreach ($n in @(@($res.style) + @($res.genre) | Where-Object { $_ })) { $tags += [pscustomobject]@{ Name = [string]$n; Weight = 1 } }
        $ev += "release '$($res.title)'"
    }
    return New-SourceResult $tags $ev
}

function Find-GenreLastFm($Terms) {
    $k = [uri]::EscapeDataString($script:Keys.LastFm)
    $track = if ($Terms.Remixer) { "$($Terms.Title) ($($Terms.Remixer) Remix)" } else { $Terms.Title }
    $a = [uri]::EscapeDataString($Terms.Artist)
    $t = [uri]::EscapeDataString($track)
    $headers = @{ 'User-Agent' = $script:UserAgent }
    $r = Invoke-Api 'Last.fm' "https://ws.audioscrobbler.com/2.0/?method=track.gettoptags&artist=${a}&track=${t}&autocorrect=1&api_key=${k}&format=json" $headers 300
    if ($null -eq $r) { return $null }
    if ($r.error -in @(10, 26)) {
        $script:DisabledSources['Last.fm'] = $true
        Write-Warning 'Last.fm rejected the API key; skipping Last.fm. Check it under Settings > Lookup services.'
        return New-SourceResult @() ''
    }
    $tags = @($r.toptags.tag | Where-Object { $_ -and [int]$_.count -ge 10 } |
        ForEach-Object { [pscustomobject]@{ Name = [string]$_.name; Weight = [int]$_.count } })
    $ev = @()
    if ($tags.Count -gt 0) { $ev += "track tags for '$track'" }
    if ($tags.Count -eq 0 -and (Test-Switch 'ArtistFallback') -and -not $Terms.Remixer) {
        $r = Invoke-Api 'Last.fm' "https://ws.audioscrobbler.com/2.0/?method=artist.gettoptags&artist=${a}&autocorrect=1&api_key=${k}&format=json" $headers 300
        if ($null -eq $r) { return $null }
        $tags = @($r.toptags.tag | Where-Object { $_ -and [int]$_.count -ge 10 } |
            ForEach-Object { [pscustomobject]@{ Name = [string]$_.name; Weight = [int]$_.count } })
        if ($tags.Count -gt 0) { $ev += 'artist overall tags (guess)' }
    }
    return New-SourceResult $tags $ev
}

function Find-GenreDiscogsText([string]$Slug) {
    $q = [uri]::EscapeDataString($Slug)
    $url = "https://api.discogs.com/database/search?type=release&q=${q}&per_page=5"
    $headers = @{ 'User-Agent' = $script:UserAgent; 'Authorization' = "Discogs token=$($script:Keys.Discogs)" }
    $r = Invoke-Api 'Discogs' $url $headers 1200
    if ($null -eq $r) { return $null }
    $slugWords = @($Slug.ToLowerInvariant() -split '\s+' | Where-Object { $_ })
    $tags = @(); $ev = @()
    foreach ($res in @($r.results)) {
        if (-not $res) { continue }
        $title = [string]$res.title
        if ($title -match '^(?i)various\b') { continue }
        $titleWords = @($title.ToLowerInvariant() -split '[^a-z0-9]+' | Where-Object { $_ })
        if ($titleWords.Count -lt 2) { continue }
        $inSlug = @($titleWords | Where-Object { $slugWords -contains $_ }).Count
        if (($inSlug / $titleWords.Count) -lt 0.9) { continue }
        foreach ($n in @(@($res.style) + @($res.genre) | Where-Object { $_ })) { $tags += [pscustomobject]@{ Name = [string]$n; Weight = 1 } }
        $ev += "release '$title'"
    }
    return New-SourceResult $tags $ev
}

function Get-EnabledSources {
    $s = @()
    if (Test-Switch 'MusicBrainz') { $s += 'MusicBrainz' }
    if ((Test-Switch 'Discogs') -and $script:Keys.Discogs) { $s += 'Discogs' }
    if ((Test-Switch 'LastFm') -and $script:Keys.LastFm) { $s += 'Last.fm' }
    return $s
}

function Invoke-Source([string]$Name, $Terms) {
    switch ($Name) {
        'MusicBrainz' { return Find-GenreMusicBrainz $Terms }
        'Discogs'     { return Find-GenreDiscogs $Terms }
        'Last.fm'     { return Find-GenreLastFm $Terms }
    }
}

function Register-Failure([string]$Name) {
    if ([int]$script:ApiFailures[$Name] -ge 3) {
        $script:DisabledSources[$Name] = $true
        Write-Warning "$Name keeps failing - skipping it for the rest of this run."
    }
}

# $true / $false, or $null when MusicBrainz could not be reached.
function Test-ArtistExists([string]$Name, [hashtable]$Cache) {
    $ck = 'ArtistExists::' + $Name.ToLowerInvariant()
    if (-not $Cache.ContainsKey($ck)) {
        $r = Invoke-MusicBrainz ('artist:"{0}"' -f ($Name -replace '[\\"]', '')) 'artist'
        if ($null -eq $r) { return $null }
        $hit = @($r.artists) | Where-Object { $_.name -ieq $Name } | Select-Object -First 1
        $Cache[$ck] = if ($hit) { '1' } else { '0' }
    }
    return ($Cache[$ck] -eq '1')
}

# A slug such as "acray bad boy cray tunez" does not say where the artist ends: try the longest
# artist name (up to 3 words) that MusicBrainz knows as an exact artist; the rest is the title.
function Resolve-Slug($Terms, [hashtable]$Cache, [hashtable]$Stats) {
    if (-not (Test-Switch 'SlugSplitting') -or -not (Test-Switch 'MusicBrainz') -or $script:DisabledSources['MusicBrainz']) { return }
    $ck = 'Split::' + $Terms.Slug
    if (-not $Cache.ContainsKey($ck)) {
        $words = @($Terms.Slug -split '\s+' | Where-Object { $_ })
        $result = ''
        $max = [math]::Min(3, $words.Count - 1)
        for ($k = $max; $k -ge 1; $k--) {
            $cand = ($words[0..($k - 1)] -join ' ')
            $exists = Test-ArtistExists $cand $Cache
            if ($null -eq $exists) { Register-Failure 'MusicBrainz'; return }
            if ($exists) { $result = $cand + '|' + ($words[$k..($words.Count - 1)] -join ' '); break }
        }
        $Cache[$ck] = $result
        $Stats.NewLookups++
    }
    if ($Cache[$ck]) {
        $parts = ([string]$Cache[$ck]) -split '\|', 2
        $Terms.Artist = $parts[0]
        $Terms.Title = $parts[1]
    }
}

# Asks every enabled service, combines the answers, then checks BPM.
# Returns Genre/Source/Evidence/Rule, or $null (reason written to the log).
function Resolve-Genre($Terms, [hashtable]$Cache, [hashtable]$Stats, $Row) {
    if (-not $Terms.Artist -and $Terms.Slug) { Resolve-Slug $Terms $Cache $Stats }

    $results = @()
    if (-not $Terms.Artist) {
        if (-not $script:Keys.Discogs -or -not (Test-Switch 'Discogs') -or $script:DisabledSources['Discogs']) { return $null }
        $ck = 'DiscogsText::' + $Terms.Slug.ToLowerInvariant()
        if (-not $Cache.ContainsKey($ck)) {
            $res = Find-GenreDiscogsText $Terms.Slug
            if ($null -eq $res) { Register-Failure 'Discogs'; return $null }
            $Cache[$ck] = ConvertTo-CacheText $res
            $Stats.NewLookups++
        }
        $results += [pscustomobject]@{ Name = 'Discogs'; Res = (ConvertFrom-CacheText $Cache[$ck]) }
    }
    else {
        foreach ($name in (Get-EnabledSources)) {
            if ($script:DisabledSources[$name]) { continue }
            $ck = '{0}::{1}' -f $name, (Get-CacheKey $Terms)
            if (-not $Cache.ContainsKey($ck)) {
                $res = Invoke-Source $name $Terms
                if ($null -eq $res) { Register-Failure $name; continue }   # failed, not "no genre": do not cache
                $Cache[$ck] = ConvertTo-CacheText $res
                $Stats.NewLookups++
            }
            $results += [pscustomobject]@{ Name = $name; Res = (ConvertFrom-CacheText $Cache[$ck]) }
        }
    }

    $combined = @{}; $ev = @(); $used = @(); $unknown = @()
    foreach ($x in $results) {
        $sc = Get-GenreScores $x.Res.Tags $Terms.Artist
        $unknown += @($sc.Unknown)
        $sum = 0.0
        foreach ($v in $sc.Scores.Values) { $sum += [double]$v }
        if ($sum -le 0) { continue }
        $w = [double]$SourceWeight[$x.Name]
        if ($w -le 0) { $w = 1.0 }
        foreach ($k in @($sc.Scores.Keys)) { $combined[$k] = [double]$combined[$k] + $w * [double]$sc.Scores[$k] / $sum }
        $used += $x.Name
        $ev += ($x.Name + ': ' + $x.Res.Evidence + ' [' + (Format-Scores $sc.Scores) + ']')
    }
    foreach ($u in @($unknown | Select-Object -Unique)) { Add-DiscoveredGenre $u 'online' '' }
    if ($used.Count -eq 0) { return $null }

    $source = $used -join '+'
    $evidence = $ev -join ' | '
    $rule = 'combined ' + (Format-Scores $combined)
    $pick = Select-Genre $combined
    if (-not $pick) {
        Write-GenreLog $Row 'Online genre lookup' $source '' 'Left in Review: sources disagree' $evidence $rule
        return $null
    }
    if (-not (Test-GenreBpm $pick.Genre $Row.BPM)) {
        Write-GenreLog $Row 'Online genre lookup' $source $pick.Genre ("Rejected: BPM $($Row.BPM) does not fit $($pick.Genre)") $evidence $rule
        return $null
    }
    return [pscustomobject]@{ Genre = $pick.Genre; Source = $source; Evidence = $evidence; Rule = $rule }
}

# MusicBrainz's list of real genre names (about 2,000), refreshed about once a month.
# Used as a dictionary only: it tells real genres apart from junk tags. It never creates folders.
function Update-GenreDictionary([switch]$Force) {
    $script:GenreDictionary = @{}
    if (-not (Test-Switch 'GenreDictionary')) { return }
    $fresh = $false
    $data = Read-JsonFile $DictionaryPath
    if ($data -and $data.Genres) {
        foreach ($g in @($data.Genres)) { $n = Get-Norm ([string]$g); if ($n) { $script:GenreDictionary[$n] = [string]$g } }
        $when = [datetime]::MinValue
        [void][datetime]::TryParse([string]$data.FetchedAt, $Inv, [Globalization.DateTimeStyles]::RoundtripKind, [ref]$when)
        $fresh = ((Get-Date) - $when).TotalDays -lt 30
    }
    if (-not (Test-Switch 'MusicBrainz') -or ($fresh -and -not $Force)) { return }
    Write-Host 'Updating the MusicBrainz genre list (about once a month)...' -ForegroundColor DarkGray
    $r = Invoke-Api 'MusicBrainz' 'https://musicbrainz.org/ws/2/genre/all?fmt=txt' @{ 'User-Agent' = $script:UserAgent; 'Accept' = 'text/plain' } 1100
    if (-not ($r -is [string])) { Write-Warning 'Could not download the MusicBrainz genre list; using the saved one (if any).'; return }
    $names = @($r -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if ($names.Count -lt 100) { Write-Warning 'The MusicBrainz genre list looked incomplete; keeping the saved one.'; return }
    $script:GenreDictionary = @{}
    foreach ($g in $names) { $n = Get-Norm $g; if ($n) { $script:GenreDictionary[$n] = $g } }
    Save-JsonAtomic ([ordered]@{ FetchedAt = (Get-Date).ToString('o'); Source = 'https://musicbrainz.org/ws/2/genre/all?fmt=txt'; Genres = $names }) $DictionaryPath -Compress
    Write-Host "Genre list updated: $($names.Count) genres." -ForegroundColor DarkGray
}

function Set-ProposalGenre($Proposal) {
    $Proposal.Row.Genre = $Proposal.Genre
    $Proposal.Row.GenreSource = $Proposal.Source
}

function Approve-Proposal($Proposal, [string]$How) {
    Set-ProposalGenre $Proposal
    Set-Decision -Decisions $script:Decisions -Row $Proposal.Row -Genre $Proposal.Genre -Source ('{0} ({1})' -f $How, $Proposal.Source)
    Write-GenreLog $Proposal.Row 'Genre suggestion decision' $Proposal.Source $Proposal.Genre $How $Proposal.Evidence $Proposal.Rule
}

# =============================================================================
# Duplicates, free space, playlists
# =============================================================================

# Same artist and title, but different files. Only reported - nothing is deleted.
function Find-Duplicates($Rows) {
    return @($Rows |
        Where-Object { $_.SourceExists -and $_.Artist -and $_.Title } |
        Group-Object { (Get-Norm $_.Artist) + '|' + (Get-Norm $_.Title) } |
        Where-Object { @($_.Group | Select-Object -ExpandProperty SourcePath -Unique).Count -gt 1 })
}

function Test-FreeSpace($Items) {
    if (-not (Test-Switch 'FreeSpaceCheck') -or $script:OutputRoot -notmatch '^[A-Za-z]:') { return $true }
    $need = [int64]0
    foreach ($it in @($Items)) { try { $need += (Get-Item -LiteralPath $it.SourcePath).Length } catch { } }
    try { $free = (New-Object IO.DriveInfo ($script:OutputRoot.Substring(0, 2))).AvailableFreeSpace }
    catch { return $true }
    $margin = [int64]500MB
    if ($need + $margin -le $free) { return $true }
    Write-Warning ('Not enough free space on {0}: the copies need {1:N1} GB, but only {2:N1} GB is free.' -f $script:OutputRoot.Substring(0, 2), ($need / 1GB), ($free / 1GB))
    return (Read-YesNo 'Copy anyway (stops when the drive is full)?' $false $false)
}

# One playlist per destination folder, pointing at your ORIGINAL files. Import them into
# Rekordbox and the tracks keep their cue points, hot cues and beatgrids.
# The _Playlists folder is rewritten on every run.
function Write-Playlists($Rows) {
    if (-not (Test-Switch 'Playlists')) { return 0 }
    $items = @($Rows | Where-Object { $_.SourceExists })
    if ($items.Count -eq 0) { return 0 }
    $groups = @($items | Group-Object DestinationFolder)
    if ($script:Preview) {
        Write-Host "Preview: would write $($groups.Count) playlists into $script:PlaylistRoot"
        return 0
    }
    New-Item -ItemType Directory -Force -Path $script:PlaylistRoot | Out-Null
    $oldDirs = @()
    foreach ($f in @(Get-ChildItem -LiteralPath $script:PlaylistRoot -Recurse -File -Filter '*.m3u8')) {
        if (Test-InsideFolder $f.FullName $script:PlaylistRoot) { Remove-Item -LiteralPath $f.FullName; $oldDirs += $f.DirectoryName }
    }
    Remove-EmptyFolders $oldDirs $script:PlaylistRoot
    $utf8 = New-Object Text.UTF8Encoding $false
    $written = 0
    foreach ($g in $groups) {
        $path = Join-Path $script:PlaylistRoot ($g.Name + '.m3u8')
        if ($path.Length -gt $MaxPathLength -or -not (Test-InsideFolder $path $script:PlaylistRoot)) { continue }
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $path) | Out-Null
        $lines = New-Object System.Collections.Generic.List[string]
        $lines.Add('#EXTM3U')
        # Sorted for mixing: by key number, then A/B, then BPM.
        $sorted = $g.Group | Sort-Object @{ Expression = { if ([string]$_.Camelot -match '^(\d+)') { [int]$Matches[1] } else { 99 } } },
            @{ Expression = { [string]$_.Camelot } },
            @{ Expression = { if ($null -ne $_.BPMRounded) { [int]$_.BPMRounded } else { 999 } } }
        foreach ($t in $sorted) {
            $label = if ($t.Artist) { "$($t.Artist) - $($t.Title)" } else { [string]$t.Title }
            $lines.Add('#EXTINF:-1,' + ($label -replace '[\r\n]', ' '))
            $lines.Add([string]$t.SourcePath)
        }
        [IO.File]::WriteAllLines($path, $lines, $utf8)
        $written++
    }
    return $written
}

# =============================================================================
# One-time import from the earlier organizer script ('<output>\_Organizer')
# =============================================================================

function Import-OldOrganizerData {
    if ($script:Unattended -or $script:Preview -or -not $script:OutputRoot) { return }
    $old = Join-Path $script:OutputRoot '_Organizer'
    if (-not (Test-Path -LiteralPath $old -PathType Container)) { return }
    Write-Host ''
    Write-Host "Found data from the earlier organizer script in: $old" -ForegroundColor Cyan
    Write-Host '  Your decisions, copy history, lookup cache and API keys can be moved into TuneTidy.'
    if (-not (Read-YesNo 'Import them now?' $true)) { return }

    New-Item -ItemType Directory -Force -Path $script:DataRoot | Out-Null
    foreach ($f in @('Decisions.csv', 'Processed.csv')) {
        $src = Join-Path $old $f
        $dst = Join-Path $script:DataRoot $f
        if ((Test-Path -LiteralPath $src) -and -not (Test-Path -LiteralPath $dst)) {
            Copy-Item -LiteralPath $src -Destination $dst
            Write-Host "  imported $f"
        }
    }
    # The old cache stored genre names, which work as tags in the new cache.
    $j = Read-JsonFile (Join-Path $old 'GenreCache-v3.json')
    if ($j) {
        $cache = Get-TagCache
        $n = 0
        foreach ($p in $j.PSObject.Properties) { if (-not $cache.ContainsKey($p.Name)) { $cache[$p.Name] = [string]$p.Value; $n++ } }
        Save-TagCache $cache
        Write-Host "  imported $n cached lookups"
    }
    $oldKeys = Join-Path $old 'ApiKeys.json'
    $j = Read-JsonFile $oldKeys
    if ($j) {
        $saved = Get-SavedKeys
        foreach ($n in @('Discogs', 'LastFm')) {
            $p = $j.PSObject.Properties[$n]
            if ($p -and $p.Value -and -not $saved[$n]) { $saved[$n] = [string]$p.Value }
        }
        Save-SavedKeys $saved
        Remove-Item -LiteralPath $oldKeys
        Write-Host '  imported your API keys (now encrypted) and deleted the old plain-text key file'
    }
    $j = Read-JsonFile (Join-Path $old 'Settings.json')
    if ($j -and $j.MbContact -and [string]$j.MbContact -ne 'none' -and -not $script:Settings.MbContact) {
        $script:Settings.MbContact = [string]$j.MbContact
        Save-Settings
        Update-UserAgent
    }
    if (Test-Path -LiteralPath (Join-Path $old 'ApiKeys.xml')) {
        Write-Host '  Note: the old folder still has ApiKeys.xml with your keys in readable form - delete it.' -ForegroundColor Yellow
    }
    try {
        Rename-Item -LiteralPath $old -NewName '_Organizer-old'
        Write-Host "  The old folder is now '_Organizer-old'. You can delete it once everything works."
    }
    catch { Write-Warning "Could not rename the old folder: $($_.Exception.Message)" }
    Update-ApiKeys
}

# =============================================================================
# Settings menu
# =============================================================================

function Show-SettingsMenu {
    while ($true) {
        $s = $script:Settings
        $sources = @(Get-EnabledSources)
        Write-Host ''
        Write-Host '=== Settings ===' -ForegroundColor Green
        Write-Host ('  1) Folders                        music: {0}' -f $s.MusicRoot)
        Write-Host ('  2) Folder layout                  {0}' -f (Get-LayoutSummary))
        Write-Host ('  3) Genres                         {0} genres, detail level: {1}' -f $script:GenreRules.Count, $s.DetailLevel)
        Write-Host ('  4) Lookup services and API keys   {0}' -f ($(if ($sources.Count) { $sources -join ', ' } else { 'all off' })))
        Write-Host '  5) On/off switches'
        Write-Host '  6) Logs'
        Write-Host '  7) Reset'
        Write-Host '  0) Back'
        switch (Read-Choice 'Choose' @('0', '1', '2', '3', '4', '5', '6', '7') '0') {
            '0' { return }
            '1' { Edit-FolderSettings }
            '2' { Edit-LayoutSettings }
            '3' { Edit-GenreSettings }
            '4' { Edit-ServiceSettings }
            '5' { Edit-Switches }
            '6' { Show-LogsMenu }
            '7' { Invoke-ResetMenu }
        }
    }
}

function Edit-FolderSettings {
    while ($true) {
        $s = $script:Settings
        Write-Host ''
        Write-Host '--- Folders ---' -ForegroundColor Cyan
        Write-Host "  Music folder:      $($s.MusicRoot)"
        Write-Host "  Output folder:     $($s.OutputFolder)"
        $nf = @($s.NewTracksFolders)
        if ($nf.Count -eq 0) { Write-Host '  New-tracks folders: (none)' }
        else { for ($i = 0; $i -lt $nf.Count; $i++) { Write-Host ('  New-tracks folder {0}: {1}' -f ($i + 1), $nf[$i]) } }
        Write-Host "  Rekordbox export:  $(if ($s.XmlPath) { $s.XmlPath } else { '(asked when you run)' })"
        Write-Host ''
        Write-Host '  1) Change music folder   2) Change output folder   3) Add a new-tracks folder'
        Write-Host '  4) Remove a new-tracks folder   5) Change Rekordbox export file   0) Back'
        switch (Read-Choice 'Choose' @('0', '1', '2', '3', '4', '5') '0') {
            '0' { return }
            '1' {
                $s.MusicRoot = Read-FolderPath 'Music folder' $s.MusicRoot $true $false
                Save-Settings
            }
            '2' {
                $new = Read-FolderPath 'Output folder for the sorted copies' $s.OutputFolder $false $false
                $problem = Test-OutputFolderOk $new
                if ($problem) { Write-Host $problem -ForegroundColor Yellow; continue }
                if ($new -ine $script:OutputRoot) {
                    $newData = Join-Path $new $DataFolderName
                    if ((Test-Path -LiteralPath $script:DataRoot) -and -not (Test-Path -LiteralPath $newData) -and
                        (Read-YesNo 'Take your decisions and copy history along to the new folder?' $true)) {
                        New-Item -ItemType Directory -Force -Path $new | Out-Null
                        Copy-Item -LiteralPath $script:DataRoot -Destination $newData -Recurse
                    }
                    $s.OutputFolder = $new
                    Save-Settings
                    Initialize-Paths
                    Write-Host 'Copies already made in the old folder stay there; the next run offers to copy tracks again.' -ForegroundColor DarkGray
                }
            }
            '3' {
                $f = Read-FolderPath 'New-tracks folder (Enter to cancel)' '' $true $true
                if ($f -and (Add-NewTracksFolder $f)) { Save-Settings }
            }
            '4' {
                if ($nf.Count -eq 0) { continue }
                $n = ConvertTo-IntOrDefault (Read-Host 'Number of the folder to remove') 0 0 $nf.Count
                if ($n -ge 1) { $s.NewTracksFolders = @($nf | Where-Object { $_ -ne $nf[$n - 1] }); Save-Settings }
            }
            '5' {
                $f = Read-FilePath 'Full path to the Rekordbox XML export (Enter to keep)'
                if ($f) { $s.XmlPath = $f; Save-Settings }
            }
        }
    }
}

function Edit-LayoutSettings {
    while ($true) {
        $s = $script:Settings
        Write-Host ''
        Write-Host '--- Folder layout ---' -ForegroundColor Cyan
        Write-Host "  Now: $(Get-LayoutSummary)"
        $groups = @($s.BpmGroups.Keys | ForEach-Object { "$_ = $($s.BpmGroups[$_])" })
        Write-Host ('  BPM groups: {0}; other families: {1} BPM bands' -f ($(if ($groups.Count) { $groups -join ', ' } else { 'none' })), $s.BpmBandWidth)
        Write-Host ''
        Write-Host '  1) Layout   2) Keys per folder   3) Key notation   4) BPM groups per family   5) BPM band width   0) Back'
        switch (Read-Choice 'Choose' @('0', '1', '2', '3', '4', '5') '0') {
            '0' { return }
            '1' {
                $keys = @($LayoutNames.Keys)
                for ($i = 0; $i -lt $keys.Count; $i++) { Write-Host ('  {0}) {1}' -f ($i + 1), $LayoutNames[$keys[$i]]) }
                $n = ConvertTo-IntOrDefault (Read-Host 'Layout number') 0 0 $keys.Count
                if ($n -ge 1) { $s.FolderLayout = $keys[$n - 1] }
            }
            '2' {
                Write-Host '  Neighbouring Camelot numbers mix well, so they can share a folder.'
                Write-Host '  1 = a folder per number (12), 2 = six folders, 3 = four folders (recommended), 4 = three, 6 = two.'
                $s.KeyGroupSize = ConvertTo-IntOrDefault (Read-Host "Keys per folder (1-6) [$($s.KeyGroupSize)]") ([int]$s.KeyGroupSize) 1 6
            }
            '3' {
                $keys = @($NotationNames.Keys)
                for ($i = 0; $i -lt $keys.Count; $i++) { Write-Host ('  {0}) {1}' -f ($i + 1), $NotationNames[$keys[$i]]) }
                $n = ConvertTo-IntOrDefault (Read-Host 'Notation number') 0 0 $keys.Count
                if ($n -ge 1) { $s.KeyNotation = $keys[$n - 1] }
            }
            '4' {
                Write-Host '  Boundaries split BPM into groups: "170,176" gives 0-169, 170-175 and 176+.'
                Write-Host ('  Families: ' + ((@($script:GenreRules | ForEach-Object { $_.Family } | Select-Object -Unique)) -join ', ')) -ForegroundColor DarkGray
                $fam = (Read-Host 'Family (for example DnB)').Trim()
                if (-not $fam) { continue }
                $txt = (Read-Host "Boundaries for $fam (Enter = remove its groups)").Trim()
                $existing = @($s.BpmGroups.Keys | Where-Object { $_ -ieq $fam })
                foreach ($k in $existing) { $s.BpmGroups.Remove($k) }
                if ($txt) {
                    if ($txt -notmatch '^\d+(\s*[,;]\s*\d+)*$') { Write-Host 'Please type numbers separated by commas, like 170,176.' -ForegroundColor Yellow; continue }
                    $s.BpmGroups[$fam] = ($txt -replace '\s', '' -replace ';', ',')
                }
            }
            '5' { $s.BpmBandWidth = ConvertTo-IntOrDefault (Read-Host "BPM band width (1-50) [$($s.BpmBandWidth)]") ([int]$s.BpmBandWidth) 1 50 }
        }
        Save-Settings
        Write-Host 'Saved. Copies made earlier are offered to move into the new layout on the next run.' -ForegroundColor DarkGray
    }
}

# Track counts per genre from the last saved plan (cheap, no XML needed).
function Get-PlanGenreCounts {
    $c = @{}
    if ($script:PlanPath -and (Test-Path -LiteralPath $script:PlanPath)) {
        try { foreach ($r in (Import-SafeCsv $script:PlanPath)) { if ($r.Genre) { $c[[string]$r.Genre] = 1 + [int]$c[[string]$r.Genre] } } } catch { }
    }
    return $c
}

function Select-GenreRule([string]$Prompt) {
    $v = (Read-Host $Prompt).Trim()
    if (-not $v) { return $null }
    $n = 0
    if ([int]::TryParse($v, [ref]$n) -and $n -ge 1 -and $n -le $script:GenreRules.Count) { return $script:GenreRules[$n - 1] }
    $r = Get-GenreRule $v
    if (-not $r) { Write-Host "No genre called '$v'." -ForegroundColor Yellow }
    return $r
}

function Show-GenreList {
    $counts = Get-PlanGenreCounts
    Write-Host ('  {0,3}  {1,-24} {2,-14} {3,-16} {4}' -f '#', 'Genre', 'Family', 'BPM', 'Tracks (last plan)')
    for ($i = 0; $i -lt $script:GenreRules.Count; $i++) {
        $r = $script:GenreRules[$i]
        $bpm = if ($r.Bpm) { $r.Bpm } elseif ($r.Family -ne $r.Name) { '(family)' } else { '-' }
        $cnt = if ($counts.ContainsKey($r.Name)) { $counts[$r.Name] } else { '' }
        Write-Host ('  {0,3}  {1,-24} {2,-14} {3,-16} {4}' -f ($i + 1), (Limit-Text $r.Name 24), (Limit-Text $r.Family 14), $bpm, $cnt)
    }
    Write-Host '  Genres are checked from top to bottom; the first one that fits a tag wins.' -ForegroundColor DarkGray
}

function Test-BpmRangeText([string]$Text) {
    if (-not $Text) { return $true }
    return ($Text -match '^\d{2,3}-\d{2,3}(,\d{2,3}-\d{2,3})*$')
}

function Edit-GenreSettings {
    while ($true) {
        Write-Host ''
        Write-Host '--- Genres ---' -ForegroundColor Cyan
        Write-Host "  $($script:GenreRules.Count) genres. Detail level: $($script:Settings.DetailLevel) (Detailed = subgenre folders, Broad = one folder per family)"
        Write-Host "  Genre list file: $RulesPath" -ForegroundColor DarkGray
        Write-Host '  1) Switch detail level     2) List genres            3) Add a genre'
        Write-Host '  4) Edit a genre            5) Merge two genres       6) Delete a genre'
        Write-Host '  7) Junk tags (ignored)     8) Update MusicBrainz genre list now'
        Write-Host '  9) Restore the default genre list                    0) Back'
        $c = Read-Choice 'Choose' @('0', '1', '2', '3', '4', '5', '6', '7', '8', '9') '0'
        switch ($c) {
            '0' { return }
            '1' {
                $script:Settings.DetailLevel = if ($script:Settings.DetailLevel -eq 'Broad') { 'Detailed' } else { 'Broad' }
                Save-Settings
                Write-Host "Detail level is now $($script:Settings.DetailLevel)."
            }
            '2' { Show-GenreList }
            '3' {
                $name = (Read-Host 'New genre name (for example Minimal DnB)').Trim()
                if (-not $name) { continue }
                if (Get-GenreRule $name) { Write-Host 'That genre already exists.' -ForegroundColor Yellow; continue }
                $guess = Get-RuleForTag $name
                $defFam = if ($guess) { $guess.Family } else { $name }
                $fam = (Read-Host "Family [$defFam]").Trim()
                if (-not $fam) { $fam = $defFam }
                $bpm = (Read-Host 'Accepted BPM ranges, like 160-180 (Enter = use the family''s)').Trim()
                if (-not (Test-BpmRangeText $bpm)) { Write-Host 'BPM ranges look like 160-180 or 160-180,80-90. Saved without BPM.' -ForegroundColor Yellow; $bpm = '' }
                Add-GenreRule $name $fam $bpm
                Save-GenreRules
                Write-Host "Added '$name'. Tags spelled exactly like the name are matched; add more patterns under Edit." -ForegroundColor Green
            }
            '4' { Edit-OneGenre }
            '5' {
                Show-GenreList
                $from = Select-GenreRule 'Merge which genre (number or name)'
                if (-not $from) { continue }
                $into = Select-GenreRule "Into which genre should '$($from.Name)' go"
                if (-not $into -or $into.Name -eq $from.Name) { continue }
                $into.Patterns = @($into.Patterns) + @($from.Patterns)
                $into.Aliases = @($into.Aliases) + @($from.Name) + @($from.Aliases)
                foreach ($r in $script:GenreRules) { if ($r.Family -eq $from.Name) { $r.Family = $into.Family } }
                $script:GenreRules = @($script:GenreRules | Where-Object { $_.Name -ne $from.Name })
                Update-GenreRuleSets
                Save-GenreRules
                $n = Update-DecisionGenre $from.Name $into.Name
                Write-Host "Merged '$($from.Name)' into '$($into.Name)' ($n remembered decisions updated)." -ForegroundColor Green
            }
            '6' {
                Show-GenreList
                $r = Select-GenreRule 'Delete which genre (number or name)'
                if (-not $r) { continue }
                if (-not (Read-YesNo "Delete '$($r.Name)'? Its tags then fall through to the next genre that fits." $false)) { continue }
                $script:GenreRules = @($script:GenreRules | Where-Object { $_.Name -ne $r.Name })
                Update-GenreRuleSets
                Save-GenreRules
                Write-Host 'Deleted. Tracks you decided by hand keep that genre (as a folder name).'
            }
            '7' { Edit-IgnoredTags }
            '8' {
                if (-not (Test-Switch 'MusicBrainz')) { Write-Host 'MusicBrainz is switched off.' -ForegroundColor Yellow; continue }
                Update-GenreDictionary -Force
            }
            '9' {
                if (-not (Read-YesNo 'Replace your genre list with the default list? Genres you added are lost.' $false)) { continue }
                $d = Get-DefaultGenreRules
                Set-GenreRules $d.Genres $d.Ignored $d.Declined
                Save-GenreRules
                Write-Host 'Default genre list restored.'
            }
        }
    }
}

function Edit-OneGenre {
    Show-GenreList
    $r = Select-GenreRule 'Edit which genre (number or name)'
    if (-not $r) { return }
    while ($true) {
        Write-Host ''
        Write-Host "  Name:     $($r.Name)"
        Write-Host "  Family:   $($r.Family)"
        Write-Host "  BPM:      $(if ($r.Bpm) { $r.Bpm } else { '(uses the family''s)' })"
        Write-Host "  Patterns: $($r.Patterns -join '   ')"
        Write-Host '  1) Rename   2) Family   3) BPM ranges   4) Patterns   0) Back'
        switch (Read-Choice 'Choose' @('0', '1', '2', '3', '4') '0') {
            '0' { return }
            '1' {
                $new = (Read-Host 'New name').Trim()
                if (-not $new -or (Get-GenreRule $new)) { Write-Host 'Empty or already used.' -ForegroundColor Yellow; continue }
                $old = $r.Name
                foreach ($x in $script:GenreRules) { if ($x.Family -eq $old) { $x.Family = $new } }
                $r.Name = $new
                foreach ($k in @($script:Settings.BpmGroups.Keys)) {
                    if ($k -ieq $old) { $v = $script:Settings.BpmGroups[$k]; $script:Settings.BpmGroups.Remove($k); $script:Settings.BpmGroups[$new] = $v; Save-Settings }
                }
                $n = Update-DecisionGenre $old $new
                Write-Host "Renamed ($n remembered decisions updated)."
            }
            '2' { $f = (Read-Host "Family [$($r.Family)]").Trim(); if ($f) { $r.Family = $f } }
            '3' {
                $b = (Read-Host 'BPM ranges like 160-180,80-90 (first = main tempo; Enter = use the family''s)').Trim()
                if (Test-BpmRangeText $b) { $r.Bpm = $b } else { Write-Host 'Not changed: use the form 160-180,80-90.' -ForegroundColor Yellow }
            }
            '4' {
                Write-Host '  Patterns are regular expressions checked against the tag in lower case, with & written as and.'
                Write-Host '  Examples: "liquid" (anywhere in the tag), "^deep dnb$" (exactly), "drum ?(and|n) ?bass".'
                $p = (Read-Host 'Patterns separated by " | " (Enter = keep)').Trim()
                if ($p) { $r.Patterns = @($p -split '\s+\|\s+' | Where-Object { $_ }) }
            }
        }
        Update-GenreRuleSets
        Save-GenreRules
        $r = Get-GenreRule $r.Name
    }
}

function Edit-IgnoredTags {
    while ($true) {
        Write-Host ''
        Write-Host '  Junk tags are never used as genres (for example "electronic", "seen live").'
        Write-Host ('  Now: ' + ($(if ($script:IgnoredRaw.Count) { $script:IgnoredRaw -join ', ' } else { '(none)' })))
        Write-Host '  1) Add   2) Remove   3) Clear the "never suggest" list   0) Back'
        switch (Read-Choice 'Choose' @('0', '1', '2', '3') '0') {
            '0' { return }
            '1' { $t = (Read-Host 'Tag to ignore').Trim(); if ($t) { $script:IgnoredRaw = @($script:IgnoredRaw) + $t } }
            '2' { $t = (Read-Host 'Tag to use again').Trim(); $script:IgnoredRaw = @($script:IgnoredRaw | Where-Object { $_ -ine $t }) }
            '3' { $script:DeclinedRaw = @(); Write-Host 'New genres you said no to will be suggested again.' }
        }
        Update-GenreRuleSets
        Save-GenreRules
    }
}

function Set-ApiKeyInteractive([string]$Name) {
    if ($Name -eq 'Discogs') { Write-Host '  Discogs token: discogs.com > log in > Settings > Developers > Generate new token' }
    else { Write-Host '  Last.fm key: last.fm/api/account/create (any app name; you only need the "API key")' }
    $plain = (ConvertFrom-SecureStringToPlain (Read-Host "  Paste your $Name key (hidden; Enter to skip)" -AsSecureString)).Trim()
    if (-not $plain) { return }
    Write-Host '  Testing the key...' -ForegroundColor DarkGray
    $ok = Test-ApiKey $Name $plain
    if ($ok -eq $true) { Write-Host '  The key works.' -ForegroundColor Green }
    elseif ($null -eq $ok) { Write-Host '  Could not check the key right now (no internet?). It is saved anyway.' -ForegroundColor Yellow }
    else {
        Write-Host '  The service rejected this key.' -ForegroundColor Yellow
        if (-not (Read-YesNo '  Save it anyway?' $false)) { return }
    }
    $saved = Get-SavedKeys
    $saved[$Name] = $plain
    Save-SavedKeys $saved
    Update-ApiKeys
    Write-Host "  Saved, encrypted, in $KeysPath"
    $envName = if ($Name -eq 'Discogs') { $EnvDiscogs } else { $EnvLastFm }
    if ([Environment]::GetEnvironmentVariable($envName)) { Write-Host "  Note: the environment variable $envName is set and is used instead." -ForegroundColor Yellow }
}

function Edit-ServiceSettings {
    while ($true) {
        $sw = $script:Settings.Switches
        Write-Host ''
        Write-Host '--- Lookup services and API keys ---' -ForegroundColor Cyan
        Write-Host ('  MusicBrainz  [{0}]  free, no key needed. Contact email: {1}' -f $(if ($sw.MusicBrainz) { 'on ' } else { 'off' }), $(if ($script:Settings.MbContact) { $script:Settings.MbContact } else { '(none - optional)' }))
        Write-Host ('  Discogs      [{0}]  key: {1}' -f $(if ($sw.Discogs) { 'on ' } else { 'off' }), $script:Keys.DiscogsFrom)
        Write-Host ('  Last.fm      [{0}]  key: {1}' -f $(if ($sw.LastFm) { 'on ' } else { 'off' }), $script:Keys.LastFmFrom)
        Write-Host '  Keys are never shown. Only artist and title are sent to these services.' -ForegroundColor DarkGray
        Write-Host ''
        Write-Host '  1) MusicBrainz on/off   2) Discogs on/off   3) Last.fm on/off'
        Write-Host '  4) Set Discogs token    5) Remove saved Discogs token    6) Test Discogs token'
        Write-Host '  7) Set Last.fm key      8) Remove saved Last.fm key      9) Test Last.fm key'
        Write-Host '  10) MusicBrainz contact email (optional)                 0) Back'
        switch (Read-Choice 'Choose' @('0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10') '0') {
            '0' { return }
            '1' { $sw.MusicBrainz = -not $sw.MusicBrainz; Save-Settings }
            '2' { $sw.Discogs = -not $sw.Discogs; Save-Settings }
            '3' { $sw.LastFm = -not $sw.LastFm; Save-Settings }
            '4' { Set-ApiKeyInteractive 'Discogs' }
            '7' { Set-ApiKeyInteractive 'LastFm' }
            { $_ -in @('5', '8') } {
                $name = if ($_ -eq '5') { 'Discogs' } else { 'LastFm' }
                $saved = Get-SavedKeys
                $saved[$name] = ''
                Save-SavedKeys $saved
                Update-ApiKeys
                Write-Host "Saved $name key removed."
            }
            { $_ -in @('6', '9') } {
                $name = if ($_ -eq '6') { 'Discogs' } else { 'LastFm' }
                $result = Test-ApiKey $name $script:Keys[$name]
                if ($result -eq $true) { Write-Host 'The key works.' -ForegroundColor Green }
                elseif ($null -eq $result) { Write-Host 'Could not check right now (no internet?).' -ForegroundColor Yellow }
                elseif (-not $script:Keys[$name]) { Write-Host 'No key set.' -ForegroundColor Yellow }
                else { Write-Host 'The service rejected the key.' -ForegroundColor Yellow }
            }
            '10' {
                Write-Host '  MusicBrainz likes apps to include a contact. TuneTidy already sends its GitHub address;'
                Write-Host '  your email is optional and only sent to MusicBrainz.'
                $v = (Read-Host 'Email (Enter = none)').Trim()
                $script:Settings.MbContact = $v
                Save-Settings
                Update-UserAgent
            }
        }
    }
}

function Edit-Switches {
    while ($true) {
        Write-Host ''
        Write-Host '--- On/off switches ---' -ForegroundColor Cyan
        $names = @($SwitchInfo.Keys)
        for ($i = 0; $i -lt $names.Count; $i++) {
            $on = Test-Switch $names[$i]
            Write-Host ('  {0,2}) [{1}] {2}' -f ($i + 1), $(if ($on) { 'on ' } else { 'off' }), $SwitchInfo[$names[$i]]) -ForegroundColor $(if ($on) { 'White' } else { 'DarkGray' })
        }
        $n = ConvertTo-IntOrDefault (Read-Host 'Number to switch on/off (Enter = back)') 0 0 $names.Count
        if ($n -lt 1) { return }
        $key = $names[$n - 1]
        $script:Settings.Switches[$key] = -not (Test-Switch $key)
        Save-Settings
    }
}

# =============================================================================
# Logs
# =============================================================================

function Show-LogSummary([string]$Path) {
    $rows = @(Import-SafeCsv $Path)
    Write-Host ''
    Write-Host "Summary of $(Split-Path -Leaf $Path)" -ForegroundColor Cyan
    $final = @($rows | Where-Object { $_.Stage -eq 'Final plan destination' })
    if ($final.Count -gt 0) {
        Write-Host "  Tracks in the plan: $($final.Count)"
        $final | Group-Object { if ($_.FinalGenre) { $_.FinalGenre } else { '(no genre -> _Review)' } } |
            Sort-Object Count -Descending | Select-Object -First 12 |
            ForEach-Object { Write-Host ('    {0,-30} {1,5}' -f (Limit-Text $_.Name 30), $_.Count) }
    }
    $initial = @($rows | Where-Object { $_.Stage -eq 'Initial classification' })
    if ($initial.Count -gt 0) {
        Write-Host '  Where genres came from at the start:'
        $initial | Group-Object Source | Sort-Object Count -Descending |
            ForEach-Object { Write-Host ('    {0,-40} {1,5}' -f (Limit-Text $_.Name 40), $_.Count) }
    }
    $online = @($rows | Where-Object { $_.Stage -in @('Online genre lookup', 'Online genre suggestion') })
    if ($online.Count -gt 0) {
        Write-Host '  Online lookups:'
        $online | Group-Object { $_.Decision -replace 'BPM \S+ does not fit .*', 'BPM does not fit the genre' } |
            Sort-Object Count -Descending |
            ForEach-Object { Write-Host ('    {0,-55} {1,5}' -f (Limit-Text $_.Name 55), $_.Count) }
    }
    $decided = @($rows | Where-Object { $_.Stage -eq 'Genre suggestion decision' })
    if ($decided.Count -gt 0) {
        Write-Host '  Your answers to suggestions:'
        $decided | Group-Object Decision | Sort-Object Count -Descending |
            ForEach-Object { Write-Host ('    {0,-55} {1,5}' -f (Limit-Text $_.Name 55), $_.Count) }
    }
    if ($rows.Count -eq 0) { Write-Host '  (empty log)' }
}

function Show-LogsMenu {
    while ($true) {
        $logs = @(Get-LogFiles)
        Write-Host ''
        Write-Host '--- Logs ---' -ForegroundColor Cyan
        Write-Host "  Log folder: $script:LogFolder"
        Write-Host ('  {0} log(s) saved, keeping the newest {1}. Genre log is {2}.' -f $logs.Count, $script:Settings.KeepLogs, $(if (Test-Switch 'GenreLog') { 'on' } else { 'off' }))
        Write-Host '  1) Summary of the latest run   2) Open the latest log in your spreadsheet app'
        Write-Host '  3) Open the log folder         4) How many logs to keep       5) Genre log on/off   0) Back'
        switch (Read-Choice 'Choose' @('0', '1', '2', '3', '4', '5') '0') {
            '0' { return }
            '1' { if ($logs.Count) { Show-LogSummary $logs[0].FullName } else { Write-Host 'No logs yet.' } }
            '2' { if ($logs.Count) { Invoke-Item -LiteralPath $logs[0].FullName } else { Write-Host 'No logs yet.' } }
            '3' { if (Test-Path -LiteralPath $script:LogFolder) { Invoke-Item -LiteralPath $script:LogFolder } else { Write-Host 'No logs yet.' } }
            '4' { $script:Settings.KeepLogs = ConvertTo-IntOrDefault (Read-Host "Logs to keep [$($script:Settings.KeepLogs)]") ([int]$script:Settings.KeepLogs) 1 1000; Save-Settings }
            '5' { $script:Settings.Switches.GenreLog = -not (Test-Switch 'GenreLog'); Save-Settings }
        }
    }
}

# =============================================================================
# Reset
# =============================================================================

function Invoke-ResetMenu {
    Write-Host ''
    Write-Host '--- Reset ---' -ForegroundColor Cyan
    Write-Host '  1) Clear the online lookup cache (the next run looks everything up again)'
    Write-Host '  2) Forget all your genre decisions (a backup copy is kept)'
    Write-Host '  3) Restore default settings (keeps your folders and API keys)'
    Write-Host '  4) Restore the default genre list'
    Write-Host '  0) Back'
    switch (Read-Choice 'Choose' @('0', '1', '2', '3', '4') '0') {
        '1' {
            if (Read-YesNo 'Clear the lookup cache?' $false) {
                foreach ($p in @($CachePath, $DictionaryPath)) { if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p } }
                Write-Host 'Lookup cache cleared.'
            }
        }
        '2' {
            if ((Test-Path -LiteralPath $script:DecisionsPath) -and (Read-YesNo 'Forget all genre decisions?' $false)) {
                $backup = Join-Path $script:DataRoot ('Decisions-backup-{0}.csv' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
                Move-Item -LiteralPath $script:DecisionsPath -Destination $backup
                $script:Decisions = @{}
                Write-Host "Decisions cleared. Backup: $backup"
            }
        }
        '3' {
            if (Read-YesNo 'Restore default settings?' $false) {
                $old = $script:Settings
                $new = Get-DefaultSettings
                foreach ($k in @('MusicRoot', 'OutputFolder', 'NewTracksFolders', 'XmlPath', 'MbContact')) { $new[$k] = $old[$k] }
                $script:Settings = $new
                Save-Settings
                Initialize-Paths
                Write-Host 'Default settings restored.'
            }
        }
        '4' {
            if (Read-YesNo 'Replace your genre list with the default list?' $false) {
                $d = Get-DefaultGenreRules
                Set-GenreRules $d.Genres $d.Ignored $d.Declined
                Save-GenreRules
                Write-Host 'Default genre list restored.'
            }
        }
    }
}

# =============================================================================
# First run
# =============================================================================

function Invoke-FirstRunSetup {
    $s = $script:Settings
    Write-Host ''
    Write-Host 'Welcome to TuneTidy! A few questions first - everything can be changed later under Settings.' -ForegroundColor Cyan
    Write-Host 'TuneTidy never moves, renames or deletes your original files. It only makes sorted copies.'
    Write-Host ''
    $s.MusicRoot = Read-FolderPath 'Your main music folder' ([Environment]::GetFolderPath('MyMusic')) $true $false
    while ($true) {
        $out = Read-FolderPath 'Folder for the sorted copies' (Join-Path $s.MusicRoot 'Organized Music') $false $false
        $problem = Test-OutputFolderOk $out
        if (-not $problem) { break }
        Write-Host $problem -ForegroundColor Yellow
    }
    $s.OutputFolder = $out
    Initialize-Paths
    Write-Host ''
    Write-Host 'Where do new downloads land before you sort them? Used by "Only new tracks". (optional)'
    while ($true) {
        $f = Read-FolderPath 'New-tracks folder (Enter to skip or finish)' '' $true $true
        if (-not $f) { break }
        if (Add-NewTracksFolder $f) { Write-Host "  added: $f" }
    }
    Write-Host ''
    Write-Host "Folder layout: $(Get-LayoutSummary)"
    Write-Host '  (change it under Settings > Folder layout)' -ForegroundColor DarkGray
    Write-Host ''
    Write-Host 'Online genre lookup: MusicBrainz works without an account. Discogs and Last.fm (free accounts)'
    Write-Host 'find many more tracks. Keys are stored encrypted, readable only by your Windows account.'
    if (Read-YesNo 'Add a Discogs token and/or Last.fm key now?' $false) {
        Set-ApiKeyInteractive 'Discogs'
        Set-ApiKeyInteractive 'LastFm'
    }
    Save-Settings
    Write-Host ''
    Write-Host "Setup saved in $SettingsPath" -ForegroundColor Green
}

# =============================================================================
# Start-up defaults (the tests load this file with ". .\TuneTidy.ps1" and stop here)
# =============================================================================

$script:Settings = Get-DefaultSettings
$defaultRules = Get-DefaultGenreRules
Set-GenreRules $defaultRules.Genres $defaultRules.Ignored $defaultRules.Declined
$script:ChannelNorms = @($DefaultChannelNames | ForEach-Object { Get-Norm $_ })
$script:OutputRoot = ''
Update-UserAgent

if ($MyInvocation.InvocationName -eq '.') { return }

try {

# =============================================================================
# STEP 0 - settings (first run asks a few questions)
# =============================================================================

Write-Host ''
Write-Host "=== TuneTidy $TuneTidyVersion ===" -ForegroundColor Green
if ($script:Preview) { Write-Host 'Preview mode (-WhatIf): no music files are copied, moved or deleted.' -ForegroundColor Yellow }

New-Item -ItemType Directory -Force -Path $AppDataRoot | Out-Null
$firstRun = -not (Test-Path -LiteralPath $SettingsPath)
$script:Settings = Read-Settings
Update-UserAgent
if ($firstRun) {
    if ($script:Unattended) { throw 'TuneTidy has not been set up yet. Run it once normally (without -Unattended) first.' }
    Invoke-FirstRunSetup
}
if (-not $script:Settings.OutputFolder) { throw "No output folder is set. Fix or delete $SettingsPath and run again." }
Initialize-Paths
Read-GenreRules
Update-ApiKeys
Import-OldOrganizerData
Remove-OldLogs

# =============================================================================
# STEP 1 - whole collection, only new tracks, or settings?
# =============================================================================

if ($script:Unattended -and -not $Mode) { $Mode = 'New' }
while (-not $Mode) {
    $nf = @($script:Settings.NewTracksFolders)
    Write-Host ''
    Write-Host '  1) Whole collection (first-time organize, or re-check everything)'
    Write-Host ('  2) Only new tracks ({0})' -f $(if ($nf.Count) { ($nf | ForEach-Object { Split-Path -Leaf $_ }) -join ', ' } else { 'no new-tracks folder set yet' }))
    Write-Host '  3) Settings'
    Write-Host '  4) Quit'
    switch (Read-Choice 'Choose 1-4 [1]' @('1', '2', '3', '4') '1') {
        '1' { $Mode = 'Initial' }
        '2' {
            if ($nf.Count -eq 0) {
                $f = Read-FolderPath 'Folder where your new tracks land (Enter to cancel)' '' $true $true
                if ($f -and (Add-NewTracksFolder $f)) { Save-Settings; $Mode = 'New' }
            }
            else { $Mode = 'New' }
        }
        '3' { Show-SettingsMenu; Initialize-Paths }
        '4' { exit 0 }
    }
}
if ($Mode -eq 'New' -and @($script:Settings.NewTracksFolders).Count -eq 0) {
    throw 'No new-tracks folder is set. Add one under Settings > Folders.'
}
$script:Decisions = Get-Decisions

# =============================================================================
# STEP 2 - which Rekordbox export?
# =============================================================================

if (-not $XmlPath) {
    $saved = [string]$script:Settings.XmlPath
    if ($saved -and (Test-Path -LiteralPath $saved -PathType Leaf)) {
        $when = (Get-Item -LiteralPath $saved).LastWriteTime
        if ($script:Unattended -or (Read-YesNo ("Use Rekordbox export '{0}' (saved {1:g})?" -f $saved, $when))) { $XmlPath = $saved }
    }
    elseif ($script:Unattended) { throw 'No saved Rekordbox export found. Pass -XmlPath or run once normally.' }
    if (-not $XmlPath) {
        $candidates = @()
        if ($PSScriptRoot) { $candidates += (Join-Path $PSScriptRoot 'export.xml') }
        $candidates += (Join-Path ([Environment]::GetFolderPath('Desktop')) 'export.xml')
        $found = $candidates | Where-Object { $_ -ne $saved -and (Test-Path -LiteralPath $_ -PathType Leaf) } | Select-Object -First 1
        if ($found -and (Read-YesNo "Use Rekordbox export '$found'?")) { $XmlPath = $found }
        while (-not $XmlPath) {
            Write-Host 'In Rekordbox: File > Export Collection in xml format.' -ForegroundColor DarkGray
            $XmlPath = Read-FilePath 'Full path to your Rekordbox XML export' $false
        }
    }
}
if (-not (Test-Path -LiteralPath $XmlPath -PathType Leaf)) { throw "XML file not found: $XmlPath" }
$XmlPath = Get-FullPathOrNull $XmlPath
if ($script:Settings.XmlPath -ne $XmlPath) { $script:Settings.XmlPath = $XmlPath; Save-Settings }
if ($Mode -eq 'New' -and ((Get-Date) - (Get-Item -LiteralPath $XmlPath).LastWriteTime).TotalHours -gt 24) {
    Write-Host 'Tip: this export is more than a day old. New tracks only show up after you import them into Rekordbox and export again.' -ForegroundColor DarkGray
}

# =============================================================================
# STEP 3 - build the list of tracks
# =============================================================================

New-Item -ItemType Directory -Force -Path $script:DataRoot | Out-Null
$tracks = @(Get-TracksFromRekordboxXml $XmlPath)
$newRoots = @($script:Settings.NewTracksFolders | ForEach-Object { Get-FullPathOrNull $_ } | Where-Object { $_ })
$script:Discovered = @{}
$skippedNetwork = 0
$skippedCopies = 0
$rows = foreach ($t in $tracks) {
    $source = [string]$t.SourcePath
    $allowed = Test-AllowedSourcePath $source
    if ($source -and -not $allowed) { $skippedNetwork++ }
    # Copies made by TuneTidy that were imported into Rekordbox are never copied again.
    if ($source -and (Test-InsideFolder $source $script:OutputRoot)) { $skippedCopies++; continue }
    if ($Mode -eq 'New') {
        if (-not $allowed) { continue }
        $inNew = $false
        foreach ($nr in $newRoots) { if (Test-InsideFolder $source $nr) { $inNew = $true } }
        if (-not $inNew) { continue }
    }
    $genreRaw = ([string]$t.Genre).Trim()
    if (-not $genreRaw -or $genreRaw -eq 'Music') { $genre = ''; $gsrc = '' }      # Rekordbox's generic "Music"
    elseif ($genreRaw -match '^(?i)loop samples?$') { $genre = $SamplesGenre; $gsrc = 'Tag' }
    else {
        $genre = Get-StandardGenre $genreRaw
        $gsrc = 'Tag'
        Register-TagDiscovery $genreRaw
    }

    [pscustomobject]@{
        PlanId            = 0
        TrackKey          = Get-TrackKey $source ([string]$t.TrackID)
        TrackID           = [string]$t.TrackID
        Artist            = [string]$t.Artist
        Title             = [string]$t.Title
        OriginalGenre     = $genreRaw
        Genre             = $genre
        BPM               = [string]$t.BPM
        Key               = [string]$t.Key
        Camelot           = Get-Camelot ([string]$t.Key)
        SourcePath        = $source
        SourceExists      = [bool]($allowed -and (Test-Path -LiteralPath $source -PathType Leaf))
        GenreSource       = $gsrc
        ParsedArtist      = ''
        ParsedTitle       = ''
        BPMRounded        = $null
        DestinationFolder = ''
        DestinationPath   = ''
        PathTooLong       = $false
    }
}
$rows = @($rows)

if ($rows.Count -eq 0) {
    if ($Mode -eq 'New') {
        Write-Warning 'No tracks in the export point to files inside your new-tracks folder(s). Import the files into Rekordbox and export a fresh XML.'
        if (-not $script:Unattended) { Read-Host 'Press Enter to close this window' | Out-Null }
        exit 0
    }
    throw 'No tracks were available to add to the plan.'
}

$n = 0
foreach ($row in $rows) { $n++; $row.PlanId = $n }

# Your earlier decisions win over tags and lookups. Delete a line in Decisions.csv to undo one.
$fromDecisions = 0
foreach ($row in $rows) {
    $d = $script:Decisions[$row.TrackKey]
    if ($d -and $d.Genre -and $d.Genre -notin @('Review', $ReviewFolder)) {
        $row.Genre = [string]$d.Genre
        $row.GenreSource = 'Decision'
        $fromDecisions++
    }
}

foreach ($row in $rows) {
    if ($row.GenreSource -eq 'Decision') {
        Write-GenreLog $row 'Initial classification' 'Your earlier decision (Decisions.csv)' $row.Genre 'Used directly' $row.Genre ''
    }
    elseif ($row.GenreSource -eq 'Tag') {
        $note = if ($row.Genre -cne $row.OriginalGenre) { 'Tag tidied to standard genre name' } else { '' }
        Write-GenreLog $row 'Initial classification' 'Rekordbox Genre field' $row.Genre 'Used directly' $row.OriginalGenre $note
    }
    else {
        Write-GenreLog $row 'Initial classification' 'No usable Rekordbox Genre' '' 'Sent to Review for further matching' $row.OriginalGenre ''
    }
}

$existing = @($rows | Where-Object SourceExists).Count
Write-Host ''
Write-Host "Loaded $($rows.Count) tracks; $existing of them found on disk."
if ($fromDecisions -gt 0) { Write-Host "$fromDecisions tracks use a genre you decided in an earlier run." }
if ($skippedCopies -gt 0) { Write-Host "$skippedCopies tracks are TuneTidy copies inside the output folder and are skipped." -ForegroundColor DarkGray }
if ($skippedNetwork -gt 0) { Write-Host "$skippedNetwork tracks are on network paths and are skipped (Settings > On/off switches)." -ForegroundColor Yellow }
if ($existing -lt $rows.Count) { Write-Host "$($rows.Count - $existing) tracks point to missing or skipped files." -ForegroundColor Yellow }

# =============================================================================
# STEP 4 - find genres for tracks that have none
# =============================================================================

$review = @($rows | Where-Object { -not $_.Genre -and $_.GenreSource -ne 'Decision' })
Write-Host ''
Write-Host "$($review.Count) of $($rows.Count) tracks have no usable genre and would go to '$ReviewFolder'."

if ($review.Count -gt 0) {

    # 4a) your own folder names (offline, instant)
    $hints = @()
    if ((Test-Switch 'FolderHints') -and -not $script:Unattended) {
        foreach ($row in $review) {
            $h = Get-FolderHint $row.SourcePath
            if (-not $h) { continue }
            $g = Get-StandardGenre $h
            if (Test-Rejected $script:Decisions $row $g) { continue }
            $hints += [pscustomobject]@{ Row = $row; Genre = $g; Source = 'Folder'; Evidence = "folder '$h'"; Rule = '' }
        }
    }
    if ($hints.Count -gt 0) {
        Write-Host ''
        Write-Host "$($hints.Count) of them already sit in your own genre folders:" -ForegroundColor Cyan
        $hints | Group-Object Genre | Sort-Object Count -Descending | ForEach-Object {
            Write-Host ('  {0,-28} {1,5} tracks' -f $_.Name, $_.Count)
        }
        if (Read-YesNo 'Use those folder names as the genre? (remembered for next time)') {
            foreach ($p in $hints) { Approve-Proposal $p 'Accepted' }
        }
        else {
            foreach ($p in $hints) {
                Set-Decision -Decisions $script:Decisions -Row $p.Row -Rejected $p.Genre -Source 'Rejected folder name'
                Write-GenreLog $p.Row 'Genre suggestion' 'Source folder name' $p.Genre 'Rejected (remembered)' $p.Evidence ''
            }
        }
        Save-Decisions $script:Decisions
    }

    # 4b) online lookup for the rest
    $remaining = @($rows | Where-Object { -not $_.Genre -and $_.GenreSource -ne 'Decision' })
    $enabled = @(Get-EnabledSources)
    $searchable = @()
    if ($remaining.Count -gt 0 -and $enabled.Count -eq 0) {
        Write-Host 'Online lookup is switched off (Settings > Lookup services and API keys).' -ForegroundColor DarkGray
    }
    elseif ($remaining.Count -gt 0) {
        foreach ($row in $remaining) {
            $terms = Get-SearchTerms $row.Artist $row.Title $row.SourcePath
            if ($terms) {
                $row.ParsedArtist = [string]$terms.Artist
                $row.ParsedTitle = [string]$terms.Title
                Write-GenreLog $row 'Search terms prepared' ([string]$terms.From) '' 'Ready for online lookup' ("Artist='$($terms.Artist)'; Title='$($terms.Title)'; Remixer='$($terms.Remixer)'; Slug='$($terms.Slug)'") ''
                $searchable += [pscustomobject]@{ Row = $row; Terms = $terms; Done = $false }
            }
            else {
                Write-GenreLog $row 'Search terms prepared' 'Rekordbox tags / filename' '' 'Could not determine artist/title (channel name or no "Artist - Title")' '' ''
            }
        }
    }

    if ($searchable.Count -gt 0) {
        Write-Host ''
        Write-Host "$($searchable.Count) remaining tracks have enough artist/title info to look up online." -ForegroundColor Cyan
        Write-Host ('Services: ' + ($enabled -join ', '))
        if (-not $script:Keys.Discogs -and -not $script:Keys.LastFm) {
            Write-Host 'Tip: MusicBrainz alone finds only some tracks. Add free Discogs / Last.fm keys under Settings > Lookup services.' -ForegroundColor DarkGray
        }

        $cache = Get-TagCache
        $pending = 0
        foreach ($item in $searchable) {
            foreach ($name in $enabled) {
                if (-not $cache.ContainsKey(('{0}::{1}' -f $name, (Get-CacheKey $item.Terms)))) { $pending++ }
            }
            if (-not $item.Terms.Artist -and $item.Terms.Slug -and -not $cache.ContainsKey('Split::' + $item.Terms.Slug)) { $pending += 3 }
        }
        if ($pending -gt 0) {
            Write-Host "Roughly $([math]::Ceiling($pending * 2 / 60)) minute(s) of new lookups (finished lookups are cached for next time)."
        }

        if (Read-YesNo 'Look up genres online now?' $true $true) {
            Update-GenreDictionary
            $found = @()
            $stats = @{ NewLookups = 0 }
            $nextSave = 20
            $aborted = $false
            $i = 0
            try {
                foreach ($item in $searchable) {
                    $i++
                    Write-Progress -Activity 'Looking up genres online' -Status "$i / $($searchable.Count): $($item.Row.Title)" -PercentComplete (100 * $i / $searchable.Count)
                    if (@($enabled | Where-Object { -not $script:DisabledSources[$_] }).Count -eq 0) { $aborted = $true; break }

                    $res = Resolve-Genre $item.Terms $cache $stats $item.Row
                    $handled = $false
                    if ($res -and (Test-Rejected $script:Decisions $item.Row $res.Genre)) {
                        Write-GenreLog $item.Row 'Online genre lookup' $res.Source $res.Genre 'Not suggested: you rejected this genre for this track before' $res.Evidence $res.Rule
                        $res = $null
                        $handled = $true
                    }
                    if ($res) {
                        $item.Done = $true
                        $label = if ($item.Terms.Artist) { "$($item.Terms.Artist) - $($item.Terms.Title)" } else { $item.Terms.Slug }
                        if ($item.Terms.Remixer) { $label += " ($($item.Terms.Remixer) Remix)" }
                        $found += [pscustomobject]@{ Row = $item.Row; Genre = $res.Genre; Source = $res.Source; Evidence = $res.Evidence; Rule = $res.Rule; Artist = $item.Terms.Artist; Label = $label }
                    }
                    elseif (-not $handled) {
                        Write-GenreLog $item.Row 'Online genre lookup' 'Enabled lookup services' '' 'No confident genre (stays in Review)' '' ''
                    }
                    if ($stats.NewLookups -ge $nextSave) { Save-TagCache $cache; $nextSave += 20 }
                }
            }
            finally {
                Save-TagCache $cache
                Write-Progress -Activity 'Looking up genres online' -Completed
            }
            if ($aborted) {
                Write-Host 'All lookup services stopped answering. Run TuneTidy again later; finished lookups are cached.' -ForegroundColor Yellow
            }

            foreach ($item in $searchable) {
                if ($item.Terms.Artist) { $item.Row.ParsedArtist = $item.Terms.Artist; $item.Row.ParsedTitle = $item.Terms.Title }
            }

            # Same-artist guess: needs at least 2 matched tracks, a 60% majority and a fitting BPM.
            if (Test-Switch 'SameArtistGuess') {
                $byArtist = @{}
                foreach ($f in $found) {
                    $ak = ([string]$f.Artist).ToLowerInvariant()
                    if (-not $ak) { continue }
                    $byArtist[$ak] = @(@($byArtist[$ak]) + $f.Genre | Where-Object { $_ })
                }
                foreach ($item in $searchable) {
                    if ($item.Done -or -not $item.Terms.Artist) { continue }
                    $ak = $item.Terms.Artist.ToLowerInvariant()
                    if (-not $byArtist.ContainsKey($ak)) { continue }
                    $genres = @($byArtist[$ak])
                    $top = $genres | Group-Object | Sort-Object Count -Descending | Select-Object -First 1
                    $share = $top.Count / $genres.Count
                    if ($genres.Count -ge 2 -and $share -ge 0.6 -and (Test-GenreBpm $top.Name $item.Row.BPM) -and -not (Test-Rejected $script:Decisions $item.Row $top.Name)) {
                        $ev = 'Based on {0} other matched tracks by this artist ({1}% {2})' -f $genres.Count, [int][math]::Round(100 * $share), $top.Name
                        $found += [pscustomobject]@{ Row = $item.Row; Genre = $top.Name; Source = 'SameArtist'; Evidence = $ev; Rule = ''; Artist = $item.Terms.Artist; Label = "$($item.Terms.Artist) - $($item.Terms.Title)" }
                    }
                }
            }

            Write-Host ''
            if ($found.Count -eq 0) {
                Write-Host 'No confident genres were found online for these tracks.' -ForegroundColor Yellow
            }
            elseif ($script:Unattended) {
                foreach ($p in $found) { Write-GenreLog $p.Row 'Online genre suggestion' $p.Source $p.Genre 'Awaiting review (unattended run; shown again on your next normal run)' $p.Evidence $p.Rule }
                Write-Host "$($found.Count) genre suggestions found; they wait for your next normal run (those tracks stay in $ReviewFolder)."
            }
            else {
                foreach ($p in $found) { Write-GenreLog $p.Row 'Online genre suggestion' $p.Source $p.Genre 'Awaiting your approval' $p.Evidence $p.Rule }
                Write-Host "Found a genre for $($found.Count) of $($searchable.Count) tracks:" -ForegroundColor Cyan
                Write-Host '  by service:'
                $found | Group-Object Source | Sort-Object Count -Descending | ForEach-Object { Write-Host ('    {0,-28} {1,5} tracks' -f $_.Name, $_.Count) }
                Write-Host '  by genre:'
                $found | Group-Object Genre | Sort-Object Count -Descending | ForEach-Object { Write-Host ('    {0,-28} {1,5} tracks' -f $_.Name, $_.Count) }
                Write-Host '  (SameArtist = guessed from the same artist''s other tracks - worth a look)'
                Write-Host ''
                Write-Host '  A = accept all   N = reject all (not remembered)   R = review one by one'
                $c = Read-Choice 'Choice [A/N/R]' @('a', 'n', 'r') 'a'

                if ($c -eq 'a') {
                    foreach ($p in $found) { Approve-Proposal $p 'Accepted' }
                }
                elseif ($c -eq 'n') {
                    foreach ($p in $found) { Write-GenreLog $p.Row 'Genre suggestion decision' $p.Source $p.Genre 'Rejected (reject all, not remembered)' $p.Evidence $p.Rule }
                }
                else {
                    Write-Host '  y = accept   n = wrong genre (remembered, never suggested again)   a = accept rest   q = skip rest (not remembered)' -ForegroundColor DarkGray
                    $acceptRest = $null
                    $idx = 0
                    foreach ($p in $found) {
                        $idx++
                        if ($acceptRest -eq $true) { Approve-Proposal $p 'Accepted (accept rest)'; continue }
                        if ($acceptRest -eq $false) {
                            Write-GenreLog $p.Row 'Genre suggestion decision' $p.Source $p.Genre 'Rejected (skip rest, not remembered)' $p.Evidence $p.Rule
                            continue
                        }
                        $b = Get-RoundedBpm ([string]$p.Row.BPM)
                        $bpmText = if ($null -ne $b) { "$b BPM" } else { 'BPM unknown' }
                        $keyText = if ($p.Row.Camelot) { [string]$p.Row.Camelot } else { 'key unknown' }
                        Write-Host ''
                        Write-Host ('[{0}/{1}] {2}' -f $idx, $found.Count, $p.Label) -ForegroundColor White
                        Write-Host ('  -> {0}   ({1}, {2}, via {3})' -f $p.Genre, $bpmText, $keyText, $p.Source) -ForegroundColor Cyan
                        if ($p.Rule) { Write-Host ('  ' + $p.Rule) -ForegroundColor DarkGray }
                        if ($p.Evidence) { Write-Host ('  evidence: ' + (Limit-Text $p.Evidence 220)) -ForegroundColor DarkGray }
                        $ans = Read-Choice 'Accept? [Y/n/a/q]' @('y', 'n', 'a', 'q') 'y'
                        switch ($ans) {
                            'y' { Approve-Proposal $p 'Accepted' }
                            'n' {
                                Set-Decision -Decisions $script:Decisions -Row $p.Row -Rejected $p.Genre -Source 'Rejected suggestion'
                                Write-GenreLog $p.Row 'Genre suggestion decision' $p.Source $p.Genre 'Rejected (remembered)' $p.Evidence $p.Rule
                            }
                            'a' { Approve-Proposal $p 'Accepted (accept rest)'; $acceptRest = $true }
                            'q' {
                                Write-GenreLog $p.Row 'Genre suggestion decision' $p.Source $p.Genre 'Rejected (skip rest, not remembered)' $p.Evidence $p.Rule
                                $acceptRest = $false
                            }
                        }
                    }
                }
                Save-Decisions $script:Decisions
            }
        }
    }
}

# 4c) new genres seen in your tags or online
if (Invoke-NewGenreReview) {
    foreach ($row in $rows) {
        if ($row.GenreSource -eq 'Tag' -and $row.Genre -ne $SamplesGenre) { $row.Genre = Get-StandardGenre $row.OriginalGenre }
    }
}

# =============================================================================
# STEP 5 - review the plan
# =============================================================================

foreach ($row in $rows) { Set-PlanDestination $row }

if (Test-Switch 'Duplicates') {
    $dupes = @(Find-Duplicates $rows)
    if ($dupes.Count -gt 0) {
        Write-Host ''
        Write-Host "$($dupes.Count) songs exist as more than one file (same artist and title). Examples:" -ForegroundColor Yellow
        foreach ($g in ($dupes | Select-Object -First 5)) {
            $first = $g.Group[0]
            Write-Host ('  {0} - {1}  ({2} files)' -f $first.Artist, (Limit-Text $first.Title 40), $g.Count)
        }
        Save-CsvAtomic @($dupes | ForEach-Object { $_.Group } | Select-Object Artist, Title, BPM, Key, SourcePath) $script:DuplicatesPath
        Write-Host "  Full list: $script:DuplicatesPath (TuneTidy copies them all; remove the extras in Rekordbox if you like)" -ForegroundColor DarkGray
    }
}

$planColumns = @('PlanId', 'Genre', 'Artist', 'Title', 'BPM', 'Key', 'Camelot', 'DestinationFolder', 'GenreSource', 'OriginalGenre', 'SourcePath')
Save-CsvAtomic @($rows | Select-Object -Property $planColumns) $script:PlanPath

Write-Host ''
Write-Host "Destination folders inside $script:OutputRoot (layout: $($LayoutNames[[string]$script:Settings.FolderLayout])):" -ForegroundColor Cyan
Show-FolderCounts $rows
Write-Host "Plan saved to: $script:PlanPath"
$tooLongCount = @($rows | Where-Object { $_.PathTooLong }).Count
if ($tooLongCount -gt 0) { Write-Host "$tooLongCount tracks would get a path longer than Windows allows and are skipped (use shorter folder names)." -ForegroundColor Yellow }

if ((Test-Switch 'OfferPlanEdit') -and (Read-YesNo 'Open the plan in your spreadsheet app to fix genres by hand?' $false $false)) {
    Invoke-Item -LiteralPath $script:PlanPath
    Write-Host 'Only the Genre column is read back. In Excel use "Save As > CSV UTF-8".' -ForegroundColor DarkGray
    Read-Host 'Edit the Genre column, save, CLOSE the spreadsheet, then press Enter here' | Out-Null

    $edited = $null
    while ($null -eq $edited) {
        try { $edited = @(Import-SafeCsv $script:PlanPath) }
        catch {
            Write-Warning "Could not read the plan (is it still open in another program?): $($_.Exception.Message)"
            Read-Host 'Close it, then press Enter to try again' | Out-Null
        }
    }
    $byId = @{}
    foreach ($r in $rows) { $byId[[string]$r.PlanId] = $r }
    $changed = 0
    $rulesChanged = $false
    foreach ($e in $edited) {
        $r = $byId[[string]$e.PlanId]
        if (-not $r -or -not $e.PSObject.Properties['Genre']) { continue }
        $typed = ([string]$e.Genre).Trim()
        if (-not $typed -or $typed -ceq [string]$r.Genre) { continue }
        if ($typed -in @('Review', $ReviewFolder)) { $r.Genre = ''; continue }
        $newGenre = Get-StandardGenre $typed
        # A genre you type yourself is added to your genre list.
        if (-not (Get-GenreRule $newGenre) -and $newGenre -ne $SamplesGenre) { Add-GenreRule $newGenre; $newGenre = (Get-StandardGenre $typed); $rulesChanged = $true }
        if ($newGenre -ceq [string]$r.Genre) { continue }
        Write-GenreLog $r 'Manual plan edit' 'Organizer plan' $newGenre 'Changed by you (remembered)' ("was: $($r.Genre)") ''
        $r.Genre = $newGenre
        $r.GenreSource = 'Decision'
        Set-Decision -Decisions $script:Decisions -Row $r -Genre $newGenre -Source 'Manual (plan edit)'
        $changed++
    }
    if ($rulesChanged) { Save-GenreRules }
    Save-Decisions $script:Decisions
    Write-Host "$changed genre change(s) applied and remembered."
    foreach ($row in $rows) { Set-PlanDestination $row }
    Save-CsvAtomic @($rows | Select-Object -Property $planColumns) $script:PlanPath
}

foreach ($row in $rows) {
    Write-GenreLog $row 'Final plan destination' 'Organizer plan' ([string]$row.Genre) ("Destination folder: $($row.DestinationFolder)") ("BPM=$($row.BPM); Key=$($row.Key); Camelot=$($row.Camelot)") ''
}

# =============================================================================
# STEP 6 - copy new tracks, move earlier copies that are now in the wrong place
# =============================================================================

$history = Get-History
$toCopy = @(); $toMove = @(); $toRecopy = @()
$inPlace = 0; $missing = 0
foreach ($r in $rows) {
    if (-not $r.SourceExists) { $missing++; continue }
    if ($r.PathTooLong -or -not $r.DestinationPath) { continue }
    $h = $history[$r.TrackKey]
    if (-not $h) { $toCopy += $r; continue }
    $copy = [string]$h.CopiedPath
    $copyOk = $copy -and (Test-InsideFolder $copy $script:OutputRoot) -and (Test-Path -LiteralPath $copy -PathType Leaf)
    if (-not $copyOk) { $toRecopy += $r; continue }
    $plannedDir = Get-FullPathOrNull (Split-Path -Parent $r.DestinationPath)
    if ((Get-FullPathOrNull (Split-Path -Parent $copy)) -ieq $plannedDir) { $inPlace++ }
    else { $toMove += [pscustomobject]@{ Row = $r; History = $h } }
}

Write-Host ''
Write-Host ("Already in the right place: {0} | new to copy: {1} | earlier copies to move: {2} | earlier copies you deleted: {3} | originals missing: {4}" -f $inPlace, $toCopy.Count, $toMove.Count, $toRecopy.Count, $missing)

$protectedFolders = @($script:DataRoot, $script:PlaylistRoot)
$moved = 0; $copied = 0; $failed = 0
try {
    if ($toMove.Count -gt 0 -and (Test-Switch 'OfferMoveCopies') -and -not $script:Unattended) {
        Write-Host ''
        Write-Host "$($toMove.Count) copies made in earlier runs are in a folder that no longer fits (corrected genre or new layout). Examples:" -ForegroundColor Cyan
        foreach ($m in ($toMove | Select-Object -First 5)) {
            $oldRel = (Split-Path -Parent $m.History.CopiedPath).Substring($script:OutputRoot.Length).TrimStart('\')
            Write-Host ("  {0}: {1} -> {2}" -f (Limit-Text $m.Row.Title 40), $oldRel, $m.Row.DestinationFolder)
        }
        if (Read-YesNo "Move these $($toMove.Count) copies? (only copies inside the output folder; originals are not touched)") {
            $emptied = @()
            foreach ($m in $toMove) {
                $from = [string]$m.History.CopiedPath
                $destDir = Split-Path -Parent $m.Row.DestinationPath
                $badDest = -not (Test-InsideFolder $destDir $script:OutputRoot)
                foreach ($pf in $protectedFolders) { if ((Test-InsideFolder $from $pf) -or (Test-InsideFolder $destDir $pf)) { $badDest = $true } }
                if (-not (Test-InsideFolder $from $script:OutputRoot) -or $badDest) {
                    Write-Warning "Skipped (outside the output folder): $from"
                    continue
                }
                if ($script:Preview) { Write-Host "Preview: would move $from -> $destDir"; continue }
                try {
                    New-Item -ItemType Directory -Force -Path $destDir | Out-Null
                    $to = Get-FreePath $m.Row.DestinationPath
                    Move-Item -LiteralPath $from -Destination $to -ErrorAction Stop
                    $emptied += (Split-Path -Parent $from)
                    $m.History.CopiedPath = $to
                    $moved++
                }
                catch { Write-Warning "Move failed for '$($m.Row.Title)': $($_.Exception.Message)"; $failed++ }
            }
            Remove-EmptyFolders $emptied $script:OutputRoot $protectedFolders
        }
    }

    if ($toRecopy.Count -gt 0 -and (Read-YesNo "$($toRecopy.Count) earlier copies are no longer in the output folder (deleted by you?). Copy them again?" $false $false)) {
        $toCopy += $toRecopy
    }

    if ($toCopy.Count -gt 0 -and (Read-YesNo "Copy $($toCopy.Count) files into '$script:OutputRoot' now? Your originals stay where they are." $true $true)) {
        if ($script:Preview) {
            foreach ($item in ($toCopy | Select-Object -First 20)) { Write-Host "Preview: would copy [$($item.DestinationFolder)]: $($item.Title)" }
            if ($toCopy.Count -gt 20) { Write-Host "Preview: ...and $($toCopy.Count - 20) more." }
        }
        elseif (Test-FreeSpace $toCopy) {
            foreach ($item in $toCopy) {
                $src = [string]$item.SourcePath
                $destDir = Split-Path -Parent $item.DestinationPath
                $badDest = -not (Test-InsideFolder $destDir $script:OutputRoot)
                foreach ($pf in $protectedFolders) { if (($destDir -ieq $pf) -or (Test-InsideFolder $destDir $pf)) { $badDest = $true } }
                if ($badDest) {
                    Write-Warning "Skipped (destination outside the output folder): $($item.DestinationFolder)"
                    $failed++
                    continue
                }
                try {
                    New-Item -ItemType Directory -Force -Path $destDir | Out-Null
                    $dest = Get-FreePath $item.DestinationPath   # never overwrite
                    Copy-Item -LiteralPath $src -Destination $dest -ErrorAction Stop
                    $copied++
                    $history[$item.TrackKey] = [pscustomobject]@{
                        TrackKey = $item.TrackKey; TrackID = $item.TrackID; Artist = $item.Artist; Title = $item.Title
                        SourcePath = $src; CopiedPath = $dest; CopiedAt = (Get-Date).ToString('s')
                    }
                    Write-Host "Copied [$($item.DestinationFolder)]: $($item.Title)"
                }
                catch {
                    Write-Warning "Copy failed for '$($item.Title)': $($_.Exception.Message)"
                    $failed++
                    if ($_.Exception -is [IO.IOException] -and $_.Exception.HResult -eq -2147024784) {   # disk full
                        Write-Warning 'The drive is full - stopping.'
                        break
                    }
                }
            }
        }
    }
}
finally {
    if ($history.Count -gt 0 -and -not $script:Preview) { Save-CsvAtomic @($history.Values) $script:HistoryPath }
}

# =============================================================================
# STEP 7 - playlists for Rekordbox (pointing at your originals, so cues are kept)
# =============================================================================

$playlists = Write-Playlists $rows
Save-Decisions $script:Decisions

Write-Host ''
Write-Host 'RUN FINISHED' -ForegroundColor Green
Write-Host "Copied: $copied | Copies moved: $moved | Failed: $failed"
Write-Host 'Original files were not moved, renamed, or deleted.'
if ($playlists -gt 0) {
    Write-Host "Playlists: $playlists in $script:PlaylistRoot" -ForegroundColor Cyan
    Write-Host '  Import them into Rekordbox (File > Import > Import Playlist) to get the same structure with all your cue points.' -ForegroundColor DarkGray
}
Write-Host "Your decisions: $script:DecisionsPath"
if ((Test-Switch 'GenreLog') -and $script:GenreLogStarted) { Write-Host "Genre matching log: $script:GenreLogPath" -ForegroundColor Cyan }
Write-Host ''
if (-not $script:Unattended) { Read-Host 'Press Enter to close this window' | Out-Null }
if ($failed -gt 0) { exit 1 }

}
catch {
    Write-Host ''
    Write-Host ('TuneTidy stopped because of an error: ' + $_.Exception.Message) -ForegroundColor Red
    if ($_.InvocationInfo) { Write-Host ('  (line ' + $_.InvocationInfo.ScriptLineNumber + ')') -ForegroundColor DarkGray }
    Write-Host '  Nothing was deleted. If this keeps happening, report it on GitHub with the message above.' -ForegroundColor DarkGray
    if (-not $script:Unattended) { Read-Host 'Press Enter to close this window' | Out-Null }
    exit 1
}
