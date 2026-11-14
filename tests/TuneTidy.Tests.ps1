# Pester 5 tests for TuneTidy. Run from the repository folder:
#   Invoke-Pester -Path .\tests
# They load the script without running it and test the parts that turn text into decisions.

BeforeAll {
    . (Join-Path (Split-Path -Parent $PSScriptRoot) 'TuneTidy.ps1')
}

Describe 'Get-Camelot' {
    It 'converts <Key> to <Expected>' -TestCases @(
        @{ Key = 'Am'; Expected = '8A' }
        @{ Key = 'A minor'; Expected = '8A' }
        @{ Key = 'Amin'; Expected = '8A' }
        @{ Key = 'C'; Expected = '8B' }
        @{ Key = 'C major'; Expected = '8B' }
        @{ Key = 'F#m'; Expected = '11A' }
        @{ Key = 'Dbm'; Expected = '12A' }
        @{ Key = '08a'; Expected = '8A' }
        @{ Key = '12B'; Expected = '12B' }
        @{ Key = '1m'; Expected = '8A' }
        @{ Key = '1d'; Expected = '8B' }
        @{ Key = '6m'; Expected = '1A' }
        @{ Key = ''; Expected = '' }
        @{ Key = 'nonsense'; Expected = '' }
    ) {
        Get-Camelot $Key | Should -Be $Expected
    }
}

Describe 'Get-RoundedBpm' {
    It 'reads <Text> as <Expected>' -TestCases @(
        @{ Text = '174.00'; Expected = 174 }
        @{ Text = '86.5'; Expected = 87 }
        @{ Text = '128'; Expected = 128 }
    ) {
        Get-RoundedBpm $Text | Should -Be $Expected
    }
    It 'returns nothing for <Text>' -TestCases @(@{ Text = '' }, @{ Text = '0' }, @{ Text = 'abc' }) {
        Get-RoundedBpm $Text | Should -BeNullOrEmpty
    }
}

Describe 'Genre rules' {
    It 'maps <Tag> to <Expected>' -TestCases @(
        @{ Tag = 'Hip-Hop'; Expected = 'Hip Hop' }
        @{ Tag = 'HipHop'; Expected = 'Hip Hop' }
        @{ Tag = 'rap'; Expected = 'Hip Hop' }
        @{ Tag = 'Drum & Bass'; Expected = 'DnB' }
        @{ Tag = "Drum'n'Bass"; Expected = 'DnB' }
        @{ Tag = 'Psy'; Expected = 'Psytrance' }
        @{ Tag = 'Tech House'; Expected = 'Tech House' }
        @{ Tag = 'House'; Expected = 'House' }
        @{ Tag = 'Hardcore Techno'; Expected = 'Hardcore' }
        @{ Tag = 'Trip Hop'; Expected = 'Downtempo' }
        @{ Tag = 'Electro Swing'; Expected = 'Electro Swing' }
    ) {
        Get-StandardGenre $Tag | Should -Be $Expected
    }

    It 'ignores junk tags' {
        Get-RuleForTag 'electronic' | Should -BeNullOrEmpty
    }

    It 'knows families' {
        Get-GenreFamily 'Tech House' | Should -Be 'House'
        Get-GenreFamily 'Psytrance' | Should -Be 'Trance'
        Get-GenreFamily 'Something New' | Should -Be 'Something New'
    }

    It 'tells another spelling apart from a possible new subgenre' {
        $techno = Get-GenreRule 'Techno'
        Test-ExactGenreName $techno 'TECHNO' | Should -BeTrue
        Test-ExactGenreName $techno 'Minimal Techno' | Should -BeFalse
    }

    It 'builds a pattern that matches only the whole name' {
        $p = New-GenrePattern 'Funk & Disco'
        (Get-Norm 'Funk & Disco') | Should -Match $p
        (Get-Norm 'Funk n Disco') | Should -Match $p
        (Get-Norm 'Funk') | Should -Not -Match $p
    }
}

Describe 'Select-Genre and BPM checks' {
    BeforeEach { $script:Settings = Get-DefaultSettings }

    It 'picks a subgenre with enough evidence' {
        (Select-Genre @{ 'Tech House' = 3.0; 'House' = 1.0 }).Genre | Should -Be 'Tech House'
    }
    It 'falls back to the family when no subgenre dominates' {
        (Select-Genre @{ 'House' = 2.0; 'Tech House' = 0.5 }).Genre | Should -Be 'House'
    }
    It 'gives up when sources disagree' {
        Select-Genre @{ 'House' = 1.0; 'Techno' = 1.0 } | Should -BeNullOrEmpty
    }
    It 'checks BPM: <Genre> at <Bpm> -> <Expected>' -TestCases @(
        @{ Genre = 'Tech House'; Bpm = '126'; Expected = $true }
        @{ Genre = 'Tech House'; Bpm = '174'; Expected = $false }
        @{ Genre = 'Dubstep'; Bpm = '70'; Expected = $true }
        @{ Genre = 'Hip Hop'; Bpm = '90'; Expected = $true }
    ) {
        Test-GenreBpm $Genre $Bpm | Should -Be $Expected
    }
}

Describe 'Folder layout' {
    BeforeEach { $script:Settings = Get-DefaultSettings }

    It 'groups keys: <Camelot> with <Size> per folder in <Notation> -> <Expected>' -TestCases @(
        @{ Camelot = '8A'; Size = 3; Notation = 'Camelot'; Expected = 'Keys 07-09' }
        @{ Camelot = '10B'; Size = 3; Notation = 'Camelot'; Expected = 'Keys 10-12' }
        @{ Camelot = '8A'; Size = 1; Notation = 'Camelot'; Expected = '08 - 8A 8B' }
        @{ Camelot = '8A'; Size = 1; Notation = 'OpenKey'; Expected = '01 - 1m 1d' }
        @{ Camelot = '8A'; Size = 1; Notation = 'Musical'; Expected = '08 - Am C' }
        @{ Camelot = '8A'; Size = 3; Notation = 'OpenKey'; Expected = 'Keys 01-03' }
        @{ Camelot = ''; Size = 3; Notation = 'Camelot'; Expected = 'No key' }
    ) {
        $script:Settings.KeyGroupSize = $Size
        $script:Settings.KeyNotation = $Notation
        Get-KeyLabel $Camelot | Should -Be $Expected
    }

    It 'makes BPM bands: <Genre> at <Bpm> -> <Expected>' -TestCases @(
        @{ Genre = 'Tech House'; Bpm = '126'; Expected = '125-129 BPM' }
        @{ Genre = 'House'; Bpm = '124'; Expected = '120-124 BPM' }
        @{ Genre = 'House'; Bpm = '64'; Expected = '125-129 BPM' }
        @{ Genre = 'House'; Bpm = ''; Expected = 'No BPM' }
    ) {
        Get-BpmLabel $Bpm $Genre | Should -Be $Expected
    }

    It 'uses custom BPM groups for a family: <Bpm> -> <Expected>' -TestCases @(
        @{ Bpm = '118'; Expected = '0-119 BPM' }
        @{ Bpm = '124'; Expected = '120-125 BPM' }
        @{ Bpm = '128'; Expected = '126-129 BPM' }
        @{ Bpm = '132'; Expected = '130+ BPM' }
    ) {
        $script:Settings.BpmGroups['House'] = '120,126,130'
        Get-BpmLabel $Bpm 'Tech House' | Should -Be $Expected
    }

    It 'builds the default one-level folder' {
        $item = [pscustomobject]@{ Genre = 'Tech House'; BPM = '126'; Camelot = '8A' }
        Get-TargetFolder $item | Should -Be 'Tech House\125-129 BPM - Keys 07-09'
    }
    It 'uses the family folder at the Broad detail level' {
        $script:Settings.DetailLevel = 'Broad'
        $item = [pscustomobject]@{ Genre = 'Tech House'; BPM = '126'; Camelot = '8A' }
        Get-TargetFolder $item | Should -Be 'House\125-129 BPM - Keys 07-09'
    }
    It 'builds two levels for GenreBpmKey' {
        $script:Settings.FolderLayout = 'GenreBpmKey'
        $item = [pscustomobject]@{ Genre = 'Techno'; BPM = '132'; Camelot = '2A' }
        Get-TargetFolder $item | Should -Be 'Techno\130-134 BPM\Keys 01-03'
    }
    It 'sends tracks without a genre to _Review' {
        Get-TargetFolder ([pscustomobject]@{ Genre = ''; BPM = '126'; Camelot = '8A' }) | Should -Be '_Review'
    }
}

Describe 'File names and search terms' {
    It 'reads "01 Artist - Title"' {
        $f = Get-NameFromFile 'C:\Music\01 Some Artist - Some Track.mp3'
        $f.Artist | Should -Be 'Some Artist'
        $f.Title | Should -Be 'Some Track'
    }
    It 'removes download-site names and turns the rest into a slug' {
        $f = Get-NameFromFile 'C:\Music\somesite.cc_artist-bad-boy-track-name-othersite.org.mp3'
        $f.Slug | Should -Be 'artist bad boy track name'
    }
    It 'keeps hyphenated words when removing a site name at the end' {
        (Get-NameFromFile 'C:\Music\Some Artist - Big Room-Anthem-site.com.mp3').Title | Should -Be 'Big Room-Anthem'
    }
    It 'drops "(Official Video)" from file names' {
        (Get-NameFromFile 'C:\Music\Artist - Title (Official Video).mp3').Title | Should -Be 'Title'
    }
    It 'drops "(Original Mix)" and finds remixers' {
        (New-Terms 'Artist' 'Track (Original Mix)' 'Tags').Title | Should -Be 'Track'
        $t = New-Terms 'Artist' 'Track (Other Artist Remix)' 'Tags'
        $t.Title | Should -Be 'Track'
        $t.Remixer | Should -Be 'Other Artist'
    }
    It 'keeps meaningful brackets' {
        (New-Terms 'ATB' '9 PM (Till I Come)' 'Tags').Title | Should -Be '9 PM (Till I Come)'
    }
    It 'ignores a channel name in the Artist field' {
        $t = Get-SearchTerms 'Proximity' 'Some Artist - Some Track'
        $t.Artist | Should -Be 'Some Artist'
        $t.Title | Should -Be 'Some Track'
    }
    It 'removes " - Topic" from YouTube artist channels' {
        (Get-SearchTerms 'ATB - Topic' '9 PM (Till I Come)').Artist | Should -Be 'ATB'
    }
}

Describe 'Small helpers' {
    It 'cleans folder names' {
        Format-FolderName 'CON' | Should -Be '_CON'
        Format-FolderName 'AC/DC' | Should -Be 'AC_DC'
        Format-FolderName '' | Should -Be 'Review'
    }
    It 'normalises text for matching' {
        Get-Norm "R$([char]0x00F6)yksopp & Robyn" | Should -Be 'royksopp and robyn'   # o with umlaut
    }
    It 'shortens very long file names but keeps the extension' {
        $p = Get-ShortPath 'C:\Music\Organized Music\Techno' (('x' * 300) + '.mp3')
        $p.Length | Should -BeLessOrEqual 250
        $p | Should -Match '\.mp3$'
    }
    It 'keeps raw tags through the cache' {
        $res = [pscustomobject]@{ Tags = @([pscustomobject]@{ Name = 'minimal techno'; Weight = 3 }); Evidence = 'test' }
        $back = ConvertFrom-CacheText (ConvertTo-CacheText $res)
        $back.Tags[0].Name | Should -Be 'minimal techno'
        $back.Tags[0].Weight | Should -Be 3
        $back.Evidence | Should -Be 'test'
    }
    It 'protects spreadsheet cells from formulas and reads them back unchanged' {
        $path = Join-Path $TestDrive 'safe.csv'
        Export-SafeCsv ([pscustomobject]@{ Title = '=HYPERLINK("x")' }) $path
        (Get-Content -LiteralPath $path -Raw) | Should -Match "'=HYPERLINK"
        (Import-SafeCsv $path)[0].Title | Should -Be '=HYPERLINK("x")'
    }
}

Describe 'Rekordbox playlists' {
    BeforeAll {
        $xmlText = @'
<?xml version="1.0" encoding="UTF-8"?>
<DJ_PLAYLISTS Version="1.0.0">
  <COLLECTION Entries="3">
    <TRACK TrackID="1" Name="One" Artist="A" Location="file://localhost/C:/Music/one.mp3"/>
    <TRACK TrackID="2" Name="Two" Artist="B" Location="file://localhost/C:/Music/two.mp3"/>
    <TRACK TrackID="3" Name="Three" Artist="C" Location="file://localhost/C:/Music/three.mp3"/>
  </COLLECTION>
  <PLAYLISTS>
    <NODE Type="0" Name="ROOT" Count="2">
      <NODE Name="#latestmix" Type="1" KeyType="0" Entries="2">
        <TRACK Key="1"/>
        <TRACK Key="2"/>
      </NODE>
      <NODE Type="0" Name="Gigs" Count="1">
        <NODE Name="Warmup" Type="1" KeyType="1" Entries="1">
          <TRACK Key="file://localhost/C:/Music/three.mp3"/>
        </NODE>
      </NODE>
    </NODE>
  </PLAYLISTS>
</DJ_PLAYLISTS>
'@
        $script:testXml = [xml]$xmlText
        $script:testXmlPath = Join-Path $TestDrive 'export.xml'
        Set-Content -LiteralPath $script:testXmlPath -Value $xmlText -Encoding UTF8
    }

    It 'reads the playlist tree with folders' {
        $tree = @(Get-RekordboxPlaylists $script:testXml)
        @($tree | Where-Object { -not $_.IsFolder }).Count | Should -Be 2
        ($tree | Where-Object { $_.Name -eq 'Warmup' }).Path | Should -Be 'Gigs/Warmup'
        ($tree | Where-Object { $_.Name -eq '#latestmix' }).Keys.Count | Should -Be 2
    }
    It 'finds playlists by name or folder path' {
        $tree = @(Get-RekordboxPlaylists $script:testXml)
        (Find-PlaylistsByName $tree @('#latestmix')).Path | Should -Be '#latestmix'
        (Find-PlaylistsByName $tree @('Gigs/Warmup')).Name | Should -Be 'Warmup'
        (Find-PlaylistsByName $tree @('warmup')).Path | Should -Be 'Gigs/Warmup'
    }
    It 'stops with a clear message for an unknown playlist' {
        $tree = @(Get-RekordboxPlaylists $script:testXml)
        { Find-PlaylistsByName $tree @('Nope') } | Should -Throw '*Nope*'
    }
    It 'reads the track count from the start of the file' {
        Get-XmlTrackCount $script:testXmlPath | Should -Be 3
    }
    It 'reads tracks with Windows paths' {
        $tracks = @(Get-TracksFromRekordboxXml $script:testXml)
        $tracks.Count | Should -Be 3
        $tracks[0].SourcePath | Should -Be 'C:\Music\one.mp3'
    }
}
