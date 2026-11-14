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
        @{ Tag = 'Drum & Bass'; Expected = 'DnB' }
        @{ Tag = 'Drum and Bass'; Expected = 'DnB' }
        @{ Tag = "Drum'n'Bass"; Expected = 'DnB' }
        @{ Tag = 'DnB'; Expected = 'DnB' }
        @{ Tag = 'Jump Up Drum & Bass'; Expected = 'Jump Up' }
        @{ Tag = 'Liquid Funk'; Expected = 'Liquid DnB' }
        @{ Tag = 'neuro'; Expected = 'Neurofunk' }
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
        Get-GenreFamily 'Neurofunk' | Should -Be 'DnB'
        Get-GenreFamily 'Tech House' | Should -Be 'House'
        Get-GenreFamily 'Something New' | Should -Be 'Something New'
    }

    It 'tells another spelling apart from a possible new subgenre' {
        $dnb = Get-GenreRule 'DnB'
        Test-ExactGenreName $dnb 'Drum & Bass' | Should -BeTrue
        Test-ExactGenreName $dnb 'Minimal Drum & Bass' | Should -BeFalse
    }

    It 'builds a pattern that matches only the whole name' {
        $p = New-GenrePattern 'Minimal Drum & Bass'
        (Get-Norm 'Minimal Drum & Bass') | Should -Match $p
        (Get-Norm 'Minimal Drum n Bass') | Should -Match $p
        (Get-Norm 'Drum & Bass') | Should -Not -Match $p
    }
}

Describe 'Select-Genre and BPM checks' {
    BeforeEach { $script:Settings = Get-DefaultSettings }

    It 'picks a subgenre with enough evidence' {
        (Select-Genre @{ 'Liquid DnB' = 3.0; 'DnB' = 1.0 }).Genre | Should -Be 'Liquid DnB'
    }
    It 'falls back to the family when no subgenre dominates' {
        (Select-Genre @{ 'DnB' = 2.0; 'Liquid DnB' = 0.5 }).Genre | Should -Be 'DnB'
    }
    It 'gives up when sources disagree' {
        Select-Genre @{ 'House' = 1.0; 'DnB' = 1.0 } | Should -BeNullOrEmpty
    }
    It 'checks BPM: <Genre> at <Bpm> -> <Expected>' -TestCases @(
        @{ Genre = 'Neurofunk'; Bpm = '174'; Expected = $true }
        @{ Genre = 'Neurofunk'; Bpm = '87'; Expected = $true }
        @{ Genre = 'Neurofunk'; Bpm = '125'; Expected = $false }
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

    It 'makes BPM groups: <Genre> at <Bpm> -> <Expected>' -TestCases @(
        @{ Genre = 'Liquid DnB'; Bpm = '174'; Expected = '170-175 BPM' }
        @{ Genre = 'Liquid DnB'; Bpm = '87'; Expected = '170-175 BPM' }
        @{ Genre = 'Neurofunk'; Bpm = '178'; Expected = '176+ BPM' }
        @{ Genre = 'DnB'; Bpm = '165'; Expected = '0-169 BPM' }
        @{ Genre = 'House'; Bpm = '124'; Expected = '120-124 BPM' }
        @{ Genre = 'House'; Bpm = ''; Expected = 'No BPM' }
    ) {
        Get-BpmLabel $Bpm $Genre | Should -Be $Expected
    }

    It 'builds the default one-level folder' {
        $item = [pscustomobject]@{ Genre = 'Liquid DnB'; BPM = '174'; Camelot = '8A' }
        Get-TargetFolder $item | Should -Be 'Liquid DnB\170-175 BPM - Keys 07-09'
    }
    It 'uses the family folder at the Broad detail level' {
        $script:Settings.DetailLevel = 'Broad'
        $item = [pscustomobject]@{ Genre = 'Liquid DnB'; BPM = '174'; Camelot = '8A' }
        Get-TargetFolder $item | Should -Be 'DnB\170-175 BPM - Keys 07-09'
    }
    It 'builds two levels for GenreBpmKey' {
        $script:Settings.FolderLayout = 'GenreBpmKey'
        $item = [pscustomobject]@{ Genre = 'Neurofunk'; BPM = '174'; Camelot = '2A' }
        Get-TargetFolder $item | Should -Be 'Neurofunk\170-175 BPM\Keys 01-03'
    }
    It 'sends tracks without a genre to _Review' {
        Get-TargetFolder ([pscustomobject]@{ Genre = ''; BPM = '174'; Camelot = '8A' }) | Should -Be '_Review'
    }
}

Describe 'File names and search terms' {
    It 'reads "01 Artist - Title"' {
        $f = Get-NameFromFile 'C:\Music\01 Netsky - Rio.mp3'
        $f.Artist | Should -Be 'Netsky'
        $f.Title | Should -Be 'Rio'
    }
    It 'turns a web-rip file name into a slug' {
        $f = Get-NameFromFile 'C:\Music\ytmp3free.cc_acray-bad-boy-cray-tunez-youtubemp3free.org.mp3'
        $f.Slug | Should -Be 'acray bad boy cray tunez'
    }
    It 'drops "(Official Video)" from file names' {
        (Get-NameFromFile 'C:\Music\Artist - Title (Official Video).mp3').Title | Should -Be 'Title'
    }
    It 'drops "(Original Mix)" and finds remixers' {
        (New-Terms 'Netsky' 'Rio (Original Mix)' 'Tags').Title | Should -Be 'Rio'
        $t = New-Terms 'Netsky' 'Rio (Camo & Krooked Remix)' 'Tags'
        $t.Title | Should -Be 'Rio'
        $t.Remixer | Should -Be 'Camo & Krooked'
    }
    It 'keeps meaningful brackets' {
        (New-Terms 'ATB' '9 PM (Till I Come)' 'Tags').Title | Should -Be '9 PM (Till I Come)'
    }
    It 'ignores a channel name in the Artist field' {
        $t = Get-SearchTerms 'DnB Allstars' 'Gino - Du Du Du'
        $t.Artist | Should -Be 'Gino'
        $t.Title | Should -Be 'Du Du Du'
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
        $p = Get-ShortPath 'C:\Music\Organized Music\DnB' (('x' * 300) + '.mp3')
        $p.Length | Should -BeLessOrEqual 250
        $p | Should -Match '\.mp3$'
    }
    It 'keeps raw tags through the cache' {
        $res = [pscustomobject]@{ Tags = @([pscustomobject]@{ Name = 'liquid funk'; Weight = 3 }); Evidence = 'test' }
        $back = ConvertFrom-CacheText (ConvertTo-CacheText $res)
        $back.Tags[0].Name | Should -Be 'liquid funk'
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
