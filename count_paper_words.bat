@echo off
setlocal
set "XRF_WORDCOUNT_SELF=%~f0"
set "XRF_WORDCOUNT_DOC=%~1"
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -Command "$src = Get-Content -LiteralPath $env:XRF_WORDCOUNT_SELF -Raw -Encoding UTF8; $code = ($src -split '(?m)^# === POWERSHELL ===\r?\n', 2)[1]; if (-not $code) { throw 'Embedded script not found' }; & ([scriptblock]::Create($code))"
echo.
pause
exit /b %errorlevel%

# === POWERSHELL ===
$ErrorActionPreference = 'Stop'
$docPath = $env:XRF_WORDCOUNT_DOC

if ([string]::IsNullOrWhiteSpace($docPath)) {
    Add-Type -AssemblyName System.Windows.Forms
    $picker = New-Object System.Windows.Forms.OpenFileDialog
    $picker.Title = 'Select your manuscript'
    $picker.Filter = 'Word documents (*.docx;*.doc)|*.docx;*.doc'
    if ($picker.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $docPath = $picker.FileName
}

if (-not (Test-Path -LiteralPath $docPath -PathType Leaf)) {
    Write-Host "File not found: $docPath"
    return
}

$word = $null
$doc = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($docPath, $false, $true)

    # Recognize the counted manuscript region and non-counted back matter.
    $intro = '(?i)^(?:\d+(?:\.\d+)*[.)]?\s*)?Introduction$'
    $conclusion = '(?i)^(?:\d+(?:\.\d+)*[.)]?\s*)?(?:Conclusions?|Concluding remarks)(?:\s+and\s+(?:outlook|future work))?$'
    $stop = '(?i)^(?:\d+(?:\.\d+)*[.)]?\s*)?(?:References|Bibliography|Acknowledg(?:e)?ments?|Funding|Author contributions?|CRediT authorship contribution statement|Data availability(?: statement)?|Declaration of (?:competing )?interests?|Conflicts? of interest|Appendix(?:\s+[A-Z0-9]+)?|Supplementary (?:material|information))$'
    $caption = '(?i)^(?:Fig(?:ure)?\.?|Table)\s*\d+[A-Za-z]?\s*[:.\-\u2013\u2014]'

    # These count as main section titles even if they were not formatted with Heading 1.
    $knownHeading = '(?i)^(?:\d+[.)]?\s*)?(?:Introduction|Materials?\s+and\s+Methods?|Methods?(?:\s+and\s+Materials?)?|Methodology|Experimental(?:\s+(?:Methods?|Procedure|Setup))?|Materials?|Methods?|Results?|Discussion|Results?\s+and\s+Discussion|Discussion\s+and\s+Results?|Conclusions?|Concluding\s+Remarks(?:\s+and\s+(?:Outlook|Future\s+Work))?|Outlook|Future\s+Work)$'

    $started = $false
    $sawConclusion = $false
    $mainWords = 0
    $captionWords = 0
    $tableWords = 0
    $countedParagraphs = 0
    $stopAt = '(end of document)'
    $sections = New-Object 'System.Collections.Generic.List[object]'
    $currentSection = $null

    foreach ($paragraph in $doc.Paragraphs) {
        $range = $paragraph.Range
        $line = [string]$range.Text
        $line = ($line -replace '[\r\n\a\v]', ' ').Trim()
        if (-not $started) {
            if ($line -match $intro) { $started = $true } else { continue }
        }

        $styleName = ''
        try { $styleName = [string]$range.Style.NameLocal } catch {}

        if ($line -match $conclusion) { $sawConclusion = $true }
        if ($line -match $stop -and ($sawConclusion -or $line -match '(?i)References|Bibliography')) {
            $stopAt = $line
            break
        }

        # Exclude any top-level heading that follows the Conclusion section.
        $isHeading1 = $styleName -match '(?i)^(?:Heading|Uberschrift|\u00dcberschrift)\s*1$'
        $isKnownHeading = $line -match $knownHeading
        $isNumberedTopLevel = ($line -match '^\d+[.)]\s+\S' -and $line.Length -le 120)
        $isTopLevelHeading = $isHeading1 -or $isKnownHeading -or $isNumberedTopLevel
        if ($sawConclusion -and $line -notmatch $conclusion -and $isTopLevelHeading) {
            $stopAt = $line
            break
        }

        # Word's own word-count algorithm (wdStatisticWords = 0).
        $words = [int]$range.ComputeStatistics(0)
        if ([bool]$range.Information(12)) { # wdWithInTable = 12
            $tableWords += $words
            continue
        }
        if ($styleName -match '(?i)Caption|Beschriftung' -or $line -match $caption) {
            $captionWords += $words
            continue
        }

        if ($isTopLevelHeading -and $line.Length -gt 0) {
            # Each major heading begins a new section; subsection text belongs to it.
            $currentSection = [pscustomobject]@{ Name = $line; Words = 0 }
            [void]$sections.Add($currentSection)
        }
        if ($null -eq $currentSection) {
            # Safety fallback; normally the Introduction heading creates the first section.
            $currentSection = [pscustomobject]@{ Name = 'Introduction'; Words = 0 }
            [void]$sections.Add($currentSection)
        }
        $currentSection.Words += $words
        $mainWords += $words
        $countedParagraphs++
    }

    if (-not $started) {
        Write-Host 'No standalone Introduction heading found. Check your section headings.'
        return
    }

    $messages = New-Object 'System.Collections.Generic.List[string]'
    [void]$messages.Add("Manuscript: $docPath")
    [void]$messages.Add('------------------------------------------------------------')
    [void]$messages.Add('COUNTED MAIN TEXT - BREAKDOWN BY SECTION')
    [void]$messages.Add('Percentages use the counted main-text total (not the 5,000-word limit).')
    [void]$messages.Add('')
    foreach ($section in $sections) {
        $pct = if ($mainWords -gt 0) { 100.0 * $section.Words / $mainWords } else { 0.0 }
        [void]$messages.Add(('{0,-42} {1,7:N0} words  {2,5:N1}%' -f $section.Name, $section.Words, $pct))
    }
    [void]$messages.Add('------------------------------------------------------------')
    [void]$messages.Add(('TOTAL COUNTED MAIN TEXT: {0:N0} words (100.0%)' -f $mainWords))
    [void]$messages.Add(('Share of 5,000-word limit: {0:N1}%' -f (100.0 * $mainWords / 5000)))
    [void]$messages.Add("Including excluded table-cell text: $($mainWords + $tableWords) words")
    [void]$messages.Add('------------------------------------------------------------')
    [void]$messages.Add("Excluded figure/table captions: $captionWords words")
    [void]$messages.Add("Excluded text inside tables:   $tableWords words")
    [void]$messages.Add('Excluded: title, abstract, keywords, references, acknowledgements,')
    [void]$messages.Add('          appendices and other sections outside the main text.')
    [void]$messages.Add('Included: major headings, subsection headings, and in-text citations.')
    [void]$messages.Add("Counted paragraphs: $countedParagraphs")
    [void]$messages.Add("Counting stopped at: $stopAt")
    if (-not $sawConclusion) {
        [void]$messages.Add('WARNING: No standalone Conclusion(s) heading detected.')
    }
    if ($sections.Count -le 1) {
        [void]$messages.Add('WARNING: Only one major section was detected. Use Heading 1 styles for major headings.')
    }

    $messages | ForEach-Object { Write-Host $_ }
    $report = [System.IO.Path]::ChangeExtension($docPath, '.wordcount.txt')
    $messages | Out-File -LiteralPath $report -Encoding UTF8
    Write-Host "`nReport saved to: $report"
} catch {
    Write-Host "ERROR: $($_.Exception.Message)"
    Write-Host 'This utility requires Microsoft Word installed on Windows.'
} finally {
    if ($null -ne $doc) { $doc.Close($false) }
    if ($null -ne $word) { $word.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
}
