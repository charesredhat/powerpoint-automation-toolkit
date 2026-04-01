param(
    [Parameter(Mandatory = $true)]
    [string]$TargetPresentation,

    [string]$OutputPresentation,
    [string]$LogPath
)

$tgt = (Resolve-Path -LiteralPath $TargetPresentation).Path

if (-not $OutputPresentation) {
    $OutputPresentation = $tgt -replace '\.pptx$', '_corrected.pptx'
}

if (-not $LogPath) {
    $LogPath = Join-Path (Split-Path $tgt) 'corrections_log.txt'
}

Write-Output "Target: $tgt"
Write-Output "Output: $OutputPresentation"
Write-Output "Log: $LogPath"

$pp = New-Object -ComObject PowerPoint.Application
$word = New-Object -ComObject Word.Application
$word.Visible = $false

try {
    $pres = $pp.Presentations.Open($tgt, [ref]$false, [ref]$false, [ref]$false)
    "Corrections log for $tgt`n" | Out-File -FilePath $LogPath -Encoding utf8

    for ($i = 1; $i -le $pres.Slides.Count; $i++) {
        $slide = $pres.Slides.Item($i)
        for ($s = 1; $s -le $slide.Shapes.Count; $s++) {
            $shape = $slide.Shapes.Item($s)
            try {
                if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                    $orig = $shape.TextFrame.TextRange.Text
                    if (-not [string]::IsNullOrWhiteSpace($orig)) {
                        $doc = $word.Documents.Add()
                        try {
                            $doc.Content.Text = $orig
                            $errors = $doc.SpellingErrors
                            if ($errors.Count -gt 0) {
                                $corrected = $orig
                                ("Slide {0} - Shape {1}: Found {2} spelling error(s)" -f $i, $s, $errors.Count) | Out-File -FilePath $LogPath -Append -Encoding utf8
                                foreach ($err in $errors) {
                                    $wrong = $err.Text
                                    $sugs = $word.GetSpellingSuggestions($wrong)
                                    if ($sugs.Count -gt 0) {
                                        $suggest = $sugs.Item(1).Name
                                        $pattern = "(?<!\\w)" + [regex]::Escape($wrong) + "(?!\\w)"
                                        $corrected = [regex]::Replace(
                                            $corrected,
                                            $pattern,
                                            [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $sugs.Item(1).Name }
                                        )
                                        ("    '{0}' -> '{1}'" -f $wrong, $suggest) | Out-File -FilePath $LogPath -Append -Encoding utf8
                                    } else {
                                        ("    '{0}' -> (no suggestion)" -f $wrong) | Out-File -FilePath $LogPath -Append -Encoding utf8
                                    }
                                }

                                $shape.TextFrame.TextRange.Text = $corrected
                                '    Applied corrected text.' | Out-File -FilePath $LogPath -Append -Encoding utf8
                            }
                        }
                        finally {
                            $doc.Close([ref]$false)
                        }
                    }
                }
            } catch {
                ("Slide {0} - Shape {1}: error processing ({2})" -f $i, $s, $_) | Out-File -FilePath $LogPath -Append -Encoding utf8
            }
        }
    }

    try {
        $pres.SaveAs($OutputPresentation)
        ("Saved corrected presentation: $OutputPresentation") | Out-File -FilePath $LogPath -Append -Encoding utf8
    } catch {
        ("Failed to save corrected presentation: $_") | Out-File -FilePath $LogPath -Append -Encoding utf8
    }
}
finally {
    if ($pres) { $pres.Close() }
    if ($pp) { $pp.Quit() }
    if ($word) { $word.Quit() }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Output "Done. Corrections log: $LogPath"
