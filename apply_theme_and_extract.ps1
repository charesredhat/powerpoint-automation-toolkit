param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePresentation,

    [Parameter(Mandatory = $true)]
    [string]$TargetPresentation,

    [string]$BackupPath,
    [string]$OutputPresentation,
    [string]$ExtractedTextPath
)

$src = (Resolve-Path -LiteralPath $SourcePresentation).Path
$tgt = (Resolve-Path -LiteralPath $TargetPresentation).Path

if (-not $BackupPath) {
    $BackupPath = $tgt -replace '\.pptx$', '_backup.pptx'
}

if (-not $OutputPresentation) {
    $OutputPresentation = $tgt -replace '\.pptx$', '_theme_applied.pptx'
}

if (-not $ExtractedTextPath) {
    $ExtractedTextPath = Join-Path (Split-Path $tgt) 'extracted_text.txt'
}

Copy-Item -LiteralPath $tgt -Destination $BackupPath -Force
Write-Output "Backup created: $BackupPath"

$pp = New-Object -ComObject PowerPoint.Application
$pp.Visible = $false

try {
    try {
        $presSrc = $pp.Presentations.Open($src, [ref]$false, [ref]$true, [ref]$false)
    } catch {
        $presSrc = $pp.Presentations.Open($src)
    }

    try {
        $presTgt = $pp.Presentations.Open($tgt, [ref]$false, [ref]$false, [ref]$false)
    } catch {
        $presTgt = $pp.Presentations.Open($tgt)
    }

    try {
        $presTgt.ApplyTemplate($src)
        Write-Output 'Applied template from source presentation.'
    } catch {
        Write-Output "ApplyTemplate failed: $_"
    }

    Set-Content -LiteralPath $ExtractedTextPath -Value '' -Encoding UTF8
    foreach ($i in 1..$presTgt.Slides.Count) {
        $slide = $presTgt.Slides.Item($i)
        Add-Content -LiteralPath $ExtractedTextPath -Value "--- Slide $i ---" -Encoding UTF8
        foreach ($shape in $slide.Shapes) {
            try {
                if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                    $text = $shape.TextFrame.TextRange.Text
                    if (-not [string]::IsNullOrWhiteSpace($text)) {
                        Add-Content -LiteralPath $ExtractedTextPath -Value $text -Encoding UTF8
                    }
                }
            } catch {
            }
        }

        Add-Content -LiteralPath $ExtractedTextPath -Value '' -Encoding UTF8
    }

    Write-Output "Extracted text written to: $ExtractedTextPath"

    try {
        $presTgt.SaveAs($OutputPresentation)
        Write-Output "Saved themed presentation as: $OutputPresentation"
    } catch {
        Write-Output "Failed to save themed presentation: $_"
    }
}
finally {
    if ($presSrc) { $presSrc.Close() }
    if ($presTgt) { $presTgt.Close() }
    if ($pp) { $pp.Quit() }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Output 'Done.'
