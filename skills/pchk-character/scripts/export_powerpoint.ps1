param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [Parameter(Mandatory = $true)][string]$OutputDirectory,
    [switch]$Pdf,
    [switch]$Preview
)
$ErrorActionPreference = "Stop"
$app = $null
$deck = $null
$hadPowerPoint = @(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count -gt 0
try {
    $app = New-Object -ComObject PowerPoint.Application
    $deck = $app.Presentations.Open([System.IO.Path]::GetFullPath($InputPath), -1, 0, 0)
    $overflow = @()
    foreach ($slide in $deck.Slides) {
        foreach ($shape in $slide.Shapes) {
            if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame2.HasText -eq -1) {
                $text = $shape.TextFrame2.TextRange
                $frame = $shape.TextFrame2
                $bottom = $text.BoundTop + $text.BoundHeight
                $right = $text.BoundLeft + $text.BoundWidth
                if ($bottom -gt ($shape.Top + $shape.Height - $frame.MarginBottom + 1.2) -or
                    $right -gt ($shape.Left + $shape.Width - $frame.MarginRight + 5.0) -or
                    $right -gt ($deck.PageSetup.SlideWidth + 0.5)) {
                    $overflow += [ordered]@{
                        slide = $slide.SlideIndex
                        shape = $shape.Name
                        actual_bottom = [math]::Round($bottom, 2)
                        allowed_bottom = [math]::Round($shape.Top + $shape.Height, 2)
                        actual_right = [math]::Round($right, 2)
                        allowed_right = [math]::Round($shape.Left + $shape.Width, 2)
                    }
                }
            }
        }
    }
    $report = [ordered]@{
        status = $(if ($overflow.Count) { "overflow" } else { "ok" })
        slides = $deck.Slides.Count
        renderer = "Microsoft PowerPoint"
        overflow = @($overflow)
    }
    $report | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $OutputDirectory "report.json") -Encoding UTF8
    if ($overflow.Count) { throw "Text exceeds its layout boxes; see report.json." }
    if ($Preview) {
        foreach ($slide in $deck.Slides) {
            $file = Join-Path $OutputDirectory ("slide-{0}.png" -f $slide.SlideIndex)
            $slide.Export($file, "PNG", 1800, 1268)
        }
    }
    if ($Pdf) { $deck.SaveAs((Join-Path $OutputDirectory "sheet.pdf"), 32) }
} catch {
    Write-Error $_
    exit 2
} finally {
    if ($deck) { $deck.Close() }
    if ($app) {
        if (-not $hadPowerPoint) { $app.Quit() }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
    }
}
