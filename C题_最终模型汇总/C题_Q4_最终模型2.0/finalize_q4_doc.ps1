param(
    [string]$DocumentPath = "",
    [string]$PdfPath = ""
)

$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $DocumentPath) {
    $builtDocument = Get-ChildItem -LiteralPath $here -Filter "*.docx" -File |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if (-not $builtDocument) {
        throw "No DOCX document found in $here"
    }
    $DocumentPath = $builtDocument.FullName
}
if (-not $PdfPath) {
    $PdfPath = [System.IO.Path]::ChangeExtension($DocumentPath, ".pdf")
}

$documentFullPath = (Resolve-Path -LiteralPath $DocumentPath).Path
$expectedRoot = (Resolve-Path -LiteralPath $here).Path
$pdfFullPath = [System.IO.Path]::GetFullPath($PdfPath)
if (-not $documentFullPath.StartsWith($expectedRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Document is outside the expected output directory: $documentFullPath"
}
if (-not $pdfFullPath.StartsWith($expectedRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "PDF is outside the expected output directory: $pdfFullPath"
}

$figures = @(
    @{ Marker = "__Q4_FIGURE_PRICE__"; File = "q4_price_profiles.svg"; WidthCm = 14.5 },
    @{ Marker = "__Q4_FIGURE_RESPONSE__"; File = "q4_optimization_response.svg"; WidthCm = 14.5 }
)

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$document = $null
try {
    $document = $word.Documents.Open($documentFullPath, $false, $false)
    foreach ($figure in $figures) {
        $source = Join-Path $here ("figures\" + $figure.File)
        if (-not (Test-Path -LiteralPath $source)) {
            throw "Missing vector figure: $source"
        }
        $range = $document.Content.Duplicate
        $found = $range.Find.Execute($figure.Marker)
        if (-not $found) {
            throw "Figure marker not found: $($figure.Marker)"
        }
        $range.Text = ""
        $picture = $document.InlineShapes.AddPicture($source, $false, $true, $range)
        $picture.LockAspectRatio = -1
        $picture.Width = $figure.WidthCm * 28.3464567
        $picture.AlternativeText = [System.IO.Path]::GetFileNameWithoutExtension($source)
    }
    $document.Save()
    if (Test-Path -LiteralPath $pdfFullPath) {
        Remove-Item -LiteralPath $pdfFullPath -Force
    }
    $document.ExportAsFixedFormat($pdfFullPath, 17)
    $pages = $document.ComputeStatistics(2)
    $document.Close($false)
    $document = $null
    Write-Output "document=$documentFullPath"
    Write-Output "pdf=$pdfFullPath"
    Write-Output "vector_figures=$($figures.Count)"
    Write-Output "pages=$pages"
}
finally {
    if ($null -ne $document) {
        $document.Close($false)
    }
    $word.Quit()
}
