param(
    [Parameter(Mandatory = $true)]
    [string]$InputDocx,
    [Parameter(Mandatory = $true)]
    [string]$OutputDir,
    [int]$Dpi = 144
)

$ErrorActionPreference = 'Stop'
$inputPath = (Resolve-Path -LiteralPath $InputDocx).Path
$outputPath = [System.IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Force -Path $outputPath | Out-Null
$pdfPath = Join-Path $outputPath 'document.pdf'

$word = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($inputPath, $false, $true)
    try {
        $doc.ExportAsFixedFormat($pdfPath, 17)
    }
    finally {
        $doc.Close(0)
    }
}
finally {
    if ($word) {
        $word.Quit() | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

$pdftoppm = Get-Command pdftoppm -ErrorAction SilentlyContinue
if (-not $pdftoppm) {
    throw 'Word 已生成 PDF，但未找到 pdftoppm，无法生成逐页 PNG。'
}
& $pdftoppm.Source -png -r $Dpi $pdfPath (Join-Path $outputPath 'page')
Get-ChildItem -LiteralPath $outputPath -Filter 'page-*.png' |
    Sort-Object Name |
    Select-Object FullName, Length
