param(
    [Parameter(Mandatory = $true)]
    [string]$InputDocx
)

$ErrorActionPreference = 'Stop'
$inputPath = (Resolve-Path -LiteralPath $InputDocx).Path
$word = $null
$doc = $null
$tocLayout = $null

function Clear-HeaderFooterStory {
    param([Parameter(Mandatory = $true)]$Story)

    try {
        $Story.LinkToPrevious = $false
    }
    catch {
        # The first section has no previous section to unlink from.
    }

    for ($index = $Story.Shapes.Count; $index -ge 1; $index--) {
        $Story.Shapes.Item($index).Delete()
    }
    for ($index = $Story.Range.InlineShapes.Count; $index -ge 1; $index--) {
        $Story.Range.InlineShapes.Item($index).Delete()
    }
    $Story.Range.Text = ''
}

function Set-TocPageLayout {
    param([Parameter(Mandatory = $true)]$Document)

    if ($Document.TablesOfContents.Count -eq 0) {
        return $null
    }

    $toc = $Document.TablesOfContents.Item(1)
    $tocSection = $toc.Range.Sections.Item(1)
    # Information(3) on the complete TOC range returns its active end page.
    # A multi-page TOC is valid, so inspect a collapsed duplicate at the
    # field's start when checking that it begins on the section's first page.
    $tocStart = $toc.Range.Duplicate
    $tocStart.Collapse(1)
    $tocPage = $tocStart.Information(3)
    $sectionStart = $tocSection.Range.Duplicate
    $sectionStart.Collapse(1)
    $sectionStartPage = $sectionStart.Information(3)

    if ($tocPage -ne $sectionStartPage) {
        throw "TOC must start on the first page of its section before first-page header/footer suppression can be applied."
    }

    # Match the reference layout: the TOC is internally page 1, but its first-page
    # header and footer are blank. The first body page therefore displays page 2.
    $tocSection.PageSetup.DifferentFirstPageHeaderFooter = -1
    Clear-HeaderFooterStory -Story $tocSection.Headers.Item(2)
    Clear-HeaderFooterStory -Story $tocSection.Footers.Item(2)

    $pageNumbers = $tocSection.Footers.Item(1).PageNumbers
    $pageNumbers.RestartNumberingAtSection = $true
    $pageNumbers.StartingNumber = 1

    return [pscustomobject]@{
        section = $tocSection.Index
        toc_page = $tocPage
        section_start_page = $sectionStartPage
        first_page_header_footer_hidden = $true
        numbering_start = 1
    }
}

function Set-TocFormatting {
    param([Parameter(Mandatory = $true)]$Document)

    $bodySection = $Document.Sections.Item($Document.Sections.Count)
    $pageSetup = $bodySection.PageSetup
    $rightTabPosition = $pageSetup.PageWidth - $pageSetup.LeftMargin - $pageSetup.RightMargin
    $styleIds = @(-20, -21, -22)
    $leftIndents = @(0.0, 20.16, 40.32)

    for ($index = 0; $index -lt $styleIds.Count; $index++) {
        $style = $Document.Styles.Item($styleIds[$index])
        $style.Font.Name = 'Times New Roman'
        $style.Font.NameAscii = 'Times New Roman'
        $style.Font.NameFarEast = '宋体'
        $style.Font.Size = 12
        $style.Font.Bold = 0
        $style.Font.Color = 0
        $style.ParagraphFormat.LeftIndent = $leftIndents[$index]
        $style.ParagraphFormat.RightIndent = 0
        $style.ParagraphFormat.FirstLineIndent = 0
        $style.ParagraphFormat.SpaceBefore = 0
        $style.ParagraphFormat.SpaceAfter = 0
        $style.ParagraphFormat.LineSpacingRule = 1
        $style.ParagraphFormat.TabStops.ClearAll()
        $style.ParagraphFormat.TabStops.Add($rightTabPosition, 2, 1) | Out-Null
    }

    foreach ($paragraph in $Document.Paragraphs) {
        $text = ($paragraph.Range.Text -replace '[\r\a]', '').Trim()
        if ($text -eq '目录') {
            $paragraph.Alignment = 1
            $paragraph.Format.FirstLineIndent = 0
            $paragraph.Format.LeftIndent = 0
            $paragraph.Format.RightIndent = 0
            $paragraph.Format.SpaceBefore = 0
            $paragraph.Format.SpaceAfter = 12
            $paragraph.Format.LineSpacingRule = 5
            $paragraph.Range.Font.Name = 'Times New Roman'
            $paragraph.Range.Font.NameAscii = 'Times New Roman'
            $paragraph.Range.Font.NameFarEast = '宋体'
            $paragraph.Range.Font.Size = 18
            $paragraph.Range.Font.Bold = 1
            $paragraph.Range.Font.Color = 0
        }
    }

    foreach ($toc in $Document.TablesOfContents) {
        $toc.Range.Font.Name = 'Times New Roman'
        $toc.Range.Font.NameAscii = 'Times New Roman'
        $toc.Range.Font.NameFarEast = '宋体'
        $toc.Range.Font.Size = 12
        $toc.Range.Font.Bold = 0
        $toc.Range.Font.Color = 0
    }
}

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($inputPath, $false, $false)

    if ($doc.TablesOfContents.Count -gt 0) {
        $tocLayout = Set-TocPageLayout -Document $doc
    }
    $doc.Repaginate()
    if ($doc.TablesOfContents.Count -gt 0) {
        foreach ($toc in $doc.TablesOfContents) {
            $toc.Update()
        }
        Set-TocFormatting -Document $doc
    }

    foreach ($storyType in 1..17) {
        try {
            $range = $doc.StoryRanges.Item($storyType)
        }
        catch {
            $range = $null
        }
        while ($null -ne $range) {
            if ($range.Fields.Count -gt 0) {
                $range.Fields.Update() | Out-Null
            }
            $range = $range.NextStoryRange
        }
    }

    $doc.Repaginate()
    if ($doc.TablesOfContents.Count -gt 0) {
        foreach ($toc in $doc.TablesOfContents) {
            $toc.UpdatePageNumbers()
        }
        Set-TocFormatting -Document $doc
        $doc.Repaginate()
        foreach ($toc in $doc.TablesOfContents) {
            $toc.UpdatePageNumbers()
        }
    }
    $doc.Save()

    [pscustomobject]@{
        path = $inputPath
        tables_of_contents = $doc.TablesOfContents.Count
        fields = $doc.Fields.Count
        toc_layout = $tocLayout
        updated = $true
    }
}
finally {
    if ($null -ne $doc) {
        $doc.Close(0)
    }
    if ($null -ne $word) {
        $word.Quit() | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
