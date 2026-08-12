param(
    [string]$Source = (Join-Path $PSScriptRoot "..\..\2026-04-28-汇报.pptx"),
    [string]$Output = (Join-Path $PSScriptRoot "..\slides\AI科研作图-公开版.pptx"),
    [string]$PdfOutput = (Join-Path $PSScriptRoot "..\slides\AI科研作图-公开版.pdf")
)

$ErrorActionPreference = "Stop"

$sourcePath = [IO.Path]::GetFullPath($Source)
$outputPath = [IO.Path]::GetFullPath($Output)
$pdfPath = [IO.Path]::GetFullPath($PdfOutput)

New-Item -ItemType Directory -Force -Path ([IO.Path]::GetDirectoryName($outputPath)) | Out-Null

$app = New-Object -ComObject PowerPoint.Application
$presentation = $null

try {
    $presentation = $app.Presentations.Open($sourcePath, $true, $false, $false)

    # Public cover: remove personal and supervisor names while keeping attribution.
    $cover = $presentation.Slides.Item(1)
    $identityGroup = $cover.Shapes.Item(4)
    $identityGroup.Delete()
    $coverIdentity = $cover.Shapes.AddTextbox(1, 320, 338, 320, 42)
    $coverIdentity.TextFrame.TextRange.Text = "整理者：nanfengy"
    $coverIdentity.TextFrame.TextRange.Font.NameFarEast = "微软雅黑"
    $coverIdentity.TextFrame.TextRange.Font.Name = "Arial"
    $coverIdentity.TextFrame.TextRange.Font.Size = 22
    $coverIdentity.TextFrame.TextRange.Font.Bold = $true
    $coverIdentity.TextFrame.TextRange.Font.Color.RGB = 0x70401B
    $coverIdentity.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    $coverIdentity.TextFrame.MarginLeft = 0
    $coverIdentity.TextFrame.MarginRight = 0

    # Correct product-name typos in the public copy without changing the source.
    foreach ($slide in $presentation.Slides) {
        foreach ($shape in $slide.Shapes) {
            try {
                if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                    [void]$shape.TextFrame.TextRange.Replace("Nona Banana Pro", "Nano Banana Pro", 0, 0)
                    [void]$shape.TextFrame.TextRange.Replace("Comfy-UI", "ComfyUI", 0, 0)
                }
            }
            catch {
                # Some media and grouped shapes do not expose text ranges.
            }
        }
    }

    # Embedded videos make the source deck larger than GitHub's 100 MB file limit.
    # Keep the slide context and remove only media objects.
    foreach ($slideNumber in @(41, 71)) {
        $slide = $presentation.Slides.Item($slideNumber)
        for ($index = $slide.Shapes.Count; $index -ge 1; $index--) {
            if ($slide.Shapes.Item($index).Type -eq 16) {
                $slide.Shapes.Item($index).Delete()
            }
        }
    }

    # Replace removed videos with readable static notices.
    $videoSlide = $presentation.Slides.Item(41)
    for ($index = $videoSlide.Shapes.Count; $index -ge 1; $index--) {
        $shape = $videoSlide.Shapes.Item($index)
        try {
            if (
                $shape.HasTextFrame -and
                $shape.TextFrame.HasText -and
                $shape.TextFrame.TextRange.Text -match "《纸手机》|《代理人》"
            ) {
                $shape.Delete()
            }
        }
        catch {
        }
    }
    $videoPanel = $videoSlide.Shapes.AddShape(5, 150, 150, 660, 300)
    $videoPanel.Fill.ForeColor.RGB = 0xF8F2EE
    $videoPanel.Line.ForeColor.RGB = 0xEAD8C8
    $videoPanel.TextFrame.TextRange.Text = "动态案例说明`r`r原演示稿在此播放 AI 视频案例。`r为控制公开文件体积并避免转载第三方视频，公开版已移除内嵌媒体。"
    $videoPanel.TextFrame.TextRange.Font.NameFarEast = "微软雅黑"
    $videoPanel.TextFrame.TextRange.Font.Name = "Arial"
    $videoPanel.TextFrame.TextRange.Font.Size = 21
    $videoPanel.TextFrame.TextRange.Font.Color.RGB = 0x70401B
    $videoPanel.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    $videoPanel.TextFrame.VerticalAnchor = 3
    $videoPanel.TextFrame.MarginLeft = 30
    $videoPanel.TextFrame.MarginRight = 30

    $caseSlide = $presentation.Slides.Item(71)
    $caseNotes = @(
        @{
            Left = 130
            Top = 350
            Width = 293
            Height = 165
            Text = "第三步｜AI 生成`r原内嵌动态案例已移除"
        },
        @{
            Left = 565
            Top = 350
            Width = 302
            Height = 170
            Text = "第四步｜剪辑精修`r公开版保留流程与静态示意"
        }
    )
    foreach ($note in $caseNotes) {
        $panel = $caseSlide.Shapes.AddShape(
            5,
            $note.Left,
            $note.Top,
            $note.Width,
            $note.Height
        )
        $panel.Fill.ForeColor.RGB = 0xF8F2EE
        $panel.Line.ForeColor.RGB = 0xEAD8C8
        $panel.TextFrame.TextRange.Text = $note.Text
        $panel.TextFrame.TextRange.Font.NameFarEast = "微软雅黑"
        $panel.TextFrame.TextRange.Font.Name = "Arial"
        $panel.TextFrame.TextRange.Font.Size = 17
        $panel.TextFrame.TextRange.Font.Bold = $true
        $panel.TextFrame.TextRange.Font.Color.RGB = 0x70401B
        $panel.TextFrame.TextRange.ParagraphFormat.Alignment = 2
        $panel.TextFrame.VerticalAnchor = 3
        $panel.TextFrame.MarginLeft = 16
        $panel.TextFrame.MarginRight = 16
    }

    # Replace the institutional closing slide with a neutral public ending.
    $closing = $presentation.Slides.Item(74)
    $closingThanks = $closing.Shapes.Item(4)
    $closingReview = $closing.Shapes.Item(3)
    $closingLogo = $closing.Shapes.Item(2)
    $closingThanks.TextFrame.TextRange.Text = "感谢阅读"
    $closingReview.TextFrame.TextRange.Text = "欢迎交流、补充与指正"
    $closingLogo.Delete()

    # Strip personal document metadata and comments from the public copy.
    try {
        $presentation.RemoveDocumentInformation(99)
    }
    catch {
        Write-Warning "PowerPoint did not remove every metadata category: $($_.Exception.Message)"
    }

    $presentation.SaveAs($outputPath, 24)
    $presentation.SaveAs($pdfPath, 32)
}
finally {
    if ($presentation) {
        $presentation.Close()
        [Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
    }
    $app.Quit()
    [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# PowerPoint can retain an empty generic comment-author record after removing
# document information. Remove the orphaned OOXML part and its registrations.
$fileStream = [IO.File]::Open($outputPath, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite)
$archive = $null
try {
    $archive = [IO.Compression.ZipArchive]::new(
        $fileStream,
        [IO.Compression.ZipArchiveMode]::Update,
        $false
    )

    $commentAuthors = $archive.GetEntry("ppt/commentAuthors.xml")
    if ($commentAuthors) {
        $commentAuthors.Delete()
    }

    $xmlTargets = @(
        @{
            Path = "ppt/_rels/presentation.xml.rels"
            XPath = "//*[local-name()='Relationship' and contains(@Type, '/commentAuthors')]"
        },
        @{
            Path = "[Content_Types].xml"
            XPath = "//*[local-name()='Override' and @PartName='/ppt/commentAuthors.xml']"
        }
    )

    foreach ($target in $xmlTargets) {
        $entry = $archive.GetEntry($target.Path)
        if (-not $entry) {
            continue
        }

        $reader = [IO.StreamReader]::new($entry.Open())
        try {
            $content = $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }

        $document = [Xml.XmlDocument]::new()
        $document.PreserveWhitespace = $true
        $document.LoadXml($content)
        $nodes = @($document.SelectNodes($target.XPath))
        foreach ($node in $nodes) {
            [void]$node.ParentNode.RemoveChild($node)
        }

        $entry.Delete()
        $newEntry = $archive.CreateEntry($target.Path)
        $writer = [IO.StreamWriter]::new($newEntry.Open(), [Text.UTF8Encoding]::new($false))
        try {
            $document.Save($writer)
        }
        finally {
            $writer.Dispose()
        }
    }
}
finally {
    if ($archive) {
        $archive.Dispose()
    }
    $fileStream.Dispose()
}

Get-Item -LiteralPath $outputPath, $pdfPath |
    Select-Object Name, Length, LastWriteTime
