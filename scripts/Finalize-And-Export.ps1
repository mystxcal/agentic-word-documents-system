# Finalize fields, save the editable DOCX, and export the matching PDF in one Word session.
param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [Parameter(Mandatory = $true)][string]$OutputDocxPath,
    [Parameter(Mandatory = $true)][string]$OutputPdfPath
)

. (Join-Path $PSScriptRoot "Word-Automation.ps1")

$source = [System.IO.Path]::GetFullPath($InputPath)
$docxTarget = [System.IO.Path]::GetFullPath($OutputDocxPath)
$pdfTarget = [System.IO.Path]::GetFullPath($OutputPdfPath)
if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw "Input DOCX not found: $source" }
if ([System.IO.Path]::GetExtension($source).ToLowerInvariant() -ne '.docx') { throw "Input must be a DOCX: $source" }
if ([System.IO.Path]::GetExtension($docxTarget).ToLowerInvariant() -ne '.docx') { throw "OutputDocxPath must be a DOCX" }
if ([System.IO.Path]::GetExtension($pdfTarget).ToLowerInvariant() -ne '.pdf') { throw "OutputPdfPath must be a PDF" }
foreach ($target in @($docxTarget, $pdfTarget)) {
    if (Test-Path -LiteralPath $target) { throw "Output already exists: $target" }
    [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($target)) | Out-Null
}
Copy-Item -LiteralPath $source -Destination $docxTarget

$word = $null
$document = $null
$wordProcessId = 0
$oldUpdateFieldsAtPrint = $null
$oldUpdateLinksAtPrint = $null
$completed = $false
try {
    $existingWordProcessIds = @(
        Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE'" -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty ProcessId
    )
    $word = New-Object -ComObject Word.Application
    $wordProcessId = Get-AgenticWordProcessId -Word $word -ExistingProcessIds $existingWordProcessIds
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $oldUpdateFieldsAtPrint = $word.Options.UpdateFieldsAtPrint
    $oldUpdateLinksAtPrint = $word.Options.UpdateLinksAtPrint
    $word.Options.UpdateFieldsAtPrint = $false
    $word.Options.UpdateLinksAtPrint = $false
    $document = $word.Documents.Open($docxTarget, $false, $false)
    foreach ($toc in $document.TablesOfContents) { $toc.Update() }
    foreach ($tableOfFigures in $document.TablesOfFigures) { $tableOfFigures.Update() }
    foreach ($storyType in 1..17) {
        try {
            $range = $document.StoryRanges.Item($storyType)
            while ($null -ne $range) {
                if ($range.Fields.Count -gt 0) { $range.Fields.Update() | Out-Null }
                $range = $range.NextStoryRange
            }
        } catch {
            # A document need not contain every Word story type.
        }
    }
    foreach ($toc in $document.TablesOfContents) { $toc.UpdatePageNumbers() }
    foreach ($tableOfFigures in $document.TablesOfFigures) { $tableOfFigures.UpdatePageNumbers() }
    # Repeated page furniture can be omitted when Word exports a package whose
    # cached pagination predates the current sections or fields.
    $document.Repaginate()
    $document.Save()
    $document.ExportAsFixedFormat($pdfTarget, 17, $false, 0, 0, 1, 99999, 0, $true, $true, 1, $true, $true, $false)
    if (-not (Test-Path -LiteralPath $pdfTarget) -or (Get-Item -LiteralPath $pdfTarget).Length -le 1000) {
        throw "Word did not produce a usable PDF: $pdfTarget"
    }
    $completed = $true
    [pscustomobject]@{
        input = $source
        docx = $docxTarget
        pdf = $pdfTarget
        fields_updated = $true
        exported = $true
        word_sessions = 1
    } | ConvertTo-Json -Compress
} finally {
    if ($null -ne $word) {
        try { if ($null -ne $oldUpdateFieldsAtPrint) { $word.Options.UpdateFieldsAtPrint = $oldUpdateFieldsAtPrint } } catch {}
        try { if ($null -ne $oldUpdateLinksAtPrint) { $word.Options.UpdateLinksAtPrint = $oldUpdateLinksAtPrint } } catch {}
    }
    Stop-AgenticWordAutomation -Word $word -Documents @($document) -WordProcessId $wordProcessId
    $document = $null
    $word = $null
    if (-not $completed) {
        foreach ($target in @($docxTarget, $pdfTarget)) {
            if (Test-Path -LiteralPath $target) { try { Remove-Item -LiteralPath $target -Force } catch {} }
        }
    }
}
