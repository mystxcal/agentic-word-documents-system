param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [Parameter(Mandatory = $true)][string]$OutputPath
)

. (Join-Path $PSScriptRoot "Word-Automation.ps1")

$source = [System.IO.Path]::GetFullPath($InputPath)
$target = [System.IO.Path]::GetFullPath($OutputPath)
if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw "Input DOCX not found: $source" }
if ([System.IO.Path]::GetExtension($source).ToLowerInvariant() -ne '.docx') { throw "Input must be a DOCX: $source" }
if ([System.IO.Path]::GetExtension($target).ToLowerInvariant() -ne '.pdf') { throw "Output must be a PDF: $target" }
if (Test-Path -LiteralPath $target) { throw "Output already exists: $target" }

[System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($target)) | Out-Null
$word = $null
$document = $null
$wordProcessId = 0
$exportCompleted = $false
$oldUpdateFieldsAtPrint = $null
$oldUpdateLinksAtPrint = $null
try {
    $existingWordProcessIds = @(
        Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE'" -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty ProcessId
    )
    $word = New-Object -ComObject Word.Application
    $wordProcessId = Get-AgenticWordProcessId -Word $word -ExistingProcessIds $existingWordProcessIds
    $word.Visible = $false
    $word.DisplayAlerts = 0
    # The compiler has already finalized and saved every field. Word's global
    # "update fields before printing" preference is machine state, not
    # document state; leaving it enabled can mutate a finalized TOC while a
    # read-only document is being exported. Export the saved field results
    # exactly as authored so the same package produces the same PDF on every
    # workstation.
    $oldUpdateFieldsAtPrint = $word.Options.UpdateFieldsAtPrint
    $oldUpdateLinksAtPrint = $word.Options.UpdateLinksAtPrint
    $word.Options.UpdateFieldsAtPrint = $false
    $word.Options.UpdateLinksAtPrint = $false
    $document = $word.Documents.Open($source, $false, $true)
    # Word can retain stale pagination in an otherwise valid DOCX. Force one
    # in-memory layout pass so repeated headers, footers, and fields are
    # exported from the current package rather than its cached page state.
    $document.Repaginate()
    # wdExportFormatPDF = 17; optimize for print; export the complete document.
    $document.ExportAsFixedFormat($target, 17, $false, 0, 0, 1, 99999, 0, $true, $true, 1, $true, $true, $false)
    $exportCompleted = (Test-Path -LiteralPath $target) -and ((Get-Item -LiteralPath $target).Length -gt 1000)
    if (-not $exportCompleted) { throw "Word did not produce a usable PDF: $target" }
    [pscustomobject]@{ input = $source; output = $target; exported = $true } | ConvertTo-Json -Compress
} catch {
    if (-not $exportCompleted -and (Test-Path -LiteralPath $target)) { try { Remove-Item -LiteralPath $target -Force } catch {} }
    throw
} finally {
    if ($null -ne $word) {
        try { if ($null -ne $oldUpdateFieldsAtPrint) { $word.Options.UpdateFieldsAtPrint = $oldUpdateFieldsAtPrint } } catch {}
        try { if ($null -ne $oldUpdateLinksAtPrint) { $word.Options.UpdateLinksAtPrint = $oldUpdateLinksAtPrint } } catch {}
    }
    Stop-AgenticWordAutomation -Word $word -Documents @($document) -WordProcessId $wordProcessId
    $document = $null
    $word = $null
}
