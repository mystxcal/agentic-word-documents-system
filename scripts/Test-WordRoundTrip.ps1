param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [Parameter(Mandatory = $true)][string]$OutputPath,
    [string]$Marker = "[AGENTIC DOCS WORD EDITABILITY CHECK]",
    [string]$FindText,
    [string]$ReplacementText
)

. (Join-Path $PSScriptRoot "Word-Automation.ps1")

$ErrorActionPreference = "Stop"
$inputFile = (Resolve-Path -LiteralPath $InputPath).Path
$outputFile = [System.IO.Path]::GetFullPath($OutputPath)
if (Test-Path -LiteralPath $outputFile) {
    throw "Refusing to overwrite round-trip output: $outputFile"
}
[System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($outputFile)) | Out-Null
$inputHashBefore = (Get-FileHash -LiteralPath $inputFile -Algorithm SHA256).Hash

$word = $null
$document = $null
$check = $null
$wordProcessId = 0
try {
    $existingWordProcessIds = @(
        Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE'" -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty ProcessId
    )
    $word = New-Object -ComObject Word.Application
    $wordProcessId = Get-AgenticWordProcessId -Word $word -ExistingProcessIds $existingWordProcessIds
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($inputFile, $false, $false)
    if ($FindText) {
        if (-not $ReplacementText) { throw "ReplacementText is required when FindText is supplied" }
        $range = $document.Content
        $find = $range.Find
        $find.ClearFormatting()
        $find.Text = $FindText
        $find.Replacement.ClearFormatting()
        $find.Replacement.Text = $ReplacementText
        $found = $find.Execute($FindText, $false, $false, $false, $false, $false, $true, 1, $false, $ReplacementText, 1)
        if (-not $found) { throw "The requested in-document edit target was not found" }
        $expectedText = $ReplacementText
        $editKind = "replace-inside-document"
    }
    else {
        $range = $document.Content
        $range.Collapse(0)
        $range.InsertAfter("`r$Marker")
        $expectedText = $Marker
        $editKind = "append"
    }
    $document.SaveAs2($outputFile, 16)
    $document.Close(0)
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
    $document = $null

    $check = $word.Documents.Open($outputFile, $false, $true)
    $markerPresent = $check.Content.Text.Contains($expectedText)
    $check.Close(0)
    if (-not $markerPresent) {
        throw "Word reopened the saved file but the requested edit was missing"
    }
}
finally {
    Stop-AgenticWordAutomation -Word $word -Documents @($document, $check) -WordProcessId $wordProcessId
    $document = $null
    $check = $null
    $word = $null
}

$inputHashAfter = (Get-FileHash -LiteralPath $inputFile -Algorithm SHA256).Hash
if ($inputHashBefore -ne $inputHashAfter) {
    throw "The editability test modified its input file"
}

[ordered]@{
    ok = $true
    input = $inputFile
    input_unchanged = $true
    output = $outputFile
    output_sha256 = (Get-FileHash -LiteralPath $outputFile -Algorithm SHA256).Hash
    marker = $Marker
    edit_kind = $editKind
    reopened_in_word = $true
} | ConvertTo-Json -Depth 4 -Compress
