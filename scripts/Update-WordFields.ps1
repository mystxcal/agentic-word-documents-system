param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [Parameter(Mandatory = $true)][string]$OutputPath,
    [ValidateRange(1, 3)][int]$Passes = 2
)

. (Join-Path $PSScriptRoot "Word-Automation.ps1")

$source = [System.IO.Path]::GetFullPath($InputPath)
$target = [System.IO.Path]::GetFullPath($OutputPath)
if (-not (Test-Path -LiteralPath $source -PathType Leaf)) {
    throw "Input DOCX not found: $source"
}
if ([System.IO.Path]::GetExtension($source).ToLowerInvariant() -ne '.docx') {
    throw "Input must be a DOCX: $source"
}
if ($source -eq $target) {
    throw 'Input and output must be different; finalization never overwrites the working DOCX.'
}
if (Test-Path -LiteralPath $target) {
    throw "Output already exists: $target"
}

$targetDirectory = [System.IO.Path]::GetDirectoryName($target)
[System.IO.Directory]::CreateDirectory($targetDirectory) | Out-Null
Copy-Item -LiteralPath $source -Destination $target

try {
    # Updating a TOC creates fresh PAGEREF fields and bookmarks. Word can leave
    # the first generated reference stale until the package is saved and opened
    # by a new Word process. The default therefore performs two independent
    # application passes; the compiler uses one pass before and one after it
    # embeds component baselines.
    foreach ($documentPass in 1..$Passes) {
        $word = $null
        $document = $null
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
            $document = $word.Documents.Open($target, $false, $false)
            foreach ($toc in $document.TablesOfContents) { $toc.Update() }
            foreach ($tableOfFigures in $document.TablesOfFigures) { $tableOfFigures.Update() }
            foreach ($fieldPass in 1..2) {
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
            }
            $document.Save()
        } finally {
            Stop-AgenticWordAutomation -Word $word -Documents @($document) -WordProcessId $wordProcessId
            $document = $null
            $word = $null
        }
    }
    [pscustomobject]@{
        input = $source
        output = $target
        fields_updated = $true
        passes = $Passes
    } | ConvertTo-Json -Compress
} catch {
    if (Test-Path -LiteralPath $target) {
        try { Remove-Item -LiteralPath $target -Force } catch {}
    }
    throw
}
