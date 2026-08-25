if (-not ("AgenticDocs.NativeMethods" -as [type])) {
    Add-Type -Namespace AgenticDocs -Name NativeMethods -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint processId);
'@
}

function Get-AgenticWordProcessId {
    param(
        [Parameter(Mandatory = $true)][object]$Word,
        [int[]]$ExistingProcessIds = @()
    )

    [uint32]$candidate = 0
    try {
        [void][AgenticDocs.NativeMethods]::GetWindowThreadProcessId(
            [System.IntPtr]$Word.Hwnd,
            [ref]$candidate
        )
    } catch {}
    if ($candidate -gt 0) {
        return [int]$candidate
    }

    # Hidden Word automation instances do not expose Hwnd on every Office
    # build. Fall back to the new /Automation process created by this call.
    for ($attempt = 0; $attempt -lt 10; $attempt++) {
        $created = @(
            Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE'" -ErrorAction SilentlyContinue |
                Where-Object {
                    $_.CommandLine -match "/Automation\s+-Embedding" -and
                    $ExistingProcessIds -notcontains [int]$_.ProcessId
                } |
                Sort-Object CreationDate -Descending
        )
        if ($created.Count -gt 0) {
            return [int]$created[0].ProcessId
        }
        Start-Sleep -Milliseconds 100
    }
    return 0
}

function Stop-AgenticWordAutomation {
    param(
        [object]$Word,
        [object[]]$Documents = @(),
        [int]$WordProcessId = 0
    )

    foreach ($item in @($Documents)) {
        if ($null -eq $item) { continue }
        try { $item.Close(0) } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($item) } catch {}
    }
    if ($null -ne $Word) {
        try { $Word.DisplayAlerts = 0 } catch {}
        try { $Word.NormalTemplate.Saved = $true } catch {}
        try { $Word.Quit(0) } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Word) } catch {}
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if ($WordProcessId -le 0) { return }
    for ($attempt = 0; $attempt -lt 10; $attempt++) {
        if ($null -eq (Get-Process -Id $WordProcessId -ErrorAction SilentlyContinue)) { return }
        Start-Sleep -Milliseconds 100
    }

    # Only terminate the exact hidden automation instance created by this
    # script. Never touch an ordinary interactive Word process.
    $process = Get-CimInstance Win32_Process -Filter "ProcessId=$WordProcessId" -ErrorAction SilentlyContinue
    if ($null -ne $process -and $process.Name -eq "WINWORD.EXE" -and $process.CommandLine -match "/Automation\s+-Embedding") {
        Stop-Process -Id $WordProcessId -Force -ErrorAction SilentlyContinue
    }
}
