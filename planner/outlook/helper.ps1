# Planner's link to classic Outlook. Started once by the app and kept running while the app is open.
#
# Protocol: one JSON request per line on stdin, one JSON response per line on stdout (UTF-8).
#   {"id": 1, "op": "hello"}                                   -> {"id": 1, "ok": true, "version": 1, "connection": 700, "startedOutlook": true}
#   {"id": 2, "op": "events", "start": "2026-11-09", "end": "2026-11-16"}
#                                                              -> {"id": 2, "ok": true, "connected": true, "connection": 700, "rows": [...]}
#   {"id": 3, "op": "quit"}                                    -> {"id": 3, "ok": true}, then the helper exits
# Failures come back as {"id": n, "ok": false, "error": "..."}; the helper keeps running.
#
# Classic Outlook started through COM runs without a window. If this helper started it, the helper closes it
# again on quit (or when the app's end of stdin closes), unless you've since opened an Outlook window yourself.

$ErrorActionPreference = 'Stop'
$utf8 = New-Object System.Text.UTF8Encoding $false
[Console]::InputEncoding = $utf8
[Console]::OutputEncoding = $utf8

$Version = 1
$ConnectWaitSeconds = 20
$MaxRows = 3000
$script:ol = $null
$script:ns = $null
$script:startedOutlook = $false

function Release($obj) {
    if ($null -ne $obj -and [Runtime.InteropServices.Marshal]::IsComObject($obj)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj)
    }
}

function Connect-Outlook {
    if ($null -ne $script:ol) { return }
    # Ours to close if nothing was running, or if what's running is a background copy started by automation
    # (-Embedding), e.g. one left behind when an earlier helper was killed. One you opened yourself is left alone.
    $running = @(Get-CimInstance Win32_Process -Filter "Name='OUTLOOK.EXE'" -ErrorAction SilentlyContinue)
    $script:startedOutlook = ($running.Count -eq 0) -or
        (@($running | Where-Object { $_.CommandLine -notmatch '(?i)[-/]embedding' }).Count -eq 0)
    $script:ol = New-Object -ComObject Outlook.Application
    $script:ns = $script:ol.GetNamespace('MAPI')
}

function Disconnect-Outlook([bool]$closeIfOurs) {
    if ($null -eq $script:ol) { return }
    try {
        if ($closeIfOurs -and $script:startedOutlook -and $script:ol.Explorers.Count -eq 0 -and
            $script:ol.Inspectors.Count -eq 0) { $script:ol.Quit() }
    } catch { }
    Release $script:ns
    Release $script:ol
    $script:ns = $null
    $script:ol = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# 0 = not an Exchange account (nothing to wait for); 500-800 = connected to the server.
function Test-Connected([int]$mode) { return ($mode -eq 0 -or $mode -ge 500) }

function Wait-Connected {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while ($true) {
        $mode = [int]$script:ns.ExchangeConnectionMode
        if ((Test-Connected $mode) -or $sw.Elapsed.TotalSeconds -ge $ConnectWaitSeconds) { return $mode }
        Start-Sleep -Milliseconds 250
    }
}

function Get-Events([string]$startText, [string]$endText) {
    $ic = [Globalization.CultureInfo]::InvariantCulture
    $s = [datetime]::ParseExact($startText, 'yyyy-MM-dd', $ic)
    $e = [datetime]::ParseExact($endText, 'yyyy-MM-dd', $ic)
    $mode = Wait-Connected
    $cal = $script:ns.GetDefaultFolder(9)          # olFolderCalendar
    $items = $cal.Items
    $items.Sort('[Start]')
    $items.IncludeRecurrences = $true
    $filter = "[Start] < '" + $e.ToString('g') + "' AND [End] > '" + $s.ToString('g') + "'"
    try { $res = $items.Restrict($filter) } catch { $res = $items }
    $rows = New-Object System.Collections.ArrayList
    try {
        foreach ($a in $res) {
            try {
                if ($a.Start -ge $e) { break }
                if ($a.End -le $s) { continue }
                [void]$rows.Add([pscustomobject]@{
                    id       = [string]$a.EntryID + '|' + $a.Start.ToString('s')
                    title    = [string]$a.Subject
                    start    = $a.Start.ToString('s')
                    end      = $a.End.ToString('s')
                    allDay   = [bool]$a.AllDayEvent
                    location = [string]$a.Location
                    busy     = [int]$a.BusyStatus
                })
            } finally { Release $a }
            if ($rows.Count -ge $MaxRows) { break }
        }
    } finally {
        if (-not [object]::ReferenceEquals($res, $items)) { Release $res }
        Release $items
        Release $cal
    }
    return @{ ok = $true; connected = (Test-Connected $mode); connection = $mode; rows = @($rows) }
}

function Send($resp) {
    [Console]::Out.WriteLine(($resp | ConvertTo-Json -Compress -Depth 6))
    [Console]::Out.Flush()
}

$done = $false
try {
    while (-not $done) {
        $line = [Console]::In.ReadLine()
        if ($null -eq $line) { break }                 # the app has gone away
        $line = $line.Trim().TrimStart([char]0xFEFF)   # ignore a byte-order mark, if a caller sends one
        if (-not $line) { continue }
        $id = $null
        try {
            $req = $line | ConvertFrom-Json
            $id = $req.id
            switch ([string]$req.op) {
                'hello' {
                    Connect-Outlook
                    $resp = @{ ok = $true; version = $Version; connection = [int]$script:ns.ExchangeConnectionMode
                               startedOutlook = $script:startedOutlook }
                }
                'events' {
                    Connect-Outlook
                    $resp = Get-Events ([string]$req.start) ([string]$req.end)
                }
                'quit' {
                    Disconnect-Outlook $true
                    $resp = @{ ok = $true }
                    $done = $true
                }
                default { $resp = @{ ok = $false; error = "Unknown request '$($req.op)'." } }
            }
        } catch {
            $msg = $_.Exception.Message
            $resp = @{ ok = $false; error = $msg }
            # Outlook was closed or crashed under us: reconnect on the next request.
            if ($msg -match '0x800706BA|0x80010108|0x800706BE|RPC server|disconnected from its clients') {
                Disconnect-Outlook $false
            }
        }
        $resp.id = $id
        Send $resp
    }
} finally {
    Disconnect-Outlook $true
}
