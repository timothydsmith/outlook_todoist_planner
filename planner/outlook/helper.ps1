# Planner's link to classic Outlook. Started once by the app and kept running while the app is open.
#
# Protocol: one JSON request per line on stdin, one JSON response per line on stdout (UTF-8).
#   {"id": 1, "op": "hello"}                                   -> {"id": 1, "ok": true, "version": 1, "connection": 700, "startedOutlook": true}
#   {"id": 2, "op": "events", "start": "2026-11-09", "end": "2026-11-16"}
#                                                              -> {"id": 2, "ok": true, "connected": true, "connection": 700, "rows": [...]}
#   {"id": 3, "op": "quit"}                                    -> {"id": 3, "ok": true}, then the helper exits
# Failures come back as {"id": n, "ok": false, "error": "..."}; the helper keeps running.
#
# Events Planner makes carry the user property PlannerTaskId. These requests only ever touch such events:
#   tagged                    -> {connected, connection, items: [item...]}   every tagged event in the calendar
#   put {entryId?, taskId, subject, body, start, end, busy}  -> {entryId, start, end}   create, or update in place
#   remove {entryId}          -> {}                   moves it to Deleted Items (nothing if already gone)
#   flush                     -> {}                   send/receive, so changes reach the server promptly
# item = {entryId, taskId, subject, start, end, allDay, busy}.
#
# `tagged` reads the calendar folder's table, not the items themselves: a long-running Outlook session can
# hand back an out-of-date cached copy of an item opened by id (even one that has since been deleted),
# whereas the table always reflects what's in the folder now. An event missing from it has been deleted
# or moved out of the calendar.
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
$TagName = 'PlannerTaskId'
# The same user property, addressed by its MAPI name, for table queries (PS_PUBLIC_STRINGS namespace).
$TagDasl = 'http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/' + $TagName
$AllDayDasl = 'http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/8215000B'
$BusyDasl = 'http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/82050003'
# A table's EntryID column can be a short-term id; this is the item's permanent EntryID (PR_LONGTERM_ENTRYID_FROM_TABLE).
$LongIdDasl = 'http://schemas.microsoft.com/mapi/proptag/0x66700102'
$Invariant = [Globalization.CultureInfo]::InvariantCulture

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
                    taskId   = (Get-Tag $a)
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

# ---- Planner's own events (tagged with PlannerTaskId) ----------------------------------------------

function Get-Tag($item) {
    $props = $null; $p = $null
    try {
        $props = $item.UserProperties
        $p = $props.Find($TagName)
        if ($null -eq $p) { return $null }
        $v = [string]$p.Value
        if ($v) { return $v } else { return $null }
    } catch { return $null }
    finally { Release $p; Release $props }
}

# MAPI_E_NOT_FOUND: the item has been permanently deleted (or never existed). Anything else is a real error.
function Test-NotFound($err) {
    $ex = $err.Exception
    while ($null -ne $ex) {
        if ($ex.HResult -eq -2147221233 -or $ex.Message -match '8004010F') { return $true }
        $ex = $ex.InnerException
    }
    return $false
}

function Find-PlannerItem([string]$entryId) {
    try { return $script:ns.GetItemFromID($entryId) }
    catch { if (Test-NotFound $_) { return $null } else { throw } }
}

function Get-Tagged {
    $mode = Wait-Connected
    $cal = $script:ns.GetDefaultFolder(9)
    $table = $null; $cols = $null
    $out = New-Object System.Collections.ArrayList
    try {
        $table = $cal.GetTable("@SQL=""$TagDasl"" IS NOT NULL")
        $cols = $table.Columns
        $cols.RemoveAll()
        foreach ($c in @($LongIdDasl, 'Subject', 'Start', 'End', $AllDayDasl, $BusyDasl, $TagDasl, 'EntryID')) { [void]$cols.Add($c) }
        while (-not $table.EndOfTable) {
            $r = $table.GetNextRow()
            try {
                $tag = [string]$r.Item($TagDasl)
                if ($tag) {
                    $eid = ''
                    try { $eid = [string]$r.BinaryToString(1) } catch { }   # column 1 = the permanent id
                    if (-not $eid) { $eid = [string]$r.Item('EntryID') }
                    [void]$out.Add(@{ entryId = $eid; taskId = $tag; subject = [string]$r.Item('Subject')
                                      start = ([datetime]$r.Item('Start')).ToString('s'); end = ([datetime]$r.Item('End')).ToString('s')
                                      allDay = [bool]$r.Item($AllDayDasl); busy = [int]$r.Item($BusyDasl) })
                }
            } finally { Release $r }
            if ($out.Count -ge $MaxRows) { break }
        }
    } finally {
        Release $cols
        Release $table
        Release $cal
    }
    return @{ ok = $true; connected = (Test-Connected $mode); connection = $mode; items = @($out) }
}

function Set-PlannerItem($req) {
    $taskId = [string]$req.taskId
    if (-not $taskId) { throw 'No task id given.' }
    $start = [datetime]::ParseExact([string]$req.start, 'yyyy-MM-ddTHH:mm:ss', $Invariant)
    $end = [datetime]::ParseExact([string]$req.end, 'yyyy-MM-ddTHH:mm:ss', $Invariant)
    if ($end -le $start) { throw 'The end time must be after the start time.' }
    $a = $null
    if ($req.entryId) {
        $a = Find-PlannerItem ([string]$req.entryId)
        if ($null -eq $a) { throw 'That Outlook event no longer exists.' }
        if ((Get-Tag $a) -ne $taskId) { Release $a; throw "That Outlook event wasn't made by Planner for this task." }
    } else {
        $a = $script:ol.CreateItem(1)                  # olAppointmentItem
        $props = $a.UserProperties
        $p = $props.Add($TagName, 1, $false)           # olText; don't add a column to the calendar folder
        $p.Value = $taskId
        Release $p; Release $props
    }
    try {
        $a.Subject = [string]$req.subject
        $a.Body = [string]$req.body
        if ($a.AllDayEvent) { $a.AllDayEvent = $false }
        $a.Start = $start
        $a.End = $end
        $a.BusyStatus = [int]$req.busy
        $a.ReminderSet = $false
        $a.Save()
        return @{ ok = $true; entryId = [string]$a.EntryID; start = $a.Start.ToString('s'); end = $a.End.ToString('s') }
    } finally { Release $a }
}

function Remove-PlannerItem([string]$entryId) {
    $a = Find-PlannerItem $entryId
    if ($null -eq $a) { return @{ ok = $true } }
    try {
        if (-not (Get-Tag $a)) { throw "That Outlook event wasn't made by Planner." }
        try { $a.Delete() } catch { if (-not (Test-NotFound $_)) { throw } }   # already gone is fine
    } finally { Release $a }
    return @{ ok = $true }
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
                'tagged' { Connect-Outlook; $resp = Get-Tagged }
                'put'    { Connect-Outlook; $resp = Set-PlannerItem $req }
                'remove' { Connect-Outlook; $resp = Remove-PlannerItem ([string]$req.entryId) }
                'flush'  {
                    Connect-Outlook
                    $mode = Wait-Connected                     # send/receive fails until Outlook has logged on
                    if (Test-Connected $mode) { $script:ns.SendAndReceive($false) }
                    $resp = @{ ok = $true; connected = (Test-Connected $mode) }
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
        # Let go of every item object this request touched, so Outlook doesn't keep (and later hand back)
        # stale cached copies of them.
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
} finally {
    Disconnect-Outlook $true
}
