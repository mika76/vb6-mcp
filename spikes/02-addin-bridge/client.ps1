# Spike 2 client. Runs NON-elevated. Finds the add-in via its discovery file and
# sends requests over loopback TCP. Usage:
#   client.ps1                      -> PING + LIST
#   client.ps1 READ Project1 Form1  -> live code of one component
param([string]$Verb = '', [string[]]$Rest = @())

$dir = Join-Path $env:LOCALAPPDATA 'vb6mcp\instances'
$files = Get-ChildItem $dir -Filter *.json -ErrorAction SilentlyContinue
if (-not $files) { "No discovery files in $dir - is the add-in loaded?"; exit 1 }
$enc = [Text.Encoding]::GetEncoding(1252)

function Invoke-Bridge($info, [string]$line) {
    $c = [Net.Sockets.TcpClient]::new()
    $c.Connect('127.0.0.1', $info.port)
    $s = $c.GetStream(); $s.ReadTimeout = 5000
    $req = $enc.GetBytes("$($info.token) $line`n"); $s.Write($req, 0, $req.Length)
    # header: "OK <n>\r\n" or "ERR <n>\r\n"
    $head = New-Object System.Text.StringBuilder
    while ($true) { $b = $s.ReadByte(); if ($b -lt 0) { break }; if ($b -eq 10) { break }; if ($b -ne 13) { [void]$head.Append([char]$b) } }
    $parts = $head.ToString().Split(' ')
    $n = [int]$parts[1]; $buf = New-Object byte[] $n; $got = 0
    while ($got -lt $n) { $r = $s.Read($buf, $got, $n - $got); if ($r -le 0) { break }; $got += $r }
    $c.Close()
    [pscustomobject]@{ Status = $parts[0]; Bytes = $n; Text = $enc.GetString($buf, 0, $got) }
}

foreach ($f in $files) {
    $info = Get-Content $f.FullName -Raw | ConvertFrom-Json
    $alive = [bool](Get-Process -Id $info.pid -ErrorAction SilentlyContinue)
    "== instance pid=$($info.pid) port=$($info.port) alive=$alive"
    if (-not $alive) { continue }
    if ($Verb) {
        $r = Invoke-Bridge $info ("$Verb " + ($Rest -join "`t"))
        "$($r.Status) ($($r.Bytes) bytes)"; $r.Text
    } else {
        $sw = [Diagnostics.Stopwatch]::StartNew(); $r = Invoke-Bridge $info 'PING'; $sw.Stop()
        "PING -> $($r.Status): $($r.Text)  [$($sw.ElapsedMilliseconds) ms]"
        $r = Invoke-Bridge $info 'BOGUS'; "BOGUS -> $($r.Status): $($r.Text)"
        "LIST (project, component, type, dirty, lines):"; (Invoke-Bridge $info 'LIST').Text
    }
}
