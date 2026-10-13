# Spike 12 client for the watchdog. Runs non-elevated. Usage: wd.ps1 PING | STATUS | SCREENSHOT <pid> [target] | BREAK <pid> | KILL <pid> force
param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Words)
$f = Join-Path $env:LOCALAPPDATA 'vb6mcp\watchdog.json'
if (-not (Test-Path $f)) { "no watchdog discovery file ($f)"; exit 1 }
$info = Get-Content $f -Raw | ConvertFrom-Json
$c = [Net.Sockets.TcpClient]::new(); $c.Connect('127.0.0.1', $info.port)
$s = $c.GetStream(); $s.ReadTimeout = 15000
$req = [Text.Encoding]::Latin1.GetBytes("$($info.token) $($Words -join ' ')`n"); $s.Write($req, 0, $req.Length)
$head = New-Object Text.StringBuilder
while (($b = $s.ReadByte()) -ge 0 -and $b -ne 10) { if ($b -ne 13) { [void]$head.Append([char]$b) } }
$n = [int]($head.ToString().Split(' ')[1]); $buf = New-Object byte[] $n; $got = 0
while ($got -lt $n) { $r = $s.Read($buf, $got, $n - $got); if ($r -le 0) { break }; $got += $r }
$c.Close()
"$($head.ToString().Split(' ')[0]): " + [Text.Encoding]::Latin1.GetString($buf, 0, $got)
