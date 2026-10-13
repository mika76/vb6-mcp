# Spike 11 listener. Run this in an ELEVATED PowerShell. It plays the role of the VB6 add-in.
$elev = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$dir = Join-Path $env:LOCALAPPDATA 'vb6mcp-spike'; New-Item -ItemType Directory -Force $dir | Out-Null
$token = [guid]::NewGuid().ToString('N')
$l = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, 0); $l.Start()
$port = $l.LocalEndpoint.Port
@{ pid = $PID; port = $port; token = $token; elevated = $elev } | ConvertTo-Json | Set-Content (Join-Path $dir 'listener.json')
"Listening on 127.0.0.1:$port  elevated=$elev  (waiting up to 120s, Ctrl+C to stop)"
$deadline = (Get-Date).AddSeconds(120)
while ((Get-Date) -lt $deadline) {
  if (-not $l.Pending()) { Start-Sleep -Milliseconds 200; continue }
  $c = $l.AcceptTcpClient(); $s = $c.GetStream()
  $r = [IO.StreamReader]::new($s); $w = [IO.StreamWriter]::new($s); $w.AutoFlush = $true
  $line = $r.ReadLine()
  if ($line -eq "AUTH $token") { $w.WriteLine("OK pong from pid=$PID elevated=$elev"); "Authenticated client served." }
  else { $w.WriteLine("DENIED"); "Rejected client with bad token." }
  $c.Close()
}
$l.Stop(); Remove-Item (Join-Path $dir 'listener.json') -ErrorAction SilentlyContinue; "Stopped."
