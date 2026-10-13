# Spike 11 client. Run NON-elevated. Connects to the elevated listener, first with a wrong token, then the right one.
$elev = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$info = Get-Content (Join-Path $env:LOCALAPPDATA 'vb6mcp-spike\listener.json') -Raw | ConvertFrom-Json
"client pid=$PID elevated=$elev -> listener pid=$($info.pid) port=$($info.port) elevated=$($info.elevated)"
foreach ($tok in @('wrong-token', $info.token)) {
  try {
    $c = [Net.Sockets.TcpClient]::new(); $c.Connect('127.0.0.1', $info.port); $s = $c.GetStream()
    $w = [IO.StreamWriter]::new($s); $w.AutoFlush = $true; $r = [IO.StreamReader]::new($s)
    $w.WriteLine("AUTH $tok"); "  sent token '$($tok.Substring(0,[Math]::Min(8,$tok.Length)))...' -> reply: $($r.ReadLine())"; $c.Close()
  } catch { "  connect failed: $($_.Exception.Message)" }
}
