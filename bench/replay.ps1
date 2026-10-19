# Benchmark replays: drive a compiled fixture EXE the way a user would and return what it shows.
# Independent of the vb6 tools under test: plain Win32 messages from this (non-elevated) process to a program it started
# itself, at the same integrity, so they arrive (spike 10 is only about elevated windows). Windowed controls take posted
# and sent messages (spike 17); the results are in TextBoxes, so their text can be read (a Label has no window).
#   . .\bench\replay.ps1; Invoke-CartReplay D:\path\Cart.exe; Invoke-ClientReplay D:\path\ClientApp.exe
if (-not ('ReplayWin32' -as [type])) {
    Add-Type -Namespace '' -Name ReplayWin32 -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumProc proc, IntPtr lParam);
public delegate bool EnumProc(IntPtr hwnd, IntPtr lParam);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr hwnd, System.Text.StringBuilder s, int n);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr SendMessage(IntPtr hwnd, int msg, IntPtr w, string l);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr SendMessage(IntPtr hwnd, int msg, IntPtr w, System.Text.StringBuilder l);
[DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr hwnd, int msg, IntPtr w, IntPtr l);
[DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hwnd, out RECT r);
[DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc proc, IntPtr lParam);
[DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int pid);
[DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hwnd);
public struct RECT { public int Left, Top, Right, Bottom; }
'@
}

function Get-ReplayText([IntPtr]$h) {
    $sb = [Text.StringBuilder]::new(1024)
    [void][ReplayWin32]::SendMessage($h, 0x000D, [IntPtr]1024, $sb)    # WM_GETTEXT
    $sb.ToString()
}

# The child windows of $h: handle, class, screen position, text
function Get-ReplayChildren([IntPtr]$h) {
    $kids = [Collections.Generic.List[object]]::new()
    $proc = [ReplayWin32+EnumProc] { param($k, $l)
        $c = [Text.StringBuilder]::new(64); [void][ReplayWin32]::GetClassName($k, $c, 64)
        $r = [ReplayWin32+RECT]::new(); [void][ReplayWin32]::GetWindowRect($k, [ref]$r)
        $kids.Add([pscustomobject]@{ h = $k; cls = $c.ToString(); top = $r.Top; left = $r.Left; text = (Get-ReplayText $k) }); $true }
    [void][ReplayWin32]::EnumChildWindows($h, $proc, [IntPtr]::Zero)
    $kids.ToArray()     # items, not the list as one object, so callers can pipe them
}

# Waits for the program's form (a visible top-level ThunderRT6FormDC) or a message box (#32770, e.g. a run-time error
# "ActiveX component can't create object"). Process.MainWindowHandle picked another window (only a VBFocusRT6 under it).
function Wait-ReplayWindow($p, [datetime]$deadline) {
    $top = [ReplayWin32+EnumProc] { param($h, $l)
        $owner = 0; [void][ReplayWin32]::GetWindowThreadProcessId($h, [ref]$owner)
        if ($owner -eq $p.Id -and [ReplayWin32]::IsWindowVisible($h)) {
            $c = [Text.StringBuilder]::new(64); [void][ReplayWin32]::GetClassName($h, $c, 64)
            if ($c.ToString() -like 'ThunderRT6Form*') { $script:replayFound = @{ form = $h }; return $false }
            if ($c.ToString() -eq '#32770') {
                $text = (Get-ReplayChildren $h | Where-Object cls -eq 'Static' | ForEach-Object text) -join ' '
                $script:replayFound = @{ dialog = "$(Get-ReplayText $h): $text".Trim() }; return $false
            }
        }
        $true }
    do {
        $script:replayFound = $null; [void][ReplayWin32]::EnumWindows($top, [IntPtr]::Zero)
        if ($script:replayFound) { return $script:replayFound }
        Start-Sleep -Milliseconds 200
    } while ((Get-Date) -lt $deadline -and -not $p.HasExited)
    @{ exited = $p.HasExited }
}

# T14: tests\fixtures\Cart (bench\make-cart-fixture.cs). 10x1, 20x2, 5x4 -> 68.00; Clear; 3x2, 7x1 -> 13.00.
function Invoke-CartReplay([string]$exe) {
    $result = [ordered]@{ ok = $false; totals = @(); expected = @('68.00', '13.00'); error = $null }
    $p = Start-Process -FilePath $exe -PassThru
    try {
        $deadline = (Get-Date).AddSeconds(15)
        $w = Wait-ReplayWindow $p $deadline
        if (-not $w.form) { $result.error = if ($w.dialog) { "dialog: $($w.dialog)" } else { 'no form within 15 s' }; return $result }
        # the form exists before Form_Load has created the controls: wait for them
        do {
            $kids = Get-ReplayChildren $w.form
            $boxes = @($kids | Where-Object cls -like '*TextBox*' | Sort-Object top, left)
            if ($boxes.Count -lt 3) { Start-Sleep -Milliseconds 200 }
        } while ($boxes.Count -lt 3 -and (Get-Date) -lt $deadline)
        if ($boxes.Count -lt 3) { $result.error = "expected 3 text boxes, found $($boxes.Count): $($kids.cls -join ', ')"; return $result }
        # The text boxes by position: price and quantity on the top row (left to right), the total below them
        $price, $qty, $total = $boxes[0].h, $boxes[1].h, $boxes[-1].h
        $buttons = @($kids | Where-Object cls -like '*CommandButton*')
        $add = ($buttons | Where-Object { $_.text -replace '&' -eq 'Add' }).h
        $clear = ($buttons | Where-Object { $_.text -replace '&' -eq 'Clear' }).h
        if (-not $add -or -not $clear) { $result.error = 'no Add or Clear button'; return $result }
        $line = { param($pr, $q) [void][ReplayWin32]::SendMessage($price, 0x000C, [IntPtr]::Zero, "$pr"); [void][ReplayWin32]::SendMessage($qty, 0x000C, [IntPtr]::Zero, "$q")   # WM_SETTEXT
                  [void][ReplayWin32]::SendMessage($add, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero) }                                                                           # BM_CLICK
        # Format$ follows the user's locale: compare the number, written with a point
        $read = { $t = Get-ReplayText $total; if ($t -match '([\d.,]+)\s*$') { $Matches[1].Replace(',', '.') } else { $t } }
        & $line 10 1; & $line 20 2; & $line 5 4; $result.totals += & $read
        [void][ReplayWin32]::SendMessage($clear, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero)
        & $line 3 2; & $line 7 1; $result.totals += & $read
        $result.ok = ($result.totals -join '|') -eq ($result.expected -join '|')
        return $result
    }
    catch { $result.error = "$_"; return $result }
    finally { if (-not $p.HasExited) { $p.Kill(); $p.WaitForExit(3000) | Out-Null } }
}

# T15: tests\fixtures\Compat (bench\make-compat-fixture.cs). The compiled ClientApp.exe creates PriceLib.clsTax at start
# and shows AddTax(100) = 120. Against a DLL that broke binary compatibility it shows a run-time error box instead.
function Invoke-ClientReplay([string]$exe) {
    $result = [ordered]@{ ok = $false; shows = $null; expected = '120'; error = $null }
    $p = Start-Process -FilePath $exe -PassThru
    try {
        $deadline = (Get-Date).AddSeconds(15)
        $w = Wait-ReplayWindow $p $deadline
        if (-not $w.form) { $result.error = if ($w.dialog) { "dialog: $($w.dialog)" } elseif ($w.exited) { 'the program ended without a form' } else { 'no form within 15 s' }; return $result }
        do {
            $box = Get-ReplayChildren $w.form | Where-Object cls -like "*TextBox*" | Select-Object -First 1
            if (-not $box -or -not $box.text) { Start-Sleep -Milliseconds 200 }
        } while ((-not $box -or -not $box.text) -and (Get-Date) -lt $deadline)
        $result.shows = if ($box) { $box.text } else { $null }
        $result.ok = $result.shows -eq $result.expected
        return $result
    }
    catch { $result.error = "$_"; return $result }
    finally { if (-not $p.HasExited) { $p.Kill(); $p.WaitForExit(3000) | Out-Null } }
}

# T16: tests\fixtures\CartTax (bench\make-carttax-fixture.cs). Books 10x1 -> 10.00; Food 10x2 -> 31.00; Standard 10x1 -> 43.00.
# The category is a Dropdown List combo: CB_SETCURSEL selects it (VB reads .Text from the selection when Add runs).
function Invoke-CartTaxReplay([string]$exe) {
    $result = [ordered]@{ ok = $false; totals = @(); expected = @('10.00', '31.00', '43.00'); error = $null }
    $p = Start-Process -FilePath $exe -PassThru
    try {
        $deadline = (Get-Date).AddSeconds(15)
        $w = Wait-ReplayWindow $p $deadline
        if (-not $w.form) { $result.error = if ($w.dialog) { "dialog: $($w.dialog)" } else { 'no form within 15 s' }; return $result }
        do {
            $kids = @(Get-ReplayChildren $w.form)
            $boxes = @($kids | Where-Object cls -like '*TextBox*' | Sort-Object top, left)
            $combo = ($kids | Where-Object cls -like '*ComboBox*' | Select-Object -First 1).h
            if ($boxes.Count -lt 3 -or -not $combo) { Start-Sleep -Milliseconds 200 }
        } while (($boxes.Count -lt 3 -or -not $combo) -and (Get-Date) -lt $deadline)
        if ($boxes.Count -lt 3 -or -not $combo) { $result.error = "controls not found: $($kids.cls -join ', ')"; return $result }
        $price, $qty, $total = $boxes[0].h, $boxes[1].h, $boxes[-1].h
        $add = ($kids | Where-Object { $_.cls -like '*CommandButton*' -and ($_.text -replace '&') -eq 'Add' }).h
        if (-not $add) { $result.error = 'no Add button'; return $result }
        $read = { $t = Get-ReplayText $total; if ($t -match '([\d.,]+)\s*$') { $Matches[1].Replace(',', '.') } else { $t } }
        foreach ($l in @(@(2, 10, 1), @(0, 10, 2), @(1, 10, 1))) {     # category index (Food 0, Standard 1, Books 2), price, quantity
            [void][ReplayWin32]::SendMessage($combo, 0x014E, [IntPtr]$l[0], [IntPtr]::Zero)             # CB_SETCURSEL
            [void][ReplayWin32]::SendMessage($price, 0x000C, [IntPtr]::Zero, "$($l[1])")                 # WM_SETTEXT
            [void][ReplayWin32]::SendMessage($qty, 0x000C, [IntPtr]::Zero, "$($l[2])")
            [void][ReplayWin32]::SendMessage($add, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero)               # BM_CLICK
            $result.totals += & $read
        }
        $result.ok = ($result.totals -join '|') -eq ($result.expected -join '|')
        return $result
    }
    catch { $result.error = "$_"; return $result }
    finally { if (-not $p.HasExited) { $p.Kill(); $p.WaitForExit(3000) | Out-Null } }
}

# T17: tests\fixtures\Shop (bench\make-shop-fixture.cs). Tea 1 -> 2.80; Coffee 2 -> 7.80; Juice 1 -> 10.80.
# The combo lists Coffee, Tea, Water, Juice (first appearance in products.dat).
function Invoke-ShopReplay([string]$exe) {
    $result = [ordered]@{ ok = $false; totals = @(); expected = @('2.80', '7.80', '10.80'); error = $null }
    $p = Start-Process -FilePath $exe -PassThru
    try {
        $deadline = (Get-Date).AddSeconds(15)
        $w = Wait-ReplayWindow $p $deadline
        if (-not $w.form) { $result.error = if ($w.dialog) { "dialog: $($w.dialog)" } else { 'no form within 15 s' }; return $result }
        do {
            $kids = @(Get-ReplayChildren $w.form)
            $boxes = @($kids | Where-Object cls -like '*TextBox*' | Sort-Object top, left)
            $combo = ($kids | Where-Object cls -like '*ComboBox*' | Select-Object -First 1).h
            if ($boxes.Count -lt 2 -or -not $combo) { Start-Sleep -Milliseconds 200 }
        } while (($boxes.Count -lt 2 -or -not $combo) -and (Get-Date) -lt $deadline)
        if ($boxes.Count -lt 2 -or -not $combo) { $result.error = "controls not found: $($kids.cls -join ', ')"; return $result }
        $qty, $total = $boxes[0].h, $boxes[-1].h
        $add = ($kids | Where-Object { $_.cls -like '*CommandButton*' -and ($_.text -replace '&') -eq 'Add' }).h
        if (-not $add) { $result.error = 'no Add button'; return $result }
        $read = { $t = Get-ReplayText $total; if ($t -match '([\d.,]+)\s*$') { $Matches[1].Replace(',', '.') } else { $t } }
        foreach ($l in @(@(1, 1), @(0, 2), @(3, 1))) {     # product index (Coffee 0, Tea 1, Water 2, Juice 3), quantity
            [void][ReplayWin32]::SendMessage($combo, 0x014E, [IntPtr]$l[0], [IntPtr]::Zero)             # CB_SETCURSEL
            [void][ReplayWin32]::SendMessage($qty, 0x000C, [IntPtr]::Zero, "$($l[1])")                   # WM_SETTEXT
            [void][ReplayWin32]::SendMessage($add, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero)               # BM_CLICK
            $result.totals += & $read
        }
        $result.ok = ($result.totals -join '|') -eq ($result.expected -join '|')
        return $result
    }
    catch { $result.error = "$_"; return $result }
    finally { if (-not $p.HasExited) { $p.Kill(); $p.WaitForExit(3000) | Out-Null } }
}

# T18: tests\fixtures\Order (bench\make-order-fixture.cs). 2x10 -> 18.00; 5x1 -> 23.00; 1x1 -> 24.00.
function Invoke-OrderReplay([string]$exe) {
    $result = [ordered]@{ ok = $false; totals = @(); expected = @('18.00', '23.00', '24.00'); error = $null }
    $p = Start-Process -FilePath $exe -PassThru
    try {
        $deadline = (Get-Date).AddSeconds(15)
        $w = Wait-ReplayWindow $p $deadline
        if (-not $w.form) { $result.error = if ($w.dialog) { "dialog: $($w.dialog)" } else { 'no form within 15 s' }; return $result }
        do {
            $kids = @(Get-ReplayChildren $w.form)
            $boxes = @($kids | Where-Object cls -like '*TextBox*' | Sort-Object top, left)
            if ($boxes.Count -lt 3) { Start-Sleep -Milliseconds 200 }
        } while ($boxes.Count -lt 3 -and (Get-Date) -lt $deadline)
        if ($boxes.Count -lt 3) { $result.error = "controls not found: $($kids.cls -join ', ')"; return $result }
        $price, $qty, $total = $boxes[0].h, $boxes[1].h, $boxes[-1].h
        $add = ($kids | Where-Object { $_.cls -like '*CommandButton*' -and ($_.text -replace '&') -eq 'Add' }).h
        if (-not $add) { $result.error = 'no Add button'; return $result }
        $read = { $t = Get-ReplayText $total; if ($t -match '([\d.,]+)\s*$') { $Matches[1].Replace(',', '.') } else { $t } }
        foreach ($l in @(@(2, 10), @(5, 1), @(1, 1))) {     # price, quantity
            [void][ReplayWin32]::SendMessage($price, 0x000C, [IntPtr]::Zero, "$($l[0])")                 # WM_SETTEXT
            [void][ReplayWin32]::SendMessage($qty, 0x000C, [IntPtr]::Zero, "$($l[1])")
            [void][ReplayWin32]::SendMessage($add, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero)               # BM_CLICK
            $result.totals += & $read
        }
        $result.ok = ($result.totals -join '|') -eq ($result.expected -join '|')
        return $result
    }
    catch { $result.error = "$_"; return $result }
    finally { if (-not $p.HasExited) { $p.Kill(); $p.WaitForExit(3000) | Out-Null } }
}
