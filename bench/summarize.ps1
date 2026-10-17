# Summarizes benchmark runs (run-bench.ps1) into a Markdown report.
# Usage: .\bench\summarize.ps1 -Results bench\results\pass1\runs.jsonl,bench\results\pass2\runs.jsonl [-Out report.md]
param([string[]]$Results, [string]$Out = '')
$ErrorActionPreference = 'Stop'
$runs = @(foreach ($f in $Results) { Get-Content $f | Where-Object { $_.Trim() } | ForEach-Object { $_ | ConvertFrom-Json -Depth 20 } })
if (-not $runs) { throw 'no runs' }

# Rescore with the current shared rules, so runs recorded before a scoring fix are judged the same way
. (Join-Path $PSScriptRoot 'scoring.ps1')
foreach ($r in $runs) {
    $correct = if ($null -ne $r.resultCorrect) { [bool]$r.resultCorrect } else { [bool]$r.success }
    $post = @{ dialogs = @($r.userInteraction.dialogsLeftOpen); unsaved = @($r.userInteraction.unsavedLeft); running = [bool]$r.userInteraction.programLeftRunning }
    $v = Get-BenchVerdict $r.task $correct "$($r.answer)" ([int]$r.userInteraction.uacPrompts) ([bool]$r.timedOut) $post
    $r | Add-Member -Force resultCorrect $correct
    $r | Add-Member -Force success $v.did
    $r | Add-Member -Force unaided $v.unaided
    $r | Add-Member -Force claimsNotDone $v.claimsNotDone
    $r.userInteraction | Add-Member -Force answerAsksUser $v.asks
}

function Avg($xs) { $v = @($xs | Where-Object { $null -ne $_ }); if ($v.Count) { ($v | Measure-Object -Average).Average } else { $null } }
function Med($xs) { $v = @($xs | Where-Object { $null -ne $_ } | Sort-Object); if (-not $v.Count) { return $null }; $m = [int][math]::Floor($v.Count / 2); if ($v.Count % 2) { $v[$m] } else { ($v[$m - 1] + $v[$m]) / 2 } }
function N($x, $digits = 0) { if ($null -eq $x) { '–' } else { ([double]$x).ToString("N$digits") } }
function Sum($xs) { (@($xs | Where-Object { $null -ne $_ }) | Measure-Object -Sum).Sum }

$sb = [Text.StringBuilder]::new()
function L([string]$s = '') { [void]$sb.AppendLine($s) }

$tasks = $runs | Group-Object task | Sort-Object Name
L "# vb6-mcp benchmark: plain Claude Code (A) vs with the vb6 MCP server (B)"
L ""
L "Model: $(($runs.model | Sort-Object -Unique) -join ', '). Runs: $($runs.Count) ($(@($runs | Where-Object config -eq 'A').Count) A, $(@($runs | Where-Object config -eq 'B').Count) B). Cost is Claude Code's API-price estimate; on a subscription nothing is billed, but it measures usage."
L ""
L "## Per task"
L ""
L "| Task | Cfg | Runs | Did task | Correct result | Unaided | Tokens (median) | Output tok | Cost~ (avg) | Time s (median) | Turns | Tool calls | vb6 calls | Tool errors | Hook blocks | Shell writes of VB6 files | UAC | Asks user | Left behind | Encoding issues | Unexpected file changes |"
L "|------|-----|------|----------|----------------|---------|-----------------|------------|-------------|-----------------|-------|------------|-----------|-------------|-------------|---------------------------|-----|-----------|-------------|-----------------|-------------------------|"
foreach ($tg in $tasks) {
    foreach ($cg in ($tg.Group | Group-Object config | Sort-Object Name)) {
        $g = $cg.Group
        $vb6 = Sum ($g | ForEach-Object { ($_.toolsByName.PSObject.Properties | Where-Object Name -like 'mcp__vb6__*' | Measure-Object Value -Sum).Sum })
        $left = Sum ($g | ForEach-Object { @($_.userInteraction.dialogsLeftOpen).Count + @($_.userInteraction.unsavedLeft).Count + [int][bool]$_.userInteraction.programLeftRunning })
        L ("| {0} {1} | {2} | {3} | {4}/{3} | {19}/{3} | {20}/{3} | {5} | {6} | {7} | {8} | {9} | {10} | {11} | {12} | {13} | {21} | {14} | {15} | {16} | {17} | {18} |" -f `
            $tg.Name, $g[0].taskName, $cg.Name, $g.Count, @($g | Where-Object success).Count,
            (N (Med $g.totalTokens)), (N (Med $g.outputTokens)), ('$' + (N (Avg $g.costUsdEstimate) 3)), (N (Med $g.wallSeconds) 0), (N (Avg $g.turns) 1),
            (N (Avg $g.toolCalls) 1), $vb6, (Sum $g.toolErrors), (Sum $g.hookBlocks), (Sum ($g | ForEach-Object { $_.userInteraction.uacPrompts })),
            @($g | Where-Object { $_.userInteraction.answerAsksUser }).Count, $left, @($g | Where-Object { -not $_.encodingOk }).Count, (Sum ($g | ForEach-Object { @($_.files.unexpected).Count })),
            @($g | Where-Object resultCorrect).Count, @($g | Where-Object unaided).Count, (Sum $g.shellWritesOfVb6Files))
    }
}

L ""
L "## Head to head where both succeeded"
L ""
L "Only tasks where both configurations succeeded at least once; medians over the successful runs."
L ""
L "| Task | A tokens | B tokens | B/A tokens | A time s | B time s | B/A time | A cost~ | B cost~ |"
L "|------|----------|----------|------------|----------|----------|----------|---------|---------|"
foreach ($tg in $tasks) {
    $a = @($tg.Group | Where-Object { $_.config -eq 'A' -and $_.success }); $b = @($tg.Group | Where-Object { $_.config -eq 'B' -and $_.success })
    if (-not $a.Count -or -not $b.Count) { continue }
    $at = Med $a.totalTokens; $bt = Med $b.totalTokens; $as = Med $a.wallSeconds; $bs = Med $b.wallSeconds
    L ("| {0} {1} | {2} | {3} | {4} | {5} | {6} | {7} | {8} | {9} |" -f $tg.Name, $tg.Group[0].taskName, (N $at), (N $bt), (N ($bt / $at) 2), (N $as), (N $bs), (N ($bs / $as) 2), ('$' + (N (Med $a.costUsdEstimate) 3)), ('$' + (N (Med $b.costUsdEstimate) 3)))
}

L ""
L "## Totals"
L ""
L "| | A | B |"
L "|---|---|---|"
$A = @($runs | Where-Object config -eq 'A'); $B = @($runs | Where-Object config -eq 'B')
L "| Runs where the agent did the task | $(@($A | Where-Object success).Count)/$($A.Count) | $(@($B | Where-Object success).Count)/$($B.Count) |"
L "| Runs done unaided (no UAC, no asking, nothing left behind) | $(@($A | Where-Object unaided).Count)/$($A.Count) | $(@($B | Where-Object unaided).Count)/$($B.Count) |"
L "| Shell writes of VB6 files (around the hook) | $(Sum $A.shellWritesOfVb6Files) | $(Sum $B.shellWritesOfVb6Files) |"
L "| Tokens (sum) | $(N (Sum $A.totalTokens)) | $(N (Sum $B.totalTokens)) |"
L "| Cost estimate (sum) | `$$(N (Sum $A.costUsdEstimate) 2) | `$$(N (Sum $B.costUsdEstimate) 2) |"
L "| Wall time (sum, min) | $(N ((Sum $A.wallSeconds) / 60) 1) | $(N ((Sum $B.wallSeconds) / 60) 1) |"
L "| Tool errors | $(Sum $A.toolErrors) | $(Sum $B.toolErrors) |"
L "| UAC prompts (user had to act) | $(Sum ($A | ForEach-Object { $_.userInteraction.uacPrompts })) | $(Sum ($B | ForEach-Object { $_.userInteraction.uacPrompts })) |"
L "| Permission denials | $(Sum ($A | ForEach-Object { $_.userInteraction.permissionDenials })) | $(Sum ($B | ForEach-Object { $_.userInteraction.permissionDenials })) |"
L "| Answers asking the user to act | $(@($A | Where-Object { $_.userInteraction.answerAsksUser }).Count) | $(@($B | Where-Object { $_.userInteraction.answerAsksUser }).Count) |"
L "| Dialogs / unsaved work / running program left | $(Sum ($A | ForEach-Object { @($_.userInteraction.dialogsLeftOpen).Count + @($_.userInteraction.unsavedLeft).Count + [int][bool]$_.userInteraction.programLeftRunning })) | $(Sum ($B | ForEach-Object { @($_.userInteraction.dialogsLeftOpen).Count + @($_.userInteraction.unsavedLeft).Count + [int][bool]$_.userInteraction.programLeftRunning })) |"
L "| Runs with encoding/integrity issues | $(@($A | Where-Object { -not $_.encodingOk }).Count) | $(@($B | Where-Object { -not $_.encodingOk }).Count) |"
L "| Unexpected file changes | $(Sum ($A | ForEach-Object { @($_.files.unexpected).Count })) | $(Sum ($B | ForEach-Object { @($_.files.unexpected).Count })) |"
L "| Timeouts | $(@($A | Where-Object timedOut).Count) | $(@($B | Where-Object timedOut).Count) |"
L "| Rate-limit retries | $(Sum $A.rateLimitRetries) | $(Sum $B.rateLimitRetries) |"

L ""
L "## Every run"
L ""
L "| Run | Did task | Correct | Unaided | Tokens | Time s | Turns | Tools | Errors | Shell writes | UAC | Asks user | Encoding | Verify / score |"
L "|-----|----------|---------|---------|--------|--------|-------|-------|--------|--------------|-----|-----------|----------|----------------|"
foreach ($r in ($runs | Sort-Object task, config, rep)) {
    $enc = if ($r.encodingOk) { 'ok' } else { ($r.encodingIssues.PSObject.Properties | ForEach-Object { "$($_.Name): $($_.Value -join ', ')" }) -join '; ' }
    $parts = @(); if ($r.verify -and $r.verify.PSObject.Properties.Count) { $parts = @($r.verify.PSObject.Properties | Where-Object Name -ne 'build' | ForEach-Object { "$($_.Name)=$($_.Value)" }); if ($r.verify.build) { $parts += "build=$($r.verify.build.ok)" } }
    elseif ($r.answerScore) { $parts = @($r.answerScore.PSObject.Properties | ForEach-Object { "$($_.Name)=$(@($_.Value) -join ',')" }) }; $v = $parts -join ' '
    L ("| {0} | {1} | {2} | {3} | {4} | {5} | {6} | {7} | {8} | {9} | {10} | {11} | {12} | {13} |" -f $r.run, $r.success, $r.resultCorrect, $r.unaided, (N $r.totalTokens), $r.wallSeconds, $r.turns, $r.toolCalls, $r.toolErrors, $r.shellWritesOfVb6Files, $r.userInteraction.uacPrompts, $r.userInteraction.answerAsksUser, $enc, $v)
}

$text = $sb.ToString()
if ($Out) { Set-Content -LiteralPath $Out -Value $text -Encoding utf8; Write-Host "Report: $Out" } else { $text }
