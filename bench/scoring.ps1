# Shared scoring rules for run-bench.ps1 and summarize.ps1 (dot-sourced), so stored runs can be rescored the same way.

# An answer that says the asked part wasn't done. A correct result is not enough when the agent couldn't do it and
# said so (pass 1: A "answered" T5 from reading the code, and left the T4 build to the user).
# T4 asks to build, not to run: "I didn't run the app" after a successful build is fine (B, rerun 2026-10-15),
# so only an unconfirmed or failed build counts.
$BenchNotDone = @{
    T4 = '(?i)(couldn''t|could not|unable to|wasn''t able to|was not able to)\s+(confirm|verify|build|compile|launch|run (the )?(build|compiler|vb6))|(didn''t|did not|haven''t|have not)\s+(confirm|verify|build|compile)|fix is unconfirmed|not confirmed'
    T5 = '(?i)(couldn''t|could not|unable to|wasn''t able to)\s+(run|launch|start|click)|from reading the code|without running|not from watching|prediction from the code'
    # T14/T16 ask to check the fix by running the program: a fix that is right but unchecked ("I couldn't run the built
    # program", A in the T16 runs) didn't do the asked part
    T14 = '(?i)(couldn''t|could not|unable to|wasn''t able to|was not able to)\s+(run|launch|start|compile|build)|(didn''t|did not|haven''t|have not)\s+(run|launch|start)\s+(the |it|Cart)'
    T16 = '(?i)(couldn''t|could not|unable to|wasn''t able to|was not able to)\s+(run|launch|start|compile|build)|(didn''t|did not|haven''t|have not)\s+(run|launch|start)\s+(the |it|Cart)'
}

# The answer asks the user to act (approve, run something themselves, click a UAC prompt).
$BenchAskUser = '(?i)\b(please|you(''ll| will)? need to|you can|manually|run (it|this|the following) (as administrator|elevated)|approve|click (yes|ok)|UAC|administrator (rights|privileges)|I (can''t|cannot|wasn''t able to|was unable to))\b'

function Get-BenchVerdict([string]$task, [bool]$resultCorrect, [string]$answer, [int]$uac, [bool]$timedOut, $post) {
    $claimsNotDone = $BenchNotDone.ContainsKey($task) -and ($answer -match $BenchNotDone[$task])
    $did = $resultCorrect -and -not $claimsNotDone
    $asks = [bool]($answer -match $BenchAskUser)
    $unaided = $did -and $uac -eq 0 -and -not $asks -and -not $timedOut -and
               @($post.dialogs).Count -eq 0 -and @($post.unsaved).Count -eq 0 -and -not $post.running
    [pscustomobject]@{ claimsNotDone = $claimsNotDone; did = $did; asks = $asks; unaided = $unaided }
}
