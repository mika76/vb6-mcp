Attribute VB_Name = "Sync"
' The IDE's own file writes, so the sync engine (watchdog, Phase 4) can tell them from changes made
' outside the IDE: a file whose time stamp changed without the IDE writing it was changed by someone
' else (git, another editor, another IDE sharing the file). Filled from FileControlEvents.AfterWriteFile
' (Connect.cls); event handlers only record (they must not do real work on the IDE thread).
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const KEEP As Long = 200
Private m_paths(KEEP - 1) As String
Private m_ticks(KEEP - 1) As Long
Private m_next As Long               ' total writes recorded; slot = m_next Mod KEEP

Public Sub RecordWrite(ByVal path As String)
    m_paths(m_next Mod KEEP) = path
    m_ticks(m_next Mod KEEP) = GetTickCount()
    m_next = m_next + 1
End Sub

' -> count US (path US tick) RS ... : writes with a sequence number >= since (the count of a previous
' call). Ticks are GetTickCount values of this machine.
Public Function WritesSince(ByVal since As Long) As String
    Dim i As Long, out As String
    If since < m_next - KEEP Then since = m_next - KEEP
    If since < 0 Then since = 0
    For i = since To m_next - 1
        If Len(out) > 0 Then out = out & RS()
        out = out & m_paths(i Mod KEEP) & US() & m_ticks(i Mod KEEP)
    Next
    WritesSince = m_next & US() & out
End Function