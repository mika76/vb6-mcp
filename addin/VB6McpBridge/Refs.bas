Attribute VB_Name = "Refs"
' References and toolbox components (Phase 5, docs/protocol.md "References"). Synchronous: adding or removing
' a reference runs no user code (unlike the designer, Design.bas). Changes need design mode. GUIDs are never
' made up here (hard rule 13): REFADD by GUID gets one the server read from the registry or a typelib, and
' VBIDE checks it against the registry itself. Paths are checked against the workspace roots by the server.
Option Explicit

' args: project -> records: name US guid US major US minor US fullPath US builtIn US isBroken US type US description
' (type 0 = typelib, 1 = project reference). Indexed loop (spike 3); works in every mode.
Public Function ListRefs(ByVal prjName As String) As String
    Dim prj As VBIDE.VBProject, i As Long, out As String
    Set prj = FindProject(prjName)
    For i = 1 To prj.References.Count
        If Len(out) > 0 Then out = out & RS()
        out = out & RefRecord(prj.References(i))
    Next
    ListRefs = out
End Function

' A broken reference raises on most properties, so each is read on its own.
Private Function RefRecord(ByVal r As VBIDE.Reference) As String
    Dim nm As String, g As String, ma As String, mi As String, fp As String, bi As String, br As String, ty As String, de As String
    On Error Resume Next
    nm = r.Name
    g = r.Guid
    ma = CStr(r.Major)
    mi = CStr(r.Minor)
    fp = r.FullPath
    bi = IIf(r.BuiltIn, "1", "0")
    br = IIf(r.IsBroken, "1", "0")
    ty = CStr(r.Type)
    de = r.Description
    RefRecord = nm & US() & g & US() & ma & US() & mi & US() & fp & US() & bi & US() & br & US() & ty & US() & de
End Function

' args: project, "file", path  or  project, "guid", {guid}, major, minor -> the new reference's record.
Public Function AddRef(args() As String) As String
    Dim prj As VBIDE.VBProject, r As VBIDE.Reference
    If Not InDesignMode() Then Err.Raise vbObjectError + 409, , "references can only be changed in design mode"
    Set prj = FindProject(args(0))
    Select Case LCase$(args(1))
    Case "file"
        If UBound(args) <> 2 Then Err.Raise vbObjectError + 400, , "REFADD file takes project, file, path"
        Set r = prj.References.AddFromFile(args(2))
    Case "guid"
        If UBound(args) <> 4 Then Err.Raise vbObjectError + 400, , "REFADD guid takes project, guid, {guid}, major, minor"
        Set r = prj.References.AddFromGuid(args(2), CLng(Val(args(3))), CLng(Val(args(4))))
    Case Else
        Err.Raise vbObjectError + 400, , "REFADD kind must be file or guid"
    End Select
    AddRef = RefRecord(r)
End Function

' args: project, reference name -> "removed". Built-in references (VBA, the VB runtime) can't be removed.
Public Function RemoveRef(ByVal prjName As String, ByVal refName As String) As String
    Dim prj As VBIDE.VBProject, i As Long, nm As String
    If Not InDesignMode() Then Err.Raise vbObjectError + 409, , "references can only be changed in design mode"
    Set prj = FindProject(prjName)
    ' "#n": by position in REFS order, for a broken reference whose Name can't be read
    If Left$(refName, 1) = "#" Then
        i = CLng(Val(Mid$(refName, 2)))
        If i < 1 Or i > prj.References.Count Then Err.Raise vbObjectError + 404, , "no reference at position " & i
        If RefIsBuiltIn(prj.References(i)) Then Err.Raise vbObjectError + 409, , "reference " & i & " is built in and can't be removed"
        prj.References.Remove prj.References(i)
        RemoveRef = "removed " & refName
        Exit Function
    End If
    For i = 1 To prj.References.Count
        nm = ""
        On Error Resume Next
        nm = prj.References(i).Name
        On Error GoTo 0
        If StrComp(nm, refName, vbTextCompare) = 0 Then
            If RefIsBuiltIn(prj.References(i)) Then Err.Raise vbObjectError + 409, , refName & " is a built-in reference and can't be removed"
            prj.References.Remove prj.References(i)
            RemoveRef = "removed " & refName
            Exit Function
        End If
    Next
    Err.Raise vbObjectError + 404, , "reference not found: " & refName
End Function

' args: project, progId [, file] -> "added". Adds an ActiveX control's component (OCX) to the project's toolbox,
' i.e. an Object= line. VBIDE has no way to list or remove toolbox components.
Public Function AddToolbox(args() As String) As String
    Dim prj As VBIDE.VBProject
    If Not InDesignMode() Then Err.Raise vbObjectError + 409, , "components can only be added in design mode"
    Set prj = FindProject(args(0))
    If UBound(args) >= 2 Then
        If Len(args(2)) > 0 Then prj.AddToolboxProgID args(1), args(2) Else prj.AddToolboxProgID args(1)
    Else
        prj.AddToolboxProgID args(1)
    End If
    AddToolbox = "added " & args(1)
End Function

' BuiltIn raises for some broken references: treat those as not built in.
Private Function RefIsBuiltIn(ByVal r As VBIDE.Reference) As Boolean
    On Error Resume Next
    RefIsBuiltIn = r.BuiltIn
End Function
