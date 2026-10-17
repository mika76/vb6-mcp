Attribute VB_Name = "ProjectVerbs"
' Project-level verbs (PLAN Phase 7, docs/protocol.md "Projects"): close with discard (CLOSE), add components
' (COMPADD), project properties (PRJGET / PRJSET) and saving one component alone (SAVECOMP).
' CLOSE, COMPADD and SAVECOMP are deferred (hard rule 14): closing can prompt, adding a form runs the code of
' UserControls on it, and saving goes through VB's file code. PRJGET / PRJSET run no user code and answer at once.
' Indexed loops only (spike 3). Paths are checked here too (hard rule 8), on the elevated side.
Option Explicit

Private Const E_BADREQUEST As Long = vbObjectError + 400
Private Const E_NOTFOUND As Long = vbObjectError + 404
Private Const E_CONFLICT As Long = vbObjectError + 409

' ---------- CLOSE ----------

' args: project, discard (0|1) -> QUEUED. Without discard, a project with unsaved changes is refused (409, the
' changed components listed). With discard (hard rule 9: the server only sends it on an explicit discard=true) the
' project is removed as it is. VBProjects.Remove SAVES changed components without asking, and IsDirty = False saves
' too (spike 31), so VBIDE has no way to drop changes: the server snapshots the files before and restores them after.
Public Function QueueClose(ByVal prjName As String, ByVal discard As String) As String
    Dim prj As VBIDE.VBProject, c As Long, changed As String
    If discard <> "0" And discard <> "1" Then Err.Raise E_BADREQUEST, , "discard must be 0 or 1"
    If Not InDesignMode() Then Err.Raise E_CONFLICT, , "the IDE isn't in design mode; end the run first"
    Set prj = FindProject(prjName)
    If discard = "0" Then
        For c = 1 To prj.VBComponents.Count
            If prj.VBComponents(c).IsDirty Then changed = changed & IIf(Len(changed) > 0, ", ", "") & prj.VBComponents(c).Name
        Next
        If Len(changed) > 0 Or prj.IsDirty Then Err.Raise E_CONFLICT, , "project " & prj.Name & " has unsaved changes (" & IIf(Len(changed) > 0, changed, "project settings") & "); save it, or close with discard"
    End If
    Enqueue "CLOSE" & US() & prj.Name & US() & discard
    QueueClose = "QUEUED" & US() & "close " & prj.Name
End Function

Public Sub CloseProject(ByVal prjName As String, ByVal discard As Boolean)
    Dim prj As VBIDE.VBProject
    On Error GoTo Fail
    Set prj = FindProject(prjName)
    g_VBE.VBProjects.Remove prj        ' saves changed components silently (spike 31); the server restores them for a discard
    LogLine "closed " & prjName
    Exit Sub
Fail:
    LogLine "close failed " & prjName & ": " & Err.Number & ": " & Err.Description
End Sub

' ---------- COMPADD ----------

' args: project, "new", kind, name  or  project, "file", path -> QUEUED.
' kind: module, class, form, mdiform, usercontrol, propertypage. The name must be a valid, unused VB name.
Public Function QueueAddComponent(args() As String) As String
    Dim prj As VBIDE.VBProject, how As String, ext As String, n As Long
    n = UBound(args) + 1
    If Not InDesignMode() Then Err.Raise E_CONFLICT, , "the IDE isn't in design mode; components can only be added in design mode"
    Set prj = FindProject(args(0))
    how = LCase$(args(1))
    If how = "new" Then
        If n <> 4 Then Err.Raise E_BADREQUEST, , "COMPADD new takes project, new, kind, name"
        If KindCode(args(2)) = 0 Then Err.Raise E_BADREQUEST, , "kind must be module, class, form, mdiform, usercontrol or propertypage"
        If Not ValidName(args(3)) Then Err.Raise E_BADREQUEST, , "'" & args(3) & "' isn't a valid VB name (a letter, then letters, digits or _, at most 40)"
        If HasComponent(prj, args(3)) Then Err.Raise E_CONFLICT, , "the project already has a component named " & args(3)
        Enqueue "COMPADD" & US() & prj.Name & US() & "new" & US() & LCase$(args(2)) & US() & args(3)
        QueueAddComponent = "QUEUED" & US() & "add " & LCase$(args(2)) & " " & args(3)
    ElseIf how = "file" Then
        If n <> 3 Then Err.Raise E_BADREQUEST, , "COMPADD file takes project, file, path"
        ext = LCase$(Mid$(args(2), InStrRev(args(2), ".") + 1))
        If ext <> "bas" And ext <> "cls" And ext <> "frm" And ext <> "ctl" And ext <> "pag" And ext <> "dsr" Then Err.Raise E_BADREQUEST, , "only .bas, .cls, .frm, .ctl, .pag and .dsr files can be added"
        If InStr(args(2), "..") > 0 Or Mid$(args(2), 2, 2) <> ":\" Then Err.Raise E_BADREQUEST, , "the path must be absolute, without '..'"
        If Not UnderRoots(args(2)) Then Err.Raise E_BADREQUEST, , args(2) & " isn't under the workspace roots (" & RootsFile() & ")"
        If Len(Dir$(args(2))) = 0 Then Err.Raise E_NOTFOUND, , "file not found: " & args(2)
        Enqueue "COMPADD" & US() & prj.Name & US() & "file" & US() & args(2)
        QueueAddComponent = "QUEUED" & US() & "add file " & args(2)
    Else
        Err.Raise E_BADREQUEST, , "COMPADD takes new or file"
    End If
End Function

Public Sub AddComponent(f() As String)
    Dim prj As VBIDE.VBProject, cmp As VBIDE.VBComponent
    On Error GoTo Fail
    Set prj = FindProject(f(1))
    If f(2) = "new" Then
        Set cmp = prj.VBComponents.Add(KindCode(f(3)))
        cmp.Name = f(4)
        LogLine "added " & f(3) & " " & cmp.Name & " to " & prj.Name
    Else
        Set cmp = prj.VBComponents.AddFile(f(3))
        LogLine "added file " & f(3) & " to " & prj.Name & " as " & cmp.Name
    End If
    Exit Sub
Fail:
    LogLine "add component failed " & f(1) & ": " & Err.Number & ": " & Err.Description
End Sub

Private Function KindCode(ByVal kind As String) As Long
    Select Case LCase$(kind)
    Case "module": KindCode = vbext_ct_StdModule
    Case "class": KindCode = vbext_ct_ClassModule
    Case "form": KindCode = vbext_ct_VBForm
    Case "mdiform": KindCode = vbext_ct_VBMDIForm
    Case "usercontrol": KindCode = vbext_ct_UserControl
    Case "propertypage": KindCode = vbext_ct_PropPage
    Case Else: KindCode = 0
    End Select
End Function

Private Function ValidName(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    If Len(s) = 0 Or Len(s) > 40 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If i = 1 Then
            If Not (ch Like "[A-Za-z]") Then Exit Function
        ElseIf Not (ch Like "[A-Za-z0-9_]") Then
            Exit Function
        End If
    Next
    ValidName = True
End Function

Private Function HasComponent(ByVal prj As VBIDE.VBProject, ByVal cmpName As String) As Boolean
    Dim c As Long
    For c = 1 To prj.VBComponents.Count
        If StrComp(prj.VBComponents(c).Name, cmpName, vbTextCompare) = 0 Then HasComponent = True: Exit Function
    Next
End Function

' ---------- PRJGET / PRJSET ----------

' args: project -> name US description US helpFile US helpContextId US startMode US buildFileName US startup US type
' startup: a component name, "Sub Main" or "(None)". Each property is read on its own: some raise for some types.
Public Function GetProjectProps(ByVal prjName As String) As String
    Dim prj As VBIDE.VBProject
    Set prj = FindProject(prjName)
    GetProjectProps = prj.Name & US() & Prop(prj, "description") & US() & Prop(prj, "helpfile") & US() & Prop(prj, "helpcontextid") & US() & _
        Prop(prj, "startmode") & US() & Prop(prj, "buildfilename") & US() & StartupText(prj) & US() & CStr(prj.Type)
End Function

Private Function Prop(ByVal prj As VBIDE.VBProject, ByVal key As String) As String
    On Error Resume Next
    Select Case key
    Case "description": Prop = prj.Description
    Case "helpfile": Prop = prj.HelpFile
    Case "helpcontextid": Prop = CStr(prj.HelpContextID)
    Case "startmode": Prop = CStr(prj.StartMode)
    Case "buildfilename": Prop = prj.BuildFileName
    End Select
End Function

' StartUpObject is a Variant: a VBComponent, or vbext_so_SubMain (0) / vbext_so_None (1). Which form VB returns is
' checked live (Phase 7 A3 spike); both are handled.
Private Function StartupText(ByVal prj As VBIDE.VBProject) As String
    Dim v As Variant
    On Error GoTo Unknown
    If IsObject(prj.VBComponents.StartUpObject) Then
        StartupText = prj.VBComponents.StartUpObject.Name
    Else
        v = prj.VBComponents.StartUpObject
        If VarType(v) = vbString Then
            StartupText = v
        ElseIf v = vbext_so_SubMain Then
            StartupText = "Sub Main"
        ElseIf v = vbext_so_None Then
            StartupText = "(None)"
        Else
            StartupText = CStr(v)
        End If
    End If
    Exit Function
Unknown:
    StartupText = "?" & Err.Description
End Function

' args: project, property, value -> the PRJGET record after the change. Design mode only.
' property: name, description, helpfile, helpcontextid, startmode (0 standalone, 1 ActiveX component),
' buildfilename, startup ("Sub Main", "(None)" or a component name).
Public Function SetProjectProp(ByVal prjName As String, ByVal key As String, ByVal value As String) As String
    Dim prj As VBIDE.VBProject, cmp As VBIDE.VBComponent
    If Not InDesignMode() Then Err.Raise E_CONFLICT, , "the IDE isn't in design mode; project settings can only be changed in design mode"
    Set prj = FindProject(prjName)
    Select Case LCase$(key)
    Case "name"
        If Not ValidName(value) Then Err.Raise E_BADREQUEST, , "'" & value & "' isn't a valid project name"
        prj.Name = value
    Case "description": prj.Description = value
    Case "helpfile": prj.HelpFile = value
    Case "helpcontextid"
        If Not IsNumeric(value) Then Err.Raise E_BADREQUEST, , "helpcontextid must be a number"
        prj.HelpContextID = CLng(value)
    Case "startmode"
        If value <> "0" And value <> "1" Then Err.Raise E_BADREQUEST, , "startmode must be 0 (standalone) or 1 (ActiveX component)"
        prj.StartMode = CLng(value)
    Case "buildfilename": prj.BuildFileName = value
    Case "startup"
        If StrComp(value, "Sub Main", vbTextCompare) = 0 Then
            prj.VBComponents.StartUpObject = vbext_so_SubMain
        ElseIf StrComp(value, "(None)", vbTextCompare) = 0 Then
            prj.VBComponents.StartUpObject = vbext_so_None
        Else
            Set cmp = FindComponent(prj.Name, value)
            SetStartupComponent prj, cmp
        End If
    Case Else
        Err.Raise E_BADREQUEST, , "unknown property " & key & " (name, description, helpfile, helpcontextid, startmode, buildfilename, startup)"
    End Select
    SetProjectProp = GetProjectProps(prj.Name)
End Function

' Tries the component itself first, then its name: which one Property Let StartUpObject accepts is checked live.
Private Sub SetStartupComponent(ByVal prj As VBIDE.VBProject, ByVal cmp As VBIDE.VBComponent)
    On Error GoTo ByName
    prj.VBComponents.StartUpObject = cmp
    Exit Sub
ByName:
    On Error GoTo 0
    prj.VBComponents.StartUpObject = cmp.Name
End Sub

' ---------- SAVECOMP ----------

' args: project, component -> QUEUED. Saves one component to its own file (VBComponent.SaveAs with its current
' path), without File > Save Project, which also rewrites the .vbp (PLAN B8). Whether SaveAs to the same path
' writes quietly is checked live (Phase 7 B8 spike).
' With a path (a component that has never been saved, e.g. one COMPADD just made): SaveAs to that new file, which
' must not exist yet (no overwriting), be absolute, without '..', and under the workspace roots.
Public Function QueueSaveComponent(args() As String) As String
    Dim cmp As VBIDE.VBComponent, path As String
    If Not InDesignMode() Then Err.Raise E_CONFLICT, , "the IDE isn't in design mode"
    Set cmp = FindComponent(args(0), args(1))
    If UBound(args) >= 2 Then path = args(2)
    If Len(path) > 0 Then
        If InStr(path, "..") > 0 Or Mid$(path, 2, 2) <> ":\" Then Err.Raise E_BADREQUEST, , "the path must be absolute, without '..'"
        If Not UnderRoots(path) Then Err.Raise E_BADREQUEST, , path & " isn't under the workspace roots (" & RootsFile() & ")"
        If Len(Dir$(path)) > 0 Then Err.Raise E_CONFLICT, , "the file already exists: " & path
    ElseIf cmp.FileCount = 0 Then
        Err.Raise E_CONFLICT, , cmp.Name & " has never been saved; pass the path for its new file"
    End If
    Enqueue "SAVECOMP" & US() & args(0) & US() & cmp.Name & US() & path
    QueueSaveComponent = "QUEUED" & US() & "save " & cmp.Name
End Function

Public Sub SaveComponent(ByVal prjName As String, ByVal cmpName As String, ByVal path As String)
    Dim cmp As VBIDE.VBComponent, ok As Boolean
    On Error GoTo Fail
    Set cmp = FindComponent(prjName, cmpName)
    If Len(path) = 0 Then path = cmp.FileNames(1)
    ok = cmp.SaveAs(path)
    LogLine "saved component " & prjName & "." & cmpName & " (" & ok & ")"
    Exit Sub
Fail:
    LogLine "save component failed " & prjName & "." & cmpName & ": " & Err.Number & ": " & Err.Description
End Sub
