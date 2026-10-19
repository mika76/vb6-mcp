Attribute VB_Name = "ProjectVerbs"
' Project-level verbs (PLAN Phase 7, docs/protocol.md "Projects"): close with discard (CLOSE), add components
' (COMPADD), project properties (PRJGET / PRJSET), saving one component alone (SAVECOMP), and new projects (PRJNEW,
' PRJSAVEAS; spike 43, deferred too).
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
    ElseIf NeverSaved(cmp) Then
        Err.Raise E_CONFLICT, , cmp.Name & " has never been saved; pass the path for its new file"
    End If
    Enqueue "SAVECOMP" & US() & args(0) & US() & cmp.Name & US() & path
    QueueSaveComponent = "QUEUED" & US() & "save " & cmp.Name
End Function

' ---------- PRJNEW / PRJSAVEAS ----------

' args: kind, name, path, mode -> QUEUED. Creates a project (VBProjects.Add) and saves it under path (spike 43).
' kind: exe, axexe, dll, ocx. mode: add (next to the open projects: the IDE becomes a group) or replace (closes the
' open projects first; refused while one has unsaved changes, since VBProjects.Remove saves them, spike 31).
' path: a new .vbp (absolute, no '..', under the roots, must not exist, its folder must exist).
Public Function QueueNewProject(args() As String) As String
    Dim mode As String, p As Long
    If Not InDesignMode() Then Err.Raise E_CONFLICT, , "the IDE isn't in design mode; end the run first"
    If ProjectTypeCode(args(0)) < 0 Then Err.Raise E_BADREQUEST, , "kind must be exe, axexe, dll or ocx"
    If Not ValidName(args(1)) Then Err.Raise E_BADREQUEST, , "'" & args(1) & "' isn't a valid project name (a letter, then letters, digits or _, at most 40)"
    CheckNewVbp args(2)
    mode = LCase$(args(3))
    If mode <> "add" And mode <> "replace" Then Err.Raise E_BADREQUEST, , "mode must be add or replace"
    If mode = "replace" Then
        For p = 1 To g_VBE.VBProjects.Count
            If ProjectChanged(g_VBE.VBProjects(p)) Then Err.Raise E_CONFLICT, , "project " & g_VBE.VBProjects(p).Name & " has unsaved changes; save or close it first"
        Next
    End If
    Enqueue "PRJNEW" & US() & LCase$(args(0)) & US() & args(1) & US() & args(2) & US() & mode
    QueueNewProject = "QUEUED" & US() & "new " & LCase$(args(0)) & " " & args(1)
End Function

' args: project, path -> QUEUED. Saves a project that has never been saved (File > Save Project would open Save As)
' under path, with its never-saved components next to it, like PRJNEW does after creating one.
Public Function QueueSaveProjectAs(ByVal prjName As String, ByVal path As String) As String
    Dim prj As VBIDE.VBProject
    If Not InDesignMode() Then Err.Raise E_CONFLICT, , "the IDE isn't in design mode"
    Set prj = FindProject(prjName)
    If Len(VbpOf(prj)) > 0 Then Err.Raise E_CONFLICT, , "project " & prj.Name & " is already saved as " & VbpOf(prj)
    CheckNewVbp path
    Enqueue "PRJSAVEAS" & US() & prj.Name & US() & path
    QueueSaveProjectAs = "QUEUED" & US() & "save " & prj.Name & " as " & path
End Function

Private Sub CheckNewVbp(ByVal path As String)
    Dim folder As String
    If InStr(path, "..") > 0 Or Mid$(path, 2, 2) <> ":\" Then Err.Raise E_BADREQUEST, , "the path must be absolute, without '..'"
    If LCase$(Right$(path, 4)) <> ".vbp" Then Err.Raise E_BADREQUEST, , "the path must be a .vbp file"
    If Not UnderRoots(path) Then Err.Raise E_BADREQUEST, , path & " isn't under the workspace roots (" & RootsFile() & ")"
    If Len(Dir$(path)) > 0 Then Err.Raise E_CONFLICT, , "the file already exists: " & path
    folder = Left$(path, InStrRev(path, "\"))
    If Len(Dir$(folder, vbDirectory)) = 0 Then Err.Raise E_NOTFOUND, , "the folder doesn't exist: " & folder
End Sub

Private Function ProjectChanged(ByVal prj As VBIDE.VBProject) As Boolean
    Dim c As Long
    If prj.IsDirty Then ProjectChanged = True: Exit Function
    For c = 1 To prj.VBComponents.Count
        If prj.VBComponents(c).IsDirty Then ProjectChanged = True: Exit Function
    Next
End Function

Private Function ProjectTypeCode(ByVal kind As String) As Long
    Select Case LCase$(kind)
    Case "exe": ProjectTypeCode = vbext_pt_StandardExe
    Case "axexe": ProjectTypeCode = vbext_pt_ActiveXExe
    Case "dll": ProjectTypeCode = vbext_pt_ActiveXDll
    Case "ocx": ProjectTypeCode = vbext_pt_ActiveXControl
    Case Else: ProjectTypeCode = -1
    End Select
End Function

Private Function VbpOf(ByVal prj As VBIDE.VBProject) As String
    On Error Resume Next                                     ' raises for a never-saved project
    VbpOf = prj.FileName
End Function

' Deferred (hard rule 14): VBProjects.Add creates default components, and saving goes through VB's file code, which
' can show dialogs.
Public Sub NewProject(f() As String)
    Dim prj As VBIDE.VBProject, i As Long
    On Error GoTo Fail
    If f(4) = "replace" Then
        For i = g_VBE.VBProjects.Count To 1 Step -1          ' indexed, never For Each (spike 3)
            g_VBE.VBProjects.Remove g_VBE.VBProjects(i)
        Next
    End If
    ' Exclusive only when nothing is open: AddFromFile(path, False) made a one-project group (spike 22)
    Set prj = g_VBE.VBProjects.Add(ProjectTypeCode(f(1)), (g_VBE.VBProjects.Count = 0))
    LogLine "new project " & prj.Name & " (" & f(1) & "), components: " & ComponentList(prj) & "; " & g_VBE.VBProjects.Count & " projects open"
    prj.Name = f(2)
    SaveProjectTo prj, f(3)
    Exit Sub
Fail:
    LogLine "new project failed " & f(2) & ": " & Err.Number & ": " & Err.Description
End Sub

Public Sub SaveProjectAs(ByVal prjName As String, ByVal path As String)
    On Error GoTo Fail
    SaveProjectTo FindProject(prjName), path
    Exit Sub
Fail:
    LogLine "save project as failed " & prjName & ": " & Err.Number & ": " & Err.Description
End Sub

' Each never-saved component goes to <folder>\<name>.<ext> first (VBComponent.SaveAs; spike 31: Save Project stops at
' a Save File As dialog for those), then the .vbp (VBProject.SaveAs). Nothing is overwritten: an existing target file
' stops the save before anything is written. Errors go up to the caller's log line.
Private Sub SaveProjectTo(ByVal prj As VBIDE.VBProject, ByVal path As String)
    Dim folder As String, c As Long, cmp As VBIDE.VBComponent, target As String
    folder = Left$(path, InStrRev(path, "\"))
    For c = 1 To prj.VBComponents.Count
        Set cmp = prj.VBComponents(c)
        If NeverSaved(cmp) Then
            target = folder & cmp.Name & ComponentExtension(cmp.Type)
            If Len(Dir$(target)) > 0 Then Err.Raise E_CONFLICT, , "the file already exists: " & target
        End If
    Next
    For c = 1 To prj.VBComponents.Count
        Set cmp = prj.VBComponents(c)
        If NeverSaved(cmp) Then
            target = folder & cmp.Name & ComponentExtension(cmp.Type)
            cmp.SaveAs target
            LogLine "saved component " & prj.Name & "." & cmp.Name & " to " & target
        End If
    Next
    prj.SaveAs path
    LogLine "saved project " & prj.Name & " as " & VbpOf(prj) & ", dirty " & prj.IsDirty & ", group file '" & GroupFile() & "'"
End Sub

' A component that has never been saved: a new project's own (Form1, Class1) reports FileCount 1 with an empty
' FileNames(1), not FileCount 0 (spike 43).
Public Function NeverSaved(ByVal cmp As VBIDE.VBComponent) As Boolean
    If cmp.FileCount = 0 Then NeverSaved = True Else NeverSaved = (Len(cmp.FileNames(1)) = 0)
End Function

Private Function ComponentList(ByVal prj As VBIDE.VBProject) As String
    Dim c As Long
    For c = 1 To prj.VBComponents.Count
        ComponentList = ComponentList & IIf(c > 1, ", ", "") & prj.VBComponents(c).Name & " (type " & prj.VBComponents(c).Type & ", files " & prj.VBComponents(c).FileCount & ")"
    Next
End Function

Private Function GroupFile() As String
    On Error Resume Next
    GroupFile = g_VBE.VBProjects.FileName
End Function

Private Function ComponentExtension(ByVal t As Long) As String
    Select Case t
    Case vbext_ct_StdModule: ComponentExtension = ".bas"
    Case vbext_ct_ClassModule: ComponentExtension = ".cls"
    Case vbext_ct_VBForm, vbext_ct_VBMDIForm: ComponentExtension = ".frm"
    Case vbext_ct_UserControl: ComponentExtension = ".ctl"
    Case vbext_ct_PropPage: ComponentExtension = ".pag"
    Case vbext_ct_DocObject: ComponentExtension = ".dob"
    Case Else: ComponentExtension = ".dsr"
    End Select
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
