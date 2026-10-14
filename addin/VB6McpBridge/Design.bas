Attribute VB_Name = "Design"
' Form designer verbs (Phase 5, docs/protocol.md "Designer"). Everything here is DEFERRED (hard rule 14):
' reading or setting a property of a UserControl whose project is open runs that UserControl's code at
' design time (Property Get/Let, InitProperties when a control is added), and user code must never run
' under the request handler. DESIGN validates, queues and replies QUEUED; the operation runs from a fresh
' dispatch (RunDesign) and stores its outcome under the request id, which RESULT returns.
' Design mode only: designer access in run mode raises error 393 and can instantiate the form (AGENTS.md).
Option Explicit

Private m_results As Collection           ' "k" & id -> "pending" | "ok" US dirty US payload | "failed" US message
Private m_order As Collection             ' ids in arrival order, to cap the store

' args: id, op, op args... -> QUEUED US id
Public Function QueueDesign(args() As String) As String
    Dim op As String, n As Long, i As Long, action As String
    n = UBound(args) + 1
    If n < 2 Then Err.Raise vbObjectError + 400, , "DESIGN needs an id and an operation"
    If Len(args(0)) = 0 Or Len(args(0)) > 40 Then Err.Raise vbObjectError + 400, , "bad request id"
    op = UCase$(args(1))
    Select Case op
    Case "CONTROLS": If n <> 4 Then Err.Raise vbObjectError + 400, , "CONTROLS takes project, component"
    Case "PROPS":    If n <> 5 Then Err.Raise vbObjectError + 400, , "PROPS takes project, component, control"
    Case "SET":      If n <> 7 Then Err.Raise vbObjectError + 400, , "SET takes project, component, control, property, value"
    Case "ADD":      If n <> 8 Then Err.Raise vbObjectError + 400, , "ADD takes project, component, progId, name, container, geometry"
    Case "REMOVE":   If n <> 5 Then Err.Raise vbObjectError + 400, , "REMOVE takes project, component, control"
    Case Else:       Err.Raise vbObjectError + 400, , "unknown designer operation " & args(1)
    End Select
    If Not InDesignMode() Then Err.Raise vbObjectError + 409, , "the IDE isn't in design mode; the designer can only be used in design mode"
    action = "DESIGN"
    For i = 0 To n - 1
        action = action & US() & args(i)
    Next
    SetResult args(0), "pending"
    Enqueue action
    QueueDesign = "QUEUED" & US() & args(0)
End Function

' args: id -> state US payload ("pending", "ok", "failed", "unknown"). A finished result is removed once read.
Public Function DesignResult(ByVal id As String) As String
    Dim v As String
    If m_results Is Nothing Then DesignResult = "unknown": Exit Function
    On Error Resume Next
    v = m_results("k" & id)
    If Err.Number <> 0 Then DesignResult = "unknown": Exit Function
    On Error GoTo 0
    If v <> "pending" Then m_results.Remove "k" & id
    DesignResult = v
End Function

Private Sub SetResult(ByVal id As String, ByVal value As String)
    If m_results Is Nothing Then Set m_results = New Collection: Set m_order = New Collection
    On Error Resume Next
    m_results.Remove "k" & id
    On Error GoTo 0
    m_results.Add value, "k" & id
    m_order.Add id
    Do While m_order.Count > 100                ' results nobody collected
        On Error Resume Next
        m_results.Remove "k" & m_order(1)
        On Error GoTo 0
        m_order.Remove 1
    Loop
End Sub

' True when no program is running or stopped: the End command (228) is enabled only in run and break mode.
Public Function InDesignMode() As Boolean
    Dim c As Object
    On Error Resume Next
    Set c = g_VBE.CommandBars.FindControl(, 228)
    If c Is Nothing Then InDesignMode = True Else InDesignMode = Not c.Enabled
End Function

' From Defer.RunQueued: f(0) = "DESIGN", f(1) = id, f(2) = op, f(3...) = op args.
Public Sub RunDesign(f() As String)
    Dim payload As String, frm As Object, cmp As VBIDE.VBComponent
    On Error GoTo Fail
    If Not InDesignMode() Then Err.Raise vbObjectError + 409, , "the IDE left design mode before the operation ran"
    Set cmp = FindComponent(f(3), f(4))
    Set frm = FormOf(cmp)
    Select Case UCase$(f(2))
    Case "CONTROLS": payload = ListControls(frm)
    Case "PROPS":    payload = ListProps(Target(cmp, frm, f(5)))
    Case "SET":      payload = SetProp(Target(cmp, frm, f(5)), f(6), f(7))
    Case "ADD":      payload = AddControl(frm, f(5), f(6), f(7), f(8))
    Case "REMOVE":   payload = RemoveControl(frm, f(5))
    End Select
    SetResult f(1), "ok" & US() & IIf(cmp.IsDirty, "1", "0") & US() & payload
    LogLine "designer " & f(2) & " " & f(3) & "." & f(4) & " done"
    Exit Sub
Fail:
    SetResult f(1), "failed" & US() & Err.Number & ": " & Err.Description
    LogLine "designer " & f(2) & " failed: " & Err.Number & ": " & Err.Description
End Sub

' The control, or the form itself when spec is the component's name (a form isn't a VBControl; its design-time
' properties are VBComponent.Properties, Phase 5 live test).
Private Function Target(ByVal cmp As VBIDE.VBComponent, ByVal frm As Object, ByVal spec As String) As Object
    If StrComp(spec, cmp.Name, vbTextCompare) = 0 Then Set Target = cmp Else Set Target = FindCtl(frm, spec)
End Function

' The designer object (VBForm) of a form, UserControl or property page.
Private Function FormOf(ByVal cmp As VBIDE.VBComponent) As Object
    Dim d As Object
    Set d = cmp.Designer
    If d Is Nothing Then Err.Raise vbObjectError + 409, , cmp.Name & " has no designer open (open it in the IDE first)"
    Set FormOf = d
End Function

' "Name" or "Name(3)" (control array element) -> VBControl. Indexed loop (spike 3).
Private Function FindCtl(ByVal frm As Object, ByVal spec As String) As Object
    Dim nm As String, idx As Long, i As Long, ctl As Object, p As Long
    nm = spec: idx = -1
    p = InStr(spec, "(")
    If p > 0 And Right$(spec, 1) = ")" Then
        nm = Left$(spec, p - 1)
        idx = CLng(Val(Mid$(spec, p + 1)))
    End If
    For i = 1 To frm.VBControls.Count
        Set ctl = frm.VBControls(i)
        If StrComp(PropText(ctl, "Name"), nm, vbTextCompare) = 0 Then
            If idx < 0 Or PropText(ctl, "Index") = CStr(idx) Then Set FindCtl = ctl: Exit Function
        End If
    Next
    Err.Raise vbObjectError + 404, , "control not found: " & spec
End Function

' A property as text, "" if the control hasn't got it.
Private Function PropText(ByVal ctl As Object, ByVal prop As String) As String
    On Error Resume Next
    PropText = FormatValue(ctl.Properties(prop).Value)
End Function

' Locale-independent text: Str$ always uses "." (CStr follows the user's locale).
Private Function FormatValue(v As Variant) As String
    Select Case VarType(v)
    Case vbString: FormatValue = v
    Case vbBoolean: FormatValue = IIf(v, "True", "False")
    Case vbInteger, vbLong, vbByte: FormatValue = CStr(v)
    Case vbSingle, vbDouble, vbCurrency, vbDecimal: FormatValue = Trim$(Str$(v))
    Case vbEmpty, vbNull: FormatValue = ""
    Case Else: FormatValue = CStr(v)
    End Select
End Function

' Controls as records: name US index US progId US className US controlType US container US left US top US width US height.
' Walks ContainedVBControls from the form down, so the container is known without VBControl.Container.
Private Function ListControls(ByVal frm As Object) As String
    Dim out As String
    Walk frm.ContainedVBControls, "", out
    ListControls = out
End Function

Private Sub Walk(ByVal coll As Object, ByVal container As String, out As String)
    Dim i As Long, ctl As Object, nm As String, idx As String, inner As Object
    For i = 1 To coll.Count
        Set ctl = coll(i)
        nm = PropText(ctl, "Name")
        idx = PropText(ctl, "Index")
        If Len(out) > 0 Then out = out & RS()
        out = out & ControlRecord(ctl, container)
        Set inner = Nothing
        On Error Resume Next
        Set inner = ctl.ContainedVBControls
        On Error GoTo 0
        If Not inner Is Nothing Then
            If inner.Count > 0 Then Walk inner, nm & IIf(idx = "" Or idx = "-1", "", "(" & idx & ")"), out
        End If
    Next
End Sub

Private Function ControlRecord(ByVal ctl As Object, ByVal container As String) As String
    ControlRecord = PropText(ctl, "Name") & US() & PropText(ctl, "Index") & US() & ctl.ProgId & US() & ctl.ClassName & US() & _
                    ctl.ControlType & US() & container & US() & PropText(ctl, "Left") & US() & PropText(ctl, "Top") & US() & _
                    PropText(ctl, "Width") & US() & PropText(ctl, "Height")
End Function

' Properties as records: name US kind US value. kind: s string, n number, b boolean, o object, i indexed (value
' not read), e error reading. A Font object is expanded into Font.Name, Font.Size, ... records.
Private Function ListProps(ByVal ctl As Object) As String
    Dim i As Long, p As Object, out As String, v As Variant, kind As String, txt As String, nm As String
    For i = 1 To ctl.Properties.Count
        Set p = ctl.Properties(i)
        kind = "": txt = ""
        nm = p.Name
        On Error Resume Next
        If p.NumIndices > 0 Then
            kind = "i": txt = IndexedText(p)
        Else
            Err.Clear
            If IsObject(p.Value) Then
                If Err.Number = 0 Then kind = "o": txt = ObjectText(nm, p.Value, out)
            Else
                Err.Clear
                v = p.Value
                If Err.Number = 0 Then kind = KindOf(v): txt = FormatValue(v)
            End If
            If Err.Number <> 0 Then kind = "e": txt = Err.Description
        End If
        On Error GoTo 0
        If Len(out) > 0 Then out = out & RS()
        out = out & nm & US() & kind & US() & txt
    Next
    ListProps = out
End Function

Private Function KindOf(v As Variant) As String
    Select Case VarType(v)
    Case vbBoolean: KindOf = "b"
    Case vbInteger, vbLong, vbByte, vbSingle, vbDouble, vbCurrency, vbDecimal: KindOf = "n"
    Case Else: KindOf = "s"
    End Select
End Function

' Text for an object property. In the designer an object property's Value is a VBIDE Properties collection of
' its sub-properties (Phase 5 live test: Font, DragIcon, MouseIcon, DataFormat), not a StdFont: its scalar
' members are appended to out as Parent.Member records (Font.Name, Font.Size, Font.Bold, ...).
Private Function ObjectText(ByVal propName As String, ByVal o As Object, out As String) As String
    Dim i As Long, sp As Object, v As Variant, nm As String
    ObjectText = TypeName(o)
    If TypeName(o) <> "Properties" Then Exit Function
    On Error Resume Next
    For i = 1 To o.Count
        Set sp = o(i)
        nm = "": nm = sp.Name
        Err.Clear
        If sp.NumIndices = 0 And Len(nm) > 0 Then
            If Not IsObject(sp.Value) And Err.Number = 0 Then
                v = sp.Value
                If Err.Number = 0 Then AddSub out, propName & "." & nm, KindOf(v), FormatValue(v)
            End If
        End If
        Err.Clear
    Next
    ObjectText = "(" & o.Count & " members)"
End Function

Private Sub AddSub(out As String, ByVal nm As String, ByVal kind As String, ByVal txt As String)
    If Len(out) > 0 Then out = out & RS()
    out = out & nm & US() & kind & US() & txt
End Sub

' The items of an indexed property (List, ItemData), LF-separated: IndexedValue(0), (1), ... until the first index
' that raises (there is no count on Property; at most 2000 items).
Private Function IndexedText(ByVal p As Object) As String
    Dim j As Long, v As Variant, out As String
    On Error GoTo Done
    For j = 0 To 1999
        v = p.IndexedValue(j)
        If j > 0 Then out = out & vbLf
        out = out & FormatValue(v)
    Next
Done:
    IndexedText = out
End Function

' Sets an indexed property from LF- or CRLF-separated items: IndexedValue(j) = item j.
Private Sub SetIndexed(ByVal p As Object, ByVal text As String)
    Dim items() As String, j As Long
    items = Split(Replace(text, vbCrLf, vbLf), vbLf)
    For j = 0 To UBound(items)
        p.IndexedValue(j) = items(j)
    Next
End Sub

' "@picture:<path>" loads a picture file under the workspace roots into an object property (Picture, Icon, ...);
' "@picture:" clears it. The IDE then writes it to the .frx itself (we never touch .frx files, hard rule 2).
Private Sub SetPicture(ByVal p As Object, ByVal path As String)
    Dim pic As Object
    If Len(path) = 0 Then
        Set pic = LoadPicture()
    Else
        If InStr(path, "..") > 0 Or Mid$(path, 2, 2) <> ":\" Then Err.Raise vbObjectError + 400, , "the picture path must be absolute, without '..'"
        If Not UnderRoots(path) Then Err.Raise vbObjectError + 400, , path & " isn't under the workspace roots"
        If Len(Dir$(path)) = 0 Then Err.Raise vbObjectError + 404, , "picture not found: " & path
        Set pic = LoadPicture(path)
    End If
    On Error Resume Next
    Set p.Object = pic
    If Err.Number <> 0 Then
        Err.Clear
        Set p.Value = pic
    End If
    If Err.Number <> 0 Then
        Dim n As Long, d As String
        n = Err.Number: d = Err.Description
        On Error GoTo 0
        Err.Raise n, , "couldn't set the picture: " & d
    End If
End Sub

' Sets a property, converting the text to the property's current type. "Parent.Member" (Font.Bold) sets a
' member of an object property through its Properties collection. Returns the value read back as records.
Private Function SetProp(ByVal ctl As Object, ByVal prop As String, ByVal text As String) As String
    Dim p As Object, dot As Long, parent As String, extra As String
    dot = InStr(prop, ".")
    If dot > 0 Then
        parent = Left$(prop, dot - 1)
        Set p = ctl.Properties(parent).Value(Mid$(prop, dot + 1))
        SetScalar p, text
        ObjectText parent, ctl.Properties(parent).Value, extra
        SetProp = extra
        Exit Function
    End If
    Set p = ctl.Properties(prop)
    If p.NumIndices > 0 Then
        SetIndexed p, text
        SetProp = prop & US() & "i" & US() & IndexedText(ctl.Properties(prop))
        Exit Function
    End If
    If Left$(text, 9) = "@picture:" Then
        SetPicture p, Mid$(text, 10)
        ObjectText prop, ctl.Properties(prop).Value, extra
        SetProp = extra
        Exit Function
    End If
    SetScalar p, text
    SetProp = prop & US() & KindOf(ctl.Properties(prop).Value) & US() & FormatValue(ctl.Properties(prop).Value)
End Function

Private Sub SetScalar(ByVal p As Object, ByVal text As String)
    Dim cur As Variant
    If p.NumIndices > 0 Then Err.Raise vbObjectError + 400, , p.Name & " is an indexed property (e.g. List); not supported"
    If IsObject(p.Value) Then Err.Raise vbObjectError + 400, , p.Name & " is an object property; set its members (e.g. Font.Bold)"
    cur = p.Value
    Select Case VarType(cur)
    Case vbString: p.Value = text
    Case vbBoolean: p.Value = ToBool(text)
    Case vbInteger: p.Value = CInt(ToNumber(text))
    Case vbLong: p.Value = CLng(ToNumber(text))
    Case vbByte: p.Value = CByte(ToNumber(text))
    Case vbSingle: p.Value = CSng(ToNumber(text))
    Case vbDouble: p.Value = ToNumber(text)
    Case vbCurrency: p.Value = CCur(ToNumber(text))
    Case Else: p.Value = text
    End Select
End Sub

' A number written the invariant way: -12, 3.5, or hex &H8000000F& (colours). Val alone turned "abc" into 0.
Private Function ToNumber(ByVal text As String) As Double
    Dim t As String, body As String
    t = Trim$(text)
    If UCase$(Left$(t, 2)) = "&H" Then
        body = Mid$(t, 3)
        If Right$(body, 1) = "&" Then body = Left$(body, Len(body) - 1)
        If Len(body) = 0 Or Len(body) > 8 Or body Like "*[!0-9A-Fa-f]*" Then Err.Raise vbObjectError + 400, , "not a number: " & text
        ToNumber = CDbl(CLng("&H" & body))
        Exit Function
    End If
    body = t
    If Left$(body, 1) = "-" Then body = Mid$(body, 2)
    If Len(body) = 0 Or body Like "*[!0-9.]*" Or body Like "*.*.*" Or body = "." Then Err.Raise vbObjectError + 400, , "not a number: " & text
    ToNumber = Val(t)
End Function

Private Function ToBool(ByVal text As String) As Boolean
    Select Case LCase$(Trim$(text))
    Case "true", "1", "-1", "yes": ToBool = True
    Case "false", "0", "no", "": ToBool = False
    Case Else: Err.Raise vbObjectError + 400, , "not a boolean: " & text
    End Select
End Function

' Adds a control (progId such as VB.CommandButton; an OCX control must be in the project's toolbox first),
' optionally inside a container control, names it and sets "left,top,width,height" (any part may be empty).
' Early-bound with RelativeVBControl passed explicitly: late-bound Add(progId) failed with error 13 (Phase 5
' live test; the omitted "Optional RelativeVBControl As VBControl = 0" went over as a missing argument).
' Returns its CONTROLS record.
Private Function AddControl(ByVal frm As Object, ByVal progId As String, ByVal nm As String, ByVal container As String, ByVal geometry As String) As String
    Dim ctl As VBIDE.VBControl, g() As String, keys As Variant, i As Long, coll As VBIDE.ContainedVBControls, rel As VBIDE.VBControl
    If Len(container) = 0 Then Set coll = frm.ContainedVBControls Else Set coll = FindCtl(frm, container).ContainedVBControls
    Set ctl = coll.Add(progId, rel, False)
    ' an unknown class makes Add return Nothing instead of raising (Phase 5 live test)
    If ctl Is Nothing Then Err.Raise vbObjectError + 400, , "VB couldn't create a '" & progId & "' control: use the class name as a .frm shows it (VB.CommandButton, MSComctlLib.ListView), and add an ActiveX control's component first (refs_add_component)"
    If Len(nm) > 0 Then ctl.Properties("Name").Value = nm
    If Len(geometry) > 0 Then
        g = Split(geometry, ",")
        keys = Array("Left", "Top", "Width", "Height")
        For i = 0 To UBound(g)
            If i > 3 Then Exit For
            If Len(Trim$(g(i))) > 0 Then SetScalar ctl.Properties(keys(i)), g(i)
        Next
    End If
    AddControl = ControlRecord(ctl, container)
End Function

' Removes a control (and the controls inside it). Its event procedures stay in the code, as when deleting by hand.
Private Function RemoveControl(ByVal frm As Object, ByVal spec As String) As String
    Dim ctl As Object
    Set ctl = FindCtl(frm, spec)
    ctl.Collection.Remove ctl
    RemoveControl = "removed " & spec
End Function
