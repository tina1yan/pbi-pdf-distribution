Attribute VB_Name = "PBI_UI_Export"
'Public example derived from an existing VBA UI Automation module.
'All report identifiers are synthetic. See docs/SETUP.md before running.
'Windows Excel VBA7 + Microsoft UI Automation reference required.
Option Explicit

Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As LongPtr, ByRef pid As Long) As Long
Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare PtrSafe Sub mouse_event Lib "user32" (ByVal flags As Long, ByVal dx As Long, ByVal dy As Long, ByVal data As Long, ByVal extra As LongPtr)
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
Private Const PBI_VerboseSnapshot As Boolean = False
Private Const PBI_Version As String = "1.14-public.1"
Private PBI_Handle As LongPtr

Private PBI_UI As UIAutomationClient.IUIAutomation
Private PBI_Window As UIAutomationClient.IUIAutomationElement
Private PBI_PID As Long
Private PBI_LogSheet As Worksheet
Private PBI_LogRow As Long
Private PBI_Busy As Boolean
Private PBI_Step As String

Private Const PBI_Name As Long = 30005
Private Const PBI_Type As Long = 30003
Private Const PBI_Offscreen As Long = 30022
Private Const PBI_AutoID As Long = 30011
Private Const PBI_Invoke As Long = 10000
Private Const PBI_Value As Long = 10002
Private Const PBI_Select As Long = 10010
Private Const PBI_Toggle As Long = 10015
Private Const PBI_Legacy As Long = 10018
Private Const PBI_Tab As Long = 50019
Private Const PBI_MenuItem As Long = 50011
Private Const PBI_TreeItem As Long = 50024
Private Const PBI_Button As Long = 50000
Private Const PBI_Edit As Long = 50004
Private Const PBI_Children As Long = 2
Private Const PBI_Descendants As Long = 4

Public Sub PBI_Test_Selected_Department()
    PBI_RunDepartments False
End Sub

Public Sub PBI_Batch_Export_Departments()
    PBI_RunDepartments True
End Sub

Private Sub PBI_RunDepartments(ByVal batchMode As Boolean)
    Dim ws As Worksheet, r As Long, vc As Long, sc As Long, rc As Long
    Dim fso As Object, source As String, base As String, dest As String
    Dim bookmark As String, volume As String, fileName As String, folder As String
    Dim tabs As Collection, tabName As Variant, wanted As Boolean, pages As Long
    Dim expected As String, e As UIAutomationClient.IUIAutomationElement
    Dim message As String, code As Long
    Dim workRows As Collection, rowItem As Variant, lastRow As Long, i As Long, completed As Long
    Dim destinations As Object
    If PBI_Busy Then Exit Sub
    PBI_Busy = True
    On Error GoTo Failed
    PBI_Step = "Read control row"
    Set PBI_LogSheet = Nothing
    Set PBI_Window = Nothing
    If Not TypeOf ActiveSheet Is Worksheet Then PBI_Fail "Select a department worksheet row."
    Set ws = ActiveSheet
    If Not (ws.Parent Is ThisWorkbook) Then PBI_Fail "Run from the ABCD workbook containing this module."
    r = ActiveCell.Row
    vc = PBI_Header(ws, "Volume Tab")
    sc = PBI_Header(ws, "Status")
    rc = PBI_Header(ws, "Run")
    If vc = 0 Or sc = 0 Or rc = 0 Then PBI_Fail "Row 19 must contain Volume Tab, Run and Status headers."
    If LCase$(Trim$(CStr(ws.Cells(19, 3).Value2))) <> "bookmark" Then PBI_Fail "C19 must be Bookmark."
    Set workRows = New Collection
    If batchMode Then
        lastRow = ws.Cells(ws.Rows.Count, rc).End(xlUp).Row
        For i = 20 To lastRow
            If UCase$(Trim$(CStr(ws.Cells(i, rc).Value2))) = "YES" Then workRows.Add i
        Next i
        If workRows.Count = 0 Then PBI_Fail "Set Run to Yes for the departments to export."
    Else
        workRows.Add PBI_TestRow(ws, r, rc)
    End If
    Set destinations = CreateObject("Scripting.Dictionary")
    destinations.CompareMode = vbTextCompare
    For Each rowItem In workRows
        r = CLng(rowItem)
        bookmark = Trim$(CStr(ws.Cells(r, 3).Value2))
        If IsNumeric(bookmark) Then bookmark = Format$(CLng(bookmark), "00")
        volume = Trim$(CStr(ws.Cells(r, vc).Value2))
        expected = PBI_ExpectedVolume(bookmark)
        If volume <> expected Then PBI_Fail "Volume Tab must be '" & expected & "' for bookmark " & bookmark & ". Blank means no volume."
        fileName = Trim$(CStr(ws.Cells(r, 1).Value2))
        If LCase$(Right$(fileName, 4)) <> ".pdf" Then fileName = fileName & ".pdf"
        PBI_ValidateName fileName
        folder = Trim$(CStr(ws.Cells(r, 2).Value2))
        Set fso = CreateObject("Scripting.FileSystemObject")
        If Not fso.FolderExists(folder) Then PBI_Fail "Output folder does not exist: " & folder
        dest = fso.BuildPath(folder, fileName)
        If fso.FileExists(dest & ".pbipartial") Then PBI_Fail "Partial copy already exists: " & dest & ".pbipartial"
        source = fso.BuildPath(Trim$(CStr(ws.Range("B16").Value2)), Trim$(CStr(ws.Range("B17").Value2)))
        If LCase$(fso.GetExtensionName(source)) <> "pbix" Then PBI_Fail "B16 + B17 must specify a PBIX file."
        If Not fso.FileExists(source) Then PBI_Fail "PBIX not found: " & source
        base = fso.GetBaseName(source)
        If destinations.Exists(dest) Then PBI_Fail "Two selected rows have the same PDF destination: " & dest
        destinations.Add dest, True
    Next rowItem
    PBI_StartLog ws
    PBI_Log "Macro version", PBI_Version
    PBI_Log "Selected departments", CStr(workRows.Count)
    For Each rowItem In workRows
        r = CLng(rowItem)
        PBI_CheckStop
        PBI_Step = "Read control row"
        bookmark = Trim$(CStr(ws.Cells(r, 3).Value2))
        If IsNumeric(bookmark) Then bookmark = Format$(CLng(bookmark), "00")
        volume = Trim$(CStr(ws.Cells(r, vc).Value2))
        expected = PBI_ExpectedVolume(bookmark)
        If volume <> expected Then PBI_Fail "Volume Tab must be '" & expected & "' for bookmark " & bookmark & ". Blank means no volume."
        fileName = Trim$(CStr(ws.Cells(r, 1).Value2))
        If LCase$(Right$(fileName, 4)) <> ".pdf" Then fileName = fileName & ".pdf"
        PBI_ValidateName fileName
        folder = Trim$(CStr(ws.Cells(r, 2).Value2))
        Set fso = CreateObject("Scripting.FileSystemObject")
        If Not fso.FolderExists(folder) Then PBI_Fail "Output folder does not exist: " & folder
        dest = fso.BuildPath(folder, fileName)
        If fso.FileExists(dest & ".pbipartial") Then PBI_Fail "Partial copy already exists: " & dest & ".pbipartial"
        source = fso.BuildPath(Trim$(CStr(ws.Range("B16").Value2)), Trim$(CStr(ws.Range("B17").Value2)))
        If LCase$(fso.GetExtensionName(source)) <> "pbix" Then PBI_Fail "B16 + B17 must specify a PBIX file."
        If Not fso.FileExists(source) Then PBI_Fail "PBIX not found: " & source
        base = fso.GetBaseName(source)
        PBI_Log "Control row", ws.Name & " / row " & CStr(r)
        PBI_Log "Source", source
        PBI_Log "Bookmark", bookmark
        PBI_Log "Volume", volume
        PBI_Log "PDF destination", dest
        ws.Cells(r, sc).Value2 = "Running " & CStr(completed + 1) & "/" & CStr(workRows.Count)

        PBI_Step = "Find Power BI window"
        Set PBI_UI = New UIAutomationClient.CUIAutomation
        Set PBI_Window = PBI_FindReport(base)
        PBI_Handle = CLngPtr(PBI_Window.GetCurrentPropertyValue(30020))
        PBI_PID = PBI_Window.CurrentProcessId
        If PBI_Handle = 0 Then PBI_Fail "Power BI window does not expose a native handle."
        PBI_Activate
        If completed = 0 Then PBI_PreparePanes
        PBI_Snapshot

        PBI_Step = "Find report tabs"
        Set tabs = PBI_ReportTabs()
        If tabs.Count > 50 Then PBI_Fail "Report has more than 50 discovered pages. Use the original template PBIX."
        If volume <> "" Then
            If Not PBI_Has(tabs, volume) Then PBI_Fail "Required volume page not found: " & volume
        End If

        PBI_Step = "Apply bookmark " & bookmark

        Set e = PBI_FindNamed(PBI_Window, bookmark, PBI_TreeItem, True)
        If e Is Nothing Then PBI_Fail "Bookmark tree item not found in the Bookmarks pane. Open View > Bookmarks and expand its group."
        PBI_TryScroll e
        PBI_Wait 0.2
        PBI_Log "Bookmark target", "Name=" & e.CurrentName & "; ID=" & e.CurrentAutomationId & "; target PID=" & CStr(e.CurrentProcessId) & "; report PID=" & CStr(PBI_PID)
        PBI_Click e, False
        PBI_Wait 5
        PBI_Log "Bookmark click sent", bookmark & " (this confirms mouse input only; filter/date correctness requires review)"

        PBI_Step = "Read all report page visibility states"
        For Each tabName In tabs
            wanted = PBI_ReadHidden(CStr(tabName))
            PBI_Log "Initial visibility", CStr(tabName) & "; hidden=" & CStr(wanted)
        Next tabName

        For Each tabName In Array("Page A", "Page B", "Page C")
            PBI_SetHidden CStr(tabName), False
        Next tabName
        For Each tabName In tabs
            If Not PBI_IsBase(CStr(tabName)) Then
                PBI_SetHidden CStr(tabName), (CStr(tabName) <> volume)
            End If
        Next tabName

        pages = 0
        For Each tabName In tabs
            wanted = PBI_IsBase(CStr(tabName)) Or (volume <> "" And CStr(tabName) = volume)
            If PBI_ReadHidden(CStr(tabName)) = wanted Then PBI_Fail "Final page visibility mismatch: " & CStr(tabName)
            If wanted Then pages = pages + 1
        Next tabName
        Set e = PBI_FindNamed(PBI_Window, "Page A", PBI_Tab, True)
        Set e = PBI_VisiblePage("Page A")
        PBI_Click e, False
        PBI_Wait 5
        ws.Cells(r, sc).Value2 = "Exporting - " & pages & " expected pages"
        PBI_Log "Pages prepared", "Expected PDF pages: " & pages
        PBI_Activate
        PBI_ExportPDF dest, base & ".pdf"
        ws.Cells(r, sc).Value2 = "PDF created - REVIEW " & pages & " pages"
        PBI_Log "PDF created", dest
        completed = completed + 1
        PBI_Log "Department complete", CStr(completed) & "/" & CStr(workRows.Count) & "; bookmark=" & bookmark
    Next rowItem
    ws.Activate
    PBI_DeleteSuccessfulLog
    MsgBox CStr(completed) & " PDF(s) saved. Review department, revenue, date and page contents before distribution.", vbInformation
    Application.StatusBar = False
    PBI_Busy = False
    Exit Sub
Failed:
    code = Err.Number
    message = Err.Description
    On Error Resume Next
    PBI_Log "STOP at " & PBI_Step, CStr(code) & ": " & message
    PBI_Snapshot
    If Not PBI_LogSheet Is Nothing Then PBI_LogSheet.Visible = xlSheetHidden
    If Not ws Is Nothing Then
        If sc > 0 And r >= 20 Then ws.Cells(r, sc).Value2 = "STOP - " & PBI_Step & ": " & message
    End If
    Application.StatusBar = False
    PBI_Busy = False
    MsgBox "EXPORT STOPPED at: " & PBI_Step & vbCrLf & message & vbCrLf & vbCrLf & _
           "No PBIX was saved by this macro. Open report state may have changed." & vbCrLf & _
           "If a PDF exists, review it before using it. Diagnostic log retained as a hidden PBI_UI worksheet.", vbExclamation
End Sub

Private Function PBI_Header(ByVal ws As Worksheet, ByVal label As String) As Long
    Dim c As Long
    For c = 3 To 12
        If LCase$(Trim$(CStr(ws.Cells(19, c).Value2))) = LCase$(label) Then
            If PBI_Header <> 0 Then PBI_Fail "Duplicate header: " & label
            PBI_Header = c
        End If
    Next c
End Function

Private Function PBI_TestRow(ByVal ws As Worksheet, ByVal selectedRow As Long, ByVal runCol As Long) As Long
    Dim lastRow As Long, i As Long, countYes As Long, onlyRow As Long
    If selectedRow >= 20 Then
        If UCase$(Trim$(CStr(ws.Cells(selectedRow, runCol).Value2))) = "YES" Then
            If Len(Trim$(CStr(ws.Cells(selectedRow, 3).Value2))) > 0 Then
                PBI_TestRow = selectedRow
                Exit Function
            End If
        End If
    End If
    lastRow = ws.Cells(ws.Rows.Count, runCol).End(xlUp).Row
    For i = 20 To lastRow
        If UCase$(Trim$(CStr(ws.Cells(i, runCol).Value2))) = "YES" Then
            countYes = countYes + 1
            onlyRow = i
        End If
    Next i
    If countYes = 0 Then PBI_Fail "Set Run to Yes for one department row (20 or later)."
    If countYes > 1 Then PBI_Fail "Multiple rows have Run=Yes. Select a cell in the department row to test."
    If Len(Trim$(CStr(ws.Cells(onlyRow, 3).Value2))) = 0 Then PBI_Fail "The only Run=Yes row has no Bookmark in column C."
    PBI_TestRow = onlyRow
End Function

Private Function PBI_ExpectedVolume(ByVal bookmark As String) As String
    'Synthetic example mapping; adapt only in a private working copy.
    Select Case bookmark
        Case "A": PBI_ExpectedVolume = "Volume A"
        Case "B": PBI_ExpectedVolume = "Volume B"
        Case "C": PBI_ExpectedVolume = "Volume C"
        Case "D": PBI_ExpectedVolume = ""
        Case Else: PBI_Fail "Unknown bookmark: " & bookmark
    End Select
End Function

Private Sub PBI_ValidateName(ByVal s As String)
    Dim ch As Variant, i As Long
    If s = ".pdf" Or Len(s) > 180 Then PBI_Fail "Missing or excessively long PDF filename."
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        If InStr(s, CStr(ch)) > 0 Then PBI_Fail "Invalid filename character: " & CStr(ch)
    Next ch
    For i = 1 To Len(s)
        If AscW(Mid$(s, i, 1)) < 32 Then PBI_Fail "Control character in filename."
    Next i
End Sub

Private Function PBI_NormalTitle(ByVal value As String) As String
    value = Replace(value, ChrW(160), " ")
    value = Replace(value, vbTab, " ")
    Do While InStr(value, "  ") > 0
        value = Replace(value, "  ", " ")
    Loop
    PBI_NormalTitle = Trim$(value)
End Function

Private Function PBI_ReportTitle(ByVal title As String) As String
    Dim appAt As Long, value As String, tail As String
    value = PBI_NormalTitle(title)
    appAt = InStrRev(value, "Power BI Desktop", -1, vbTextCompare)

    If appAt > 0 And appAt + Len("Power BI Desktop") - 1 = Len(value) Then
        value = Trim$(Left$(value, appAt - 1))
        If Len(value) > 0 Then
            tail = Right$(value, 1)
            If tail = "-" Or tail = ChrW(&H2013) Or tail = ChrW(&H2014) Then value = Trim$(Left$(value, Len(value) - 1))
        End If
    End If
    If Right$(value, 1) = "*" Then value = Trim$(Left$(value, Len(value) - 1))
    If LCase$(Right$(value, 5)) = ".pbix" Then value = Trim$(Left$(value, Len(value) - 5))
    PBI_ReportTitle = value
End Function

Private Function PBI_FindReport(ByVal base As String) As UIAutomationClient.IUIAutomationElement
    Dim a As UIAutomationClient.IUIAutomationElementArray
    Dim e As UIAutomationClient.IUIAutomationElement, match As UIAutomationClient.IUIAutomationElement
    Dim i As Long, title As String, reportTitle As String, candidates As Long, matches As Long
    PBI_Log "Expected report title", base
    Set a = PBI_UI.GetRootElement.FindAll(PBI_Children, PBI_UI.CreateTrueCondition)
    For i = 0 To a.Length - 1
        Set e = a.GetElement(i)
        title = e.CurrentName
        If InStr(1, title, "Power BI", vbTextCompare) > 0 Or _
           InStr(1, PBI_NormalTitle(title), PBI_NormalTitle(base), vbTextCompare) > 0 Then
            candidates = candidates + 1
            PBI_Log "Observed candidate window", title
        End If
        reportTitle = PBI_ReportTitle(title)
        If Len(reportTitle) > 0 Then
            If StrComp(reportTitle, PBI_NormalTitle(base), vbTextCompare) = 0 Then
                matches = matches + 1
                Set match = e
            End If
        End If
    Next i
    If matches > 1 Then PBI_Fail "Multiple matching Power BI windows. Leave only the intended report open."
    If matches = 0 Then
        If candidates = 0 Then PBI_Log "Window diagnostic", "No candidate Power BI/report title was exposed by UI Automation."
        PBI_Fail "Matching Power BI report window not found: " & base & ". Compare B17 with Observed candidate window rows in this log; the report may be open under a different title."
    End If
    PBI_Log "Matched window", match.CurrentName
    Set PBI_FindReport = match
End Function

Private Function PBI_FindNamed(ByVal root As UIAutomationClient.IUIAutomationElement, ByVal label As String, ByVal controlType As Long, Optional ByVal includeOffscreen As Boolean = False) As UIAutomationClient.IUIAutomationElement
    Dim cond As UIAutomationClient.IUIAutomationCondition
    Dim a As UIAutomationClient.IUIAutomationElementArray
    Dim e As UIAutomationClient.IUIAutomationElement, i As Long
    Set cond = PBI_UI.CreatePropertyCondition(PBI_Name, label)
    If controlType = PBI_Tab Then

        Set cond = PBI_UI.CreateOrCondition(cond, _
            PBI_UI.CreatePropertyCondition(PBI_Name, ChrW(&HF6AC) & " " & label))
    End If
    If controlType <> 0 Then Set cond = PBI_UI.CreateAndCondition(cond, PBI_UI.CreatePropertyCondition(PBI_Type, controlType))
    If Not includeOffscreen Then Set cond = PBI_UI.CreateAndCondition(cond, PBI_UI.CreatePropertyCondition(PBI_Offscreen, False))
    Set a = root.FindAll(PBI_Descendants, cond)
    For i = 0 To a.Length - 1
        Set e = a.GetElement(i)
        If e.CurrentIsEnabled <> 0 Then
            If Not PBI_FindNamed Is Nothing Then PBI_Fail "Ambiguous UI control: " & label
            Set PBI_FindNamed = e
        End If
    Next i
End Function

Private Function PBI_PageName(ByVal rawName As String) As String

    If Left$(rawName, 2) = ChrW(&HF6AC) & " " Then
        PBI_PageName = Mid$(rawName, 3)
    Else
        PBI_PageName = rawName
    End If
End Function

Private Function PBI_ReportTabs() As Collection
    Dim e As UIAutomationClient.IUIAutomationElement, parent As UIAutomationClient.IUIAutomationElement
    Dim a As UIAutomationClient.IUIAutomationElementArray
    Dim found As Collection, i As Long, depth As Long, label As String
    Set e = PBI_FindNamed(PBI_Window, "Page A", PBI_Tab, True)
    If e Is Nothing Then PBI_Fail "Cannot identify report page tabs. Keep the page bar visible; see PBI_UI log."
    Set parent = e
    For depth = 1 To 4
        Set parent = PBI_UI.ControlViewWalker.GetParentElement(parent)
        If parent Is Nothing Then Exit For
        Set a = parent.FindAll(PBI_Descendants, PBI_UI.CreatePropertyCondition(PBI_Type, PBI_Tab))
        Set found = New Collection
        For i = 0 To a.Length - 1
            Set e = a.GetElement(i)
            label = PBI_PageName(e.CurrentName)
            If Len(label) = 0 Then PBI_Fail "Unnamed tab encountered."
            If PBI_Has(found, label) Then PBI_Fail "Duplicate tab name: " & label
            found.Add label
        Next i
        If PBI_Has(found, "Page A") And PBI_Has(found, "Page B") And PBI_Has(found, "Page C") Then
            If PBI_Has(found, "Home") Or PBI_Has(found, "View") Then PBI_Fail "Page-tab container includes ribbon tabs; cannot reliably identify pages."
            For i = 1 To found.Count
                PBI_Log "Discovered page", CStr(found(i))
            Next i
            Set PBI_ReportTabs = found
            Exit Function
        End If
    Next depth
    PBI_Fail "Could not locate one page-tab container containing all 3 base pages."
End Function

Private Function PBI_Has(ByVal items As Collection, ByVal value As String) As Boolean
    Dim v As Variant
    For Each v In items
        If CStr(v) = value Then PBI_Has = True: Exit Function
    Next v
End Function

Private Function PBI_IsBase(ByVal label As String) As Boolean
    PBI_IsBase = (label = "Page A" Or label = "Page B" Or label = "Page C")
End Function

Private Sub PBI_Activate()
    ShowWindow PBI_Handle, 3
    SetForegroundWindow PBI_Handle
    PBI_Wait 0.5
    PBI_CheckForeground
End Sub

Private Sub PBI_CheckForeground()
    Dim pid As Long
    GetWindowThreadProcessId GetForegroundWindow(), pid
    If pid <> PBI_PID Then PBI_Fail "Power BI lost focus. Activate the report and retry."
End Sub

Private Function PBI_FindID(ByVal id As String, ByVal kind As Long) As UIAutomationClient.IUIAutomationElement
    Dim cond As UIAutomationClient.IUIAutomationCondition
    Dim a As UIAutomationClient.IUIAutomationElementArray, e As UIAutomationClient.IUIAutomationElement
    Dim i As Long
    Set cond = PBI_UI.CreateAndCondition(PBI_UI.CreatePropertyCondition(PBI_AutoID, id), PBI_UI.CreatePropertyCondition(PBI_Type, kind))
    Set cond = PBI_UI.CreateAndCondition(cond, PBI_UI.CreatePropertyCondition(PBI_Offscreen, False))
    Set a = PBI_Window.FindAll(PBI_Descendants, cond)
    For i = 0 To a.Length - 1
        Set e = a.GetElement(i)
        If e.CurrentIsEnabled <> 0 Then
            If Not PBI_FindID Is Nothing Then PBI_Fail "Multiple visible controls with ID: " & id
            Set PBI_FindID = e
        End If
    Next i
End Function

Private Function PBI_HasBookmarkItems(ByVal root As UIAutomationClient.IUIAutomationElement) As Boolean
    Dim a As UIAutomationClient.IUIAutomationElementArray, i As Long, label As String
    Set a = root.FindAll(PBI_Descendants, PBI_UI.CreatePropertyCondition(PBI_Type, PBI_TreeItem))
    For i = 0 To a.Length - 1
        label = a.GetElement(i).CurrentName
        Select Case label
            Case "A", "B", "C", "D"
                PBI_HasBookmarkItems = True
                Exit Function
        End Select
    Next i
End Function

Private Function PBI_PaneForCollapse(ByVal button As UIAutomationClient.IUIAutomationElement) As UIAutomationClient.IUIAutomationElement
    Dim node As UIAutomationClient.IUIAutomationElement, options As UIAutomationClient.IUIAutomationElement
    Dim a As UIAutomationClient.IUIAutomationElementArray, cond As UIAutomationClient.IUIAutomationCondition
    Dim depth As Long, box As UIAutomationClient.tagRECT, host As UIAutomationClient.tagRECT
    Set node = button
    host = PBI_Window.CurrentBoundingRectangle
    Set cond = PBI_UI.CreateAndCondition(PBI_UI.CreatePropertyCondition(PBI_Name, "Collapse"), PBI_UI.CreatePropertyCondition(PBI_Type, PBI_Button))
    For depth = 1 To 6
        Set node = PBI_UI.ControlViewWalker.GetParentElement(node)
        If node Is Nothing Then Exit Function
        If PBI_UI.CompareElements(node, PBI_Window) <> 0 Then Exit Function
        Set a = node.FindAll(PBI_Descendants, cond)
        If a.Length > 1 Then Exit Function
        Set options = node.FindFirst(PBI_Descendants, PBI_UI.CreateAndCondition(PBI_UI.CreatePropertyCondition(PBI_Name, "Pane options"), PBI_UI.CreatePropertyCondition(PBI_Type, PBI_Button)))
        If Not options Is Nothing Then
            box = node.CurrentBoundingRectangle
            If box.Left >= host.Left + (host.Right - host.Left) \ 2 And box.Right > box.Left Then
                Set PBI_PaneForCollapse = node
            End If
            Exit Function
        End If
    Next depth
End Function

Private Sub PBI_CollapseOtherPanes()
    Dim cond As UIAutomationClient.IUIAutomationCondition, a As UIAutomationClient.IUIAutomationElementArray
    Dim e As UIAutomationClient.IUIAutomationElement, pane As UIAutomationClient.IUIAutomationElement
    Dim i As Long, pass As Long, changed As Boolean
    Set cond = PBI_UI.CreateAndCondition(PBI_UI.CreatePropertyCondition(PBI_Name, "Collapse"), PBI_UI.CreatePropertyCondition(PBI_Type, PBI_Button))
    Set cond = PBI_UI.CreateAndCondition(cond, PBI_UI.CreatePropertyCondition(PBI_Offscreen, False))
    For pass = 1 To 12
        changed = False
        Set a = PBI_Window.FindAll(PBI_Descendants, cond)
        For i = 0 To a.Length - 1
            Set e = a.GetElement(i)
            Set pane = PBI_PaneForCollapse(e)
            If Not pane Is Nothing Then
                If Not PBI_HasBookmarkItems(pane) Then
                    PBI_Log "Collapse editor pane", pane.CurrentName & "; ID=" & pane.CurrentAutomationId
                    PBI_Click e, False
                    PBI_Wait 0.5
                    changed = True
                    Exit For
                End If
            End If
        Next i
        If Not changed Then Exit Sub
    Next pass
    PBI_Fail "Other editor panes did not collapse after 12 attempts."
End Sub

Private Function PBI_BookmarksVisible() As Boolean
    Dim a As UIAutomationClient.IUIAutomationElementArray, e As UIAutomationClient.IUIAutomationElement
    Dim cond As UIAutomationClient.IUIAutomationCondition, i As Long
    Set cond = PBI_UI.CreateAndCondition(PBI_UI.CreatePropertyCondition(PBI_Type, PBI_TreeItem), PBI_UI.CreatePropertyCondition(PBI_Offscreen, False))
    Set a = PBI_Window.FindAll(PBI_Descendants, cond)
    For i = 0 To a.Length - 1
        Set e = a.GetElement(i)
        Select Case e.CurrentName
            Case "A", "B", "C", "D"
                PBI_BookmarksVisible = True
                Exit Function
        End Select
    Next i
End Function

Private Sub PBI_PreparePanes()
    Dim e As UIAutomationClient.IUIAutomationElement, i As Long
    PBI_Step = "Prepare Bookmarks pane"
    PBI_Log "Window preparation", "Maximized; preparing Bookmarks as the expanded editor pane"
    PBI_CollapseOtherPanes
    If Not PBI_BookmarksVisible() Then
        Set e = PBI_FindID("Bookmarks", PBI_Button)
        If e Is Nothing Then
            Set e = PBI_FindID("view", PBI_Tab)
            If e Is Nothing Then Set e = PBI_FindNamed(PBI_Window, "View", PBI_Tab)
            If e Is Nothing Then PBI_Fail "View ribbon tab was not found for opening Bookmarks."
            PBI_Click e, False
            PBI_Wait 0.5
            Set e = PBI_FindID("bookmarksPane", PBI_Button)
        End If
        If e Is Nothing Then PBI_Fail "Bookmarks pane button was not exposed."
        PBI_Click e, False
        For i = 1 To 12
            PBI_Wait 0.25
            If PBI_BookmarksVisible() Then Exit For
        Next i
        If Not PBI_BookmarksVisible() Then PBI_Fail "Could not verify the expanded Bookmarks pane. Hidden PBI_UI log retained."
    End If
    PBI_CollapseOtherPanes
    PBI_Log "Bookmarks pane ready", "Bookmarks expanded; other recognized editor panes collapsed"
End Sub

Private Sub PBI_TryScroll(ByVal target As UIAutomationClient.IUIAutomationElement)
    Dim pattern As UIAutomationClient.IUIAutomationScrollItemPattern
    If target Is Nothing Then Exit Sub
    If target.CurrentIsOffscreen = 0 Then Exit Sub
    On Error Resume Next
    Set pattern = target.GetCurrentPattern(10017)
    If Not pattern Is Nothing Then pattern.ScrollIntoView
    On Error GoTo 0
End Sub

Private Function PBI_VisiblePage(ByVal page As String) As UIAutomationClient.IUIAutomationElement
    Dim e As UIAutomationClient.IUIAutomationElement, arrow As UIAutomationClient.IUIAutomationElement
    Dim direction As Variant, i As Long
    Set e = PBI_FindNamed(PBI_Window, page, PBI_Tab, True)
    If e Is Nothing Then PBI_Fail "Report page not exposed: " & page
    PBI_TryScroll e
    PBI_Wait 0.15
    Set e = PBI_FindNamed(PBI_Window, page, PBI_Tab)
    If Not e Is Nothing Then
        Set PBI_VisiblePage = e
        Exit Function
    End If
    For Each direction In Array("Next pages", "Previous pages")
        For i = 1 To 40
            Set arrow = PBI_FindNamed(PBI_Window, CStr(direction), PBI_Button)
            If arrow Is Nothing Then Exit For
            PBI_Click arrow, False
            Set e = PBI_FindNamed(PBI_Window, page, PBI_Tab)
            If Not e Is Nothing Then
                Set PBI_VisiblePage = e
                Exit Function
            End If
        Next i
    Next direction
    PBI_Fail "Could not scroll report page into view: " & page
End Function

Private Function PBI_InReportTree(ByVal target As UIAutomationClient.IUIAutomationElement) As Boolean
    Dim node As UIAutomationClient.IUIAutomationElement, depth As Long

    Set node = target
    For depth = 0 To 64
        If node Is Nothing Then Exit Function
        If PBI_UI.CompareElements(node, PBI_Window) <> 0 Then
            PBI_InReportTree = True
            Exit Function
        End If
        If node.CurrentProcessId = PBI_PID Then
            PBI_InReportTree = True
            Exit Function
        End If
        Set node = PBI_UI.RawViewWalker.GetParentElement(node)
    Next depth
End Function

Private Sub PBI_Click(ByVal e As UIAutomationClient.IUIAutomationElement, ByVal rightButton As Boolean)
    Dim box As UIAutomationClient.tagRECT
    Dim enabledState As Long, offscreenState As Long
    PBI_CheckStop
    If e Is Nothing Then PBI_Fail "Required UI control was not found."
    PBI_CheckForeground
    If e.CurrentProcessId <> PBI_PID Then
        If Not PBI_InReportTree(e) Then PBI_Fail "UI target is outside the matched report window tree."
        PBI_Log "Embedded UI accepted", e.CurrentName & "; target PID=" & CStr(e.CurrentProcessId) & "; report PID=" & CStr(PBI_PID)
    End If
    enabledState = CLng(e.CurrentIsEnabled)
    offscreenState = CLng(e.CurrentIsOffscreen)
    PBI_Log "Click target state", e.CurrentName & "; Enabled=" & CStr(enabledState) & "; Offscreen=" & CStr(offscreenState)
    If enabledState = 0 Then PBI_Fail "UI target is disabled: " & e.CurrentName
    If offscreenState <> 0 Then PBI_Fail "UI target is offscreen: " & e.CurrentName & ". Scroll it into view and retry."
    box = e.CurrentBoundingRectangle
    If box.Right <= box.Left Or box.Bottom <= box.Top Then PBI_Fail "UI control has no clickable bounds: " & e.CurrentName
    SetCursorPos box.Left + (box.Right - box.Left) \ 2, box.Top + (box.Bottom - box.Top) \ 2
    PBI_Wait 0.1
    PBI_CheckForeground
    If rightButton Then
        mouse_event &H8, 0, 0, 0, 0
        mouse_event &H10, 0, 0, 0, 0
    Else
        mouse_event &H2, 0, 0, 0, 0
        mouse_event &H4, 0, 0, 0, 0
    End If
    PBI_Wait 0.35
End Sub

Private Function PBI_MenuItemForPage(ByVal page As String) As UIAutomationClient.IUIAutomationElement
    Dim e As UIAutomationClient.IUIAutomationElement, label As Variant, kind As Variant
    Dim started As Double, elapsed As Double
    Set e = PBI_VisiblePage(page)
    If e Is Nothing Then PBI_Fail "Tab is missing or offscreen: " & page
    PBI_Click e, True
    started = Timer
    Do
        For Each label In Array("Hide", "Hide page", "Hide Page", "Unhide", "Unhide page", "Unhide Page", "Show page", "Show Page")
            For Each kind In Array(50002, PBI_MenuItem, PBI_Button)
                Set e = PBI_InProcess(CStr(label), CLng(kind))
                If Not e Is Nothing Then
                    PBI_Log "Visibility command", "Page=" & page & "; Name=" & e.CurrentName & "; Type=" & CStr(e.CurrentControlType)
                    Set PBI_MenuItemForPage = e
                    Exit Function
                End If
            Next kind
        Next label
        elapsed = Timer - started
        If elapsed < 0 Then elapsed = elapsed + 86400#
        If elapsed >= 4 Then Exit Do
        PBI_Wait 0.2
    Loop
    PBI_Fail "Hide/show menu command was not exposed for " & page & ". See PBI_UI log including checkbox/text controls."
End Function

Private Function PBI_ReadHidden(ByVal page As String) As Boolean
    Dim e As UIAutomationClient.IUIAutomationElement, rawName As String
    Set e = PBI_FindNamed(PBI_Window, page, PBI_Tab, True)
    If e Is Nothing Then PBI_Fail "Cannot read page visibility; tab was not exposed by Power BI: " & page
    rawName = e.CurrentName
    If rawName = ChrW(&HF6AC) & " " & page Then
        PBI_ReadHidden = True
    ElseIf rawName = page Then
        PBI_ReadHidden = False
    Else
        PBI_Fail "Unrecognized page visibility label: " & rawName
    End If
End Function

Private Sub PBI_SetHidden(ByVal page As String, ByVal hidden As Boolean)
    Dim e As UIAutomationClient.IUIAutomationElement, initial As Boolean
    Dim started As Double, elapsed As Double
    PBI_Step = "Set page visibility: " & page
    initial = PBI_ReadHidden(page)
    If initial = hidden Then
        PBI_Log "Verified page", page & "; hidden=" & CStr(hidden) & "; no change needed"
        Exit Sub
    End If
    Set e = PBI_MenuItemForPage(page)
    PBI_Click e, False
    PBI_Wait 0.4
    PBI_CheckForeground
    Application.SendKeys "{ESC}", True
    PBI_Wait 0.2
    started = Timer
    Do
        If PBI_ReadHidden(page) = hidden Then
            PBI_Log "Verified page", page & "; hidden=" & CStr(hidden)
            Exit Sub
        End If
        elapsed = Timer - started
        If elapsed < 0 Then elapsed = elapsed + 86400#
        If elapsed >= 4 Then Exit Do
        PBI_Wait 0.25
    Loop
    PBI_Fail "Hide/show command did not produce expected state for " & page
End Sub

Private Function PBI_InProcess(ByVal label As String, ByVal controlType As Long) As UIAutomationClient.IUIAutomationElement
    Dim windows As UIAutomationClient.IUIAutomationElementArray
    Dim w As UIAutomationClient.IUIAutomationElement, e As UIAutomationClient.IUIAutomationElement
    Dim i As Long
    Set windows = PBI_UI.GetRootElement.FindAll(PBI_Children, PBI_UI.CreateTrueCondition)
    For i = 0 To windows.Length - 1
        Set w = windows.GetElement(i)
        If w.CurrentProcessId = PBI_PID Then
            Set e = PBI_FindNamed(w, label, controlType)
            If Not e Is Nothing Then
                If Not PBI_InProcess Is Nothing Then PBI_Fail "Multiple process controls named " & label
                Set PBI_InProcess = e
            End If
        End If
    Next i
End Function

Private Function PBI_ExportTarget(ByVal label As String, ByVal kind As Long) As UIAutomationClient.IUIAutomationElement
    Dim cond As UIAutomationClient.IUIAutomationCondition
    Dim a As UIAutomationClient.IUIAutomationElementArray
    Dim e As UIAutomationClient.IUIAutomationElement
    Dim box As UIAutomationClient.tagRECT, chosen As UIAutomationClient.tagRECT, host As UIAutomationClient.tagRECT
    Dim i As Long, centerX As Long, centerY As Long, samePlace As Boolean
    Set cond = PBI_UI.CreateAndCondition(PBI_UI.CreatePropertyCondition(PBI_Name, label), PBI_UI.CreatePropertyCondition(PBI_Type, kind))
    Set cond = PBI_UI.CreateAndCondition(cond, PBI_UI.CreatePropertyCondition(PBI_Offscreen, False))
    If label = "File" Then
        Set cond = PBI_UI.CreateAndCondition(cond, PBI_UI.CreatePropertyCondition(PBI_AutoID, "Ribbon-file"))
    End If
    host = PBI_Window.CurrentBoundingRectangle
    Set a = PBI_Window.FindAll(PBI_Descendants, cond)
    For i = 0 To a.Length - 1
        Set e = a.GetElement(i)
        If e.CurrentIsEnabled <> 0 Then
            box = e.CurrentBoundingRectangle
            If box.Right > box.Left And box.Bottom > box.Top Then
                centerX = box.Left + (box.Right - box.Left) \ 2
                centerY = box.Top + (box.Bottom - box.Top) \ 2
                If centerX >= host.Left And centerX < host.Right And centerY >= host.Top And centerY < host.Bottom Then
                    PBI_Log "Export candidate", label & "; Type=" & CStr(kind) & "; ID=" & e.CurrentAutomationId & "; bounds=" & CStr(box.Left) & "," & CStr(box.Top) & "," & CStr(box.Right) & "," & CStr(box.Bottom)
                    If PBI_ExportTarget Is Nothing Then
                        Set PBI_ExportTarget = e
                        chosen.Left = box.Left
                        chosen.Top = box.Top
                        chosen.Right = box.Right
                        chosen.Bottom = box.Bottom
                    Else
                        samePlace = (box.Left = chosen.Left) And (box.Top = chosen.Top) And (box.Right = chosen.Right) And (box.Bottom = chosen.Bottom)
                        If Not samePlace Then PBI_Fail "Export command has multiple screen locations: " & label & ". Send Export candidate log rows."
                    End If
                End If
            End If
        End If
    Next i
End Function

Private Sub PBI_ExportCommand(ByVal labels As Variant)
    Dim label As Variant, kind As Variant, e As UIAutomationClient.IUIAutomationElement
    For Each label In labels
        For Each kind In Array(PBI_MenuItem, PBI_Button, 50019, 50007)
            Set e = PBI_ExportTarget(CStr(label), CLng(kind))
            If Not e Is Nothing Then
                PBI_Click e, False
                Exit Sub
            End If
        Next kind
    Next label
    PBI_Fail "Export UI command not found: " & Join(labels, " / ")
End Sub

Private Sub PBI_ExportPDF(ByVal dest As String, ByVal reportPDFName As String)
    Dim fso As Object, beforeJobs As Object, tempRoot As String, tempBase As String
    Dim sourcePDF As String, started As Date
    Set fso = CreateObject("Scripting.FileSystemObject")
    PBI_Step = "Check temporary PDF folder"
    tempBase = Environ$("TEMP")
    If Len(tempBase) = 0 Then tempBase = fso.GetSpecialFolder(2).Path
    tempRoot = fso.BuildPath(tempBase, "Power BI Desktop")
    If Not fso.FolderExists(tempRoot) Then PBI_Fail "Power BI temporary folder not found: " & tempRoot
    If fso.FileExists(dest & ".pbipartial") Then PBI_Fail "An earlier partial copy exists. Review or move it first: " & dest & ".pbipartial"
    PBI_Log "Temporary PDF root", tempRoot
    PBI_Step = "Open File menu"
    PBI_ExportCommand Array("File")
    PBI_Wait 0.5
    PBI_Step = "Open Export menu"
    PBI_ExportCommand Array("Export")
    PBI_Wait 0.5
    Set beforeJobs = PBI_PrintJobs(tempRoot)
    PBI_Log "Existing print jobs", CStr(beforeJobs.Count)
    started = Now
    PBI_Log "Export started", Format$(started, "yyyy-mm-dd hh:nn:ss") & "; expected file=" & reportPDFName
    PBI_Step = "Select Export to PDF"
    PBI_ExportCommand Array("Export to PDF", "Export to PDF (Preview)")
    PBI_Step = "Wait for newly exported PDF"
    sourcePDF = PBI_WaitNewPDF(tempRoot, beforeJobs, started, 240, reportPDFName)
    PBI_Log "New PDF source", sourcePDF
    PBI_Step = "Copy and verify PDF"
    PBI_CopyPDF sourcePDF, dest
    PBI_Log "PDF copy verified", dest
End Sub

Private Function PBI_PrintJobs(ByVal root As String) As Object
    Dim fso As Object, result As Object, folder As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    For Each folder In fso.GetFolder(root).SubFolders
        If LCase$(Left$(folder.Name, 10)) = "print-job-" Then
            If (folder.Attributes And 1024) <> 0 Then PBI_Fail "Unexpected linked print-job folder: " & folder.Path
            result.Add folder.Path, True
        End If
    Next folder
    Set PBI_PrintJobs = result
End Function

Private Function PBI_NewPDF(ByVal root As String, ByVal beforeJobs As Object, ByVal reportPDFName As String) As String
    Dim currentJobs As Object, key As Variant, jobPath As String, newCount As Long
    Dim fso As Object, candidate As String
    Set currentJobs = PBI_PrintJobs(root)
    For Each key In currentJobs.Keys
        If Not beforeJobs.Exists(CStr(key)) Then
            newCount = newCount + 1
            jobPath = CStr(key)
        End If
    Next key
    If newCount > 1 Then PBI_Fail "More than one new print job appeared. Run only one Power BI export at a time on this Windows account."
    If newCount = 0 Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    candidate = fso.BuildPath(jobPath, reportPDFName)
    If fso.FileExists(candidate) Then PBI_NewPDF = candidate
End Function

Private Function PBI_PDFReady(ByVal path As String) As Boolean
    Dim ff As Integer, opened As Boolean, length As Long, tailLength As Long
    Dim header As String, tail As String
    On Error GoTo Finished
    ff = FreeFile
    Open path For Binary Access Read Shared As #ff
    opened = True
    length = LOF(ff)
    If length < 100 Then GoTo Finished
    header = Space$(5)
    Get #ff, 1, header
    If header <> "%PDF-" Then GoTo Finished
    tailLength = length
    If tailLength > 2048 Then tailLength = 2048
    tail = Space$(tailLength)
    Get #ff, length - tailLength + 1, tail
    tail = Replace(Replace(Replace(Replace(tail, vbCr, ""), vbLf, ""), vbTab, ""), Chr$(0), "")
    PBI_PDFReady = (Right$(RTrim$(tail), 5) = "%%EOF")
Finished:
    On Error Resume Next
    If opened Then Close #ff
End Function

Private Function PBI_WaitNewPDF(ByVal root As String, ByVal beforeJobs As Object, ByVal started As Date, ByVal seconds As Long, ByVal reportPDFName As String) As String
    Dim fso As Object, file As Object, path As String, previous As String, stamp As String, priorStamp As String
    Dim stableSince As Date, readyToCopy As Boolean
    Set fso = CreateObject("Scripting.FileSystemObject")
    stableSince = Now
    Do
        PBI_Wait 1
        path = PBI_NewPDF(root, beforeJobs, reportPDFName)
        readyToCopy = False
        If Len(path) > 0 Then
            Set file = fso.GetFile(path)
            stamp = CStr(file.Size) & "|" & CStr(CDbl(file.DateLastModified))
            If StrComp(path, previous, vbTextCompare) <> 0 Or stamp <> priorStamp Then
                stableSince = Now
                previous = path
                priorStamp = stamp
            ElseIf DateDiff("s", stableSince, Now) >= 10 Then
                readyToCopy = PBI_PDFReady(path)
            End If
            If readyToCopy Then
                If StrComp(PBI_NewPDF(root, beforeJobs, reportPDFName), path, vbTextCompare) <> 0 Then PBI_Fail "New PDF changed during verification."
                PBI_WaitNewPDF = path
                Exit Function
            End If
        Else
            previous = ""
            priorStamp = ""
            stableSince = Now
        End If
        Application.StatusBar = "Waiting for this export's PDF: " & CStr(DateDiff("s", started, Now)) & " seconds"
        If DateDiff("s", started, Now) >= seconds Then PBI_Fail "No complete new PDF named " & reportPDFName & " found under " & root & ". Older print jobs were not used."
    Loop
End Function

Private Sub PBI_CopyPDF(ByVal source As String, ByVal dest As String)
    Dim fso As Object, staging As String, sourceHandle As Integer, copyHandle As Integer
    Dim sourceOpen As Boolean, copyOpen As Boolean, total As Long, position As Long, count As Long
    Dim sourceBytes() As Byte, copyBytes() As Byte, j As Long
    Dim errorNumber As Long, errorText As String
    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    staging = dest & ".pbipartial"
    If fso.FileExists(staging) Then PBI_Fail "Partial copy exists; review it first: " & staging
    sourceHandle = FreeFile
    Open source For Binary Access Read Lock Write As #sourceHandle
    sourceOpen = True
    fso.CopyFile source, staging, False
    PBI_Log "Staged PDF copy", staging
    copyHandle = FreeFile
    Open staging For Binary Access Read Lock Write As #copyHandle
    copyOpen = True
    total = LOF(sourceHandle)
    If total <> LOF(copyHandle) Then PBI_Fail "PDF copy size does not match source."
    position = 1
    Do While position <= total
        PBI_CheckStop
        count = total - position + 1
        If count > 65536 Then count = 65536
        ReDim sourceBytes(0 To count - 1)
        ReDim copyBytes(0 To count - 1)
        Get #sourceHandle, position, sourceBytes
        Get #copyHandle, position, copyBytes
        For j = 0 To count - 1
            If sourceBytes(j) <> copyBytes(j) Then PBI_Fail "PDF copy differs from source."
        Next j
        position = position + count
        DoEvents
    Loop
    Close #copyHandle
    copyOpen = False
    Close #sourceHandle
    sourceOpen = False
    If Not PBI_PDFReady(staging) Then PBI_Fail "Copied file failed PDF header/end-marker check."
    PBI_InstallPDF staging, dest
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    On Error Resume Next
    If copyOpen Then Close #copyHandle
    If sourceOpen Then Close #sourceHandle
    On Error GoTo 0
    Err.Raise errorNumber, "ABCD PDF copy", errorText
End Sub

Private Sub PBI_InstallPDF(ByVal staging As String, ByVal dest As String)
    Dim fso As Object, backup As String, suffix As Long, movedOld As Boolean
    Dim errorNumber As Long, errorText As String, restoreError As String
    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(dest) Then
        backup = dest & ".pbibak-" & Format$(Now, "yyyymmdd-hhnnss")
        Do While fso.FileExists(backup)
            suffix = suffix + 1
            backup = dest & ".pbibak-" & Format$(Now, "yyyymmdd-hhnnss") & "-" & CStr(suffix)
        Loop
        Name dest As backup
        movedOld = True
    End If
    Name staging As dest
    On Error GoTo 0
    If movedOld Then
        On Error Resume Next
        Kill backup
        errorNumber = Err.Number
        Err.Clear
        On Error GoTo 0
        If errorNumber <> 0 Then PBI_Log "Backup retained", backup
        PBI_Log "Existing PDF replaced", dest
    End If
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    If movedOld Then
        On Error Resume Next
        Err.Clear
        If Not fso.FileExists(dest) Then
            Name backup As dest
            If Err.Number <> 0 Then restoreError = Err.Description
        Else
            restoreError = "Destination exists; old PDF remains in backup."
        End If
        On Error GoTo 0
        If Len(restoreError) > 0 Then
            errorText = errorText & vbCrLf & "Previous PDF backup: " & backup & vbCrLf & restoreError
        Else
            errorText = errorText & vbCrLf & "Previous PDF restored."
        End If
    End If
    Err.Raise errorNumber, "ABCD PDF replace", errorText
End Sub

Private Sub PBI_CheckStop()
    If GetAsyncKeyState(27) < 0 Then PBI_Fail "Stopped with ESC."
End Sub

Private Sub PBI_Wait(ByVal seconds As Double)
    Dim started As Double, elapsed As Double
    started = Timer
    Do
        DoEvents
        PBI_CheckStop
        elapsed = Timer - started
        If elapsed < 0 Then elapsed = elapsed + 86400#
        If elapsed >= seconds Then Exit Do
        Sleep 50
    Loop
End Sub

Private Sub PBI_StartLog(ByVal controlSheet As Worksheet)
    Dim n As String, suffix As Long, other As Worksheet
    Dim oldScreen As Boolean, oldEvents As Boolean, code As Long, message As String
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    n = "PBI_UI_" & Format$(Now, "hhnnss")
    Do
        Set other = Nothing
        On Error Resume Next
        Set other = ThisWorkbook.Worksheets(n)
        On Error GoTo Failed
        If other Is Nothing Then Exit Do
        suffix = suffix + 1
        n = "PBI_UI_" & Format$(Now, "hhnnss") & "_" & suffix
    Loop
    Set PBI_LogSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    controlSheet.Activate
    PBI_LogSheet.Visible = xlSheetHidden
    PBI_LogSheet.Name = n
    PBI_LogSheet.Cells.NumberFormat = "@"
    PBI_LogSheet.Cells(1, 1).Value2 = "Time"
    PBI_LogSheet.Cells(1, 2).Value2 = "Step"
    PBI_LogSheet.Cells(1, 3).Value2 = "Detail"
    PBI_LogSheet.Columns(1).ColumnWidth = 12
    PBI_LogSheet.Columns(2).ColumnWidth = 30
    PBI_LogSheet.Columns(3).ColumnWidth = 100
    PBI_LogRow = 2
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Exit Sub
Failed:
    code = Err.Number
    message = Err.Description
    On Error Resume Next
    controlSheet.Activate
    If Not PBI_LogSheet Is Nothing Then PBI_LogSheet.Visible = xlSheetHidden
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    Err.Raise code, "ABCD hidden log", message
End Sub

Private Sub PBI_DeleteSuccessfulLog()
    Dim oldAlerts As Boolean, oldEvents As Boolean, code As Long, message As String
    If PBI_LogSheet Is Nothing Then Exit Sub
    oldAlerts = Application.DisplayAlerts
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    PBI_LogSheet.Delete
    Set PBI_LogSheet = Nothing
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Exit Sub
Failed:
    code = Err.Number
    message = Err.Description
    On Error Resume Next
    PBI_LogSheet.Visible = xlSheetHidden
    PBI_Log "Log cleanup failed", CStr(code) & ": " & message
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
End Sub

Private Sub PBI_Log(ByVal stepName As String, ByVal detail As String)
    If PBI_LogSheet Is Nothing Then Exit Sub
    PBI_LogSheet.Cells(PBI_LogRow, 1).Value2 = Format$(Now, "hh:nn:ss")
    PBI_LogSheet.Cells(PBI_LogRow, 2).Value2 = stepName
    PBI_LogSheet.Cells(PBI_LogRow, 3).Value2 = Left$(detail, 30000)
    PBI_LogRow = PBI_LogRow + 1
End Sub

Private Sub PBI_Snapshot()
    Dim windows As UIAutomationClient.IUIAutomationElementArray, a As UIAutomationClient.IUIAutomationElementArray
    Dim w As UIAutomationClient.IUIAutomationElement, e As UIAutomationClient.IUIAutomationElement
    Dim i As Long, j As Long, kind As Long, count As Long
    If Not PBI_VerboseSnapshot Then Exit Sub
    If PBI_UI Is Nothing Then Exit Sub
    If PBI_Window Is Nothing Then Exit Sub
    On Error Resume Next
    Set windows = PBI_UI.GetRootElement.FindAll(PBI_Children, PBI_UI.CreateTrueCondition)
    For i = 0 To windows.Length - 1
        Set w = windows.GetElement(i)
        If w.CurrentProcessId = PBI_PID Then
            PBI_Log "UI window", w.CurrentName
            Set a = w.FindAll(PBI_Descendants, PBI_UI.CreatePropertyCondition(PBI_Offscreen, False))
            For j = 0 To a.Length - 1
                Set e = a.GetElement(j)
                kind = e.CurrentControlType
                Select Case kind
                    Case PBI_Tab, PBI_MenuItem, PBI_TreeItem, PBI_Button, PBI_Edit, 50002, 50020, 50007, 50032
                        PBI_Log "UI control", "Type=" & kind & "; Name=" & e.CurrentName & "; ID=" & e.CurrentAutomationId
                        count = count + 1
                End Select
                If count >= 1000 Then Exit Sub
            Next j
        End If
    Next i
    On Error GoTo 0
End Sub

Private Sub PBI_Fail(ByVal message As String)
    Err.Raise vbObjectError + 2301, "ABCD PBI test", message
End Sub
