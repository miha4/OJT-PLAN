Attribute VB_Name = "OJTPlanner"
Option Explicit

Private Const SETTINGS_SHEET As String = "Nastavitve"
Private Const PLAN_SHEET As String = "OJT Plan"
Private Const SETTINGS_GROUP_ROW As Long = 3
Private Const SETTINGS_FIRST_GROUP_COL As Long = 3 'C

Private Enum GroupIdx
    giGroupName = 1
    giSrcSheetName = 2
    giIdCol = 3
    giIdRowStart = 4
    giIdRowEnd = 5
    giPlanColStart = 6
    giPlanColEnd = 7
    giDateColStart = 8
    giDateColEnd = 9
    giDateRow = 10
    giDayRow = 11
    giCandIdRowStart = 12
    giCandIdRowEnd = 13
    giPlanStartCol = 14
End Enum

Public Sub Build_OJT_Plan()
    Dim errMsg As String
    Dim trackerWb As Workbook
    Dim plannerWb As Workbook
    Dim wsPlan As Worksheet
    Dim trackerPath As String
    Dim nextOutRow As Long
    Dim groups As Collection
    Dim g As Variant
    Dim i As Long

    On Error GoTo EH
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set plannerWb = ThisWorkbook
    trackerPath = GetTrackerPath(plannerWb.Worksheets(SETTINGS_SHEET))
    Set groups = LoadGroups(plannerWb.Worksheets(SETTINGS_SHEET))
    If groups.Count = 0 Then Err.Raise 9001, , "V Nastavitve (vrstica 3, od stolpca C naprej) ni nobene skupine."

    EnsurePlanSheet plannerWb
    Set wsPlan = plannerWb.Worksheets(PLAN_SHEET)
    ResetPlanSheet wsPlan

    Set trackerWb = OpenTrackerWorkbook(trackerPath)

    nextOutRow = 1
    For i = 1 To groups.Count
        g = groups(i)
        nextOutRow = CopyGroupToPlan(GetWorksheetOrFail(trackerWb, CStr(g(giSrcSheetName))), wsPlan, g, nextOutRow)
    Next i


Cleanup:
    On Error Resume Next
    If Not trackerWb Is Nothing Then trackerWb.Close SaveChanges:=False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    On Error GoTo 0
    Exit Sub

EH:
    errMsg = "Planiraj_OJT napaka " & Err.Number & ": " & Err.Description
    Debug.Print errMsg
    MsgBox errMsg, vbCritical
    Resume Cleanup
End Sub

Public Sub Planiraj_OJT()
    Dim errMsg As String
    Dim trackerWb As Workbook
    Dim plannerWb As Workbook
    Dim wsPlan As Worksheet
    Dim wsSettings As Worksheet
    Dim trackerPath As String
    Dim groups As Collection
    Dim thresholds As Object
    Dim assignments As Collection
    Dim i As Long
    Dim g As Variant

    On Error GoTo EH
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set plannerWb = ThisWorkbook
    Set wsSettings = plannerWb.Worksheets(SETTINGS_SHEET)
    trackerPath = GetTrackerPath(wsSettings)
    Set groups = LoadGroups(wsSettings)
    If groups.Count = 0 Then Err.Raise 9001, , "V Nastavitve (vrstica 3, od stolpca C naprej) ni nobene skupine."
    Set thresholds = LoadThresholds(wsSettings)

    EnsurePlanSheet plannerWb
    Set wsPlan = plannerWb.Worksheets(PLAN_SHEET)
    ResetPlanSheet wsPlan

    Set trackerWb = OpenTrackerWorkbook(trackerPath)

    Dim nextOutRow As Long
    nextOutRow = 1
    For i = 1 To groups.Count
        g = groups(i)
        nextOutRow = CopyGroupToPlan(GetWorksheetOrFail(trackerWb, CStr(g(giSrcSheetName))), wsPlan, g, nextOutRow)
    Next i

    Set assignments = New Collection
    For i = 1 To groups.Count
        g = groups(i)
        CollectAssignments GetWorksheetOrFail(trackerWb, CStr(g(giSrcSheetName))), wsPlan, g, thresholds, assignments
    Next i

    WriteAssignments wsPlan, assignments
    MsgBox "Zaključeno. Dodelitev: " & assignments.Count, vbInformation


Cleanup:
    On Error Resume Next
    If Not trackerWb Is Nothing Then trackerWb.Close SaveChanges:=False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    On Error GoTo 0
    Exit Sub

EH:
    errMsg = "Planiraj_OJT napaka " & Err.Number & ": " & Err.Description
    Debug.Print errMsg
    MsgBox errMsg, vbCritical
    Resume Cleanup
End Sub

Private Sub CollectAssignments(ByVal wsSrc As Worksheet, ByVal wsPlan As Worksheet, ByVal g As Variant, ByVal thresholds As Object, ByRef assignments As Collection)
    Dim rowId As Long, colDate As Long
    Dim candId As String
    Dim candPhase As Long
    Dim cellValue As String
    Dim availableInstructors As Collection
    Dim chosenInstr As String
    Dim shiftCode As String

    For colDate = CLng(g(giPlanColStart)) To CLng(g(giPlanColEnd))
        For rowId = CLng(g(giCandIdRowStart)) To CLng(g(giCandIdRowEnd)) Step 3
            candId = Trim$(CStr(wsSrc.Cells(rowId + 1, CLng(g(giIdCol))).Value2))
            If Len(candId) = 0 Then GoTo NextCandidate

            cellValue = UCase$(Trim$(CStr(wsSrc.Cells(rowId + 1, colDate).Value2)))
            If cellValue <> "XS" Then GoTo NextCandidate

            candPhase = ResolvePhase(wsSrc, g, rowId + 1, colDate, thresholds)
            Set availableInstructors = GetAvailableInstructors(wsSrc, g, rowId, colDate, candPhase)

            If availableInstructors.Count > 0 Then
                If PromptAssignment(wsSrc, g, rowId + 1, colDate, candPhase, availableInstructors, chosenInstr, shiftCode) Then
                    assignments.Add CreateAssignmentItem(CStr(g(giGroupName)), wsSrc.Cells(CLng(g(giDateRow)), colDate).Value2, 2 + (colDate - CLng(g(giPlanColStart)) + 1), candId, rowId + 1, chosenInstr, shiftCode, rowId)
                End If
            End If
NextCandidate:
        Next rowId
    Next colDate
End Sub

Private Function ResolvePhase(ByVal wsSrc As Worksheet, ByVal g As Variant, ByVal candRow As Long, ByVal colDate As Long, ByVal thresholds As Object) As Long
    Dim totalHours As Double
    Dim t As Variant
    Dim reserve As Double

    totalHours = Round(CDbl(Val(wsSrc.Cells(candRow, colDate).Value2)), 0)
    t = thresholds(GetTrackType(CStr(g(giGroupName))))
    reserve = 8#

    If totalHours < (CDbl(t(1)) + reserve) Then
        ResolvePhase = 1
    ElseIf totalHours < (CDbl(t(1)) + CDbl(t(2)) - reserve) Then
        ResolvePhase = 2
    Else
        ResolvePhase = 3
    End If
End Function

Private Function GetAvailableInstructors(ByVal wsSrc As Worksheet, ByVal g As Variant, ByVal tripleStartRow As Long, ByVal colDate As Long, ByVal phase As Long) As Collection
    Dim c As New Collection
    Dim r As Long
    Dim v As String

    If phase = 2 Then
        For r = CLng(g(giIdRowStart)) To CLng(g(giIdRowEnd))
            v = UCase$(Trim$(CStr(wsSrc.Cells(r, colDate).Value2)))
            If v = "X1" Or v = "X2" Or v = "X3" Then
                c.Add Trim$(CStr(wsSrc.Cells(r, CLng(g(giIdCol))).Value2))
            End If
        Next r
    Else
        v = UCase$(Trim$(CStr(wsSrc.Cells(tripleStartRow, colDate).Value2)))
        If v = "X1" Or v = "X2" Or v = "X3" Then c.Add Trim$(CStr(wsSrc.Cells(tripleStartRow, CLng(g(giIdCol))).Value2))

        v = UCase$(Trim$(CStr(wsSrc.Cells(tripleStartRow + 2, colDate).Value2)))
        If v = "X1" Or v = "X2" Or v = "X3" Then c.Add Trim$(CStr(wsSrc.Cells(tripleStartRow + 2, CLng(g(giIdCol))).Value2))
    End If

    Set GetAvailableInstructors = c
End Function

Private Function PromptAssignment(ByVal wsSrc As Worksheet, ByVal g As Variant, ByVal candRow As Long, ByVal colDate As Long, ByVal phase As Long, ByVal instr As Collection, ByRef chosenInstr As String, ByRef shiftCode As String) As Boolean
    Dim msg As String
    Dim i As Long
    Dim pick As Variant

    msg = "Skupina: " & CStr(g(giGroupName)) & vbCrLf & _
          "Datum: " & Format$(wsSrc.Cells(CLng(g(giDateRow)), colDate).Value2, "dd.mm.") & vbCrLf & _
          "Faza: " & phase & vbCrLf & _
          "Kandidat: " & wsSrc.Cells(candRow, CLng(g(giIdCol))).Value2 & vbCrLf & _
          "Instruktorji:" & vbCrLf

    For i = 1 To instr.Count
        msg = msg & i & ") " & instr(i) & vbCrLf
    Next i
    msg = msg & vbCrLf & "Vpiši številko instruktorja ali 0 za brez dodelitve:"

    pick = Application.InputBox(msg, "OJT dodelitev", Type:=1)
    If pick = False Then Exit Function
    If CLng(pick) = 0 Then Exit Function

    If CLng(pick) < 1 Or CLng(pick) > instr.Count Then
        MsgBox "Napačna izbira.", vbExclamation
        Exit Function
    End If

    chosenInstr = instr(CLng(pick))
    shiftCode = UCase$(Trim$(CStr(Application.InputBox("Vpiši izmeno (npr A9):", "Izmena", Type:=2))))
    If Len(shiftCode) = 0 Then Exit Function

    PromptAssignment = True
End Function

Private Sub WriteAssignments(ByVal wsPlan As Worksheet, ByVal assignments As Collection)
    Dim i As Long
    Dim a As Variant
    Dim rowCand As Long, rowInstr As Long

    For i = 1 To assignments.Count
        a = assignments(i)
        rowCand = FindIdRow(wsPlan, CStr(a(3)))
        rowInstr = FindIdRow(wsPlan, CStr(a(6)))

        If rowCand > 0 Then
            wsPlan.Cells(rowCand, CLng(a(2))).Value2 = CStr(a(7)) & "s"
            AddOrReplaceComment wsPlan.Cells(rowCand, CLng(a(2))), "OJT: " & CStr(a(6)) & " - " & CStr(a(3))
        End If

        If rowInstr > 0 Then
            wsPlan.Cells(rowInstr, CLng(a(2))).Value2 = CStr(a(7)) & "i"
            AddOrReplaceComment wsPlan.Cells(rowInstr, CLng(a(2))), "OJT: " & CStr(a(6)) & " - " & CStr(a(3))
        End If
    Next i
End Sub

Private Sub AddOrReplaceComment(ByVal cell As Range, ByVal text As String)
    On Error Resume Next
    cell.CommentThreaded.Delete
    cell.ClearComments
    On Error GoTo 0
    cell.AddCommentThreaded text
End Sub

Private Function CreateAssignmentItem(ByVal groupName As String, ByVal planDate As Variant, ByVal colDate As Long, ByVal candId As String, ByVal candRow As Long, ByVal instrId As String, ByVal shiftCode As String, ByVal tripleStart As Long) As Variant
    Dim a(1 To 8) As Variant
    a(1) = groupName
    a(2) = colDate
    a(3) = candId
    a(4) = planDate
    a(5) = candRow
    a(6) = instrId
    a(7) = shiftCode
    a(8) = tripleStart
    CreateAssignmentItem = a
End Function

Private Function FindIdRow(ByVal ws As Worksheet, ByVal idValue As String) As Long
    Dim f As Range

    If Len(Trim$(idValue)) = 0 Then Exit Function

    On Error Resume Next
    Set f = ws.UsedRange.Find(What:=idValue, LookIn:=xlValues, LookAt:=xlWhole)
    On Error GoTo 0

    If Not f Is Nothing Then
        FindIdRow = f.Row
        Exit Function
    End If

    Set f = ws.Columns(1).Find(What:=idValue, LookIn:=xlValues, LookAt:=xlWhole)
    If Not f Is Nothing Then FindIdRow = f.Row
End Function

Private Function CopyGroupToPlan(ByVal wsSrc As Worksheet, ByVal wsPlan As Worksheet, ByVal g As Variant, ByVal outRow As Long) As Long
    Dim idCol As Long, nameCol As Long
    Dim rowStart As Long, rowEnd As Long
    Dim planStartCol As Long, planEndCol As Long
    Dim rowsCount As Long, planCols As Long
    Dim valsId As Variant, valsName As Variant, valsPlan As Variant

    idCol = CLng(g(giIdCol))
    nameCol = idCol + 1
    rowStart = CLng(g(giDateRow))
    rowEnd = CLng(g(giIdRowEnd))
    planStartCol = CLng(g(giPlanColStart))
    planEndCol = CLng(g(giPlanColEnd))

    wsPlan.Cells(outRow, 1).Value2 = CStr(g(giGroupName))
    wsPlan.Cells(outRow, 1).Font.Bold = True
    outRow = outRow + 1

    rowsCount = rowEnd - rowStart + 1
    planCols = planEndCol - planStartCol + 1

    valsId = wsSrc.Range(wsSrc.Cells(rowStart, idCol), wsSrc.Cells(rowEnd, idCol)).Value2
    valsName = wsSrc.Range(wsSrc.Cells(rowStart, nameCol), wsSrc.Cells(rowEnd, nameCol)).Value2
    valsPlan = wsSrc.Range(wsSrc.Cells(rowStart, planStartCol), wsSrc.Cells(rowEnd, planEndCol)).Value2

    wsPlan.Cells(outRow, 1).Resize(rowsCount, 1).Value2 = valsId
    wsPlan.Cells(outRow, 2).Resize(rowsCount, 1).Value2 = valsName
    wsPlan.Cells(outRow, 3).Resize(rowsCount, planCols).Value2 = valsPlan

    CopyGroupToPlan = outRow + rowsCount + 2
End Function

Private Function OpenTrackerWorkbook(ByVal trackerPath As String) As Workbook
    Dim wb As Workbook
    Dim normalized As String

    normalized = NormalizeSharePointPath(trackerPath)
    Debug.Print "[OJT] Odpiram tracker: "; normalized

    Set wb = FindOpenWorkbook(normalized)
    If Not wb Is Nothing Then
        Debug.Print "[OJT] Tracker že odprt: "; wb.FullName
        Set OpenTrackerWorkbook = wb
        Exit Function
    End If

    On Error Resume Next
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Filename:=normalized, UpdateLinks:=0, ReadOnly:=True, Notify:=False, AddToMru:=False)
    Application.DisplayAlerts = True
    On Error GoTo 0

    If wb Is Nothing Then
        Err.Raise 9101, , "OJTracker se ni odprl. Če uporabljaš SharePoint URL, najprej odpri datoteko ročno iz istega Office računa ali v Nastavitve!C32 vpiši lokalno sync pot (OneDrive)."
    End If

    Debug.Print "[OJT] Tracker odprt: "; wb.FullName
    Set OpenTrackerWorkbook = wb
End Function


Private Function GetWorksheetOrFail(ByVal wb As Workbook, ByVal wsName As String) As Worksheet
    Dim ws As Worksheet

    If wb Is Nothing Then
        Err.Raise 9102, , "Tracker workbook objekt ni nastavljen."
    End If

    On Error Resume Next
    Set ws = wb.Worksheets(wsName)
    On Error GoTo 0

    If ws Is Nothing Then
        Err.Raise 9103, , "Na trackerju manjka list: '" & wsName & "'."
    End If

    Set GetWorksheetOrFail = ws
End Function

Private Function GetTrackerPath(ByVal wsSettings As Worksheet) As String
    Dim v As String
    On Error Resume Next
    If wsSettings.Cells(32, 3).Hyperlinks.Count > 0 Then
        v = wsSettings.Cells(32, 3).Hyperlinks(1).Address
    End If
    On Error GoTo 0

    If Len(Trim$(v)) = 0 Then v = Trim$(CStr(wsSettings.Cells(32, 3).Value2))
    v = NormalizeSharePointPath(v)

    If Len(v) = 0 Then Err.Raise 5, , "Manjka pot do OJTracker (Nastavitve C32)."
    GetTrackerPath = v
End Function

Private Function NormalizeSharePointPath(ByVal rawPath As String) As String
    Dim p As String
    p = Trim$(rawPath)
    If Len(p) = 0 Then
        NormalizeSharePointPath = ""
        Exit Function
    End If

    If InStr(1, p, "?", vbTextCompare) > 0 Then p = Split(p, "?")(0)
    p = Replace(p, "%20", " ")
    NormalizeSharePointPath = p
End Function

Private Function FindOpenWorkbook(ByVal fullPathOrUrl As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If LCase$(NormalizeSharePointPath(wb.FullName)) = LCase$(NormalizeSharePointPath(fullPathOrUrl)) Then
            Set FindOpenWorkbook = wb
            Exit Function
        End If
    Next wb
End Function

Private Function LoadGroups(ByVal wsSettings As Worksheet) As Collection
    Dim groups As New Collection
    Dim c As Long
    Dim lastCol As Long
    Dim g(1 To 14) As Variant
    Dim groupName As String
    Dim warnings As String

    lastCol = wsSettings.Cells(SETTINGS_GROUP_ROW, wsSettings.Columns.Count).End(xlToLeft).Column
    If lastCol < SETTINGS_FIRST_GROUP_COL Then
        Set LoadGroups = groups
        Exit Function
    End If

    For c = SETTINGS_FIRST_GROUP_COL To lastCol
        groupName = Trim$(CStr(wsSettings.Cells(SETTINGS_GROUP_ROW, c).Value2))
        If Len(groupName) = 0 Then
            If HasAnyGroupConfig(wsSettings, c) Then
                warnings = warnings & "- Stolpec " & Split(wsSettings.Cells(1, c).Address(False, False), "$")(0) & " ima delne podatke brez imena skupine." & vbCrLf
            End If
            GoTo NextCol
        End If

        Debug.Print "[OJT] Group col " & c & ": " & groupName

        If Not IsGroupColumnValid(wsSettings, c, warnings, groupName) Then GoTo NextCol

        g(giGroupName) = groupName
        g(giSrcSheetName) = groupName
        g(giIdCol) = CLng(Val(wsSettings.Cells(4, c).Value2))
        g(giIdRowStart) = CLng(Val(wsSettings.Cells(5, c).Value2))
        g(giIdRowEnd) = CLng(Val(wsSettings.Cells(6, c).Value2))
        g(giPlanColStart) = ColToNum(CStr(wsSettings.Cells(7, c).Value2))
        g(giPlanColEnd) = ColToNum(CStr(wsSettings.Cells(8, c).Value2))
        g(giDateColStart) = ColToNum(CStr(wsSettings.Cells(11, c).Value2))
        g(giDateColEnd) = ColToNum(CStr(wsSettings.Cells(12, c).Value2))
        g(giDateRow) = CLng(Val(wsSettings.Cells(13, c).Value2))
        g(giDayRow) = CLng(Val(wsSettings.Cells(14, c).Value2))
        g(giCandIdRowStart) = CLng(Val(wsSettings.Cells(15, c).Value2))
        g(giCandIdRowEnd) = CLng(Val(wsSettings.Cells(16, c).Value2))
        g(giPlanStartCol) = ColToNum(CStr(wsSettings.Cells(17, c).Value2))

        groups.Add g
NextCol:
    Next c

    If Len(warnings) > 0 Then
        MsgBox "Nekatere skupine niso izpolnjene in bodo preskočene:" & vbCrLf & vbCrLf & warnings, vbExclamation
    End If

    Debug.Print "[OJT] Skupin naloženih: " & groups.Count
    Set LoadGroups = groups
End Function

Private Function IsGroupColumnValid(ByVal ws As Worksheet, ByVal c As Long, ByRef warnings As String, ByVal groupName As String) As Boolean
    Dim missing As String

    If CLng(Val(ws.Cells(4, c).Value2)) <= 0 Then missing = missing & "STOLPEC ID-JEV, "
    If CLng(Val(ws.Cells(5, c).Value2)) <= 0 Then missing = missing & "ZAČETNA VRSTICA ID-JEV, "
    If CLng(Val(ws.Cells(6, c).Value2)) <= 0 Then missing = missing & "KONČNA VRSTICA ID-JEV, "
    If Len(Trim$(CStr(ws.Cells(7, c).Value2))) = 0 Then missing = missing & "ZAČETNI STOLPEC PLANA, "
    If Len(Trim$(CStr(ws.Cells(8, c).Value2))) = 0 Then missing = missing & "KONČNI STOLPEC PLANA, "
    If Len(Trim$(CStr(ws.Cells(11, c).Value2))) = 0 Then missing = missing & "ZAČETNI STOLPEC DATUMOV, "
    If Len(Trim$(CStr(ws.Cells(12, c).Value2))) = 0 Then missing = missing & "KONČNI STOLPEC DATUMOV, "
    If CLng(Val(ws.Cells(13, c).Value2)) <= 0 Then missing = missing & "VRSTICA DATUMOV, "
    If CLng(Val(ws.Cells(14, c).Value2)) <= 0 Then missing = missing & "VRSTICA DNEVOV, "
    If CLng(Val(ws.Cells(15, c).Value2)) <= 0 Then missing = missing & "ZAČETNA VRSTICA ID-JEV KAND, "
    If CLng(Val(ws.Cells(16, c).Value2)) <= 0 Then missing = missing & "KONČNA VRSTICA ID-JEV KAND, "
    If Len(Trim$(CStr(ws.Cells(17, c).Value2))) = 0 Then missing = missing & "STOLPEC ZAČETKA PLANIRANJA OJT, "

    If Len(missing) > 0 Then
        missing = Left$(missing, Len(missing) - 2)
        warnings = warnings & "- " & groupName & ": manjkajo " & missing & "." & vbCrLf
        IsGroupColumnValid = False
    Else
        IsGroupColumnValid = True
    End If
End Function

Private Function HasAnyGroupConfig(ByVal ws As Worksheet, ByVal c As Long) As Boolean
    Dim r As Long
    For r = 4 To 17
        If Len(Trim$(CStr(ws.Cells(r, c).Value2))) > 0 Then
            HasAnyGroupConfig = True
            Exit Function
        End If
    Next r
End Function

Private Function LoadThresholds(ByVal wsSettings As Worksheet) As Object
    Dim d As Object
    Dim aps(1 To 3) As Double
    Dim acs(1 To 3) As Double

    Set d = CreateObject("Scripting.Dictionary")

    aps(1) = CDbl(Val(wsSettings.Cells(5, 12).Value2))
    aps(2) = CDbl(Val(wsSettings.Cells(6, 12).Value2))
    aps(3) = CDbl(Val(wsSettings.Cells(7, 12).Value2))

    acs(1) = CDbl(Val(wsSettings.Cells(11, 12).Value2))
    acs(2) = CDbl(Val(wsSettings.Cells(12, 12).Value2))
    acs(3) = CDbl(Val(wsSettings.Cells(13, 12).Value2))

    d.Add "APS", aps
    d.Add "ACS", acs

    Set LoadThresholds = d
End Function

Private Function GetTrackType(ByVal groupName As String) As String
    If InStr(1, UCase$(groupName), "ACS", vbTextCompare) > 0 Then
        GetTrackType = "ACS"
    Else
        GetTrackType = "APS"
    End If
End Function

Private Function ColToNum(ByVal colRef As String) As Long
    ColToNum = ThisWorkbook.Worksheets(SETTINGS_SHEET).Range(UCase$(Trim$(colRef)) & "1").Column
End Function

Private Sub EnsurePlanSheet(ByVal wb As Workbook)
    If WorksheetExists(wb, PLAN_SHEET) = False Then
        wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)).Name = PLAN_SHEET
    End If
End Sub

Private Sub ResetPlanSheet(ByVal ws As Worksheet)
    ws.Cells.Clear
End Sub

Private Function WorksheetExists(ByVal wb As Workbook, ByVal wsName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(wsName)
    WorksheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Public Sub DryRun_OJT()
    MsgBox "Dry run trenutno uporablja enako logiko kot Planiraj_OJT brez vpisa v plan (pripravi se v naslednji iteraciji).", vbInformation
End Sub
