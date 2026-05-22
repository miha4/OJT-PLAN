Attribute VB_Name = "OJTPlanner"
Option Explicit

Private Type GroupConfig
    GroupName As String
    SrcSheetName As String
    IdCol As Long
    IdRowStart As Long
    IdRowEnd As Long
    PlanColStart As Long
    PlanColEnd As Long
    DateColStart As Long
    DateColEnd As Long
    DateRow As Long
    DayRow As Long
    CandIdRowStart As Long
    CandIdRowEnd As Long
    PlanStartCol As Long
End Type

Private Type ThresholdConfig
    F1Hours As Double
    F2Hours As Double
    F3Hours As Double
End Type

Private Const SETTINGS_SHEET As String = "Nastavitve"
Private Const PLAN_SHEET As String = "OJT Plan"
Private Const SETTINGS_GROUP_ROW As Long = 3
Private Const SETTINGS_LABEL_COL As Long = 2 'B
Private Const SETTINGS_FIRST_GROUP_COL As Long = 3 'C

Public Sub Build_OJT_Plan()
    Dim trackerWb As Workbook
    Dim plannerWb As Workbook
    Dim wsPlan As Worksheet
    Dim trackerPath As String
    Dim nextOutRow As Long
    Dim groups As Collection
    Dim g As GroupConfig
    Dim i As Long

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set plannerWb = ThisWorkbook
    trackerPath = GetTrackerPath(plannerWb.Worksheets(SETTINGS_SHEET))
    Set groups = LoadGroups(plannerWb.Worksheets(SETTINGS_SHEET))

    EnsurePlanSheet plannerWb
    Set wsPlan = plannerWb.Worksheets(PLAN_SHEET)
    ResetPlanSheet wsPlan

    Set trackerWb = OpenTrackerWorkbook(trackerPath)

    nextOutRow = 1
    For i = 1 To groups.Count
        g = groups(i)
        nextOutRow = CopyGroupToPlan(trackerWb.Worksheets(g.SrcSheetName), wsPlan, g, nextOutRow)
    Next i

    trackerWb.Close SaveChanges:=False

Cleanup:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Sub Planiraj_OJT()
    Dim trackerWb As Workbook
    Dim plannerWb As Workbook
    Dim wsPlan As Worksheet
    Dim wsSettings As Worksheet
    Dim trackerPath As String
    Dim groups As Collection
    Dim thresholds As Object
    Dim assignments As Collection
    Dim i As Long
    Dim g As GroupConfig

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set plannerWb = ThisWorkbook
    Set wsSettings = plannerWb.Worksheets(SETTINGS_SHEET)
    trackerPath = GetTrackerPath(wsSettings)
    Set groups = LoadGroups(wsSettings)
    Set thresholds = LoadThresholds(wsSettings)

    EnsurePlanSheet plannerWb
    Set wsPlan = plannerWb.Worksheets(PLAN_SHEET)
    ResetPlanSheet wsPlan

    Set trackerWb = OpenTrackerWorkbook(trackerPath)

    Build_OJT_Plan

    Set assignments = New Collection
    For i = 1 To groups.Count
        g = groups(i)
        CollectAssignments trackerWb.Worksheets(g.SrcSheetName), wsPlan, wsSettings, g, thresholds, assignments
    Next i

    WriteAssignments wsPlan, assignments

    trackerWb.Close SaveChanges:=False

    MsgBox "Zaključeno. Dodelitev: " & assignments.Count, vbInformation

Cleanup:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub CollectAssignments(ByVal wsSrc As Worksheet, ByVal wsPlan As Worksheet, ByVal wsSettings As Worksheet, ByRef g As GroupConfig, ByVal thresholds As Object, ByRef assignments As Collection)
    Dim rowId As Long, colDate As Long
    Dim candId As String
    Dim candPhase As Long
    Dim cellValue As String
    Dim availableInstructors As Collection
    Dim chosenInstr As String
    Dim shiftCode As String
    Dim key As String

    For colDate = g.PlanColStart To g.PlanColEnd
        For rowId = g.CandIdRowStart To g.CandIdRowEnd Step 3
            candId = Trim$(CStr(wsSrc.Cells(rowId + 1, g.IdCol).Value2))
            If Len(candId) = 0 Then GoTo NextCandidate

            cellValue = UCase$(Trim$(CStr(wsSrc.Cells(rowId + 1, colDate).Value2)))
            If cellValue <> "XS" Then GoTo NextCandidate

            candPhase = ResolvePhase(wsSrc, g, rowId + 1, colDate, thresholds)
            Set availableInstructors = GetAvailableInstructors(wsSrc, g, rowId, colDate, candPhase)

            If availableInstructors.Count > 0 Then
                If PromptAssignment(wsSrc, g, rowId + 1, colDate, candPhase, availableInstructors, chosenInstr, shiftCode) Then
                    key = g.GroupName & "|" & colDate & "|" & candId
                    assignments.Add CreateAssignmentItem(g.GroupName, wsSrc.Cells(g.DateRow, colDate).Value2, colDate, candId, rowId + 1, chosenInstr, shiftCode, rowId)
                End If
            End If
NextCandidate:
        Next rowId
    Next colDate
End Sub

Private Function ResolvePhase(ByVal wsSrc As Worksheet, ByRef g As GroupConfig, ByVal candRow As Long, ByVal colDate As Long, ByVal thresholds As Object) As Long
    Dim totalHours As Double
    Dim t As ThresholdConfig
    Dim reserve As Double

    totalHours = Round(CDbl(Val(wsSrc.Cells(candRow, colDate).Value2)), 0)
    t = thresholds(GetTrackType(g.GroupName))
    reserve = 8#

    If totalHours < (t.F1Hours + reserve) Then
        ResolvePhase = 1
    ElseIf totalHours < (t.F1Hours + t.F2Hours - reserve) Then
        ResolvePhase = 2
    Else
        ResolvePhase = 3
    End If
End Function

Private Function GetAvailableInstructors(ByVal wsSrc As Worksheet, ByRef g As GroupConfig, ByVal tripleStartRow As Long, ByVal colDate As Long, ByVal phase As Long) As Collection
    Dim c As New Collection
    Dim r As Long
    Dim v As String

    If phase = 2 Then
        For r = g.IdRowStart To g.IdRowEnd
            v = UCase$(Trim$(CStr(wsSrc.Cells(r, colDate).Value2)))
            If v = "X1" Or v = "X2" Or v = "X3" Then
                c.Add Trim$(CStr(wsSrc.Cells(r, g.IdCol).Value2))
            End If
        Next r
    Else
        v = UCase$(Trim$(CStr(wsSrc.Cells(tripleStartRow, colDate).Value2)))
        If v = "X1" Or v = "X2" Or v = "X3" Then c.Add Trim$(CStr(wsSrc.Cells(tripleStartRow, g.IdCol).Value2))

        v = UCase$(Trim$(CStr(wsSrc.Cells(tripleStartRow + 2, colDate).Value2)))
        If v = "X1" Or v = "X2" Or v = "X3" Then c.Add Trim$(CStr(wsSrc.Cells(tripleStartRow + 2, g.IdCol).Value2))
    End If

    Set GetAvailableInstructors = c
End Function

Private Function PromptAssignment(ByVal wsSrc As Worksheet, ByRef g As GroupConfig, ByVal candRow As Long, ByVal colDate As Long, ByVal phase As Long, ByVal instr As Collection, ByRef chosenInstr As String, ByRef shiftCode As String) As Boolean
    Dim msg As String
    Dim i As Long
    Dim pick As Variant

    msg = "Skupina: " & g.GroupName & vbCrLf & _
          "Datum: " & Format$(wsSrc.Cells(g.DateRow, colDate).Value2, "dd.mm.") & vbCrLf & _
          "Faza: " & phase & vbCrLf & _
          "Kandidat: " & wsSrc.Cells(candRow, g.IdCol).Value2 & vbCrLf & _
          "Instruktorji:" & vbCrLf

    For i = 1 To instr.Count
        msg = msg & i & ") " & instr(i) & vbCrLf
    Next i
    msg = msg & vbCrLf & "Vpiši številko instruktorja ali 0 za brez dodelitve:"

    pick = Application.InputBox(msg, "OJT dodelitev", Type:=1)
    If pick = False Then
        PromptAssignment = False
        Exit Function
    End If

    If CLng(pick) = 0 Then
        PromptAssignment = False
        Exit Function
    End If

    If CLng(pick) < 1 Or CLng(pick) > instr.Count Then
        MsgBox "Napačna izbira.", vbExclamation
        PromptAssignment = False
        Exit Function
    End If

    chosenInstr = instr(CLng(pick))

    shiftCode = UCase$(Trim$(CStr(Application.InputBox("Vpiši izmeno (npr A9):", "Izmena", Type:=2))))
    If Len(shiftCode) = 0 Then
        PromptAssignment = False
        Exit Function
    End If

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
    Set f = ws.Columns(1).Find(What:=idValue, LookIn:=xlValues, LookAt:=xlWhole)
    If Not f Is Nothing Then FindIdRow = f.Row
End Function

Private Function CopyGroupToPlan(ByVal wsSrc As Worksheet, ByVal wsPlan As Worksheet, ByRef g As GroupConfig, ByVal outRow As Long) As Long
    Dim rng As Range
    Dim rowsCount As Long, colsCount As Long

    wsPlan.Cells(outRow, 1).Value2 = g.GroupName
    wsPlan.Cells(outRow, 1).Font.Bold = True
    outRow = outRow + 1

    Set rng = wsSrc.Range(wsSrc.Cells(g.DateRow, g.PlanColStart), wsSrc.Cells(g.IdRowEnd, g.PlanColEnd))
    rowsCount = rng.Rows.Count
    colsCount = rng.Columns.Count

    wsPlan.Cells(outRow, 1).Resize(rowsCount, colsCount).Value2 = rng.Value2
    CopyGroupToPlan = outRow + rowsCount + 2
End Function

Private Function OpenTrackerWorkbook(ByVal trackerPath As String) As Workbook
    Dim localPath As String
    localPath = trackerPath
    Set OpenTrackerWorkbook = Workbooks.Open(localPath, ReadOnly:=True)
End Function

Private Function GetTrackerPath(ByVal wsSettings As Worksheet) As String
    GetTrackerPath = Trim$(CStr(wsSettings.Cells(32, 3).Value2))
    If Len(GetTrackerPath) = 0 Then Err.Raise 5, , "Manjka pot do OJTracker (Nastavitve C32)."
End Function

Private Function LoadGroups(ByVal wsSettings As Worksheet) As Collection
    Dim groups As New Collection
    Dim c As Long
    Dim g As GroupConfig

    c = SETTINGS_FIRST_GROUP_COL
    Do While Len(Trim$(CStr(wsSettings.Cells(SETTINGS_GROUP_ROW, c).Value2))) > 0
        g.GroupName = Trim$(CStr(wsSettings.Cells(SETTINGS_GROUP_ROW, c).Value2))
        g.SrcSheetName = g.GroupName
        g.IdCol = CLng(Val(wsSettings.Cells(4, c).Value2))
        g.IdRowStart = CLng(Val(wsSettings.Cells(5, c).Value2))
        g.IdRowEnd = CLng(Val(wsSettings.Cells(6, c).Value2))
        g.PlanColStart = ColToNum(CStr(wsSettings.Cells(7, c).Value2))
        g.PlanColEnd = ColToNum(CStr(wsSettings.Cells(8, c).Value2))
        g.DateColStart = ColToNum(CStr(wsSettings.Cells(11, c).Value2))
        g.DateColEnd = ColToNum(CStr(wsSettings.Cells(12, c).Value2))
        g.DateRow = CLng(Val(wsSettings.Cells(13, c).Value2))
        g.DayRow = CLng(Val(wsSettings.Cells(14, c).Value2))
        g.CandIdRowStart = CLng(Val(wsSettings.Cells(15, c).Value2))
        g.CandIdRowEnd = CLng(Val(wsSettings.Cells(16, c).Value2))
        g.PlanStartCol = ColToNum(CStr(wsSettings.Cells(17, c).Value2))

        groups.Add g
        c = c + 1
    Loop

    Set LoadGroups = groups
End Function

Private Function LoadThresholds(ByVal wsSettings As Worksheet) As Object
    Dim d As Object
    Dim aps As ThresholdConfig, acs As ThresholdConfig

    Set d = CreateObject("Scripting.Dictionary")

    aps.F1Hours = CDbl(Val(wsSettings.Cells(5, 12).Value2))
    aps.F2Hours = CDbl(Val(wsSettings.Cells(6, 12).Value2))
    aps.F3Hours = CDbl(Val(wsSettings.Cells(7, 12).Value2))

    acs.F1Hours = CDbl(Val(wsSettings.Cells(11, 12).Value2))
    acs.F2Hours = CDbl(Val(wsSettings.Cells(12, 12).Value2))
    acs.F3Hours = CDbl(Val(wsSettings.Cells(13, 12).Value2))

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
    ColToNum = Range(UCase$(Trim$(colRef)) & "1").Column
End Function

Private Sub EnsurePlanSheet(ByVal wb As Workbook)
    On Error Resume Next
    wb.Worksheets(PLAN_SHEET).Name = PLAN_SHEET
    On Error GoTo 0
    If WorksheetExists(wb, PLAN_SHEET) = False Then
        wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)).Name = PLAN_SHEET
    End If
End Sub

Private Sub ResetPlanSheet(ByVal ws As Worksheet)
    ws.Cells.Clear
End Sub

Private Function WorksheetExists(ByVal wb As Workbook, ByVal wsName As String) As Boolean
    On Error Resume Next
    WorksheetExists = Not wb.Worksheets(wsName) Is Nothing
    On Error GoTo 0
End Function

Public Sub DryRun_OJT()
    MsgBox "Dry run trenutno uporablja enako logiko kot Planiraj_OJT brez vpisa v plan (pripravi se v naslednji iteraciji).", vbInformation
End Sub
