Attribute VB_Name = "OJTPlanner"
Option Explicit

Private Const SETTINGS_SHEET As String = "Nastavitve"
Private Const PLAN_SHEET As String = "OJT Plan"
Private Const SETTINGS_GROUP_ROW As Long = 3
Private Const SETTINGS_FIRST_GROUP_COL As Long = 3 'C
Private mPlanRowMap As Object

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
    giHoursRowStart = 15
    giHoursRowEnd = 16
    giPlanEnabled = 17
End Enum

Public Sub Build_OJT_Plan()
    Dim errMsg As String
    Dim trackerWb As Workbook
    Dim closeTrackerOnExit As Boolean
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

    Set trackerWb = OpenTrackerWorkbook(trackerPath, closeTrackerOnExit)

    Set mPlanRowMap = CreateObject("Scripting.Dictionary")
    Set mPlanRowMap = CreateObject("Scripting.Dictionary")
    nextOutRow = 1
    For i = 1 To groups.Count
        g = groups(i)
        Debug.Print "[OJT] Build copy group: " & CStr(g(giGroupName))
        nextOutRow = CopyGroupToPlan(GetWorksheetOrFail(trackerWb, CStr(g(giSrcSheetName))), wsPlan, g, nextOutRow)
    Next i


Cleanup:
    On Error Resume Next
    If closeTrackerOnExit And Not trackerWb Is Nothing Then trackerWb.Close SaveChanges:=False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    On Error GoTo 0
    Exit Sub

EH:
    errMsg = "Build_OJT_Plan napaka " & Err.Number & ": " & Err.Description
    Debug.Print errMsg
    MsgBox errMsg, vbCritical
    Resume Cleanup
End Sub

Public Sub Planiraj_OJT()
    Dim errMsg As String
    Dim trackerWb As Workbook
    Dim closeTrackerOnExit As Boolean
    Dim plannerWb As Workbook
    Dim wsPlan As Worksheet
    Dim wsSettings As Worksheet
    Dim trackerPath As String
    Dim groups As Collection
    Dim thresholds As Object
    Dim assignments As Collection
    Dim liveHours As Object
    Dim history As Collection
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

    Set trackerWb = OpenTrackerWorkbook(trackerPath, closeTrackerOnExit)

    Dim nextOutRow As Long
    nextOutRow = 1
    For i = 1 To groups.Count
        g = groups(i)
        Debug.Print "[OJT] Build copy group: " & CStr(g(giGroupName))
        nextOutRow = CopyGroupToPlan(GetWorksheetOrFail(trackerWb, CStr(g(giSrcSheetName))), wsPlan, g, nextOutRow)
    Next i

    Set assignments = New Collection
    Set liveHours = CreateObject("Scripting.Dictionary")
    Set history = New Collection
    For i = 1 To groups.Count
        g = groups(i)
        CollectAssignments GetWorksheetOrFail(trackerWb, CStr(g(giSrcSheetName))), wsPlan, g, thresholds, assignments, liveHours, wsPlan, history
    Next i

    MsgBox "Zaključeno. Dodelitev: " & assignments.Count, vbInformation


Cleanup:
    On Error Resume Next
    If closeTrackerOnExit And Not trackerWb Is Nothing Then trackerWb.Close SaveChanges:=False
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

Private Sub CollectAssignments(ByVal wsSrc As Worksheet, ByVal wsPlan As Worksheet, ByVal g As Variant, ByVal thresholds As Object, ByRef assignments As Collection, ByRef liveHours As Object, ByVal wsPlanOut As Worksheet, ByRef history As Collection)
    Dim rowId As Long, colDate As Long
    Dim candId As String
    Dim candPhase As Long
    Dim cellValue As String
    Dim availableInstructors As Collection
    Dim chosenInstr As String
    Dim shiftCode As String
    Dim hoursRow As Long

    For colDate = CLng(g(giPlanColStart)) To CLng(g(giPlanColEnd))
        For rowId = CLng(g(giIdRowStart)) To CLng(g(giIdRowEnd))
            cellValue = UCase$(Trim$(CStr(wsSrc.Cells(rowId, colDate).Value2)))
            If cellValue <> "XS" Then GoTo NextCandidate

            candId = Trim$(CStr(wsSrc.Cells(rowId, CLng(g(giIdCol))).Value2))
            If Len(candId) = 0 Then GoTo NextCandidate

            hoursRow = FindHoursRowById(wsSrc, g, candId)
            If hoursRow = 0 Then GoTo NextCandidate

            candPhase = ResolvePhaseLive(wsSrc, g, hoursRow, colDate, thresholds, liveHours, candId)
            Set availableInstructors = GetAvailableInstructors(wsSrc, g, rowId - 1, colDate, candPhase)

            If availableInstructors.Count > 0 Then
                If PromptAssignmentUnified(wsSrc, g, rowId, colDate, candPhase, availableInstructors, chosenInstr, shiftCode, liveHours, candId) Then
                    Dim addH As Double
                    addH = ShiftHoursForDate(wsSrc, g, colDate)
                    IncrementLiveHours liveHours, candId, addH
                    Dim itm As Variant
                    Dim instrSrcRow As Long
                    instrSrcRow = FindSourceRowById(wsSrc, g, chosenInstr)
                    itm = CreateAssignmentItem(CStr(g(giGroupName)), wsSrc.Cells(CLng(g(giDateRow)), colDate).Value2, 2 + (colDate - CLng(g(giPlanColStart)) + 1), colDate, candId, rowId, chosenInstr, shiftCode, instrSrcRow, CDbl(liveHours(UCase$(candId))), addH)
                    assignments.Add itm
                    ApplySingleAssignment wsPlanOut, itm
                    RefreshPlanView wsPlanOut
                    history.Add itm
                ElseIf UCase$(chosenInstr) = "__BACK__" Then
                    Dim undone As Variant
                    If UndoLastAssignment(wsPlanOut, history, assignments, liveHours, undone) Then
                        If CStr(undone(1)) = CStr(g(giGroupName)) Then
                            Dim backCol As Long, backRow As Long, backCandId As String, backPhase As Long
                            Dim backAvail As Collection, backShift As String, backInstr As String
                            backCol = CLng(undone(3))
                            backRow = CLng(undone(6))
                            backCandId = CStr(undone(4))
                            hoursRow = FindHoursRowById(wsSrc, g, backCandId)
                            If hoursRow > 0 Then
                                backPhase = ResolvePhaseLive(wsSrc, g, hoursRow, backCol, thresholds, liveHours, backCandId)
                                Set backAvail = GetAvailableInstructors(wsSrc, g, backRow - 1, backCol, backPhase)
                                If backAvail.Count > 0 Then
                                    If PromptAssignmentUnified(wsSrc, g, backRow, backCol, backPhase, backAvail, backInstr, backShift, liveHours, backCandId) Then
                                        Dim backAddH As Double, backInstrSrcRow As Long, backItem As Variant
                                        backAddH = ShiftHoursForDate(wsSrc, g, backCol)
                                        IncrementLiveHours liveHours, backCandId, backAddH
                                        backInstrSrcRow = FindSourceRowById(wsSrc, g, backInstr)
                                        backItem = CreateAssignmentItem(CStr(g(giGroupName)), wsSrc.Cells(CLng(g(giDateRow)), backCol).Value2, 2 + (backCol - CLng(g(giPlanColStart)) + 1), backCol, backCandId, backRow, backInstr, backShift, backInstrSrcRow, CDbl(liveHours(UCase$(backCandId))), backAddH)
                                        assignments.Add backItem
                                        ApplySingleAssignment wsPlanOut, backItem
                                        history.Add backItem
                                    End If
                                End If
                            End If
                            rowId = rowId - 1
                        Else
                            rowId = rowId - 1
                        End If
                    End If
                    RefreshPlanView wsPlanOut
                    chosenInstr = ""
                    GoTo NextCandidate
                ElseIf UCase$(chosenInstr) = "__END__" Then
                    Exit Sub
                End If
            End If
NextCandidate:
        Next rowId
    Next colDate
End Sub



Private Function PromptAssignmentUnified(ByVal wsSrc As Worksheet, ByVal g As Variant, ByVal candRow As Long, ByVal colDate As Long, ByVal phase As Long, ByVal instrList As Collection, ByRef chosenInstr As String, ByRef shiftCode As String, ByRef liveHours As Object, ByVal candId As String) As Boolean
    Dim msg As String, inputText As Variant, parts() As String
    Dim i As Long, idx As Long

    msg = "Skupina: " & CStr(g(giGroupName)) & vbCrLf & _
          "**DATUM**: " & Format$(wsSrc.Cells(CLng(g(giDateRow)), colDate).Value2, "dd.mm.") & vbCrLf & _
          "Faza: " & phase & vbCrLf & _
          "**Kandidat**: " & wsSrc.Cells(candRow, CLng(g(giIdCol))).Value2 & " - " & wsSrc.Cells(candRow, CLng(g(giIdCol)) + 1).Value2 & vbCrLf & _
          "🔴 PREDVIDENE URE: " & CStr(CDbl(liveHours(UCase$(candId)))) & vbCrLf & _
          "Instruktorji:" & vbCrLf
    For i = 1 To instrList.Count
        msg = msg & i & ") " & instrList(i) & vbCrLf
    Next i
    msg = msg & vbCrLf & "Vnos: indeks;izmena (npr 1;A9)" & vbCrLf & "0 = preskoči, B = nazaj, K = končaj"

    inputText = Application.InputBox(msg, "OJT dodelitev", Type:=2)
    If inputText = False Then Exit Function
    inputText = UCase$(Trim$(CStr(inputText)))
    If inputText = "" Or inputText = "0" Then Exit Function
    If inputText = "B" Then chosenInstr = "__BACK__": Exit Function
    If inputText = "K" Then chosenInstr = "__END__": Exit Function

    parts = Split(CStr(inputText), ";")
    If UBound(parts) <> 1 Then MsgBox "Uporabi format indeks;izmena (npr 1;A9)", vbExclamation: Exit Function
    idx = CLng(Val(parts(0)))
    If idx < 1 Or idx > instrList.Count Then MsgBox "Napačen indeks inštruktorja.", vbExclamation: Exit Function

    chosenInstr = CStr(instrList(idx))
    shiftCode = Trim$(parts(1))
    If Len(shiftCode) = 0 Then MsgBox "Manjka izmena.", vbExclamation: Exit Function
    PromptAssignmentUnified = True
End Function

Private Sub ApplySingleAssignment(ByVal wsPlan As Worksheet, ByVal a As Variant)
    Dim rowCand As Long, rowInstr As Long
    rowCand = GetPlanRowFromSource(CStr(a(1)), CLng(a(6)), wsPlan, CStr(a(4)))
    rowInstr = GetPlanRowFromSource(CStr(a(1)), CLng(a(9)), wsPlan, CStr(a(7)))
    If rowCand > 0 Then
        wsPlan.Cells(rowCand, CLng(a(2))).Value2 = CStr(a(8)) & "s"
        AddOrReplaceComment wsPlan.Cells(rowCand, CLng(a(2))), "OJT: " & CStr(a(7)) & " - " & CStr(a(4)) & " | predvidene ure: " & CStr(a(10))
    End If
    If rowInstr > 0 Then
        wsPlan.Cells(rowInstr, CLng(a(2))).Value2 = CStr(a(8)) & "i"
        AddOrReplaceComment wsPlan.Cells(rowInstr, CLng(a(2))), "OJT: " & CStr(a(7)) & " - " & CStr(a(4))
    End If
End Sub


Private Function GetPlanRowFromSource(ByVal groupName As String, ByVal srcRow As Long, ByVal wsPlan As Worksheet, ByVal fallbackId As String) As Long
    Dim k As String
    If Not mPlanRowMap Is Nothing Then
        k = groupName & "|" & CStr(srcRow)
        If mPlanRowMap.Exists(k) Then
            GetPlanRowFromSource = CLng(mPlanRowMap(k))
            Exit Function
        End If
    End If
    GetPlanRowFromSource = FindIdRow(wsPlan, fallbackId)
End Function

Private Function UndoLastAssignment(ByVal wsPlan As Worksheet, ByRef history As Collection, ByRef assignments As Collection, ByRef liveHours As Object, ByRef undoneItem As Variant) As Boolean
    Dim a As Variant, rowCand As Long, rowInstr As Long
    If history.Count = 0 Then Exit Function
    a = history(history.Count)
    history.Remove history.Count
    If assignments.Count > 0 Then assignments.Remove assignments.Count
    rowCand = GetPlanRowFromSource(CStr(a(1)), CLng(a(6)), wsPlan, CStr(a(4)))
    rowInstr = GetPlanRowFromSource(CStr(a(1)), CLng(a(9)), wsPlan, CStr(a(7)))
    If rowCand > 0 Then wsPlan.Cells(rowCand, CLng(a(2))).ClearContents
    If rowInstr > 0 Then wsPlan.Cells(rowInstr, CLng(a(2))).ClearContents
    IncrementLiveHours liveHours, CStr(a(4)), -CDbl(a(11))
    undoneItem = a
    UndoLastAssignment = True
End Function


Private Function FindSourceRowById(ByVal wsSrc As Worksheet, ByVal g As Variant, ByVal empId As String) As Long
    Dim r As Long
    For r = CLng(g(giIdRowStart)) To CLng(g(giIdRowEnd))
        If UCase$(Trim$(CStr(wsSrc.Cells(r, CLng(g(giIdCol))).Value2))) = UCase$(Trim$(empId)) Then
            FindSourceRowById = r
            Exit Function
        End If
    Next r
End Function

Private Function FindHoursRowById(ByVal wsSrc As Worksheet, ByVal g As Variant, ByVal candId As String) As Long
    Dim r As Long
    For r = CLng(g(giCandIdRowStart)) To CLng(g(giCandIdRowEnd))
        If UCase$(Trim$(CStr(wsSrc.Cells(r, CLng(g(giIdCol))).Value2))) = UCase$(candId) Then
            FindHoursRowById = r
            Exit Function
        End If
    Next r
End Function

Private Function ResolvePhaseLive(ByVal wsSrc As Worksheet, ByVal g As Variant, ByVal candRow As Long, ByVal colDate As Long, ByVal thresholds As Object, ByRef liveHours As Object, ByVal candId As String) As Long
    Dim key As String
    Dim baseHours As Double
    Dim reserveHours As Double

    key = UCase$(candId)
    If Not liveHours.Exists(key) Then
        baseHours = Round(CDbl(Val(wsSrc.Cells(candRow, colDate).Value2)), 0)
        liveHours.Add key, baseHours
    End If

    reserveHours = ShiftHoursForDate(wsSrc, g, colDate)
    ResolvePhaseLive = ResolvePhaseFromHours(CDbl(liveHours(key)), thresholds, GetTrackType(CStr(g(giGroupName))), reserveHours)
End Function

Private Sub IncrementLiveHours(ByRef liveHours As Object, ByVal candId As String, ByVal addHours As Double)
    Dim key As String
    key = UCase$(candId)
    If Not liveHours.Exists(key) Then liveHours.Add key, 0#
    liveHours(key) = CDbl(liveHours(key)) + addHours
End Sub

Private Function ShiftHoursForDate(ByVal wsSrc As Worksheet, ByVal g As Variant, ByVal colDate As Long) As Double
    Dim r As Long
    Dim v As Variant
    r = CLng(g(giCandIdRowStart)) - 1
    If r > 0 Then
        v = wsSrc.Cells(r, colDate).Value2
        If IsNumeric(v) Then
            ShiftHoursForDate = CDbl(v)
            Exit Function
        End If
    End If
    ShiftHoursForDate = 8#
End Function

Private Function ResolvePhaseFromHours(ByVal totalHours As Double, ByVal thresholds As Object, ByVal trackType As String, Optional ByVal reserveHours As Double = 0#) As Long
    Dim t As Variant
    Dim phase1Limit As Double
    Dim phase2Limit As Double
    
    t = thresholds(trackType)

    ' Fazo določimo izključno iz seštevka ur za izbran track (APP/APS/ACS)
    ' + rezerva ene izmene (dinamično iz glave ur za datum), brez prištevanja drugih trackov.
    phase1Limit = CDbl(t(1))
    phase2Limit = phase1Limit + CDbl(t(2))

    If totalHours < (phase1Limit + reserveHours) Then
        ResolvePhaseFromHours = 1
    ElseIf totalHours < (phase2Limit - reserveHours) Then
        ResolvePhaseFromHours = 2
    Else
        ResolvePhaseFromHours = 3
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

Private Function PromptAssignment(ByVal wsSrc As Worksheet, ByVal g As Variant, ByVal candRow As Long, ByVal colDate As Long, ByVal phase As Long, ByVal instrList As Collection, ByRef chosenInstr As String, ByRef shiftCode As String) As Boolean
    Dim msg As String
    Dim i As Long
    Dim pick As Variant

    msg = "Skupina: " & CStr(g(giGroupName)) & vbCrLf & _
          "Datum: " & Format$(wsSrc.Cells(CLng(g(giDateRow)), colDate).Value2, "dd.mm.") & vbCrLf & _
          "Faza: " & phase & vbCrLf & _
          "Kandidat: " & wsSrc.Cells(candRow, CLng(g(giIdCol))).Value2 & " - " & wsSrc.Cells(candRow, CLng(g(giIdCol)) + 1).Value2 & vbCrLf & _
          "Instruktorji:" & vbCrLf

    For i = 1 To instrList.Count
        msg = msg & i & ") " & instrList(i) & vbCrLf
    Next i
    msg = msg & vbCrLf & "Vpiši številko instruktorja ali 0 za brez dodelitve:"

    pick = Application.InputBox(msg, "OJT dodelitev", Type:=1)
    If pick = False Then Exit Function
    If CLng(pick) = 0 Then Exit Function

    If CLng(pick) < 1 Or CLng(pick) > instrList.Count Then
        MsgBox "Napačna izbira.", vbExclamation
        Exit Function
    End If

    chosenInstr = instrList(CLng(pick))
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
        rowCand = FindIdRow(wsPlan, CStr(a(4)))
        rowInstr = FindIdRow(wsPlan, CStr(a(7)))

        If rowCand > 0 Then
            wsPlan.Cells(rowCand, CLng(a(2))).Value2 = CStr(a(8)) & "s"
            AddOrReplaceComment wsPlan.Cells(rowCand, CLng(a(2))), "OJT: " & CStr(a(7)) & " - " & CStr(a(4)) & " | predvidene ure: " & CStr(a(10))
        End If

        If rowInstr > 0 Then
            wsPlan.Cells(rowInstr, CLng(a(2))).Value2 = CStr(a(8)) & "i"
            AddOrReplaceComment wsPlan.Cells(rowInstr, CLng(a(2))), "OJT: " & CStr(a(7)) & " - " & CStr(a(4))
        End If
    Next i
End Sub


Private Sub RefreshPlanView(ByVal ws As Worksheet)
    Dim prevUpd As Boolean
    prevUpd = Application.ScreenUpdating
    Application.ScreenUpdating = True
    ws.Calculate
    DoEvents
    Application.ScreenUpdating = prevUpd
End Sub

Private Sub AddOrReplaceComment(ByVal cell As Range, ByVal text As String)
    On Error Resume Next
    cell.CommentThreaded.Delete
    cell.ClearComments
    On Error GoTo 0
    cell.AddCommentThreaded text
End Sub

Private Function CreateAssignmentItem(ByVal groupName As String, ByVal planDate As Variant, ByVal colDate As Long, ByVal srcColDate As Long, ByVal candId As String, ByVal candRow As Long, ByVal instrId As String, ByVal shiftCode As String, ByVal tripleStart As Long, ByVal liveHoursAfter As Double, ByVal hoursAdded As Double) As Variant
    Dim a(1 To 11) As Variant
    a(1) = groupName
    a(2) = colDate
    a(3) = srcColDate
    a(4) = candId
    a(5) = planDate
    a(6) = candRow
    a(7) = instrId
    a(8) = shiftCode
    a(9) = tripleStart
    a(10) = liveHoursAfter
    a(11) = hoursAdded
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
    Dim planStartCol As Long, planEndCol As Long
    Dim dateRow As Long, dayRow As Long
    Dim idRowStart As Long, idRowEnd As Long
    Dim planCols As Long
    Dim rowCount As Long, r As Long, c As Long
    Dim outData() As Variant
    Dim seenId As Object
    Dim includeRows() As Boolean
    Dim srcId As String

    idCol = CLng(g(giIdCol))
    nameCol = idCol + 1
    planStartCol = CLng(g(giPlanColStart))
    planEndCol = CLng(g(giPlanColEnd))
    dateRow = CLng(g(giDateRow))
    dayRow = CLng(g(giDayRow))
    idRowStart = CLng(g(giIdRowStart))
    idRowEnd = CLng(g(giIdRowEnd))
    planCols = planEndCol - planStartCol + 1

    wsPlan.Cells(outRow, 1).Value2 = CStr(g(giGroupName))
    wsPlan.Cells(outRow, 1).Font.Bold = True
    outRow = outRow + 1

    ReDim outData(1 To (idRowEnd - idRowStart + 4), 1 To planCols + 2)

    rowCount = 0

    ' Glava 1: datumi
    rowCount = rowCount + 1
    outData(rowCount, 1) = "ID"
    outData(rowCount, 2) = "Ime in priimek"
    For c = 1 To planCols
        outData(rowCount, c + 2) = wsSrc.Cells(dateRow, planStartCol + c - 1).Value2
    Next c

    ' Glava 2: dnevi
    rowCount = rowCount + 1
    outData(rowCount, 1) = ""
    outData(rowCount, 2) = ""
    For c = 1 To planCols
        outData(rowCount, c + 2) = wsSrc.Cells(dayRow, planStartCol + c - 1).Value2
    Next c

    ' Podatki: samo vrstice z ID-jem (odstrani prazne vrstice)
    ReDim includeRows(idRowStart To idRowEnd)
    Set seenId = CreateObject("Scripting.Dictionary")
    For r = idRowStart To idRowEnd
        srcId = UCase$(Trim$(CStr(wsSrc.Cells(r, idCol).Value2)))
        If Len(srcId) > 0 Then
            includeRows(r) = True
            If seenId.Exists(srcId) Then
                includeRows(CLng(seenId(srcId))) = False ' pri podvojitvi obdrži zadnjo vrstico
            End If
            seenId(srcId) = r
        End If
    Next r

    For r = idRowStart To idRowEnd
        If includeRows(r) Then
            rowCount = rowCount + 1
            outData(rowCount, 1) = wsSrc.Cells(r, idCol).Value2
            mPlanRowMap(CStr(g(giGroupName)) & "|" & CStr(r)) = outRow + rowCount - 1
            outData(rowCount, 2) = wsSrc.Cells(r, nameCol).Value2
            For c = 1 To planCols
                outData(rowCount, c + 2) = wsSrc.Cells(r, planStartCol + c - 1).Value2
            Next c
        End If
    Next r

    If rowCount > 0 Then
        wsPlan.Cells(outRow, 1).Resize(rowCount, planCols + 2).Value2 = outData
        wsPlan.Cells(outRow, 1).Resize(1, planCols + 2).Font.Bold = True
        wsPlan.Cells(outRow, 3).Resize(1, planCols).Font.Color = RGB(0, 0, 255)
    End If

    Dim hrS As Long, hrE As Long, hrRows As Long
    hrS = CLng(g(giHoursRowStart)): hrE = CLng(g(giHoursRowEnd))
    If hrS > 0 And hrE >= hrS Then
        hrRows = hrE - hrS + 1
        wsPlan.Cells(outRow + rowCount + 1, 1).Value2 = "URE (kopija)"
        wsPlan.Cells(outRow + rowCount + 1, 1).Font.Bold = True
        wsPlan.Cells(outRow + rowCount + 1, 1).Font.Color = RGB(255, 0, 0)
        wsPlan.Cells(outRow + rowCount + 2, 1).Resize(hrRows, 1).Value2 = wsSrc.Range(wsSrc.Cells(hrS, idCol), wsSrc.Cells(hrE, idCol)).Value2
        wsPlan.Cells(outRow + rowCount + 2, 2).Resize(hrRows, 1).Value2 = wsSrc.Range(wsSrc.Cells(hrS, nameCol), wsSrc.Cells(hrE, nameCol)).Value2
        wsPlan.Cells(outRow + rowCount + 2, 3).Resize(hrRows, planCols).Value2 = wsSrc.Range(wsSrc.Cells(hrS, planStartCol), wsSrc.Cells(hrE, planEndCol)).Value2
        wsPlan.Cells(outRow + rowCount + 2, 1).Resize(hrRows, planCols + 2).Font.Bold = True
        wsPlan.Cells(outRow + rowCount + 2, 1).Resize(hrRows, planCols + 2).Font.Color = RGB(255, 0, 0)
        CopyGroupToPlan = outRow + rowCount + hrRows + 4
    Else
        CopyGroupToPlan = outRow + rowCount + 2
    End If
End Function

Private Function OpenTrackerWorkbook(ByVal trackerPath As String, ByRef closeOnExit As Boolean) As Workbook
    Dim wb As Workbook
    Dim normalized As String

    normalized = NormalizeSharePointPath(trackerPath)
    Debug.Print "[OJT] Odpiram tracker: "; normalized

    Set wb = FindOpenWorkbook(normalized)
    If Not wb Is Nothing Then
        closeOnExit = False
        Debug.Print "[OJT] Tracker že odprt: "; wb.FullName
        Set OpenTrackerWorkbook = wb
        Exit Function
    End If

    On Error Resume Next
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Filename:=normalized, UpdateLinks:=0, ReadOnly:=True, Notify:=False, AddToMru:=False)
    If Not wb Is Nothing Then closeOnExit = True
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
    Dim g(1 To 17) As Variant
    Dim groupName As String
    Dim warnings As String
    Dim hasPlanToggle As Boolean
    Dim planFlag As String

    lastCol = wsSettings.Cells(SETTINGS_GROUP_ROW, wsSettings.Columns.Count).End(xlToLeft).Column
    If lastCol < SETTINGS_FIRST_GROUP_COL Then
        Set LoadGroups = groups
        Exit Function
    End If

    hasPlanToggle = HasAnyPlanToggle(wsSettings, lastCol)

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
        planFlag = UCase$(Trim$(CStr(wsSettings.Cells(18, c).Value2)))
        g(giPlanEnabled) = planFlag
        g(giHoursRowStart) = CLng(Val(wsSettings.Cells(19, c).Value2))
        g(giHoursRowEnd) = CLng(Val(wsSettings.Cells(20, c).Value2))

        If hasPlanToggle Then
            If planFlag <> "DA" Then GoTo NextCol
        End If

        groups.Add g
NextCol:
    Next c

    If Len(warnings) > 0 Then
        MsgBox "Nekatere skupine niso izpolnjene in bodo preskočene:" & vbCrLf & vbCrLf & warnings, vbExclamation
    End If

    Debug.Print "[OJT] Skupin naloženih: " & groups.Count
    Set LoadGroups = groups
End Function

Private Function HasAnyPlanToggle(ByVal ws As Worksheet, ByVal lastCol As Long) As Boolean
    Dim c As Long
    Dim v As String
    For c = SETTINGS_FIRST_GROUP_COL To lastCol
        v = UCase$(Trim$(CStr(ws.Cells(18, c).Value2)))
        If v = "DA" Or v = "NE" Then
            HasAnyPlanToggle = True
            Exit Function
        End If
    Next c
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
    Dim app(1 To 3) As Double

    Set d = CreateObject("Scripting.Dictionary")

    app(1) = CDbl(Val(wsSettings.Cells(5, 11).Value2))
    app(2) = CDbl(Val(wsSettings.Cells(6, 11).Value2))
    app(3) = CDbl(Val(wsSettings.Cells(7, 11).Value2))

    aps(1) = CDbl(Val(wsSettings.Cells(8, 11).Value2))
    aps(2) = CDbl(Val(wsSettings.Cells(9, 11).Value2))
    aps(3) = CDbl(Val(wsSettings.Cells(10, 11).Value2))

    acs(1) = CDbl(Val(wsSettings.Cells(14, 11).Value2))
    acs(2) = CDbl(Val(wsSettings.Cells(15, 11).Value2))
    acs(3) = CDbl(Val(wsSettings.Cells(16, 11).Value2))

    d.Add "APP", app
    d.Add "APS", aps
    d.Add "ACS", acs

    Set LoadThresholds = d
End Function

Private Function GetTrackType(ByVal groupName As String) As String
    If InStr(1, UCase$(groupName), "APP", vbTextCompare) > 0 Then
        GetTrackType = "APP"
    ElseIf InStr(1, UCase$(groupName), "APS", vbTextCompare) > 0 Then
        GetTrackType = "APS"
    ElseIf InStr(1, UCase$(groupName), "ACS", vbTextCompare) > 0 Then
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
    On Error Resume Next
    ws.Cells.ClearContents
    ws.Cells.ClearComments
    On Error GoTo 0
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
