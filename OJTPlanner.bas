Option Explicit

Private Const SETTINGS_SHEET As String = "Nastavitve"
Private Const PLAN_SHEET As String = "OJT Plan"
Private Const SETTINGS_GROUP_ROW As Long = 3
Private Const SETTINGS_FIRST_GROUP_COL As Long = 3 'C
Private Const GROUP_RESERVED_BLANK_ROWS As Long = 2
Private Const GROUP_RESERVED_HOURS_ROWS As Long = 5
Private Const CANDIDATE_PANEL_ROWS As Long = 4
Private mPlanRowMap As Object
Private mCandidatePanelIndex As Object

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
    Set groups = LoadGroups(plannerWb.Worksheets(SETTINGS_SHEET), True)
    If groups.Count = 0 Then Err.Raise 9001, , "V Nastavitve (vrstica 3, od stolpca C naprej) ni nobene skupine."

    EnsurePlanSheet plannerWb
    Set wsPlan = plannerWb.Worksheets(PLAN_SHEET)
    ResetPlanSheet wsPlan

    Set trackerWb = OpenTrackerWorkbook(trackerPath, closeTrackerOnExit)

    Set mPlanRowMap = CreateObject("Scripting.Dictionary")
    Set mCandidatePanelIndex = CreateObject("Scripting.Dictionary")
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
    Dim activeGroupsCount As Long
    Dim totalPrompts As Long
    Dim groupXs As Long
    Dim groupMissingHours As Long
    Dim groupNoInstructors As Long
    Dim diagnostics As String

    On Error GoTo EH
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set plannerWb = ThisWorkbook
    Set wsSettings = plannerWb.Worksheets(SETTINGS_SHEET)
    trackerPath = GetTrackerPath(wsSettings)
    Set groups = LoadGroups(wsSettings, True)
    If groups.Count = 0 Then Err.Raise 9001, , "V Nastavitve (vrstica 3, od stolpca C naprej) ni nobene skupine."
    Set thresholds = LoadThresholds(wsSettings)

    EnsurePlanSheet plannerWb
    Set wsPlan = plannerWb.Worksheets(PLAN_SHEET)
    ResetPlanSheet wsPlan

    Set trackerWb = OpenTrackerWorkbook(trackerPath, closeTrackerOnExit)

    Set mPlanRowMap = CreateObject("Scripting.Dictionary")
    Set mCandidatePanelIndex = CreateObject("Scripting.Dictionary")
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
        If IsGroupPlanningEnabled(g) Then
            activeGroupsCount = activeGroupsCount + 1
            groupXs = 0
            groupMissingHours = 0
            groupNoInstructors = 0
            Debug.Print "[OJT] Planiram group: " & CStr(g(giGroupName)) & " (PLANIRAJ=" & CStr(g(giPlanEnabled)) & ")"
            totalPrompts = totalPrompts + CollectAssignments( _
                GetWorksheetOrFail(trackerWb, CStr(g(giSrcSheetName))), _
                wsPlan, g, thresholds, assignments, liveHours, wsPlan, history, _
                groupXs, groupMissingHours, groupNoInstructors)
            diagnostics = diagnostics & CStr(g(giGroupName)) & ": Xs=" & groupXs & ", brez vrstice ur=" & groupMissingHours & ", brez inštruktorjev=" & groupNoInstructors & vbCrLf
        Else
            Debug.Print "[OJT] Preskočim group: " & CStr(g(giGroupName)) & " (PLANIRAJ=" & CStr(g(giPlanEnabled)) & ")"
        End If
    Next i

    If activeGroupsCount = 0 Then
        MsgBox "Ni aktivnih skupin za planiranje. V vrstici PLANIRAJ nastavi DA pri skupini, ki jo želiš planirati.", vbExclamation
    ElseIf totalPrompts = 0 Then
        MsgBox "Zaključeno. Dodelitev: 0" & vbCrLf & vbCrLf & _
               "Makro ni našel nobenega kandidata, za katerega bi lahko ponudil izbor izmene." & vbCrLf & _
               "Preveri diagnostiko po skupinah:" & vbCrLf & diagnostics, vbExclamation
    Else
        MsgBox "Zaključeno. Dodelitev: " & assignments.Count, vbInformation
    End If


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

Private Function CollectAssignments( _
    ByVal wsSrc As Worksheet, _
    ByVal wsPlan As Worksheet, _
    ByVal g As Variant, _
    ByVal thresholds As Object, _
    ByRef assignments As Collection, _
    ByRef liveHours As Object, _
    ByVal wsPlanOut As Worksheet, _
    ByRef history As Collection, _
    ByRef foundXs As Long, _
    ByRef missingHoursRows As Long, _
    ByRef noInstructorCandidates As Long) As Long
    Dim rowId As Long, colDate As Long
    Dim rowStart As Long, rowEnd As Long, colStart As Long, colEnd As Long
    Dim candId As String
    Dim candPhase As Long
    Dim cellValue As String
    Dim availableInstructors As Collection
    Dim chosenInstr As String
    Dim shiftCode As String
    Dim hoursRow As Long
    Dim undone As Variant
    Dim candName As String
    Dim hoursRowCache As Object
    Dim requireCandidateMarker As Boolean

    rowStart = CLng(g(giIdRowStart))
    rowEnd = CLng(g(giIdRowEnd))
    colStart = GetPlanningStartColumn(g)
    colEnd = CLng(g(giPlanColEnd))
    rowId = rowStart
    colDate = colStart
    Set hoursRowCache = CreateObject("Scripting.Dictionary")
    requireCandidateMarker = HasCandidateTrainingRows(wsSrc, g)

    Do While colDate <= colEnd
        If rowId > rowEnd Then
            rowId = rowStart
            colDate = colDate + 1
            GoTo ContinueLoop
        End If

            If requireCandidateMarker Then
                If Not IsCandidateTrainingRow(wsSrc, g, rowId) Then GoTo NextCandidate
            End If

            cellValue = NormalizeScheduleCode(wsSrc.Cells(rowId, colDate).Value2)
            If cellValue <> "XS" Then GoTo NextCandidate

            foundXs = foundXs + 1
            candId = Trim$(CStr(wsSrc.Cells(rowId, CLng(g(giIdCol))).Value2))
            If Len(candId) = 0 Then GoTo NextCandidate
            candName = Trim$(CStr(wsSrc.Cells(rowId, CLng(g(giIdCol)) + 1).Value2))

            hoursRow = GetCachedHoursRow(wsSrc, g, candId, candName, hoursRowCache)
            If hoursRow = 0 Then
                missingHoursRows = missingHoursRows + 1
                EnsureLiveHours liveHours, candId, 0#
                candPhase = ResolvePhaseFromHours(0#, thresholds, GetTrackType(CStr(g(giGroupName))), ShiftHoursForDate(wsSrc, g, colDate))
                Debug.Print "[OJT] Kandidat " & candId & " nima vrstice ur; nadaljujem z 0 urami. Skupina: " & CStr(g(giGroupName))
            Else
                candPhase = ResolvePhaseLive(wsSrc, g, hoursRow, colDate, thresholds, liveHours, candId)
            End If
            Set availableInstructors = GetAvailableInstructors(wsSrc, g, rowId - 1, colDate, candPhase, assignments)

            If availableInstructors.Count > 0 Then
                CollectAssignments = CollectAssignments + 1
                HighlightPlanCell wsPlanOut, g, rowId, colDate, candId, True
                HighlightInstructorCandidates wsPlanOut, wsSrc, g, rowId, availableInstructors, colDate, candPhase, True
                RefreshPlanView wsPlanOut
                If PromptAssignmentUnified(wsSrc, g, rowId, colDate, candPhase, availableInstructors, chosenInstr, shiftCode, liveHours, candId) Then
                    Dim addH As Double
                    addH = ShiftHoursForDate(wsSrc, g, colDate)
                    IncrementLiveHours liveHours, candId, addH
                    Dim itm As Variant
                    Dim instrSrcRow As Long
                    Dim planCol As Long
                    Dim candPlanRow As Long
                    Dim instrPlanRow As Long
                    Dim candOriginalValue As Variant
                    Dim instrOriginalValue As Variant
                    instrSrcRow = FindInstructorRowForAssignment(wsSrc, g, rowId, colDate, candPhase, chosenInstr)
                    planCol = 2 + (colDate - CLng(g(giPlanColStart)) + 1)
                    candPlanRow = GetPlanRowFromSource(CStr(g(giGroupName)), rowId, wsPlanOut, candId)
                    instrPlanRow = GetPlanRowFromSource(CStr(g(giGroupName)), instrSrcRow, wsPlanOut, chosenInstr)
                    If candPlanRow > 0 Then candOriginalValue = wsPlanOut.Cells(candPlanRow, planCol).Value2
                    If instrPlanRow > 0 Then instrOriginalValue = wsPlanOut.Cells(instrPlanRow, planCol).Value2
                    itm = CreateAssignmentItem(CStr(g(giGroupName)), wsSrc.Cells(CLng(g(giDateRow)), colDate).Value2, planCol, colDate, candId, rowId, chosenInstr, shiftCode, instrSrcRow, CDbl(liveHours(UCase$(candId))), addH, hoursRow, candPhase, candOriginalValue, instrOriginalValue)
                    assignments.Add itm
                    ApplySingleAssignment wsPlanOut, itm
                    HighlightPlanCell wsPlanOut, g, rowId, colDate, candId, False
                    HighlightInstructorCandidates wsPlanOut, wsSrc, g, rowId, availableInstructors, colDate, candPhase, False
                    RefreshPlanView wsPlanOut
                    history.Add itm
                ElseIf UCase$(chosenInstr) = "__BACK__" Then
                    HighlightPlanCell wsPlanOut, g, rowId, colDate, candId, False
                    HighlightInstructorCandidates wsPlanOut, wsSrc, g, rowId, availableInstructors, colDate, candPhase, False
                    If UndoLastAssignment(wsPlanOut, history, assignments, liveHours, undone) Then
                        If CStr(undone(1)) = CStr(g(giGroupName)) Then
                            colDate = CLng(undone(3))
                            rowId = CLng(undone(6))
                            RefreshPlanView wsPlanOut
                            chosenInstr = ""
                            GoTo ContinueLoop
                        End If
                    End If
                    RefreshPlanView wsPlanOut
                    chosenInstr = ""
                    GoTo NextCandidate
                ElseIf UCase$(chosenInstr) = "__END__" Then
                    HighlightPlanCell wsPlanOut, g, rowId, colDate, candId, False
                    HighlightInstructorCandidates wsPlanOut, wsSrc, g, rowId, availableInstructors, colDate, candPhase, False
                    Exit Function
                Else
                    HighlightPlanCell wsPlanOut, g, rowId, colDate, candId, False
                    HighlightInstructorCandidates wsPlanOut, wsSrc, g, rowId, availableInstructors, colDate, candPhase, False
                End If
            Else
                noInstructorCandidates = noInstructorCandidates + 1
                Debug.Print "[OJT] Kandidat " & candId & " ima Xs, vendar ni prostih inštruktorjev za fazo " & candPhase & ". Skupina: " & CStr(g(giGroupName))
            End If
NextCandidate:
        rowId = rowId + 1
ContinueLoop:
    Loop
End Function



Private Function PromptAssignmentUnified(ByVal wsSrc As Worksheet, ByVal g As Variant, ByVal candRow As Long, ByVal colDate As Long, ByVal phase As Long, ByVal instrList As Collection, ByRef chosenInstr As String, ByRef shiftCode As String, ByRef liveHours As Object, ByVal candId As String) As Boolean
    Dim msg As String, inputText As Variant, parts() As String
    Dim i As Long, idx As Long
    Dim defaultShift As String

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

    defaultShift = GetCurrentShiftInputHint(wsSrc.Cells(candRow, colDate).Value2)
    inputText = Application.InputBox(msg, "OJT dodelitev", IIf(Len(defaultShift) > 0, "1;" & defaultShift, ""), Type:=2)
    If VarType(inputText) = vbBoolean Then
        If inputText = False Then Exit Function
    End If
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

Private Function GetCurrentShiftInputHint(ByVal rawValue As Variant) As String
    Dim v As String
    If IsError(rawValue) Then Exit Function
    If IsNull(rawValue) Then Exit Function
    v = Trim$(rawValue)
    If Len(v) = 0 Then Exit Function
    If Right$(v, 1) = "s" Or Right$(v, 1) = "i" Then
        v = Left$(v, Len(v) - 1)
    End If
    GetCurrentShiftInputHint = Trim$(v)
End Function

Private Sub HighlightPlanCell(ByVal wsPlan As Worksheet, ByVal g As Variant, ByVal srcRow As Long, ByVal srcCol As Long, ByVal fallbackId As String, ByVal active As Boolean)
    Dim rowPlan As Long
    Dim colPlan As Long
    rowPlan = GetPlanRowFromSource(CStr(g(giGroupName)), srcRow, wsPlan, fallbackId)
    colPlan = 2 + (srcCol - CLng(g(giPlanColStart)) + 1)
    If rowPlan <= 0 Or colPlan <= 0 Then Exit Sub
    If active Then
        wsPlan.Cells(rowPlan, colPlan).Font.Color = RGB(255, 0, 0)
    Else
        wsPlan.Cells(rowPlan, colPlan).Font.Color = RGB(0, 0, 0)
    End If
End Sub

Private Sub HighlightInstructorCandidates(ByVal wsPlan As Worksheet, ByVal wsSrc As Worksheet, ByVal g As Variant, ByVal candRow As Long, ByVal instrList As Collection, ByVal srcCol As Long, ByVal phase As Long, ByVal active As Boolean)
    Dim i As Long
    Dim instrId As String
    Dim srcRow As Long
    Dim planRow As Long
    Dim planCol As Long

    planCol = 2 + (srcCol - CLng(g(giPlanColStart)) + 1)
    For i = 1 To instrList.Count
        instrId = CStr(instrList(i))
        srcRow = FindInstructorRowForAssignment(wsSrc, g, candRow, srcCol, phase, instrId)
        If srcRow > 0 Then
            planRow = GetPlanRowFromSource(CStr(g(giGroupName)), srcRow, wsPlan, instrId)
            If planRow > 0 Then
                If active Then
                    wsPlan.Cells(planRow, planCol).Font.Color = RGB(255, 0, 0)
                Else
                    wsPlan.Cells(planRow, planCol).Font.Color = RGB(0, 0, 0)
                End If
            End If
        End If
    Next i
End Sub

Private Function FindInstructorRowForAssignment(ByVal wsSrc As Worksheet, ByVal g As Variant, ByVal candRow As Long, ByVal colDate As Long, ByVal phase As Long, ByVal instrId As String) As Long
    Dim r As Long
    Dim v As String
    Dim idCol As Long
    Dim normalizedId As String

    idCol = CLng(g(giIdCol))
    normalizedId = NormalizeLookupText(instrId)
    If Len(normalizedId) = 0 Then Exit Function

    If phase <> 2 Then
        r = candRow - 1
        If RowHasAvailableInstructor(wsSrc, r, idCol, colDate, normalizedId) Then
            FindInstructorRowForAssignment = r
            Exit Function
        End If

        r = candRow + 1
        If RowHasAvailableInstructor(wsSrc, r, idCol, colDate, normalizedId) Then
            FindInstructorRowForAssignment = r
            Exit Function
        End If
    End If

    For r = CLng(g(giIdRowStart)) To CLng(g(giIdRowEnd))
        If NormalizeLookupText(wsSrc.Cells(r, idCol).Value2) = normalizedId Then
            v = NormalizeScheduleCode(wsSrc.Cells(r, colDate).Value2)
            If IsInstructorAvailabilityCode(v) Then
                FindInstructorRowForAssignment = r
                Exit Function
            End If
        End If
    Next r

    FindInstructorRowForAssignment = FindSourceRowById(wsSrc, g, instrId)
End Function

Private Function RowHasAvailableInstructor(ByVal wsSrc As Worksheet, ByVal rowNumber As Long, ByVal idCol As Long, ByVal colDate As Long, ByVal normalizedInstrId As String) As Boolean
    If rowNumber <= 0 Then Exit Function
    If NormalizeLookupText(wsSrc.Cells(rowNumber, idCol).Value2) <> normalizedInstrId Then Exit Function
    RowHasAvailableInstructor = IsInstructorAvailabilityCode(NormalizeScheduleCode(wsSrc.Cells(rowNumber, colDate).Value2))
End Function

Private Function HasCandidateTrainingRows(ByVal wsSrc As Worksheet, ByVal g As Variant) As Boolean
    Dim r As Long

    For r = CLng(g(giIdRowStart)) To CLng(g(giIdRowEnd))
        If IsCandidateTrainingRow(wsSrc, g, r) Then
            HasCandidateTrainingRows = True
            Exit Function
        End If
    Next r
End Function

Private Function IsCandidateTrainingRow(ByVal wsSrc As Worksheet, ByVal g As Variant, ByVal rowNumber As Long) As Boolean
    Dim marker As String
    Dim roleCol As Long

    If rowNumber <= 0 Then Exit Function

    roleCol = CLng(g(giIdCol)) + 2
    marker = NormalizeLookupText(wsSrc.Cells(rowNumber, roleCol).Value2)
    If Len(marker) = 0 Then Exit Function

    IsCandidateTrainingRow = (Left$(marker, 8) = "KANDIDAT")
End Function

Private Sub ApplySingleAssignment(ByVal wsPlan As Worksheet, ByVal a As Variant)
    Dim rowCand As Long, rowInstr As Long, rowHours As Long
    rowCand = GetPlanRowFromSource(CStr(a(1)), CLng(a(6)), wsPlan, CStr(a(4)))
    rowInstr = GetPlanRowFromSource(CStr(a(1)), CLng(a(9)), wsPlan, CStr(a(7)))
    rowHours = GetPlanRowFromSource(CStr(a(1)), CLng(a(12)), wsPlan, CStr(a(4)))
    If rowCand > 0 Then
        wsPlan.Cells(rowCand, CLng(a(2))).Value2 = CStr(a(8)) & "s"
        AddOrReplaceComment wsPlan.Cells(rowCand, CLng(a(2))), BuildOjtComment(CStr(a(7)), GetPlanPersonName(wsPlan, rowInstr))
    End If
    If rowInstr > 0 Then
        wsPlan.Cells(rowInstr, CLng(a(2))).Value2 = CStr(a(8)) & "i"
        AddOrReplaceComment wsPlan.Cells(rowInstr, CLng(a(2))), BuildOjtComment(CStr(a(4)), GetPlanPersonName(wsPlan, rowCand))
    End If
    WriteCandidateHoursPanel wsPlan, CStr(a(1)), CStr(a(4)), CLng(a(2)), CDbl(a(10)), CLng(a(13))
End Sub


Private Function GetPlanRowFromSource(ByVal groupName As String, ByVal srcRow As Long, ByVal wsPlan As Worksheet, ByVal fallbackId As String) As Long
    Dim k As String
    Dim rowInGroup As Long

    If Not mPlanRowMap Is Nothing Then
        k = groupName & "|" & CStr(srcRow)
        If mPlanRowMap.Exists(k) Then
            GetPlanRowFromSource = CLng(mPlanRowMap(k))
            Exit Function
        End If
    End If

    rowInGroup = FindIdRowInGroup(wsPlan, groupName, fallbackId)
    If rowInGroup > 0 Then
        GetPlanRowFromSource = rowInGroup
        Exit Function
    End If

    GetPlanRowFromSource = FindIdRow(wsPlan, fallbackId)
End Function

Private Function FindIdRowInGroup(ByVal wsPlan As Worksheet, ByVal groupName As String, ByVal idValue As String) As Long
    Dim firstRow As Long
    Dim lastRow As Long
    Dim r As Long
    Dim targetId As String

    targetId = NormalizeLookupText(idValue)
    If Len(targetId) = 0 Then Exit Function

    firstRow = GetGroupFirstPlanRow(groupName)
    lastRow = GetGroupLastPlanRow(groupName)
    If firstRow <= 0 Or lastRow < firstRow Then Exit Function

    For r = firstRow To lastRow
        If NormalizeLookupText(wsPlan.Cells(r, 1).Value2) = targetId Then
            FindIdRowInGroup = r
            Exit Function
        End If
    Next r
End Function

Private Function UndoLastAssignment(ByVal wsPlan As Worksheet, ByRef history As Collection, ByRef assignments As Collection, ByRef liveHours As Object, ByRef undoneItem As Variant) As Boolean
    Dim a As Variant, rowCand As Long, rowInstr As Long
    Dim candOriginalValue As Variant
    Dim instrOriginalValue As Variant
    If history.Count = 0 Then Exit Function
    a = history(history.Count)
    history.Remove history.Count
    If assignments.Count > 0 Then assignments.Remove assignments.Count
    rowCand = GetPlanRowFromSource(CStr(a(1)), CLng(a(6)), wsPlan, CStr(a(4)))
    rowInstr = GetPlanRowFromSource(CStr(a(1)), CLng(a(9)), wsPlan, CStr(a(7)))

    If UBound(a) >= 14 Then candOriginalValue = a(14)
    If UBound(a) >= 15 Then instrOriginalValue = a(15)

    If rowCand > 0 Then
        wsPlan.Cells(rowCand, CLng(a(2))).Value2 = candOriginalValue
        ClearCellComments wsPlan.Cells(rowCand, CLng(a(2)))
    End If
    If rowInstr > 0 Then
        wsPlan.Cells(rowInstr, CLng(a(2))).Value2 = instrOriginalValue
        ClearCellComments wsPlan.Cells(rowInstr, CLng(a(2)))
    End If
    IncrementLiveHours liveHours, CStr(a(4)), -CDbl(a(11))
    ClearCandidateHoursPanelEntry wsPlan, CStr(a(1)), CStr(a(4)), CLng(a(2)), CLng(a(13))
    undoneItem = a
    UndoLastAssignment = True
End Function

Private Sub ClearCandidateHoursPanelEntry(ByVal wsPlan As Worksheet, ByVal groupName As String, ByVal candId As String, ByVal planCol As Long, ByVal phase As Long)
    Dim lastGroupRow As Long
    Dim baseRow As Long
    Dim rowHours As Long
    Dim candidateIdx As Long

    lastGroupRow = GetGroupLastPlanRow(groupName)
    If lastGroupRow <= 0 Then Exit Sub

    candidateIdx = EnsureCandidatePanelIndex(groupName, candId)
    baseRow = lastGroupRow + 2 + (candidateIdx - 1) * CANDIDATE_PANEL_ROWS
    rowHours = baseRow + phase

    If rowHours > 0 And planCol > 0 Then wsPlan.Cells(rowHours, planCol).ClearContents
    If Not CandidatePanelHasHours(wsPlan, baseRow) Then
        wsPlan.Cells(baseRow, 1).Resize(CANDIDATE_PANEL_ROWS, 2).ClearContents
    End If
End Sub

Private Function CandidatePanelHasHours(ByVal wsPlan As Worksheet, ByVal baseRow As Long) As Boolean
    Dim lastCol As Long
    Dim r As Long
    Dim c As Long
    Dim v As Variant

    If baseRow <= 0 Then Exit Function

    lastCol = wsPlan.UsedRange.Column + wsPlan.UsedRange.Columns.Count - 1
    If lastCol < 3 Then Exit Function

    For r = baseRow + 1 To baseRow + 3
        For c = 3 To lastCol
            v = wsPlan.Cells(r, c).Value2
            If Not IsError(v) Then
                If Len(Trim$(CStr(v))) > 0 Then
                    CandidatePanelHasHours = True
                    Exit Function
                End If
            End If
        Next c
    Next r
End Function

Private Sub WriteCandidateHoursPanel(ByVal wsPlan As Worksheet, ByVal groupName As String, ByVal candId As String, ByVal planCol As Long, ByVal hoursVal As Double, ByVal phase As Long)
    Dim lastGroupRow As Long
    Dim rowName As Long, rowHours As Long, baseRow As Long, candidateIdx As Long
    Dim rowCand As Long
    Dim candLabel As String

    lastGroupRow = GetGroupLastPlanRow(groupName)
    If lastGroupRow <= 0 Then Exit Sub

    candidateIdx = EnsureCandidatePanelIndex(groupName, candId)
    baseRow = lastGroupRow + 2 + (candidateIdx - 1) * CANDIDATE_PANEL_ROWS
    rowName = baseRow
    rowHours = baseRow + phase
    rowCand = FindIdRow(wsPlan, candId)
    candLabel = candId
    If rowCand > 0 Then
        If Len(Trim$(CStr(wsPlan.Cells(rowCand, 2).Value2))) > 0 Then
            candLabel = candId & " - " & CStr(wsPlan.Cells(rowCand, 2).Value2)
        End If
    End If

    wsPlan.Cells(rowName, 1).Value2 = "Kandidat"
    wsPlan.Cells(rowName, 2).Value2 = candLabel
    wsPlan.Cells(rowHours, 1).Value2 = "URE F" & CStr(phase)
    wsPlan.Cells(rowHours, planCol).Value2 = hoursVal
End Sub

Private Function EnsureCandidatePanelIndex(ByVal groupName As String, ByVal candId As String) As Long
    Dim k As String
    k = groupName & "|" & UCase$(candId)
    If mCandidatePanelIndex Is Nothing Then Set mCandidatePanelIndex = CreateObject("Scripting.Dictionary")
    If Not mCandidatePanelIndex.Exists(k) Then mCandidatePanelIndex(k) = mCandidatePanelIndex.Count + 1
    EnsureCandidatePanelIndex = CLng(mCandidatePanelIndex(k))
End Function

Private Function GetGroupLastPlanRow(ByVal groupName As String) As Long
    Dim k As Variant
    Dim p As Long
    Dim grp As String

    If mPlanRowMap Is Nothing Then Exit Function
    For Each k In mPlanRowMap.Keys
        p = InStr(1, CStr(k), "|", vbTextCompare)
        If p > 0 Then
            grp = Left$(CStr(k), p - 1)
            If StrComp(grp, groupName, vbTextCompare) = 0 Then
                If CLng(mPlanRowMap(k)) > GetGroupLastPlanRow Then
                    GetGroupLastPlanRow = CLng(mPlanRowMap(k))
                End If
            End If
        End If
    Next k
End Function

Private Function GetGroupFirstPlanRow(ByVal groupName As String) As Long
    Dim k As Variant
    Dim p As Long
    Dim grp As String
    Dim mappedRow As Long

    If mPlanRowMap Is Nothing Then Exit Function
    For Each k In mPlanRowMap.Keys
        p = InStr(1, CStr(k), "|", vbTextCompare)
        If p > 0 Then
            grp = Left$(CStr(k), p - 1)
            If StrComp(grp, groupName, vbTextCompare) = 0 Then
                mappedRow = CLng(mPlanRowMap(k))
                If GetGroupFirstPlanRow = 0 Or mappedRow < GetGroupFirstPlanRow Then
                    GetGroupFirstPlanRow = mappedRow
                End If
            End If
        End If
    Next k
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

Private Function GetCachedHoursRow(ByVal wsSrc As Worksheet, ByVal g As Variant, ByVal candId As String, ByVal candName As String, ByRef hoursRowCache As Object) As Long
    Dim key As String
    key = NormalizeLookupText(candId) & "|" & NormalizeLookupText(candName)

    If hoursRowCache Is Nothing Then Set hoursRowCache = CreateObject("Scripting.Dictionary")
    If Not hoursRowCache.Exists(key) Then
        hoursRowCache.Add key, FindHoursRowByCandidate(wsSrc, g, candId, candName)
    End If

    GetCachedHoursRow = CLng(hoursRowCache(key))
End Function

Private Function FindHoursRowByCandidate(ByVal wsSrc As Worksheet, ByVal g As Variant, ByVal candId As String, ByVal candName As String) As Long
    FindHoursRowByCandidate = FindHoursRowByCandidateInRange(wsSrc, g, candId, candName, CLng(g(giHoursRowStart)), CLng(g(giHoursRowEnd)), True)
    If FindHoursRowByCandidate > 0 Then Exit Function

    ' If settings point at the schedule rows, do not accept an ordinary candidate
    ' row as a cumulative-hours row. Only explicit "URE <candidate>" labels are
    ' safe outside the configured hours range.
    FindHoursRowByCandidate = FindHoursRowByCandidateInRange(wsSrc, g, candId, candName, CLng(g(giCandIdRowStart)), CLng(g(giCandIdRowEnd)), False)
    If FindHoursRowByCandidate > 0 Then Exit Function

    FindHoursRowByCandidate = FindHoursRowByCandidateInUsedRange(wsSrc, candId, candName)
End Function

Private Function FindHoursRowByCandidateInRange(ByVal wsSrc As Worksheet, ByVal g As Variant, ByVal candId As String, ByVal candName As String, ByVal rowStart As Long, ByVal rowEnd As Long, ByVal allowExactIdentity As Boolean) As Long
    Dim r As Long
    Dim idCol As Long
    Dim nameCol As Long
    Dim idVal As Variant
    Dim nameVal As Variant

    If rowStart <= 0 Or rowEnd < rowStart Then Exit Function

    idCol = CLng(g(giIdCol))
    nameCol = idCol + 1

    For r = rowStart To rowEnd
        idVal = wsSrc.Cells(r, idCol).Value2
        nameVal = wsSrc.Cells(r, nameCol).Value2

        If IsExplicitCandidateHoursLabel(idVal, candId, candName) Or _
           IsExplicitCandidateHoursLabel(nameVal, candId, candName) Then
            FindHoursRowByCandidateInRange = r
            Exit Function
        End If

        If allowExactIdentity Then
            If TextMatchesCandidate(idVal, candId, candName) Or _
               TextMatchesCandidate(nameVal, candId, candName) Then
                If RowLooksLikeCumulativeHours(wsSrc, g, r) Then
                    FindHoursRowByCandidateInRange = r
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

Private Function RowLooksLikeCumulativeHours(ByVal wsSrc As Worksheet, ByVal g As Variant, ByVal rowNumber As Long) As Boolean
    Dim c As Long
    Dim v As Variant
    Dim numericCells As Long
    Dim textCells As Long
    Dim colStart As Long
    Dim colEnd As Long

    colStart = CLng(g(giPlanColStart))
    colEnd = CLng(g(giPlanColEnd))
    If colStart <= 0 Or colEnd < colStart Then Exit Function

    For c = colStart To colEnd
        v = wsSrc.Cells(rowNumber, c).Value2
        If IsError(v) Then
            textCells = textCells + 1
            Exit For
        ElseIf Len(Trim$(CStr(v))) > 0 Then
            If IsNumeric(v) Then
                numericCells = numericCells + 1
            Else
                textCells = textCells + 1
                If textCells > 0 Then Exit For
            End If
        End If
    Next c

    RowLooksLikeCumulativeHours = (numericCells > 0 And textCells = 0)
End Function

Private Function FindHoursRowByCandidateInUsedRange(ByVal wsSrc As Worksheet, ByVal candId As String, ByVal candName As String) As Long
    Dim cell As Range

    On Error Resume Next
    For Each cell In wsSrc.UsedRange.Cells
        If IsExplicitCandidateHoursLabel(cell.Value2, candId, candName) Then
            FindHoursRowByCandidateInUsedRange = cell.Row
            Exit Function
        End If
    Next cell
    On Error GoTo 0
End Function

Private Function TextMatchesCandidate(ByVal rawValue As Variant, ByVal candId As String, ByVal candName As String) As Boolean
    Dim textValue As String
    textValue = NormalizeLookupText(rawValue)
    If Len(textValue) = 0 Then Exit Function

    If Len(NormalizeLookupText(candId)) > 0 Then
        If textValue = NormalizeLookupText(candId) Then
            TextMatchesCandidate = True
            Exit Function
        End If
    End If

    If Len(NormalizeLookupText(candName)) > 0 Then
        TextMatchesCandidate = (textValue = NormalizeLookupText(candName))
    End If
End Function

Private Function IsExplicitCandidateHoursLabel(ByVal rawValue As Variant, ByVal candId As String, ByVal candName As String) As Boolean
    Dim label As String
    Dim candidatePart As String
    Dim separator As String

    label = NormalizeLookupText(rawValue)
    If Len(label) < 4 Then Exit Function
    If Left$(label, 3) <> "URE" Then Exit Function
    separator = Mid$(label, 4, 1)
    If separator <> " " And separator <> ":" And separator <> "-" And separator <> "." Then Exit Function

    candidatePart = Trim$(Mid$(label, 4))
    Do While Left$(candidatePart, 1) = ":" Or Left$(candidatePart, 1) = "-" Or Left$(candidatePart, 1) = "."
        candidatePart = Trim$(Mid$(candidatePart, 2))
    Loop

    If Len(candidatePart) = 0 Then Exit Function
    IsExplicitCandidateHoursLabel = TextMatchesCandidate(candidatePart, candId, candName)
End Function

Private Function NormalizeLookupText(ByVal rawValue As Variant) As String
    Dim s As String
    If IsError(rawValue) Then Exit Function
    If IsNull(rawValue) Then Exit Function
    s = CStr(rawValue)
    s = Replace(s, Chr$(160), " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Trim$(s)
    Do While InStr(1, s, "  ", vbBinaryCompare) > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeLookupText = UCase$(s)
End Function

Private Sub EnsureLiveHours(ByRef liveHours As Object, ByVal candId As String, ByVal defaultHours As Double)
    Dim key As String
    key = UCase$(candId)
    If Not liveHours.Exists(key) Then liveHours.Add key, defaultHours
End Sub

Private Function NormalizeScheduleCode(ByVal rawValue As Variant) As String
    If IsError(rawValue) Then Exit Function
    If IsNull(rawValue) Then Exit Function
    NormalizeScheduleCode = UCase$(Replace(Trim$(Replace(CStr(rawValue), Chr$(160), " ")), " ", ""))
End Function

Private Function ResolvePhaseLive(ByVal wsSrc As Worksheet, ByVal g As Variant, ByVal candRow As Long, ByVal colDate As Long, ByVal thresholds As Object, ByRef liveHours As Object, ByVal candId As String) As Long
    Dim key As String
    Dim baseHours As Double
    Dim reserveHours As Double

    key = UCase$(candId)
    If Not liveHours.Exists(key) Then
        baseHours = GetCumulativeHoursAtDate(wsSrc, g, candRow, colDate)
        liveHours.Add key, baseHours
    End If

    reserveHours = ShiftHoursForDate(wsSrc, g, colDate)
    ResolvePhaseLive = ResolvePhaseFromHours(CDbl(liveHours(key)), thresholds, GetTrackType(CStr(g(giGroupName))), reserveHours)
End Function

Private Function GetCumulativeHoursAtDate(ByVal wsSrc As Worksheet, ByVal g As Variant, ByVal hoursRow As Long, ByVal colDate As Long) As Double
    Dim searchStartCol As Long
    Dim exactHours As Double
    Dim exactIsNumeric As Boolean
    Dim previousHours As Double
    Dim hasPreviousHours As Boolean
    Dim c As Long
    Dim v As Variant

    If hoursRow <= 0 Then Exit Function

    v = wsSrc.Cells(hoursRow, colDate).Value2
    If Not IsError(v) Then exactIsNumeric = (Len(Trim$(CStr(v))) > 0 And IsNumeric(v))
    If exactIsNumeric Then
        exactHours = Round(CDbl(v), 0)
        If exactHours <> 0# Then
            GetCumulativeHoursAtDate = exactHours
            Exit Function
        End If
    End If

    ' If the selected date cell is blank or shows a relative 0 at the planning
    ' boundary, use the last known cumulative value before that date. The
    ' displayed hours must remain cumulative across months, while the planning
    ' loop may still count only newly added hours from here on.
    searchStartCol = 1
    For c = colDate - 1 To searchStartCol Step -1
        v = wsSrc.Cells(hoursRow, c).Value2
        If Not IsError(v) Then
            If Len(Trim$(CStr(v))) > 0 And IsNumeric(v) Then
                previousHours = Round(CDbl(v), 0)
                hasPreviousHours = True
                Exit For
            End If
        End If
    Next c

    If exactIsNumeric Then
        If exactHours = 0# And hasPreviousHours And previousHours > 0# Then
            GetCumulativeHoursAtDate = previousHours
        Else
            GetCumulativeHoursAtDate = exactHours
        End If
    ElseIf hasPreviousHours Then
        GetCumulativeHoursAtDate = previousHours
    End If
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
        If Not IsError(v) Then
            If Len(Trim$(CStr(v))) > 0 And IsNumeric(v) Then
                ShiftHoursForDate = CDbl(v)
                Exit Function
            End If
        End If
    End If
    ShiftHoursForDate = 8#
End Function

Private Function ResolvePhaseFromHours(ByVal totalHours As Double, ByVal thresholds As Object, ByVal trackType As String, Optional ByVal reserveHours As Double = 0#) As Long
    Dim t As Variant
    Dim app As Variant
    Dim baseBeforeTrack As Double
    Dim trackHours As Double
    Dim appTotal As Double
    Dim phase1Limit As Double
    Dim phase2Limit As Double
    
    If trackType = "ACS" And thresholds.Exists("APP_ACS") Then
        app = thresholds("APP_ACS")
    ElseIf thresholds.Exists("APP") Then
        app = thresholds("APP")
    End If
    appTotal = CDbl(app(1)) + CDbl(app(2)) + CDbl(app(3))

    ' Ure v sledilniku so kumulativne.
    ' Kandidat znotraj APS/ACS skupine je lahko še vedno v APP bloku.
    ' Če skupni total še ni presegel APP bloka, fazo določamo po APP pravilih.
    ' Ko APP blok preseže, za APS/ACS odštejemo APP in fazo določamo po track urah.
    If trackType = "APS" Or trackType = "ACS" Then
        If totalHours < appTotal Then
            If trackType = "ACS" And thresholds.Exists("APP_ACS") Then
                t = thresholds("APP_ACS")
            Else
                t = thresholds("APP")
            End If
            baseBeforeTrack = 0#
        Else
            t = thresholds(trackType)
            baseBeforeTrack = appTotal
        End If
    Else
        t = thresholds(trackType)
        baseBeforeTrack = 0#
    End If
    trackHours = totalHours - baseBeforeTrack
    If trackHours < 0# Then trackHours = 0#

    phase1Limit = CDbl(t(1))
    phase2Limit = phase1Limit + CDbl(t(2))

    ' Konzervativno planiranje:
    ' - prehod 1 -> 2 zamaknemo za eno izmeno (reserveHours),
    '   da kandidat ostane še en termin pri primarnih inštruktorjih.
    ' - prehod 2 -> 3 ostane konservativen kot prej.
    If trackHours < (phase1Limit + reserveHours) Then
        ResolvePhaseFromHours = 1
    ElseIf trackHours < (phase2Limit - reserveHours) Then
        ResolvePhaseFromHours = 2
    Else
        ResolvePhaseFromHours = 3
    End If
End Function

Private Function GetAvailableInstructors(ByVal wsSrc As Worksheet, ByVal g As Variant, ByVal tripleStartRow As Long, ByVal colDate As Long, ByVal phase As Long, ByVal assignments As Collection) As Collection
    Dim c As New Collection
    Dim seen As Object
    Dim r As Long
    Dim v As String
    Dim planDate As Variant
    Set seen = CreateObject("Scripting.Dictionary")
    planDate = wsSrc.Cells(CLng(g(giDateRow)), colDate).Value2

    If phase = 2 Then
        For r = CLng(g(giIdRowStart)) To CLng(g(giIdRowEnd))
            v = NormalizeScheduleCode(wsSrc.Cells(r, colDate).Value2)
            If IsInstructorAvailabilityCode(v) Then
                AddAvailableInstructor c, seen, assignments, planDate, wsSrc.Cells(r, CLng(g(giIdCol))).Value2
            End If
        Next r
    Else
        v = NormalizeScheduleCode(wsSrc.Cells(tripleStartRow, colDate).Value2)
        If IsInstructorAvailabilityCode(v) Then
            AddAvailableInstructor c, seen, assignments, planDate, wsSrc.Cells(tripleStartRow, CLng(g(giIdCol))).Value2
        End If

        v = NormalizeScheduleCode(wsSrc.Cells(tripleStartRow + 2, colDate).Value2)
        If IsInstructorAvailabilityCode(v) Then
            AddAvailableInstructor c, seen, assignments, planDate, wsSrc.Cells(tripleStartRow + 2, CLng(g(giIdCol))).Value2
        End If
    End If

    Set GetAvailableInstructors = c
End Function

Private Sub AddAvailableInstructor(ByRef instrList As Collection, ByRef seen As Object, ByVal assignments As Collection, ByVal planDate As Variant, ByVal rawInstrId As Variant)
    Dim instrId As String
    Dim key As String

    If IsError(rawInstrId) Then Exit Sub
    If IsNull(rawInstrId) Then Exit Sub
    instrId = Trim$(CStr(rawInstrId))
    key = NormalizeLookupText(instrId)
    If Len(key) = 0 Then Exit Sub
    If InstructorAlreadyAssignedOnDate(assignments, instrId, planDate) Then Exit Sub

    If Not seen.Exists(key) Then
        seen.Add key, True
        instrList.Add instrId
    End If
End Sub

Private Function InstructorAlreadyAssignedOnDate(ByVal assignments As Collection, ByVal instrId As String, ByVal planDate As Variant) As Boolean
    Dim i As Long
    Dim a As Variant
    Dim assignedInstr As String

    If assignments Is Nothing Then Exit Function

    For i = 1 To assignments.Count
        a = assignments(i)
        assignedInstr = CStr(a(7))
        If NormalizeLookupText(assignedInstr) = NormalizeLookupText(instrId) Then
            If SamePlanningDate(a(5), planDate) Then
                InstructorAlreadyAssignedOnDate = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function SamePlanningDate(ByVal leftDate As Variant, ByVal rightDate As Variant) As Boolean
    On Error GoTo SafeExit

    If IsError(leftDate) Or IsError(rightDate) Then Exit Function

    If IsNumeric(leftDate) And IsNumeric(rightDate) Then
        SamePlanningDate = (CLng(CDbl(leftDate)) = CLng(CDbl(rightDate)))
    ElseIf IsDate(leftDate) And IsDate(rightDate) Then
        SamePlanningDate = (DateValue(CDate(leftDate)) = DateValue(CDate(rightDate)))
    Else
        SamePlanningDate = (NormalizeLookupText(leftDate) = NormalizeLookupText(rightDate))
    End If

SafeExit:
End Function

Private Function IsInstructorAvailabilityCode(ByVal normalizedCode As String) As Boolean
    IsInstructorAvailabilityCode = (normalizedCode = "X1" Or normalizedCode = "X2" Or normalizedCode = "X3")
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
    If VarType(pick) = vbBoolean Then
        If pick = False Then Exit Function
    End If
    If CLng(pick) = 0 Then Exit Function

    If CLng(pick) < 1 Or CLng(pick) > instrList.Count Then
        MsgBox "Napačna izbira.", vbExclamation
        Exit Function
    End If

    chosenInstr = instrList(CLng(pick))
    pick = Application.InputBox("Vpiši izmeno (npr A9):", "Izmena", Type:=2)
    If VarType(pick) = vbBoolean Then
        If pick = False Then Exit Function
    End If
    shiftCode = UCase$(Trim$(CStr(pick)))
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
            AddOrReplaceComment wsPlan.Cells(rowCand, CLng(a(2))), BuildOjtComment(CStr(a(7)), GetPlanPersonName(wsPlan, rowInstr))
        End If

        If rowInstr > 0 Then
            wsPlan.Cells(rowInstr, CLng(a(2))).Value2 = CStr(a(8)) & "i"
            AddOrReplaceComment wsPlan.Cells(rowInstr, CLng(a(2))), BuildOjtComment(CStr(a(4)), GetPlanPersonName(wsPlan, rowCand))
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
    ClearCellComments cell
    cell.AddCommentThreaded text
End Sub

Private Function BuildOjtComment(ByVal relatedId As String, ByVal relatedFullName As Variant) As String
    BuildOjtComment = "OJT|" & NormalizeCommentText(relatedId) & "|" & FormatSurnameName(relatedFullName)
End Function

Private Function GetPlanPersonName(ByVal wsPlan As Worksheet, ByVal planRow As Long) As Variant
    If planRow <= 0 Then
        GetPlanPersonName = vbNullString
        Exit Function
    End If

    GetPlanPersonName = wsPlan.Cells(planRow, 2).Value2
End Function

Private Function FormatSurnameName(ByVal rawName As Variant) As String
    Dim nameText As String
    Dim parts() As String
    Dim i As Long
    Dim formatted As String

    nameText = NormalizeCommentText(rawName)
    If Len(nameText) = 0 Then Exit Function

    parts = Split(nameText, " ")
    If UBound(parts) = 0 Then
        FormatSurnameName = parts(0)
        Exit Function
    End If

    For i = 1 To UBound(parts)
        If Len(formatted) > 0 Then formatted = formatted & "_"
        formatted = formatted & parts(i)
    Next i
    FormatSurnameName = formatted & "_" & parts(0)
End Function

Private Function NormalizeCommentText(ByVal rawValue As Variant) As String
    Dim s As String

    If IsError(rawValue) Then Exit Function
    If IsNull(rawValue) Then Exit Function

    s = CStr(rawValue)
    s = Replace(s, Chr$(160), " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Trim$(s)

    Do While InStr(1, s, "  ", vbBinaryCompare) > 0
        s = Replace(s, "  ", " ")
    Loop

    NormalizeCommentText = s
End Function

Private Sub ClearCellComments(ByVal cell As Range)
    On Error Resume Next
    cell.CommentThreaded.Delete
    cell.ClearComments
    On Error GoTo 0
End Sub

Private Function CreateAssignmentItem(ByVal groupName As String, ByVal planDate As Variant, ByVal colDate As Long, ByVal srcColDate As Long, ByVal candId As String, ByVal candRow As Long, ByVal instrId As String, ByVal shiftCode As String, ByVal tripleStart As Long, ByVal liveHoursAfter As Double, ByVal hoursAdded As Double, ByVal hoursRow As Long, ByVal candPhase As Long, ByVal candOriginalValue As Variant, ByVal instrOriginalValue As Variant) As Variant
    Dim a(1 To 15) As Variant
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
    a(12) = hoursRow
    a(13) = candPhase
    a(14) = candOriginalValue
    a(15) = instrOriginalValue
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
    Dim actualEndRow As Long
    Dim reservedEndRow As Long
    Dim scheduleRowsReserved As Long

    idCol = CLng(g(giIdCol))
    nameCol = idCol + 1
    planStartCol = CLng(g(giPlanColStart))
    planEndCol = CLng(g(giPlanColEnd))
    dateRow = CLng(g(giDateRow))
    dayRow = CLng(g(giDayRow))
    idRowStart = CLng(g(giIdRowStart))
    idRowEnd = CLng(g(giIdRowEnd))
    planCols = planEndCol - planStartCol + 1

    If mPlanRowMap Is Nothing Then Set mPlanRowMap = CreateObject("Scripting.Dictionary")

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
        For r = hrS To hrE
            mPlanRowMap(CStr(g(giGroupName)) & "|" & CStr(r)) = outRow + rowCount + 2 + (r - hrS)
        Next r
        wsPlan.Cells(outRow + rowCount + 2, 1).Resize(hrRows, planCols + 2).Font.Bold = True
        wsPlan.Cells(outRow + rowCount + 2, 1).Resize(hrRows, planCols + 2).Font.Color = RGB(255, 0, 0)
        actualEndRow = outRow + rowCount + hrRows + GROUP_RESERVED_BLANK_ROWS + 1
    Else
        actualEndRow = outRow + rowCount
    End If

    scheduleRowsReserved = (idRowEnd - idRowStart + 1) + 3 ' naslov skupine + 2 glavi + vrstice urnika
    reservedEndRow = outRow + scheduleRowsReserved + _
                     GROUP_RESERVED_BLANK_ROWS + _
                     GROUP_RESERVED_HOURS_ROWS + _
                     GROUP_RESERVED_BLANK_ROWS - 1

    CopyGroupToPlan = WorksheetFunction.Max(actualEndRow, reservedEndRow) + 1
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

Private Function LoadGroups(ByVal wsSettings As Worksheet, Optional ByVal includeDisabled As Boolean = False) As Collection
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
        planFlag = NormalizePlanFlag(wsSettings.Cells(18, c).Value2)
        If Not hasPlanToggle And Len(planFlag) = 0 Then planFlag = "DA"
        g(giPlanEnabled) = planFlag
        g(giHoursRowStart) = CLng(Val(wsSettings.Cells(19, c).Value2))
        g(giHoursRowEnd) = CLng(Val(wsSettings.Cells(20, c).Value2))

        If Not includeDisabled Then
            If hasPlanToggle Then
                If planFlag <> "DA" Then GoTo NextCol
            End If
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


Private Function IsGroupPlanningEnabled(ByVal g As Variant) As Boolean
    Dim planFlag As String
    planFlag = UCase$(Trim$(Replace(CStr(g(giPlanEnabled)), Chr$(160), " ")))

    IsGroupPlanningEnabled = (planFlag = "DA")
End Function

Private Function NormalizePlanFlag(ByVal rawValue As Variant) As String
    NormalizePlanFlag = UCase$(Trim$(Replace(CStr(rawValue), Chr$(160), " ")))
End Function

Private Function GetPlanningStartColumn(ByVal g As Variant) As Long
    Dim configuredStart As Long
    Dim planStart As Long
    Dim planEnd As Long

    configuredStart = CLng(g(giPlanStartCol))
    planStart = CLng(g(giPlanColStart))
    planEnd = CLng(g(giPlanColEnd))

    If configuredStart < planStart Then configuredStart = planStart
    If configuredStart > planEnd Then configuredStart = planStart

    GetPlanningStartColumn = configuredStart
End Function

Private Function HasAnyPlanToggle(ByVal ws As Worksheet, ByVal lastCol As Long) As Boolean
    Dim c As Long
    Dim v As String
    For c = SETTINGS_FIRST_GROUP_COL To lastCol
        v = NormalizePlanFlag(ws.Cells(18, c).Value2)
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
    Dim appAcs(1 To 3) As Double

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

    appAcs(1) = CDbl(Val(wsSettings.Cells(11, 11).Value2))
    appAcs(2) = CDbl(Val(wsSettings.Cells(12, 11).Value2))
    appAcs(3) = CDbl(Val(wsSettings.Cells(13, 11).Value2))

    d.Add "APP", app
    d.Add "APS", aps
    d.Add "ACS", acs
    d.Add "APP_ACS", appAcs

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
