Attribute VB_Name = "OrderPipeline_MultiChannel"
' ============================================================
' Weekly radiopharmaceutical order processing pipeline (Portfolio / Anonymized Version)
' Real hospital/client codes, the distributor name, and colleague names have
' been replaced with generic placeholders (HOSP01-24, PARTNER_CO, GroupA/B/C).
' Business logic and structure are preserved as-is.
' ============================================================

Option Explicit

' ============================================================
' CapsuleData_FirstRound_Generate (v6)
'  - v5에서 "저용량 Group B/Group A 통합" 했던 걸 원복
'  - 저용량(2/5mCi)도 다시 Group A(LDHCM)/Group B(LDHN1)로 구분해서 각자 칸에 채움
'  - 나머지(요일 자동생성, 병합 앵커행 SUMIFS, 칸 부족시 행삽입)는 v4/v5 그대로 유지
' ============================================================

Sub CapsuleData_FirstRound_Generate()

    Dim ws1 As Worksheet, wsOrder As Worksheet, wsOrder2 As Worksheet
    Set ws1 = Sheets("Sheet1")
    Set wsOrder = Sheets("주문데이터")
    Set wsOrder2 = Sheets("2차주문데이터")

    Const groupAList As String = "HOSP01,HOSP02,HOSP03,HOSP04,HOSP05,HOSP06"
    Const groupBList As String = "HOSP07,HOSP08,HOSP09,HOSP10,HOSP11,HOSP12,HOSP13,HOSP14,HOSP15,HOSP16,HOSP17,HOSP18,HOSP19,HOSP20,HOSP21,HOSP22,HOSP23,HOSP24"

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim inputWk As String
    inputWk = InputBox("Week 번호를 입력하세요 (예: 27)", "Week 입력", "27")
    If inputWk = "" Then
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
End If

    ws1.Cells(1, 4).Value = Val(inputWk)

    Dim buckets As Object: Set buckets = CreateObject("Scripting.Dictionary")
    Dim warnList As Collection: Set warnList = New Collection

    ' ── 주문데이터 → (그룹|요일) 버킷 수집 (순서 유지) ──
    Dim lastRow As Long, i As Long
    lastRow = wsOrder.Cells(wsOrder.Rows.Count, 1).End(xlUp).Row

    Dim cID As String, cDay As String, cMci As Double, cQty As Long, isGroupA As Boolean
    For i = 2 To lastRow
        cID = Trim(CStr(wsOrder.Cells(i, 1).Value))
        cDay = Trim(CStr(wsOrder.Cells(i, 2).Value))
        cMci = Val(wsOrder.Cells(i, 3).Value)
        cQty = Val(wsOrder.Cells(i, 4).Value)
        If cID = "" Or cQty = 0 Then GoTo NextRow

        isGroupA = (InStr("," & groupAList & ",", "," & cID & ",") > 0)

        Select Case cMci
            Case 100, 50, 30, 25, 200
                Call AddToBucket(buckets, IIf(isGroupA, "GroupA", "GroupB"), cDay, cID)
            Case 5, 2
                ' 저용량도 다시 Group A/B 구분해서 각자 칸으로
                Call AddToBucket(buckets, IIf(isGroupA, "GroupA_Low", "GroupB_Low"), cDay, cID)
            Case 10
                warnList.Add "10mCi 저용량 (템플릿에 mCi단 없음): " & cDay & "요일 / " & cID
            Case Else
                warnList.Add "미분류 mCi(" & cMci & "): " & cDay & "요일 / " & cID
        End Select
NextRow:
    Next i

       Dim lastRow2 As Long, c2ID As String, c2Day As String
    lastRow2 = wsOrder2.Cells(wsOrder2.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow2
        c2ID = Trim(CStr(wsOrder2.Cells(i, 1).Value))
        c2Day = Trim(CStr(wsOrder2.Cells(i, 2).Value))
        If c2ID = "" Or Val(wsOrder2.Cells(i, 4).Value) = 0 Then GoTo NextRow2
        If InStr("," & groupBList & ",", "," & c2ID & ",") = 0 Then
            warnList.Add "2차주문데이터: groupBList 목록에 없는 병원(확인 필요) - " & c2Day & "요일 / " & c2ID
            GoTo NextRow2
        End If
        Call AddToBucket(buckets, "GroupC", c2Day, c2ID)
NextRow2:
    Next i

    Dim offset As Long: offset = 0

    Call ProcessGroup(ws1, "GroupA", 4, Array(Array("일", 4), Array("토", 5), Array("금", 6), Array("목", 8), Array("수", 9), Array("화", 13), Array("월", 17)), "주문데이터", buckets, offset)

    Call ProcessGroup(ws1, "GroupB", 4, Array(Array("일", 21), Array("토", 25), Array("금", 29), Array("목", 33), Array("수", 37), Array("화", 41), Array("월", 45)), "주문데이터", buckets, offset)

    Call ProcessGroup(ws1, "GroupB_Low", 2, Array(Array("금", 59), Array("목", 62), Array("수", 63), Array("화", 67), Array("월", 72)), "주문데이터", buckets, offset)

    Call ProcessGroup(ws1, "GroupA_Low", 2, Array(Array("금", 86), Array("목", 87), Array("수", 88), Array("화", 89), Array("월", 90)), "주문데이터", buckets, offset)

    ' ── Solution section (fixed rows 96-100, PARTNER_CO) ──
    Dim solStart As Long: solStart = 96 + offset
    Dim rr As Long
    For rr = solStart To solStart + 4
        ws1.Cells(rr, 5).Formula = "=SUMIFS(주문데이터!$D:$D,주문데이터!$A:$A,$C" & rr & ",주문데이터!$B:$B,$A" & rr & ",주문데이터!$C:$C,$D" & rr & ")"
        ws1.Cells(rr, 7).Formula = "=SUMIFS(주문데이터!$D:$D,주문데이터!$A:$A,$C" & rr & ",주문데이터!$B:$B,$A" & rr & ",주문데이터!$C:$C,$F" & rr & ")"
    Next rr

    ' ── HN2(Capsule-2): 화요일은 원래 빈 5칸을 최초 1회 초기화 ──
    Dim hwaRow As Long: hwaRow = 119 + offset
    Call EnsureHwaBlockInit(ws1, hwaRow, 5)

    Call ProcessGroup(ws1, "GroupC", 4, Array(Array("토", 105), Array("금", 109), Array("목", 113), Array("수", 115), Array("화", 119)), "2차주문데이터", buckets, offset)

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    If warnList.Count > 0 Then
        Dim msg As String, w As Variant
        For Each w In warnList
            msg = msg & w & vbCrLf
        Next w
        MsgBox "캡슐데이터 생성 완료. 아래 항목은 확인이 필요합니다:" & vbCrLf & vbCrLf & msg, vbExclamation
    Else
        MsgBox "캡슐데이터 자동생성 완료! (저용량 Group A/B 다시 구분)", vbInformation
    End If
End Sub

' ------------------------------------------------------------
Sub ProcessGroup(ws As Worksheet, grp As String, tierCount As Integer, predefinedDays As Variant, sheetSrc As String, buckets As Object, ByRef offset As Long)
    Dim d As Variant
    Dim lastRowOfGroup As Long
    Dim knownDays As Object: Set knownDays = CreateObject("Scripting.Dictionary")

    For Each d In predefinedDays
        Dim dayLbl As String, origStart As Long, actualStart As Long, curCap As Long
        dayLbl = d(0): origStart = d(1)
        actualStart = origStart + offset
        knownDays(dayLbl) = True

        If grp = "GroupC" And dayLbl = "화" Then
            Call EnsureHwaBlockInit(ws, actualStart, 5)
        End If

        curCap = GetCurrentCapacity(ws, actualStart)

        Dim key As String: key = grp & "|" & dayLbl
        Dim ids() As String, idCnt As Long
        Call GetBucketIdsInOrder(buckets, key, ids, idCnt)

        If idCnt > curCap Then
            Dim extra As Long: extra = idCnt - curCap
            Call GrowBlock(ws, actualStart, curCap, extra)
            curCap = curCap + extra
            offset = offset + extra
        End If

        Call FillBlock(ws, actualStart, curCap, ids, idCnt, tierCount, sheetSrc)
        lastRowOfGroup = actualStart + curCap - 1
    Next d

    ' ── 이 그룹에 애초에 슬롯이 없던 요일: 주문 있으면 자동 생성(재실행시 중복 방지) ──
    Dim allDays As Variant: allDays = Array("일", "토", "금", "목", "수", "화", "월")
    Dim dd As Variant
    For Each dd In allDays
        If Not knownDays.exists(CStr(dd)) Then
            Dim k2 As String: k2 = grp & "|" & CStr(dd)
            Dim ids2() As String, idCnt2 As Long
            Call GetBucketIdsInOrder(buckets, k2, ids2, idCnt2)
            If idCnt2 > 0 Then
                Dim marker As String: marker = "AUTO:" & grp & "|" & CStr(dd)
                Dim foundCell As Range
                Set foundCell = ws.Columns(26).Find(marker, LookAt:=xlWhole)

                If Not foundCell Is Nothing Then
                    Dim exStart As Long: exStart = foundCell.Row
                    Dim exCap As Long: exCap = GetCurrentCapacity(ws, exStart)
                    If idCnt2 > exCap Then
                        Dim exExtra As Long: exExtra = idCnt2 - exCap
                        Call GrowBlock(ws, exStart, exCap, exExtra)
                        exCap = exCap + exExtra
                        offset = offset + exExtra
                    End If
                    Call FillBlock(ws, exStart, exCap, ids2, idCnt2, tierCount, sheetSrc)
                    lastRowOfGroup = exStart + exCap - 1
                Else
                    Dim newStart As Long: newStart = lastRowOfGroup + 1
                    Call AutoCreateBlock(ws, lastRowOfGroup, CStr(dd), tierCount, idCnt2, offset, grp)
                    Call FillBlock(ws, newStart, idCnt2, ids2, idCnt2, tierCount, sheetSrc)
                End If
            End If
        End If
    Next dd
End Sub

Sub AddToBucket(buckets As Object, grp As String, dayLbl As String, id As String)
    Dim key As String: key = grp & "|" & dayLbl
    If Not buckets.exists(key) Then Set buckets(key) = New Collection
    Dim col As Collection: Set col = buckets(key)
    Dim v As Variant, dup As Boolean: dup = False
    For Each v In col
        If CStr(v) = id Then dup = True: Exit For
    Next v
    If Not dup Then col.Add id
End Sub

Sub GetBucketIdsInOrder(buckets As Object, key As String, ByRef ids() As String, ByRef cnt As Long)
    cnt = 0
    If buckets.exists(key) Then
        Dim col As Collection: Set col = buckets(key)
        cnt = col.Count
        If cnt > 0 Then
            ReDim ids(1 To cnt)
            Dim n As Long: n = 1
            Dim v As Variant
            For Each v In col
                ids(n) = CStr(v): n = n + 1
            Next v
        End If
    End If
End Sub

Function GetCurrentCapacity(ws As Worksheet, r As Long) As Long
    Dim c As Range: Set c = ws.Cells(r, 1)
    If c.MergeCells Then
        GetCurrentCapacity = c.MergeArea.Rows.Count
    Else
        GetCurrentCapacity = 1
    End If
End Function

Sub GrowBlock(ws As Worksheet, startRow As Long, curCap As Long, extra As Long)
    Dim lastRowOfBlock As Long: lastRowOfBlock = startRow + curCap - 1
    Dim insertAt As Long: insertAt = lastRowOfBlock + 1

    ws.Rows(insertAt).Resize(extra).Insert Shift:=xlDown

    ws.Rows(lastRowOfBlock).Copy
    ws.Rows(insertAt).Resize(extra).PasteSpecial xlPasteAll
    Application.CutCopyMode = False

    Dim aCell As Range: Set aCell = ws.Cells(startRow, 1)
    Dim bCell As Range: Set bCell = ws.Cells(startRow, 2)
    Dim aFormula As String, bFormula As String
    aFormula = aCell.Formula
    bFormula = bCell.Formula

    If aCell.MergeCells Then aCell.MergeArea.UnMerge
    If bCell.MergeCells Then bCell.MergeArea.UnMerge

    ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + curCap + extra - 1, 1)).Merge
    ws.Range(ws.Cells(startRow, 2), ws.Cells(startRow + curCap + extra - 1, 2)).Merge

    aCell.Formula = aFormula
    bCell.Formula = bFormula
End Sub

Sub AutoCreateBlock(ws As Worksheet, ByRef afterRow As Long, dayLbl As String, tierCount As Integer, cap As Long, ByRef offset As Long, grp As String)
    Dim insertAt As Long: insertAt = afterRow + 1

    ws.Rows(insertAt).Resize(cap).Insert Shift:=xlDown

    ws.Rows(afterRow).Copy
    ws.Rows(insertAt).Resize(cap).PasteSpecial xlPasteAll
    Application.CutCopyMode = False

    ws.Range(ws.Cells(insertAt, 1), ws.Cells(insertAt + cap - 1, 1)).Merge
    ws.Range(ws.Cells(insertAt, 2), ws.Cells(insertAt + cap - 1, 2)).Merge
    ws.Cells(insertAt, 1).Formula = "=TEXT(B" & insertAt & ",""aaa"")"
    ws.Cells(insertAt, 2).Formula = DayBFormula(dayLbl)
    ws.Cells(insertAt, 26).Value = "AUTO:" & grp & "|" & dayLbl

    offset = offset + cap
    afterRow = insertAt + cap - 1
End Sub

Function DayBFormula(dayLbl As String) As String
    Select Case dayLbl
        Case "일": DayBFormula = "=B17+6"
        Case "토": DayBFormula = "=B17+5"
        Case "금": DayBFormula = "=B17+4"
        Case "목": DayBFormula = "=B17+3"
        Case "수": DayBFormula = "=B17+2"
        Case "화": DayBFormula = "=B17+1"
        Case "월": DayBFormula = "=DATE(YEAR(TODAY()),1,1)-WEEKDAY(DATE(YEAR(TODAY()),1,1),2)+1+($D$1*7)"
    End Select
End Function

Sub FillBlock(ws As Worksheet, startRow As Long, curCap As Long, ids() As String, idCnt As Long, tierCount As Integer, sheetName As String)
    Dim r As Long, i As Long, t As Integer, mciCol As Integer, eaCol As Integer
    For i = 0 To curCap - 1
        r = startRow + i
        If i < idCnt Then
            ws.Cells(r, 3).Value = ids(i + 1)
            For t = 1 To tierCount
                mciCol = 4 + 2 * (t - 1)
                eaCol = mciCol + 1
                ws.Cells(r, eaCol).Formula = "=SUMIFS('" & sheetName & "'!$D:$D,'" & sheetName & _
                    "'!$A:$A,$C" & r & ",'" & sheetName & "'!$B:$B,$A" & startRow & ",'" & sheetName & _
                    "'!$C:$C,$" & Chr(64 + mciCol) & r & ")"
            Next t
        Else
            ws.Cells(r, 3).Value = "-"
            For t = 1 To tierCount
                mciCol = 4 + 2 * (t - 1)
                eaCol = mciCol + 1
                ws.Cells(r, eaCol).Value = "-"
            Next t
        End If
    Next i
End Sub

Sub EnsureHwaBlockInit(ws As Worksheet, startRow As Long, cap As Long)
    Dim aCell As Range: Set aCell = ws.Cells(startRow, 1)
    If aCell.MergeCells Then Exit Sub
    If Trim(CStr(aCell.Value)) <> "" Then Exit Sub

    ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + cap - 1, 1)).Merge
    ws.Range(ws.Cells(startRow, 2), ws.Cells(startRow + cap - 1, 2)).Merge
    ws.Cells(startRow, 1).Formula = "=TEXT(B" & startRow & ",""aaa"")"
    ws.Cells(startRow, 2).Formula = "=B17+1"
End Sub


Sub Matrix_신규Week_추가()

    Dim wsM As Worksheet, wsOrder As Worksheet
    Set wsM = Sheets("Matrix w29-w37")
    Set wsOrder = Sheets("주문데이터")

    Const HDR_ROW As Long = 3
    Const FIRST_DATA_ROW As Long = 4
    Const NEW_COLOR As Long = 5296274
    Const DROP_COLOR As Long = 13551615

    Application.ScreenUpdating = False

    Call InsertHospitalAtPosition(wsM, "HOSP24", 19, HDR_ROW, FIRST_DATA_ROW)

    Dim wkStr As String
    wkStr = InputBox("추가할 Week 번호를 입력하세요 (예: 38)", "Week 입력")
    If wkStr = "" Then GoTo CleanUp
    Dim wkNum As Long: wkNum = Val(wkStr)
    Dim diagHdr As String, therHdr As String, solHdr As String
    diagHdr = "W" & wkNum & vbLf & "Diag"
    therHdr = "W" & wkNum & vbLf & "Ther"
    solHdr = "W" & wkNum & vbLf & "Solution"

    Dim chkCol As Range
    Set chkCol = wsM.Rows(HDR_ROW).Find(What:=diagHdr, LookAt:=xlWhole, LookIn:=xlValues)
    If Not chkCol Is Nothing Then
        MsgBox "W" & wkNum & " 컬럼이 이미 있습니다. 중복 추가하지 않았습니다.", vbExclamation
        GoTo CleanUp
    End If

    Dim totalCapsCell As Range
    Set totalCapsCell = wsM.Rows(HDR_ROW).Find(What:="Total", LookAt:=xlPart, LookIn:=xlValues, MatchCase:=False)
    If totalCapsCell Is Nothing Then
        MsgBox "'Total' 헤더를 찾을 수 없어 중단합니다.", vbCritical
        Application.ScreenUpdating = True
        Exit Sub
    End If
    Dim insertCol As Long: insertCol = totalCapsCell.Column

    Dim prevEndCol As Long: prevEndCol = insertCol - 1
    Dim prevStartCol As Long: prevStartCol = prevEndCol
    Dim prevPrefix As String
    prevPrefix = Split(CStr(wsM.Cells(HDR_ROW, prevEndCol).Value), vbLf)(0)
    Do While prevStartCol > 1
        If Split(CStr(wsM.Cells(HDR_ROW, prevStartCol - 1).Value), vbLf)(0) = prevPrefix Then
            prevStartCol = prevStartCol - 1
        Else
            Exit Do
        End If
    Loop

    Dim titleRange As Range: Set titleRange = wsM.Range("A1")
    Dim titleMergeAddr As String: titleMergeAddr = titleRange.MergeArea.Address
    titleRange.MergeArea.UnMerge

    wsM.Range(wsM.Cells(1, insertCol), wsM.Cells(1, insertCol + 2)).EntireColumn.Insert Shift:=xlToRight

    wsM.Range(wsM.Cells(1, prevStartCol), wsM.Cells(wsM.UsedRange.Row + wsM.UsedRange.Rows.Count - 1, prevEndCol)).Copy
    wsM.Cells(1, insertCol).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False

    Dim newLastCol As Long
    newLastCol = wsM.Range(titleMergeAddr).Column + wsM.Range(titleMergeAddr).Columns.Count - 1 + 3
    wsM.Range(wsM.Cells(1, 1), wsM.Cells(1, newLastCol)).Merge
    wsM.Cells(1, 1).Value = titleRange.Value

    Dim dCol As Long, tCol As Long, sCol As Long
    dCol = insertCol: tCol = insertCol + 1: sCol = insertCol + 2
    wsM.Cells(HDR_ROW, dCol).Value = diagHdr
    wsM.Cells(HDR_ROW, tCol).Value = therHdr
    wsM.Cells(HDR_ROW, sCol).Value = solHdr

    Dim subtotalRow As Long, totalRow As Long
    Dim fCell As Range
    Set fCell = wsM.Cells.Find(What:="Sub total", LookAt:=xlPart, LookIn:=xlValues, MatchCase:=False)
    If fCell Is Nothing Then
        MsgBox "'Sub total' 행을 찾지 못했습니다. 구조를 확인해주세요.", vbCritical
        Application.ScreenUpdating = True
        Exit Sub
    End If
    subtotalRow = fCell.Row
    totalRow = subtotalRow + 1
    Dim lastDataRow As Long: lastDataRow = subtotalRow - 1

    Set totalCapsCell = wsM.Rows(HDR_ROW).Find(What:="Total", LookAt:=xlPart, LookIn:=xlValues, MatchCase:=False)
    Dim xCol As Long: xCol = totalCapsCell.Column
    Dim xColLtr As String: xColLtr = Split(wsM.Cells(1, xCol).Address, "$")(1)
    Dim dColLtr As String: dColLtr = Split(wsM.Cells(1, dCol).Address, "$")(1)
    Dim tColLtr As String: tColLtr = Split(wsM.Cells(1, tCol).Address, "$")(1)
    Dim sColLtr As String: sColLtr = Split(wsM.Cells(1, sCol).Address, "$")(1)

    Dim diagDict As Object, therDict As Object, solDict As Object
    Set diagDict = CreateObject("Scripting.Dictionary")
    Set therDict = CreateObject("Scripting.Dictionary")
    Set solDict = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long, i As Long
    lastRow = wsOrder.Cells(wsOrder.Rows.Count, 1).End(xlUp).Row
    Dim cID As String, cMci As Double, cQty As Long, keyU As String
    Dim warnList As Collection: Set warnList = New Collection

    For i = 2 To lastRow
        cID = Trim(CStr(wsOrder.Cells(i, 1).Value))
        cMci = Val(wsOrder.Cells(i, 3).Value)
        cQty = Val(wsOrder.Cells(i, 4).Value)
        If cID = "" Or cQty = 0 Then GoTo NextRow
        keyU = UCase(cID)

        If keyU = "PARTNER_CO" Then
            solDict(keyU) = IIf(solDict.exists(keyU), solDict(keyU), 0) + cQty
        Else
            Select Case cMci
                Case 2, 5, 10
                    diagDict(keyU) = IIf(diagDict.exists(keyU), diagDict(keyU), 0) + cQty
                Case 25, 30, 50, 100, 200
                    therDict(keyU) = IIf(therDict.exists(keyU), therDict(keyU), 0) + cQty
                Case Else
                    warnList.Add "미분류 mCi(" & cMci & "): " & cID
            End Select
        End If
NextRow:
    Next i

    Dim r As Long, hosp As String, hospU As String
    Dim d As Double, t As Double, s As Double, prevTotal As Double, thisTotal As Double
    Dim matchedKeys As Object: Set matchedKeys = CreateObject("Scripting.Dictionary")

    For r = FIRST_DATA_ROW To lastDataRow
        hosp = Trim(CStr(wsM.Cells(r, 2).Value))
        If hosp = "" Then GoTo NextHospRow
        hospU = UCase(hosp)
        matchedKeys(hospU) = True

        d = IIf(diagDict.exists(hospU), diagDict(hospU), 0)
        t = IIf(therDict.exists(hospU), therDict(hospU), 0)
        s = IIf(solDict.exists(hospU), solDict(hospU), 0)

        wsM.Cells(r, dCol).Value = IIf(d = 0, "-", d)
        wsM.Cells(r, tCol).Value = IIf(t = 0, "-", t)
        wsM.Cells(r, sCol).Value = IIf(s = 0, "-", s)

        prevTotal = 0
        Dim pc As Long
        For pc = prevStartCol To prevEndCol
            prevTotal = prevTotal + Val(wsM.Cells(r, pc).Value)
        Next pc
        thisTotal = d + t + s

        If thisTotal > 0 And prevTotal = 0 Then
            wsM.Range(wsM.Cells(r, dCol), wsM.Cells(r, sCol)).Interior.Color = NEW_COLOR
            wsM.Cells(r, 2).Interior.Color = NEW_COLOR
        ElseIf thisTotal = 0 And prevTotal > 0 Then
            wsM.Range(wsM.Cells(r, dCol), wsM.Cells(r, sCol)).Interior.Color = DROP_COLOR
            wsM.Cells(r, 2).Interior.Color = DROP_COLOR
        End If

        wsM.Cells(r, xCol).Formula = "=SUM(C" & r & ":" & sColLtr & r & ")"
NextHospRow:
    Next r

    Dim allKeys As Object: Set allKeys = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In diagDict.Keys: allKeys(k) = True: Next k
    For Each k In therDict.Keys: allKeys(k) = True: Next k
    For Each k In solDict.Keys: allKeys(k) = True: Next k

    Dim newHospCount As Long: newHospCount = 0
    For Each k In allKeys.Keys
        If Not matchedKeys.exists(CStr(k)) Then
            newHospCount = newHospCount + 1
            wsM.Rows(lastDataRow + 1).Insert Shift:=xlDown
            wsM.Rows(lastDataRow).Copy
            wsM.Rows(lastDataRow + 1).PasteSpecial xlPasteFormats
            Application.CutCopyMode = False

            wsM.Cells(lastDataRow + 1, 1).Value = wsM.Cells(lastDataRow, 1).Value + 1
            wsM.Cells(lastDataRow + 1, 2).Value = k
            Dim col As Long
            For col = 3 To xCol - 1
                If col < dCol Or col > sCol Then wsM.Cells(lastDataRow + 1, col).Value = "-"
            Next col
            wsM.Cells(lastDataRow + 1, dCol).Value = IIf(diagDict.exists(CStr(k)), diagDict(CStr(k)), "-")
            wsM.Cells(lastDataRow + 1, tCol).Value = IIf(therDict.exists(CStr(k)), therDict(CStr(k)), "-")
            wsM.Cells(lastDataRow + 1, sCol).Value = IIf(solDict.exists(CStr(k)), solDict(CStr(k)), "-")
            wsM.Cells(lastDataRow + 1, xCol).Formula = "=SUM(C" & (lastDataRow + 1) & ":" & sColLtr & (lastDataRow + 1) & ")"

            wsM.Range(wsM.Cells(lastDataRow + 1, dCol), wsM.Cells(lastDataRow + 1, sCol)).Interior.Color = NEW_COLOR
            wsM.Cells(lastDataRow + 1, 2).Interior.Color = NEW_COLOR

            lastDataRow = lastDataRow + 1
            subtotalRow = subtotalRow + 1
            totalRow = totalRow + 1
        End If
    Next k

    wsM.Cells(subtotalRow, dCol).Formula = "=SUM(" & dColLtr & FIRST_DATA_ROW & ":" & dColLtr & lastDataRow & ")"
    wsM.Cells(subtotalRow, tCol).Formula = "=SUM(" & tColLtr & FIRST_DATA_ROW & ":" & tColLtr & lastDataRow & ")"
    wsM.Cells(subtotalRow, sCol).Formula = "=SUM(" & sColLtr & FIRST_DATA_ROW & ":" & sColLtr & lastDataRow & ")"
    wsM.Cells(subtotalRow, xCol).Formula = "=SUM(C" & subtotalRow & ":" & sColLtr & subtotalRow & ")"

    wsM.Range(wsM.Cells(totalRow, dCol), wsM.Cells(totalRow, sCol)).Merge
    wsM.Cells(totalRow, dCol).Formula = "=SUM(" & dColLtr & subtotalRow & ":" & sColLtr & subtotalRow & ")"
    wsM.Cells(totalRow, xCol).Formula = "=SUM(C" & totalRow & ":" & sColLtr & totalRow & ")"

CleanUp:
    Application.ScreenUpdating = True
    Dim msg As String, w As Variant
    If newHospCount > 0 Then msg = msg & "신규 병원 " & newHospCount & "건 자동 추가(연두색)." & vbCrLf
    If Not warnList Is Nothing Then
        If warnList.Count > 0 Then
            For Each w In warnList
                msg = msg & w & vbCrLf
            Next w
        End If
    End If
    If msg = "" Then
        MsgBox "W" & wkNum & " 열 추가 완료!", vbInformation
    Else
        MsgBox "W" & wkNum & " 열 추가 완료." & vbCrLf & vbCrLf & msg, vbExclamation
    End If

End Sub

Sub InsertHospitalAtPosition(ws As Worksheet, hospitalName As String, position As Long, HDR_ROW As Long, FIRST_DATA_ROW As Long)
    Dim lastDataRowNow As Long
    Dim fCell As Range
    Set fCell = ws.Cells.Find(What:="Sub total", LookAt:=xlPart, LookIn:=xlValues, MatchCase:=False)
    If fCell Is Nothing Then Exit Sub
    lastDataRowNow = fCell.Row - 1

    Dim already As Range
    Set already = ws.Range(ws.Cells(FIRST_DATA_ROW, 2), ws.Cells(lastDataRowNow, 2)).Find(What:=hospitalName, LookAt:=xlWhole, LookIn:=xlValues)
    If Not already Is Nothing Then Exit Sub

    Dim targetRow As Long: targetRow = FIRST_DATA_ROW + position - 1
    If targetRow > lastDataRowNow + 1 Then targetRow = lastDataRowNow + 1

    Dim lastCol As Long
    lastCol = ws.Cells(HDR_ROW, ws.Columns.Count).End(xlToLeft).Column

    ws.Rows(targetRow).Insert Shift:=xlDown
    ws.Rows(targetRow + 1).Copy
    ws.Rows(targetRow).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False

    ws.Cells(targetRow, 2).Value = hospitalName

    Dim c As Long
    For c = 3 To lastCol
        ws.Cells(targetRow, c).Value = "-"
    Next c

    Dim r As Long, seq As Long: seq = 1
    For r = FIRST_DATA_ROW To lastDataRowNow + 1
        ws.Cells(r, 1).Value = seq
        seq = seq + 1
    Next r
End Sub


' ============================================================
' Step3_PerHospital_추가
'  - "Per hospital" 시트(1월~누적 로그) 맨 아래에 이번 주 데이터만 추가
'  - 기존 데이터는 절대 안 건드림 (append 전용)
'  - Month는 "그 주 일요일이 속한 달" 기준 (ISO 8601, Week01=2025-12-29 월)
'  - 리비전 이력(=558+113 같은 수식)은 반영 안 함, 최종 확정 수량만 숫자로 기록
' ============================================================

Sub Step3_PerHospital_추가()

    Dim wsPH As Worksheet, wsOrder As Worksheet, wsOrder2 As Worksheet
    Set wsPH = Sheets("Per hospital")
    Set wsOrder = Sheets("주문데이터")
    Set wsOrder2 = Sheets("2차주문데이터")

    Const groupAList As String = "HOSP01,HOSP02,HOSP03,HOSP04,HOSP05,HOSP06"

    Dim wkStr As String
    wkStr = InputBox("Week 번호를 입력하세요 (예: 38)", "Week 입력")
    If wkStr = "" Then Exit Sub
    Dim wkNum As Long: wkNum = Val(wkStr)
    Dim wkLabel As String: wkLabel = "Week " & Format(wkNum, "00")

    ' ── 중복 실행 방지 ──
    Dim dupCheck As Range
    Set dupCheck = wsPH.Columns(3).Find(What:=wkLabel, LookAt:=xlWhole, LookIn:=xlValues)
    If Not dupCheck Is Nothing Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox(wkLabel & " 데이터가 이미 있습니다. 그래도 추가하시겠습니까? (중복될 수 있습니다)", vbYesNo + vbExclamation)
        If ans = vbNo Then Exit Sub
    End If

    ' ── Month 계산: ISO 8601, Week01 = 2025-12-29(월) / Month = 그 주 일요일이 속한 달 ──
    Dim anchorMon As Date: anchorMon = DateSerial(2025, 12, 29)   ' Week01 월요일 (2026년 사이클 기준)
    Dim weekMon As Date: weekMon = anchorMon + (wkNum - 1) * 7
    Dim weekSun As Date: weekSun = weekMon + 6
    Dim monthLabel As String: monthLabel = MonthName(Month(weekSun), False)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' ── 주문데이터/2차주문데이터 -> (Shipment, 병원, mCi) 별 수량 합산 ──
    Dim shipOrder As Variant: shipOrder = Array("GroupB- 1st", "GroupA- 1st", "GroupB- 2nd")
    Dim hospOrder As Object: Set hospOrder = CreateObject("Scripting.Dictionary")   ' key: shipment -> Collection(병원명, 순서유지)
    Dim mciDict As Object: Set mciDict = CreateObject("Scripting.Dictionary")       ' key: shipment|병원 -> Collection(mCi값)
    Dim qtyDict As Object: Set qtyDict = CreateObject("Scripting.Dictionary")       ' key: shipment|병원|mCi -> 합산수량

    Dim s As Variant
    For Each s In shipOrder
        Set hospOrder(s) = New Collection
    Next s

    Dim lastRow As Long, i As Long, cID As String, cMci As Double, cQty As Long, shipKey As String
    lastRow = wsOrder.Cells(wsOrder.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        cID = Trim(CStr(wsOrder.Cells(i, 1).Value))
        cMci = Val(wsOrder.Cells(i, 3).Value)
        cQty = Val(wsOrder.Cells(i, 4).Value)
        If cID = "" Or cQty = 0 Then GoTo NextRow1
        shipKey = IIf(InStr("," & groupAList & ",", "," & cID & ",") > 0, "GroupA- 1st", "GroupB- 1st")
        Call Accumulate(hospOrder, mciDict, qtyDict, shipKey, cID, cMci, cQty)
NextRow1:
    Next i

    Dim lastRow2 As Long, c2ID As String, c2Mci As Double, c2Qty As Long
    lastRow2 = wsOrder2.Cells(wsOrder2.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow2
        c2ID = Trim(CStr(wsOrder2.Cells(i, 1).Value))
        c2Mci = Val(wsOrder2.Cells(i, 3).Value)
        c2Qty = Val(wsOrder2.Cells(i, 4).Value)
        If c2ID = "" Or c2Qty = 0 Then GoTo NextRow2
        Call Accumulate(hospOrder, mciDict, qtyDict, "GroupB- 2nd", c2ID, c2Mci, c2Qty)
NextRow2:
    Next i

    ' ── Appended at the bottom of the sheet in order (GroupB-1st -> GroupA-1st -> GroupB-2nd) ──
    Dim startRow As Long
    startRow = wsPH.Cells(wsPH.Rows.Count, 2).End(xlUp).Row + 1
    Dim r As Long: r = startRow
    Dim totalRowsAdded As Long: totalRowsAdded = 0

    For Each s In shipOrder
        Dim hospCol As Collection: Set hospCol = hospOrder(s)
        If hospCol.Count = 0 Then GoTo NextShip

        Dim groupStart As Long: groupStart = r
        Dim h As Variant
        For Each h In hospCol
            Dim mciKey As String: mciKey = CStr(s) & "|" & CStr(h)
            Dim mciCol As Collection: Set mciCol = mciDict(mciKey)
            Dim mciArr() As Double
            Call SortMciDescending(mciCol, mciArr)

            Dim mi As Long
            For mi = LBound(mciArr) To UBound(mciArr)
                Dim mciVal As Double: mciVal = mciArr(mi)
                Dim qKey As String: qKey = mciKey & "|" & mciVal
                Dim qty As Long: qty = qtyDict(qKey)

                wsPH.Cells(r, 2).Value = monthLabel
                wsPH.Cells(r, 3).Value = wkLabel
                wsPH.Cells(r, 4).Value = s
                wsPH.Cells(r, 5).Value = h
                wsPH.Cells(r, 6).Value = FormatMci(mciVal)
                wsPH.Cells(r, 7).Value = qty
                r = r + 1
            Next mi
        Next h

        Dim groupEnd As Long: groupEnd = r - 1
        wsPH.Cells(groupStart, 8).Formula = "=SUM(G" & groupStart & ":G" & groupEnd & ")"
        totalRowsAdded = totalRowsAdded + (groupEnd - groupStart + 1)
NextShip:
    Next s

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    If totalRowsAdded = 0 Then
        MsgBox wkLabel & "에 해당하는 주문데이터가 없어 추가된 행이 없습니다.", vbExclamation
    Else
        MsgBox wkLabel & " (" & monthLabel & ") Per hospital 시트에 " & totalRowsAdded & "행 추가 완료!" & vbCrLf & _
               (startRow) & "행부터 " & (r - 1) & "행까지.", vbInformation

        ' ── Summary Pivot 원본 범위 확장 + 새로고침 ──
        Call Step3b_PerHospital_Pivot_갱신
    End If
End Sub

' ------------------------------------------------------------
Sub Step3b_PerHospital_Pivot_갱신()
    Dim wsPH As Worksheet
    Set wsPH = Sheets("Per hospital")
    Dim lastRow As Long
    lastRow = wsPH.Cells(wsPH.Rows.Count, 2).End(xlUp).Row

    Dim pt As PivotTable
    Set pt = Sheets("Summary Pivot- hospitals global").PivotTables(1)

    pt.ChangePivotCache ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:="'Per hospital'!B3:H" & lastRow)

    pt.RefreshTable
    MsgBox "Summary Pivot 갱신 완료! (Per hospital B3:H" & lastRow & " 기준)", vbInformation
End Sub

' ------------------------------------------------------------
Sub Accumulate(hospOrder As Object, mciDict As Object, qtyDict As Object, shipKey As String, hospID As String, mciVal As Double, qty As Long)
    Dim hospCol As Collection: Set hospCol = hospOrder(shipKey)
    Dim exists As Boolean: exists = False
    Dim v As Variant
    For Each v In hospCol
        If CStr(v) = hospID Then exists = True: Exit For
    Next v
    If Not exists Then hospCol.Add hospID

    Dim mciKey As String: mciKey = shipKey & "|" & hospID
    If Not mciDict.exists(mciKey) Then Set mciDict(mciKey) = New Collection
    Dim mciCol As Collection: Set mciCol = mciDict(mciKey)
    Dim mciExists As Boolean: mciExists = False
    For Each v In mciCol
        If CDbl(v) = mciVal Then mciExists = True: Exit For
    Next v
    If Not mciExists Then mciCol.Add mciVal

    Dim qKey As String: qKey = mciKey & "|" & mciVal
    qtyDict(qKey) = IIf(qtyDict.exists(qKey), qtyDict(qKey), 0) + qty
End Sub

Sub SortMciDescending(mciCol As Collection, ByRef arr() As Double)
    Dim n As Long: n = mciCol.Count
    ReDim arr(0 To n - 1)
    Dim idx As Long: idx = 0
    Dim v As Variant
    For Each v In mciCol
        arr(idx) = CDbl(v): idx = idx + 1
    Next v
    Dim p As Long, q As Long, tmp As Double
    For p = 0 To n - 2
        For q = 0 To n - 2 - p
            If arr(q) < arr(q + 1) Then
                tmp = arr(q): arr(q) = arr(q + 1): arr(q + 1) = tmp
            End If
        Next q
    Next p
End Sub

Function FormatMci(v As Double) As String
    If v = Int(v) Then
        FormatMci = CStr(CLng(v)) & "mCi"
    Else
        FormatMci = CStr(v) & "mCi"
    End If
End Function


' ============================================================
' Step4_ContainerLog_추가
'  - "2026 Containers+면장" 시트 맨 아래에 이번 주 A/B/C/D/F만 추가
'  - E(컨테이너 수), G(면장)는 절대 안 건드림 - 본인이 손으로 입력하는 영역
'  - 기존 누적 데이터(1~현재 주차)는 절대 안 건드림 (append 전용)
' ============================================================

Sub Step4_ContainerLog_추가()

    Dim wsCL As Worksheet, wsOrder As Worksheet, wsOrder2 As Worksheet
    Set wsCL = Sheets("2026 Containers+면장")
    Set wsOrder = Sheets("주문데이터")
    Set wsOrder2 = Sheets("2차주문데이터")

    Const groupAList As String = "HOSP01,HOSP02,HOSP03,HOSP04,HOSP05,HOSP06"

    Dim wkStr As String
    wkStr = InputBox("Week 번호를 입력하세요 (예: 38)", "Week 입력")
    If wkStr = "" Then Exit Sub
    Dim wkNum As Long: wkNum = Val(wkStr)
    Dim wkLabel As String: wkLabel = Format(wkNum, "00")   ' 이 시트는 "38"처럼 2자리만 씀 (Per hospital의 "Week 38"과 다름)

    ' ── 중복 실행 방지 ──
    Dim dupCheck As Range
    Set dupCheck = wsCL.Columns(2).Find(What:=wkLabel, LookAt:=xlWhole, LookIn:=xlValues)
    If Not dupCheck Is Nothing Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox("Week " & wkLabel & " 데이터가 이미 있습니다. 그래도 추가하시겠습니까?", vbYesNo + vbExclamation)
        If ans = vbNo Then Exit Sub
    End If

    ' ── Month 계산: Step3와 동일한 "일요일이 속한 달" 규칙, 대문자로 ──
    Dim anchorMon As Date: anchorMon = DateSerial(2025, 12, 29)   ' Week01 월요일 (2026년 사이클 기준, Step3와 동일 앵커)
    Dim weekMon As Date: weekMon = anchorMon + (wkNum - 1) * 7
    Dim weekSun As Date: weekSun = weekMon + 6
    Dim monthLabel As String: monthLabel = UCase(MonthName(Month(weekSun), False))   ' "JANUARY" 형태

    Application.ScreenUpdating = False

    ' ── 주문데이터/2차주문데이터 -> 선적별 총 수량 합산 (mCi 구분 없이 전부 합침, Step3와 다른 점) ──
    Dim shipOrder As Variant: shipOrder = Array("GroupB-01", "GroupA-01", "GroupB-02")   ' 이 시트의 표기 그대로
    Dim shipQty As Object: Set shipQty = CreateObject("Scripting.Dictionary")
    Dim s As Variant
    For Each s In shipOrder: shipQty(CStr(s)) = 0: Next s

    Dim lastRow As Long, i As Long, cID As String, cQty As Long
    lastRow = wsOrder.Cells(wsOrder.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        cID = Trim(CStr(wsOrder.Cells(i, 1).Value))
        cQty = Val(wsOrder.Cells(i, 4).Value)
        If cID = "" Or cQty = 0 Then GoTo NextRow1
        If InStr("," & groupAList & ",", "," & cID & ",") > 0 Then
            shipQty("GroupA-01") = shipQty("GroupA-01") + cQty
        Else
            shipQty("GroupB-01") = shipQty("GroupB-01") + cQty
        End If
NextRow1:
    Next i

    Dim lastRow2 As Long, c2ID As String, c2Qty As Long
    lastRow2 = wsOrder2.Cells(wsOrder2.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow2
        c2ID = Trim(CStr(wsOrder2.Cells(i, 1).Value))
        c2Qty = Val(wsOrder2.Cells(i, 4).Value)
        If c2ID = "" Or c2Qty = 0 Then GoTo NextRow2
        shipQty("GroupB-02") = shipQty("GroupB-02") + c2Qty
NextRow2:
    Next i

    ' ── Appends 3 rows at the bottom (GroupB-01/GroupA-01/GroupB-02), skipping any shipment with zero quantity (matches legacy file behavior) ──
    Dim startRow As Long
    startRow = wsCL.Cells(wsCL.Rows.Count, 1).End(xlUp).Row + 1
    Dim r As Long: r = startRow
    Dim firstRowOfWeek As Long: firstRowOfWeek = startRow
    Dim rowsAdded As Long: rowsAdded = 0

    For Each s In shipOrder
        Dim qty As Long: qty = shipQty(CStr(s))
        If qty > 0 Then
            wsCL.Cells(r, 4).Value = qty              ' D: Number of Caps/Sol
            wsCL.Cells(r, 3).Value = CStr(s)           ' C: #Shipment
            ' E(컨테이너 수), G(면장)는 비워둠 - 본인 입력 영역
            r = r + 1
            rowsAdded = rowsAdded + 1
        End If
    Next s

    If rowsAdded = 0 Then
        MsgBox "Week " & wkLabel & "에 해당하는 주문데이터가 없어 추가된 행이 없습니다.", vbExclamation
        GoTo CleanUp
    End If

    wsCL.Cells(firstRowOfWeek, 1).Value = monthLabel   ' A: MONTH (그 주 첫 줄에만)
    wsCL.Cells(firstRowOfWeek, 2).Value = wkLabel       ' B: Week (그 주 첫 줄에만)
    wsCL.Cells(firstRowOfWeek, 6).Formula = "=SUM(E" & firstRowOfWeek & ":E" & (r - 1) & ")"   ' F: 그 주 컨테이너 합계 (E는 본인이 채운 뒤 이 수식이 자동 합산)

CleanUp:
    Application.ScreenUpdating = True
    If rowsAdded > 0 Then
        MsgBox "Week " & wkLabel & " (" & monthLabel & ") Container Log에 " & rowsAdded & "개 선적 행 추가 완료!" & vbCrLf & _
               "E열(컨테이너 수)와 G열(면장)은 직접 입력해주세요.", vbInformation
    End If
End Sub



' ============================================================
' Step5_Summary_추가
'  - "Summary" 시트 맨 아래(Week37- SEPTEMBER 다음)에 새 주차만 추가
'  - 기존 Week12~37 블록은 절대 안 건드림 (append 전용)
'  - 25mCi는 별도 한 줄, 나머지 High dose(30/50/100/200mCi)는
'    Group A/B/1차/2차 구분 없이 전부 합쳐서 "Capsules- All shipments"로 기록
'  - 최근(Week36~37) 형식 그대로: 실제 셀 참조 수식 사용, 손으로 더한 문자열 수식 안 씀
' ============================================================

Sub Step5_Summary_추가()

    Dim wsSum As Worksheet, wsOrder As Worksheet, wsOrder2 As Worksheet
    Set wsSum = Sheets("Summary")
    Set wsOrder = Sheets("주문데이터")
    Set wsOrder2 = Sheets("2차주문데이터")

    Dim wkStr As String
    wkStr = InputBox("Week 번호를 입력하세요 (예: 40)", "Week 입력")
    If wkStr = "" Then Exit Sub
    Dim wkNum As Long: wkNum = Val(wkStr)
    Dim wkShort As String: wkShort = "Week " & wkNum

    ' ── Month 계산: Step3/4와 동일한 "일요일이 속한 달" 규칙, 대문자 ──
    Dim anchorMon As Date: anchorMon = DateSerial(2025, 12, 29)   ' Week01 월요일 (2026년 사이클 기준)
    Dim weekMon As Date: weekMon = anchorMon + (wkNum - 1) * 7
    Dim weekSun As Date: weekSun = weekMon + 6
    Dim monthLabel As String: monthLabel = UCase(MonthName(Month(weekSun), False))
    Dim wkLabelFull As String: wkLabelFull = wkShort & "- " & monthLabel

    ' ── 중복 실행 방지 ──
    Dim dupCheck As Range
    Set dupCheck = wsSum.Columns(3).Find(What:=wkLabelFull, LookAt:=xlWhole, LookIn:=xlValues)
    If Not dupCheck Is Nothing Then
        MsgBox wkLabelFull & " 블록이 이미 있습니다. 중복 추가하지 않았습니다.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' ── 주문데이터 + 2차주문데이터 통합, Group A/B/1차/2차 구분 없이 mCi별 합산 ──
    Dim mci25 As Double, mciHighOthers As Double, mci5 As Double, mci2 As Double, solutionQty As Double
    mci25 = 0: mciHighOthers = 0: mci5 = 0: mci2 = 0: solutionQty = 0

    Dim srcSheets As Variant: srcSheets = Array(wsOrder, wsOrder2)
    Dim ws As Variant, lastRow As Long, i As Long
    Dim cID As String, cMci As Double, cQty As Long

    For Each ws In srcSheets
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastRow
            cID = Trim(CStr(ws.Cells(i, 1).Value))
            cMci = Val(ws.Cells(i, 3).Value)
            cQty = Val(ws.Cells(i, 4).Value)
            If cID = "" Or cQty = 0 Then GoTo NextRow

            If UCase(cID) = "PARTNER_CO" Then
                solutionQty = solutionQty + cQty
            Else
                Select Case cMci
                    Case 25: mci25 = mci25 + cQty
                    Case 30, 50, 100, 200: mciHighOthers = mciHighOthers + cQty
                    Case 5: mci5 = mci5 + cQty
                    Case 2: mci2 = mci2 + cQty
                End Select
            End If
NextRow:
        Next i
    Next ws

    ' ── 맨 아래 블록 위치 찾기 (Capsules- All shipments 마지막 줄 다음 3칸 띄고 새 블록) ──
    Dim lastAllShipRow As Long: lastAllShipRow = 0
    Dim r As Long
    For r = wsSum.Cells(wsSum.Rows.Count, 3).End(xlUp).Row To 1 Step -1
        If Trim(CStr(wsSum.Cells(r, 3).Value)) = "Capsules- All shipments" Then
            lastAllShipRow = r
            Exit For
        End If
    Next r
    If lastAllShipRow = 0 Then
        MsgBox "'Capsules- All shipments' 마지막 줄을 찾지 못했습니다. 시트 구조를 확인해주세요.", vbCritical
        Application.ScreenUpdating = True
        Exit Sub
    End If

    Dim hdrRow As Long: hdrRow = lastAllShipRow + 4   ' 기존 패턴과 동일: 3줄 띄고 헤더
    Dim r25 As Long: r25 = hdrRow + 1
    Dim r5 As Long: r5 = hdrRow + 2
    Dim r2 As Long: r2 = hdrRow + 3
    Dim rSol As Long: rSol = hdrRow + 4
    Dim rAll As Long: rAll = hdrRow + 5

    ' ── 헤더 ──
    wsSum.Cells(hdrRow, 3).Value = wkLabelFull
    wsSum.Cells(hdrRow, 6).Value = wkShort

    ' ── 25mCi ──
    wsSum.Cells(r25, 3).Value = "25mCi"
    wsSum.Cells(r25, 4).Value = mci25
    wsSum.Cells(r25, 6).Value = "High dose"
    wsSum.Cells(r25, 7).Formula = "=D" & rAll & "+D" & r25

    ' ── 5mCi ──
    wsSum.Cells(r5, 3).Value = "5mCi"
    wsSum.Cells(r5, 4).Value = mci5
    wsSum.Cells(r5, 6).Value = "Low dose"
    wsSum.Cells(r5, 7).Formula = "=D" & r5 & "+D" & r2

    ' ── 2mCi ──
    wsSum.Cells(r2, 3).Value = "2mCi"
    wsSum.Cells(r2, 4).Value = mci2
    wsSum.Cells(r2, 6).Value = "Solution"
    wsSum.Cells(r2, 7).Formula = "=D" & rSol

    ' ── Solution ──
    wsSum.Cells(rSol, 3).Value = "Solution"
    wsSum.Cells(rSol, 4).Value = solutionQty

    ' ── Capsules- All shipments (25mCi 제외 나머지 High dose 전부) ──
    wsSum.Cells(rAll, 3).Value = "Capsules- All shipments"
    wsSum.Cells(rAll, 4).Value = mciHighOthers

CleanUp:
    Application.ScreenUpdating = True
    MsgBox wkLabelFull & " 블록 추가 완료! (" & hdrRow & "~" & rAll & "행)", vbInformation
End Sub
