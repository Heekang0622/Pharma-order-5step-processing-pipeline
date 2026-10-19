Attribute VB_Name = "OrderPipeline_MultiChannel"
' ============================================================
' Weekly radiopharmaceutical order processing pipeline (Portfolio / Anonymized Version)
' Real hospital/client codes, the distributor name, and colleague names have
' been replaced with generic placeholders (HOSP01-25, PARTNER_CO, GroupA/B/C).
' Business logic and structure are preserved as-is.
' ============================================================
Option Explicit

'==============================================================
' 공용 설정 (여기만 고치면 Step 1~5 전체에 반영됩니다)
'==============================================================
Private Const GROUP_A_LIST As String = "HOSP01,HOSP02,HOSP03,HOSP04,HOSP05,HOSP06"
Private Const GROUP_B_LIST As String = "HOSP07,HOSP08,HOSP09,HOSP10,HOSP11,HOSP12,HOSP13,HOSP14,HOSP15,HOSP16,HOSP17,HOSP18,HOSP19,HOSP20,HOSP21,HOSP22,HOSP23,HOSP24,HOSP25"
Private Const VALID_DAYS As String = "월화수목금토일"
Private Const SOLUTION_KEY As String = "PARTNER_CO"

Private Const SH_FORM As String = "Sheet1"
Private Const SH_ORDER1 As String = "주문데이터"
Private Const SH_ORDER2 As String = "2차주문데이터"
Private Const SH_MATRIX As String = "Matrix w29-w37"
Private Const SH_PERHOSP As String = "Per hospital"
Private Const SH_PIVOT As String = "Summary Pivot- hospitals global"
Private Const SH_CONTAINER As String = "2026 Containers+면장"
Private Const SH_SUMMARY As String = "Summary"

' Matrix(Step 2)에 2차주문데이터도 포함할지 (Step 3,4,5와 맞추려면 True)
Private Const MATRIX_INCLUDE_2ND As Boolean = True
' Step 3~5의 주차 기준 월요일 (1주차 월요일)
Private Const ANCHOR_MON As Date = #12/29/2025#

Private gEa As Range
Private gLastWeek As Long

Private Type OrderRec
    id As String
    mci As Double
    qty As Long
    is2nd As Boolean
    isSol As Boolean
End Type
Private gOrd() As OrderRec
Private gOrdN As Long
Private gSolNote As String

'==============================================================
' 공용 함수
'==============================================================
Function Norm(s As String) As String
    Norm = UCase(Replace(Trim(s), " ", ""))
End Function

Function InList(lst As String, id As String) As Boolean
    InList = (InStr("," & UCase(Replace(lst, " ", "")) & ",", "," & Norm(id) & ",") > 0)
End Function

Sub AddId(b As Object, key As String, id As String)
    Dim col As Collection, v As Variant
    If Not b.exists(key) Then Set b(key) = New Collection
    Set col = b(key)
    For Each v In col
        If Norm(CStr(v)) = Norm(id) Then Exit Sub
    Next v
    col.Add id
End Sub

Function ColL(c As Long) As String
    ColL = Split(Cells(1, c).Address(True, False), "$")(0)
End Function

Private Sub AddQty(d As Object, ByVal key As String, ByVal q As Double)
    If d.exists(key) Then
        d(key) = d(key) + q
    Else
        d(key) = q
    End If
End Sub

Private Function GetQty(d As Object, ByVal key As String) As Double
    If d.exists(key) Then GetQty = d(key)
End Function

Private Function WeekMonday(ByVal wk As Long) As Date
    WeekMonday = ANCHOR_MON + (wk - 1) * 7
End Function

Private Function MonthEn(ByVal m As Long) As String
    MonthEn = Choose(m, "January", "February", "March", "April", "May", "June", _
                        "July", "August", "September", "October", "November", "December")
End Function

Private Function LastRowAny(ws As Worksheet, ByVal c1 As Long, ByVal c2 As Long) As Long
    Dim c As Long, r As Long
    For c = c1 To c2
        r = ws.Cells(ws.rows.Count, c).End(xlUp).Row
        If r > LastRowAny Then LastRowAny = r
    Next c
End Function

Private Function AskWeek() As Long
    Dim s As String, d As String
    If gLastWeek > 0 Then d = CStr(gLastWeek)
    s = Trim(InputBox("Week 번호를 입력하세요 (예: 38)", "Week 입력", d))
    If s = "" Then Exit Function
    If Not IsNumeric(s) Then
        MsgBox "Week는 숫자로 입력해주세요.", vbExclamation
        Exit Function
    End If
    If Val(s) < 1 Or Val(s) > 53 Then
        MsgBox "Week는 1~53 사이로 입력해주세요.", vbExclamation
        Exit Function
    End If
    AskWeek = CLng(Val(s))
    gLastWeek = AskWeek
End Function

' 월 라벨: 시트에 이미 쓰인 방식(영문 "September" / 한글 "10월")을 따라갑니다.
' 피벗의 MONTH 필터가 라벨 문자열 기준이라, 방식이 섞이면 새 주차가 피벗에서 빠집니다.
Private Function UseKoreanMonth(ws As Worksheet, ByVal col As Long) As Boolean
    Dim r As Long, t As String
    For r = ws.Cells(ws.rows.Count, col).End(xlUp).Row To 1 Step -1
        t = Trim(CStr(ws.Cells(r, col).Value))
        If t <> "" Then
            If col = 3 And ws.Name = SH_SUMMARY Then
                ' Summary: "Week 40- 10월" 형태의 헤더만 대상
                If Left(t, 5) = "Week " And InStr(t, "- ") > 0 Then
                    UseKoreanMonth = (Right(t, 1) = ChrW(&HC6D4))
                    Exit Function
                End If
            Else
                If InStr(t, "Week") = 0 Then
                    UseKoreanMonth = (Right(t, 1) = ChrW(&HC6D4))
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

Private Function MonthText(ByVal m As Long, ByVal korean As Boolean) As String
    If korean Then
        MonthText = CStr(m) & ChrW(&HC6D4)
    Else
        MonthText = MonthEn(m)
    End If
End Function

' ws의 col열 값이 label인 행을 모두 삭제 (아래에서 위로). 삭제한 행 수를 반환.
' notAtBottom: 삭제한 주차 아래에 다른 데이터가 더 있었는지
Private Function DeleteRowsWithLabel(ws As Worksheet, ByVal col As Long, ByVal label As String, _
                                     ByVal firstRow As Long, ByRef notAtBottom As Boolean) As Long
    Dim r As Long, lastR As Long, maxMatch As Long, n As Long
    lastR = ws.Cells(ws.rows.Count, col).End(xlUp).Row
    For r = lastR To firstRow Step -1
        If Trim(CStr(ws.Cells(r, col).Value)) = label Then
            If maxMatch = 0 Then maxMatch = r
            ws.rows(r).Delete
            n = n + 1
        End If
    Next r
    notAtBottom = (maxMatch > 0 And maxMatch < lastR)
    DeleteRowsWithLabel = n
End Function

Private Sub Report(ByVal title As String, ByVal msg As String, ByVal style As Long)
    MsgBox msg, style, title
End Sub

' 주문데이터 + 2차주문데이터를 한 번에 읽어 gOrd에 저장
Private Sub LoadOrders()
    Dim names As Variant, k As Long, i As Long, lr As Long
    Dim ws As Worksheet, cID As String, cQty As Long
    names = Array(SH_ORDER1, SH_ORDER2)
    gOrdN = 0
    ReDim gOrd(1 To 1)
    For k = 0 To 1
        Set ws = Sheets(CStr(names(k)))
        lr = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
        For i = 2 To lr
            cID = Trim(CStr(ws.Cells(i, 1).Value))
            cQty = Val(ws.Cells(i, 4).Value)
            If cID <> "" And cQty <> 0 Then
                gOrdN = gOrdN + 1
                ReDim Preserve gOrd(1 To gOrdN)
                gOrd(gOrdN).id = cID
                gOrd(gOrdN).mci = Val(ws.Cells(i, 3).Value)
                gOrd(gOrdN).qty = cQty
                gOrd(gOrdN).is2nd = (k = 1)
            End If
        Next i
    Next k
    Call ClassifySolution
End Sub

' 용액(Solution) 판별 규칙
'   1) 병원명이 PARTNER_CO (예전 방식)
'   2) 규격이 200mCi인 행 (200mCi는 용액에만 쓰임)
'   3) 주문데이터 첫 두 행 중 200mCi가 있으면, 그 두 행의 100mCi도 용액
'      (최근 W39~W41: 첫 두 행이 HOSP11 200mCi, 100mCi)
'   첫 두 행에 200mCi가 없으면(예: W42) 100mCi는 일반 캡슐로 둡니다.
Private Sub ClassifySolution()
    Dim i As Long, cnt As Long, idx(1 To 2) As Long, has200 As Boolean, j As Long
    gSolNote = ""
    For i = 1 To gOrdN
        gOrd(i).isSol = (Norm(gOrd(i).id) = SOLUTION_KEY Or gOrd(i).mci = 200)
        If Not gOrd(i).is2nd And cnt < 2 Then
            cnt = cnt + 1
            idx(cnt) = i
            If gOrd(i).mci = 200 Then has200 = True
        End If
    Next i
    If has200 Then
        For j = 1 To cnt
            If gOrd(idx(j)).mci = 100 Then gOrd(idx(j)).isSol = True
        Next j
    End If
    For i = 1 To gOrdN
        If gOrd(i).isSol Then
            gSolNote = gSolNote & gOrd(i).id & " " & gOrd(i).mci & "mCi " & gOrd(i).qty & "ea; "
        End If
    Next i
    If gSolNote <> "" Then gSolNote = "[용액으로 분류] " & gSolNote & "(맞는지 확인해주세요)"
End Sub

Private Function ExpectedTotal(incl2nd As Boolean, inclSol As Boolean) As Double
    Dim i As Long, t As Double
    For i = 1 To gOrdN
        If incl2nd Or Not gOrd(i).is2nd Then
            If inclSol Or Not gOrd(i).isSol Then t = t + gOrd(i).qty
        End If
    Next i
    ExpectedTotal = t
End Function

'==============================================================
' 실행용 매크로 (Step 1~5를 각각 따로 실행, 실행할 때마다 주차를 입력)
'==============================================================
Sub 캡슐데이터_회사폼_자동생성()
    Dim wk As Long
    wk = AskWeek()
    If wk = 0 Then Exit Sub
    Step1_Run wk
End Sub

Sub Matrix_신규Week_추가()
    Dim wk As Long
    wk = AskWeek()
    If wk = 0 Then Exit Sub
    Matrix_Run wk
End Sub

Sub Step3_PerHospital_추가()
    Dim wk As Long
    wk = AskWeek()
    If wk = 0 Then Exit Sub
    Step3_Run wk
End Sub

Sub Step4_ContainerLog_추가()
    Dim wk As Long
    wk = AskWeek()
    If wk = 0 Then Exit Sub
    Step4_Run wk
End Sub

Sub Step5_Summary_추가()
    Dim wk As Long
    wk = AskWeek()
    If wk = 0 Then Exit Sub
    Step5_Run wk
End Sub

'==============================================================
' Step 1: 회사폼 자동생성 (로직은 기존과 동일, 주차 입력/공용 상수만 변경)
'==============================================================
Private Sub Step1_Run(ByVal wkNum As Long)

    Const STEP_NAME As String = "1. 회사폼"

    Dim ws1 As Worksheet, wsOrder As Worksheet, wsOrder2 As Worksheet
    Dim errMsg As String
    Dim warnList As Collection: Set warnList = New Collection
    Dim buckets As Object: Set buckets = CreateObject("Scripting.Dictionary")
    Dim expTotal As Double
    Set gEa = Nothing

    On Error GoTo ErrH
    Set ws1 = Sheets(SH_FORM)
    Set wsOrder = Sheets(SH_ORDER1)
    Set wsOrder2 = Sheets(SH_ORDER2)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    '---------- 1) 기존 내용 초기화 ----------
    ws1.Range("A:O").UnMerge
    ws1.Range("A:O").FormatConditions.Delete
    ws1.Range("A:O").Clear

    ws1.Cells(1, 1).Value = "Partner_Co"
    ws1.Cells(1, 3).Value = "Week"
    ws1.Cells(1, 5).Value = wkNum

    '---------- 2) 주문 분류 ----------
    Dim i As Long, lastRow As Long
    Dim cID As String, cDay As String, cMci As Double, cQty As Long, isGroupA As Boolean

    lastRow = wsOrder.Cells(wsOrder.rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        cID = Trim(CStr(wsOrder.Cells(i, 1).Value))
        cDay = Trim(CStr(wsOrder.Cells(i, 2).Value))
        cMci = Val(wsOrder.Cells(i, 3).Value)
        cQty = Val(wsOrder.Cells(i, 4).Value)
        If cID = "" Or cQty = 0 Then GoTo NextRow

        If Len(cDay) <> 1 Or InStr(VALID_DAYS, cDay) = 0 Then
            warnList.Add "[요일 오류] 주문데이터 " & i & "행: '" & cDay & "' (" & cID & ")"
            GoTo NextRow
        End If

        isGroupA = InList(GROUP_A_LIST, cID)

        Select Case cMci
            Case 100, 50, 30, 25
                AddId buckets, IIf(isGroupA, "GroupA", "GroupB") & "|" & cDay, cID
                expTotal = expTotal + cQty
            Case 5, 2
                AddId buckets, IIf(isGroupA, "GroupA_Low", "GroupB_Low") & "|" & cDay, cID
                expTotal = expTotal + cQty
            Case 200
                ' Solution 구간에서 처리
            Case 10
                warnList.Add "[10mCi] 표에 없음: " & cDay & " / " & cID & " / " & cQty & "ea"
            Case Else
                warnList.Add "[미분류 mCi " & cMci & "] " & cDay & " / " & cID & " / " & cQty & "ea"
        End Select
NextRow:
    Next i

    Dim lastRow2 As Long
    lastRow2 = wsOrder2.Cells(wsOrder2.rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow2
        cID = Trim(CStr(wsOrder2.Cells(i, 1).Value))
        cDay = Trim(CStr(wsOrder2.Cells(i, 2).Value))
        cMci = Val(wsOrder2.Cells(i, 3).Value)
        cQty = Val(wsOrder2.Cells(i, 4).Value)
        If cID = "" Or cQty = 0 Then GoTo NextRow2

        If Len(cDay) <> 1 Or InStr(VALID_DAYS, cDay) = 0 Then
            warnList.Add "[요일 오류] 2차주문데이터 " & i & "행: '" & cDay & "' (" & cID & ")"
            GoTo NextRow2
        End If
        If Not InList(GROUP_B_LIST, cID) Then
            warnList.Add "[GroupB 목록 없음] 2차주문데이터: " & cDay & " / " & cID
            GoTo NextRow2
        End If
        Select Case cMci
            Case 100, 50, 30, 25
                AddId buckets, "GroupC|" & cDay, cID
                expTotal = expTotal + cQty
            Case Else
                warnList.Add "[2차 mCi " & cMci & " 표에 없음] " & cDay & " / " & cID & " / " & cQty & "ea"
        End Select
NextRow2:
    Next i

    AddId buckets, "SOL|월", "Partner_Co"

    '---------- 3) 위에서부터 차례로 그리기 ----------
    Dim curRow As Long: curRow = 4
    Dim capTot As String, allTot As String, t As String

    t = WriteSection(ws1, curRow, "", "GroupA", "GroupA", "일,토,금,목,수,화,월", Array(100, 50, 30, 25), SH_ORDER1, buckets, 2, "total")
    capTot = t: allTot = t

    t = WriteSection(ws1, curRow, "Capsule-1", "GroupB", "GroupB", "일,토,금,목,수,화,월", Array(100, 50, 30, 25), SH_ORDER1, buckets, 2, "1st")
    capTot = capTot & "," & t: allTot = allTot & "," & t

    t = WriteSection(ws1, curRow, "Low-dose Capsule", "GroupB", "GroupB_Low", "금,목,수,화,월", Array(5, 2), SH_ORDER1, buckets, 2, "1st")
    capTot = capTot & "," & t: allTot = allTot & "," & t

    t = WriteSection(ws1, curRow, "", "GroupA", "GroupA_Low", "금,목,수,화,월", Array(5, 2), SH_ORDER1, buckets, 2, "1st")
    capTot = capTot & "," & t: allTot = allTot & "," & t

    t = WriteSection(ws1, curRow, "Solution", "", "SOL", "월", Array(200, 100), SH_ORDER1, buckets, 1, "1st")
    allTot = allTot & "," & t

    t = WriteSection(ws1, curRow, "Capsule- 2", "GroupC", "GroupC", "토,금,목,수,화", Array(100, 50, 30, 25), SH_ORDER2, buckets, 2, "2nd")
    capTot = capTot & "," & t: allTot = allTot & "," & t

    ws1.Cells(curRow, 9).Value = "Grand Total"
    ws1.Cells(curRow, 11).Formula = "=" & Replace(allTot, ",", "+")
    ws1.Range(ws1.Cells(curRow, 9), ws1.Cells(curRow, 11)).Font.Bold = True

    '---------- 4) 계산 후 노란색 표시 + 합계 검증 ----------
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate

    Dim c As Range
    If Not gEa Is Nothing Then
        For Each c In gEa.Cells
            If VarType(c.Value) = vbDouble Then
                If c.Value > 0 Then c.Interior.Color = RGB(255, 255, 0)
            End If
        Next c
    End If

    Dim actual As Double, p As Variant
    For Each p In Split(capTot, ",")
        actual = actual + Val(ws1.Range(CStr(p)).Value)
    Next p
    If actual <> expTotal Then
        warnList.Add "[합계 불일치] 주문데이터 기준 " & expTotal & " / 표 " & actual
    End If

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    If errMsg <> "" Then
        Report STEP_NAME, errMsg, vbCritical
    ElseIf warnList.Count > 0 Then
        Dim msg As String, w As Variant, n As Long
        For Each w In warnList
            n = n + 1
            If n <= 25 Then msg = msg & w & vbCrLf
        Next w
        If n > 25 Then msg = msg & "... 외 " & (n - 25) & "건"
        Report STEP_NAME, "완료. 아래 항목은 확인이 필요합니다:" & vbCrLf & msg, vbExclamation
    Else
        Report STEP_NAME, "캡슐데이터 자동생성 완료! (주문 수량 전부 일치)", vbInformation
    End If
    Exit Sub

ErrH:
    errMsg = "오류 " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Sub

Function WriteSection(ws As Worksheet, ByRef curRow As Long, title As String, hdr As String, _
                      grp As String, defDays As String, tiers As Variant, sheetName As String, _
                      buckets As Object, minRows As Long, totLabel As String) As String

    Dim nT As Long: nT = UBound(tiers) - LBound(tiers) + 1
    Dim k As Long, r As Long, d As Variant
    Dim key As String, n As Long, rows As Long, hasIds As Boolean

    If title <> "" Then
        ws.Cells(curRow, 2).Value = title
        ws.Cells(curRow, 2).Font.Bold = True
        ws.Cells(curRow, 2).Font.Color = RGB(255, 0, 0)
        curRow = curRow + 1
    End If

    Dim hdrRow As Long: hdrRow = curRow
    If hdr <> "" Then
        ws.Cells(hdrRow, 1).Value = hdr
        ws.Range(ws.Cells(hdrRow, 1), ws.Cells(hdrRow, 3)).Interior.Color = RGB(221, 235, 247)
        ws.Cells(hdrRow, 1).Font.Bold = True
        ws.Cells(hdrRow, 1).Font.Color = RGB(0, 70, 140)
    End If
    For k = 1 To nT
        ws.Cells(hdrRow, 4 + 2 * (k - 1)).Value = "mCi"
        ws.Cells(hdrRow, 5 + 2 * (k - 1)).Value = "ea"
        ws.Cells(hdrRow + k - 1, 12).Value = tiers(LBound(tiers) + k - 1)
    Next k
    ws.Range(ws.Cells(hdrRow, 1), ws.Cells(hdrRow, 3 + 2 * nT)).Font.Bold = True
    ws.Range(ws.Cells(hdrRow, 4), ws.Cells(hdrRow, 3 + 2 * nT)).HorizontalAlignment = xlCenter

    Dim dStart As Long: dStart = hdrRow + 1
    curRow = dStart

    Dim ids As Collection
    For Each d In Split("일,토,금,목,수,화,월", ",")
        key = grp & "|" & CStr(d)
        hasIds = buckets.exists(key)
        If hasIds Or InStr(defDays, CStr(d)) > 0 Then
            If hasIds Then
                Set ids = buckets(key)
                n = ids.Count
            Else
                Set ids = Nothing
                n = 0
            End If
            rows = n
            If rows < minRows Then rows = minRows
            Call WriteBlock(ws, curRow, rows, CStr(d), ids, n, tiers, sheetName)
            curRow = curRow + rows
        End If
    Next d

    Dim dEnd As Long: dEnd = curRow - 1

    ' 오른쪽 요약(L: mCi, M: 합계)
    For k = 1 To nT
        ws.Cells(hdrRow + k - 1, 13).Formula = "=SUM($" & ColL(5 + 2 * (k - 1)) & "$" & dStart & ":$" & ColL(5 + 2 * (k - 1)) & "$" & dEnd & ")"
    Next k
    ws.Range(ws.Cells(hdrRow, 12), ws.Cells(hdrRow + nT - 1, 13)).HorizontalAlignment = xlCenter

    ' 합계 줄
    ws.Cells(curRow, 11).Value = totLabel
    ws.Cells(curRow, 12).Value = "total"
    ws.Cells(curRow, 13).Formula = "=SUM($M$" & hdrRow & ":$M$" & hdrRow + nT - 1 & ")"
    With ws.Range(ws.Cells(curRow, 11), ws.Cells(curRow, 13))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
    End With
    ws.Cells(curRow, 11).Font.Color = RGB(0, 70, 140)
    ws.Cells(curRow, 12).Font.Color = RGB(0, 70, 140)

    WriteSection = ws.Cells(curRow, 13).Address(False, False)
    curRow = curRow + 2
End Function

Sub WriteBlock(ws As Worksheet, top As Long, rows As Long, dayLbl As String, _
               ids As Collection, n As Long, tiers As Variant, sheetName As String)

    Dim nT As Long: nT = UBound(tiers) - LBound(tiers) + 1
    Dim i As Long, k As Long, r As Long, mciCol As Long, eaCol As Long
    Dim lastCol As Long: lastCol = 3 + 2 * nT

    ' 요일/날짜: 병합 후 맨 윗칸에만 입력
    ws.Range(ws.Cells(top, 1), ws.Cells(top + rows - 1, 1)).Merge
    ws.Range(ws.Cells(top, 2), ws.Cells(top + rows - 1, 2)).Merge
    ws.Cells(top, 1).Value = dayLbl
    ws.Cells(top, 2).Formula = "=IFERROR(DATE(YEAR(TODAY()),1,1)+($E$1*7)-WEEKDAY(DATE(YEAR(TODAY()),1,1),2)+MATCH(A" & top & ",{""월"",""화"",""수"",""목"",""금"",""토"",""일""},0),"""")"
    ws.Cells(top, 2).NumberFormat = "mm/dd"
    With ws.Range(ws.Cells(top, 1), ws.Cells(top + rows - 1, 2))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(242, 242, 242)
        .Font.Bold = True
    End With

    For i = 0 To rows - 1
        r = top + i
        If i < n Then
            ws.Cells(r, 3).Value = ids(i + 1)
        Else
            ws.Cells(r, 3).Value = "-"
        End If

        For k = 1 To nT
            mciCol = 4 + 2 * (k - 1)
            eaCol = mciCol + 1
            ws.Cells(r, mciCol).Value = tiers(LBound(tiers) + k - 1)
            ws.Cells(r, mciCol).Interior.Color = RGB(217, 217, 217)

            If i < n Then
                ws.Cells(r, eaCol).Formula = "=SUMIFS('" & sheetName & "'!$D:$D," & _
                    "'" & sheetName & "'!$A:$A,$C" & r & "," & _
                    "'" & sheetName & "'!$B:$B,$A$" & top & "," & _
                    "'" & sheetName & "'!$C:$C,$" & ColL(mciCol) & r & ")"
                ws.Cells(r, eaCol).NumberFormat = "#,##0;-#,##0;""-"""
                If gEa Is Nothing Then
                    Set gEa = ws.Cells(r, eaCol)
                Else
                    Set gEa = Union(gEa, ws.Cells(r, eaCol))
                End If
            Else
                ws.Cells(r, eaCol).Value = "-"
            End If
        Next k
    Next i

    ws.Range(ws.Cells(top, 3), ws.Cells(top + rows - 1, lastCol)).HorizontalAlignment = xlCenter
    With ws.Range(ws.Cells(top, 1), ws.Cells(top + rows - 1, lastCol))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    With ws.Range(ws.Cells(top, 1), ws.Cells(top + rows - 1, lastCol))
        .BorderAround xlContinuous, xlMedium
    End With
End Sub

'==============================================================
' Step 2: Matrix 신규 Week 추가
'   - 2차주문데이터 포함(MATRIX_INCLUDE_2ND)
'   - 병원명 공백/대소문자 무시 매칭
'   - 이전 주차 서식 복사 후 신규/중단 색 초기화
'   - HOSP24 삽입은 주차 입력/중복 확인 이후에 실행
'   - 취소/중복/오류 시 "완료" 메시지가 뜨지 않음
'==============================================================
Private Sub Matrix_Run(ByVal wkNum As Long)

    Const STEP_NAME As String = "2. Matrix"
    Const HDR_ROW As Long = 3
    Const FIRST_DATA_ROW As Long = 4
    Const NEW_COLOR As Long = 5296274
    Const DROP_COLOR As Long = 13551615

    Dim wsM As Worksheet
    Dim diagHdr As String, therHdr As String, solHdr As String
    Dim chkCol As Range, totalCapsCell As Range, fCell As Range, titleRange As Range
    Dim insertCol As Long, prevEndCol As Long, prevStartCol As Long, prevPrefix As String
    Dim titleMergeAddr As String, newLastCol As Long
    Dim dCol As Long, tCol As Long, sCol As Long, xCol As Long
    Dim dColLtr As String, tColLtr As String, sColLtr As String
    Dim subtotalRow As Long, totalRow As Long, lastDataRow As Long
    Dim diagDict As Object, therDict As Object, solDict As Object, nameDict As Object
    Dim matchedKeys As Object, allKeys As Object
    Dim i As Long, r As Long, pc As Long, col As Long
    Dim keyU As String, hosp As String, hospU As String
    Dim d As Double, t As Double, s As Double, prevTotal As Double, thisTotal As Double
    Dim k As Variant
    Dim newHospCount As Long
    Dim sumAll As Double, unclassSum As Double, expected As Double
    Dim msg As String, errMsg As String, skipped As Boolean
    Dim ans As VbMsgBoxResult, overwrite As Boolean

    On Error GoTo ErrH
    Application.ScreenUpdating = False
    Set wsM = Sheets(SH_MATRIX)
    LoadOrders

    diagHdr = "W" & wkNum & vbLf & "Diag"
    therHdr = "W" & wkNum & vbLf & "Ther"
    solHdr = "W" & wkNum & vbLf & "Solution"

    ' ---- 중복 확인 (가장 먼저): 이미 있으면 그 열을 새 값으로 덮어쓸지 묻는다 ----
    Set chkCol = wsM.rows(HDR_ROW).Find(What:=diagHdr, LookAt:=xlWhole, LookIn:=xlValues)
    If Not chkCol Is Nothing Then
        ans = MsgBox("W" & wkNum & " 컬럼이 이미 있습니다. 새로 계산한 값으로 덮어쓸까요?", vbYesNo + vbQuestion, STEP_NAME)
        If ans = vbNo Then
            skipped = True
            msg = "W" & wkNum & " 컬럼이 이미 있어 아무것도 바꾸지 않았습니다."
            GoTo CleanUp
        End If
        overwrite = True
        dCol = chkCol.Column
        If CStr(wsM.Cells(HDR_ROW, dCol + 1).Value) <> therHdr Or CStr(wsM.Cells(HDR_ROW, dCol + 2).Value) <> solHdr Then
            errMsg = "W" & wkNum & " 컬럼이 Diag/Ther/Solution 3열 구성이 아니라서 덮어쓸 수 없습니다. 직접 지우고 다시 실행해주세요."
            GoTo CleanUp
        End If
    End If

    ' ---- HOSP24 행 (이미 있으면 아무 일도 하지 않음) ----
    Call InsertHospitalAtPosition(wsM, "HOSP24", 19, HDR_ROW, FIRST_DATA_ROW)

    ' ---- 삽입 위치(Total 열)와 직전 주차 열 범위 ----
    Set totalCapsCell = wsM.rows(HDR_ROW).Find(What:="Total", LookAt:=xlPart, LookIn:=xlValues, MatchCase:=False)
    If totalCapsCell Is Nothing Then
        errMsg = "'Total' 헤더를 찾을 수 없어 중단합니다."
        GoTo CleanUp
    End If
    If overwrite Then
        insertCol = dCol
    Else
        insertCol = totalCapsCell.Column
    End If

    prevEndCol = insertCol - 1
    prevStartCol = prevEndCol
    prevPrefix = Split(CStr(wsM.Cells(HDR_ROW, prevEndCol).Value), vbLf)(0)
    Do While prevStartCol > 1
        If Split(CStr(wsM.Cells(HDR_ROW, prevStartCol - 1).Value), vbLf)(0) = prevPrefix Then
            prevStartCol = prevStartCol - 1
        Else
            Exit Do
        End If
    Loop

    If Not overwrite Then
        ' ---- 제목 병합 해제 -> 열 삽입 -> 서식 복사 -> 제목 재병합 ----
        Set titleRange = wsM.Range("A1")
        titleMergeAddr = titleRange.MergeArea.Address
        titleRange.MergeArea.UnMerge

        wsM.Range(wsM.Cells(1, insertCol), wsM.Cells(1, insertCol + 2)).EntireColumn.Insert Shift:=xlToRight

        wsM.Range(wsM.Cells(1, prevStartCol), wsM.Cells(wsM.UsedRange.Row + wsM.UsedRange.rows.Count - 1, prevEndCol)).Copy
        wsM.Cells(1, insertCol).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False

        newLastCol = wsM.Range(titleMergeAddr).Column + wsM.Range(titleMergeAddr).Columns.Count - 1 + 3
        wsM.Range(wsM.Cells(1, 1), wsM.Cells(1, newLastCol)).Merge
        wsM.Cells(1, 1).Value = titleRange.Value

        dCol = insertCol: tCol = insertCol + 1: sCol = insertCol + 2
        wsM.Cells(HDR_ROW, dCol).Value = diagHdr
        wsM.Cells(HDR_ROW, tCol).Value = therHdr
        wsM.Cells(HDR_ROW, sCol).Value = solHdr
    End If
    tCol = dCol + 1: sCol = dCol + 2

    ' ---- Sub total / Total 행 찾기 ----
    Set fCell = wsM.Cells.Find(What:="Sub total", LookAt:=xlPart, LookIn:=xlValues, MatchCase:=False)
    If fCell Is Nothing Then
        errMsg = "'Sub total' 행을 찾지 못했습니다. 구조를 확인해주세요."
        GoTo CleanUp
    End If
    subtotalRow = fCell.Row
    totalRow = subtotalRow + 1
    lastDataRow = subtotalRow - 1

    Set totalCapsCell = wsM.rows(HDR_ROW).Find(What:="Total", LookAt:=xlPart, LookIn:=xlValues, MatchCase:=False)
    xCol = totalCapsCell.Column
    dColLtr = ColL(dCol): tColLtr = ColL(tCol): sColLtr = ColL(sCol)

    ' ---- 직전 주차에서 복사된 신규(연두)/중단(빨강) 색 제거 ----
    For r = FIRST_DATA_ROW To lastDataRow
        For col = dCol To sCol
            If wsM.Cells(r, col).Interior.Color = NEW_COLOR Or wsM.Cells(r, col).Interior.Color = DROP_COLOR Then
                wsM.Cells(r, col).Interior.ColorIndex = xlNone
            End If
        Next col
    Next r

    If gSolNote <> "" Then msg = msg & gSolNote & vbCrLf
    ' ---- 주문 집계 (병원명 공백/대소문자 무시) ----
    Set diagDict = CreateObject("Scripting.Dictionary")
    Set therDict = CreateObject("Scripting.Dictionary")
    Set solDict = CreateObject("Scripting.Dictionary")
    Set nameDict = CreateObject("Scripting.Dictionary")

    For i = 1 To gOrdN
        If gOrd(i).is2nd And Not MATRIX_INCLUDE_2ND Then GoTo NextOrd
        keyU = Norm(gOrd(i).id)
        If Not nameDict.exists(keyU) Then nameDict(keyU) = gOrd(i).id

        If gOrd(i).isSol Then
            AddQty solDict, keyU, gOrd(i).qty
        Else
            Select Case gOrd(i).mci
                Case 2, 5, 10
                    AddQty diagDict, keyU, gOrd(i).qty
                Case 25, 30, 50, 100, 200
                    AddQty therDict, keyU, gOrd(i).qty
                Case Else
                    msg = msg & "미분류 mCi(" & gOrd(i).mci & "): " & gOrd(i).id & " / " & gOrd(i).qty & "ea" & vbCrLf
                    unclassSum = unclassSum + gOrd(i).qty
            End Select
        End If
NextOrd:
    Next i

    ' ---- 기존 병원 행 채우기 ----
    Set matchedKeys = CreateObject("Scripting.Dictionary")

    For r = FIRST_DATA_ROW To lastDataRow
        hosp = Trim(CStr(wsM.Cells(r, 2).Value))
        If hosp <> "" Then
            hospU = Norm(hosp)
            If matchedKeys.exists(hospU) Then
                msg = msg & "[중복 병원명] Matrix " & r & "행: " & hosp & " (수량은 첫 행에만 반영)" & vbCrLf
                d = 0: t = 0: s = 0
            Else
                matchedKeys(hospU) = True
                d = GetQty(diagDict, hospU)
                t = GetQty(therDict, hospU)
                s = GetQty(solDict, hospU)
            End If

            wsM.Cells(r, dCol).Value = IIf(d = 0, "-", d)
            wsM.Cells(r, tCol).Value = IIf(t = 0, "-", t)
            wsM.Cells(r, sCol).Value = IIf(s = 0, "-", s)

            prevTotal = 0
            For pc = prevStartCol To prevEndCol
                prevTotal = prevTotal + Val(wsM.Cells(r, pc).Value)
            Next pc
            thisTotal = d + t + s
            sumAll = sumAll + thisTotal

            If thisTotal > 0 And prevTotal = 0 Then
                wsM.Range(wsM.Cells(r, dCol), wsM.Cells(r, sCol)).Interior.Color = NEW_COLOR
                wsM.Cells(r, 2).Interior.Color = NEW_COLOR
            ElseIf thisTotal = 0 And prevTotal > 0 Then
                wsM.Range(wsM.Cells(r, dCol), wsM.Cells(r, sCol)).Interior.Color = DROP_COLOR
                wsM.Cells(r, 2).Interior.Color = DROP_COLOR
            End If

            wsM.Cells(r, xCol).Formula = "=SUM(C" & r & ":" & sColLtr & r & ")"
        End If
    Next r

    ' ---- Matrix에 없는 신규 병원 추가 ----
    Set allKeys = CreateObject("Scripting.Dictionary")
    For Each k In diagDict.Keys: allKeys(k) = True: Next k
    For Each k In therDict.Keys: allKeys(k) = True: Next k
    For Each k In solDict.Keys: allKeys(k) = True: Next k

    For Each k In allKeys.Keys
        If Not matchedKeys.exists(CStr(k)) Then
            newHospCount = newHospCount + 1
            wsM.rows(lastDataRow + 1).Insert Shift:=xlDown
            wsM.rows(lastDataRow).Copy
            wsM.rows(lastDataRow + 1).PasteSpecial xlPasteFormats
            Application.CutCopyMode = False

            wsM.Cells(lastDataRow + 1, 1).Value = wsM.Cells(lastDataRow, 1).Value + 1
            wsM.Cells(lastDataRow + 1, 2).Value = nameDict(CStr(k))
            For col = 3 To xCol - 1
                If col < dCol Or col > sCol Then wsM.Cells(lastDataRow + 1, col).Value = "-"
            Next col

            d = GetQty(diagDict, CStr(k))
            t = GetQty(therDict, CStr(k))
            s = GetQty(solDict, CStr(k))
            wsM.Cells(lastDataRow + 1, dCol).Value = IIf(d = 0, "-", d)
            wsM.Cells(lastDataRow + 1, tCol).Value = IIf(t = 0, "-", t)
            wsM.Cells(lastDataRow + 1, sCol).Value = IIf(s = 0, "-", s)
            sumAll = sumAll + d + t + s
            wsM.Cells(lastDataRow + 1, xCol).Formula = "=SUM(C" & (lastDataRow + 1) & ":" & sColLtr & (lastDataRow + 1) & ")"

            wsM.Range(wsM.Cells(lastDataRow + 1, dCol), wsM.Cells(lastDataRow + 1, sCol)).Interior.Color = NEW_COLOR
            wsM.Cells(lastDataRow + 1, 2).Interior.Color = NEW_COLOR

            lastDataRow = lastDataRow + 1
            subtotalRow = subtotalRow + 1
            totalRow = totalRow + 1
        End If
    Next k

    ' ---- Sub total / Total ----
    wsM.Cells(subtotalRow, dCol).Formula = "=SUM(" & dColLtr & FIRST_DATA_ROW & ":" & dColLtr & lastDataRow & ")"
    wsM.Cells(subtotalRow, tCol).Formula = "=SUM(" & tColLtr & FIRST_DATA_ROW & ":" & tColLtr & lastDataRow & ")"
    wsM.Cells(subtotalRow, sCol).Formula = "=SUM(" & sColLtr & FIRST_DATA_ROW & ":" & sColLtr & lastDataRow & ")"
    wsM.Cells(subtotalRow, xCol).Formula = "=SUM(C" & subtotalRow & ":" & sColLtr & subtotalRow & ")"

    wsM.Range(wsM.Cells(totalRow, dCol), wsM.Cells(totalRow, sCol)).Merge
    wsM.Cells(totalRow, dCol).Formula = "=SUM(" & dColLtr & subtotalRow & ":" & sColLtr & subtotalRow & ")"
    wsM.Cells(totalRow, xCol).Formula = "=SUM(C" & totalRow & ":" & sColLtr & totalRow & ")"

    ' ---- 합계 검증 ----
    expected = ExpectedTotal(MATRIX_INCLUDE_2ND, True) - unclassSum
    If sumAll <> expected Then
        msg = msg & "[합계 불일치] 주문데이터 기준 " & expected & " / Matrix 반영 " & sumAll & vbCrLf
    End If

CleanUp:
    Application.CutCopyMode = False
    Application.ScreenUpdating = True

    If errMsg <> "" Then
        Report STEP_NAME, errMsg, vbCritical
    ElseIf skipped Then
        Report STEP_NAME, msg, vbExclamation
    Else
        If newHospCount > 0 Then msg = "신규 병원 " & newHospCount & "건 자동 추가(연두색)." & vbCrLf & msg
        If msg = "" Then
            Report STEP_NAME, "W" & wkNum & IIf(overwrite, " 열 덮어쓰기 완료!", " 열 추가 완료!"), vbInformation
        Else
            Report STEP_NAME, "W" & wkNum & IIf(overwrite, " 열 덮어쓰기 완료.", " 열 추가 완료.") & vbCrLf & vbCrLf & msg, vbExclamation
        End If
    End If
    Exit Sub

ErrH:
    errMsg = "오류 " & Err.Number & ": " & Err.Description
    Resume CleanUp
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

    ws.rows(targetRow).Insert Shift:=xlDown
    ws.rows(targetRow + 1).Copy
    ws.rows(targetRow).PasteSpecial xlPasteFormats
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

'==============================================================
' Step 3: Per hospital 추가
'==============================================================
Private Sub Step3_Run(ByVal wkNum As Long)

    Const STEP_NAME As String = "3. Per hospital"

    Dim wsPH As Worksheet
    Dim wkLabel As String, monthLabel As String
    Dim weekSun As Date
    Dim dupCheck As Range, ans As VbMsgBoxResult
    Dim shipOrder As Variant, s As Variant, h As Variant
    Dim hospOrder As Object, mciDict As Object, qtyDict As Object, nameDict As Object
    Dim i As Long, shipKey As String, hk As String
    Dim startRow As Long, r As Long, totalRowsAdded As Long
    Dim groupStart As Long, groupEnd As Long, mi As Long
    Dim hospCol As Collection, mciCol As Collection
    Dim mciArr() As Double
    Dim mciKey As String, qKey As String, mciVal As Double, qty As Double
    Dim sumWritten As Double, expected As Double
    Dim msg As String, errMsg As String, pivotMsg As String, skipped As Boolean
    Dim delCount As Long, notBottom As Boolean

    On Error GoTo ErrH
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Set wsPH = Sheets(SH_PERHOSP)
    LoadOrders

    wkLabel = "Week " & Format(wkNum, "00")

    Set dupCheck = wsPH.Columns(3).Find(What:=wkLabel, LookAt:=xlWhole, LookIn:=xlValues)
    If Not dupCheck Is Nothing Then
        ans = MsgBox(wkLabel & " 데이터가 이미 있습니다. 기존 행을 지우고 새로 입력할까요?", vbYesNo + vbQuestion, STEP_NAME)
        If ans = vbNo Then
            skipped = True
            msg = wkLabel & " 데이터가 이미 있어 아무것도 바꾸지 않았습니다."
            GoTo CleanUp
        End If
        delCount = DeleteRowsWithLabel(wsPH, 3, wkLabel, 4, notBottom)
        msg = msg & "[덮어쓰기] 기존 " & wkLabel & " " & delCount & "행을 지우고 다시 입력했습니다." & vbCrLf
        If notBottom Then msg = msg & "[순서] 이 주차 아래에 다른 주차 데이터가 있어 새 행은 맨 아래에 붙습니다." & vbCrLf
    End If

    weekSun = WeekMonday(wkNum) + 6
    monthLabel = MonthText(Month(weekSun), UseKoreanMonth(wsPH, 2))

    shipOrder = Array("GroupB- 1st", "GroupA- 1st", "GroupB- 2nd")
    Set hospOrder = CreateObject("Scripting.Dictionary")
    Set mciDict = CreateObject("Scripting.Dictionary")
    Set qtyDict = CreateObject("Scripting.Dictionary")
    Set nameDict = CreateObject("Scripting.Dictionary")

    For Each s In shipOrder
        Set hospOrder(s) = New Collection
    Next s

    For i = 1 To gOrdN
        If gOrd(i).is2nd Then
            shipKey = "GroupB- 2nd"
            If Not InList(GROUP_B_LIST, gOrd(i).id) Then
                If InStr(msg, "[GroupB 목록 없음] " & gOrd(i).id & vbCrLf) = 0 Then
                    msg = msg & "[GroupB 목록 없음] " & gOrd(i).id & vbCrLf
                End If
            End If
        ElseIf InList(GROUP_A_LIST, gOrd(i).id) Then
            shipKey = "GroupA- 1st"
        Else
            shipKey = "GroupB- 1st"
        End If
        hk = Norm(gOrd(i).id)
        If Not nameDict.exists(hk) Then nameDict(hk) = gOrd(i).id
        Call Accumulate(hospOrder, mciDict, qtyDict, shipKey, hk, gOrd(i).mci, gOrd(i).qty)
    Next i

    startRow = wsPH.Cells(wsPH.rows.Count, 2).End(xlUp).Row + 1
    r = startRow

    For Each s In shipOrder
        Set hospCol = hospOrder(s)
        If hospCol.Count > 0 Then
            groupStart = r
            For Each h In hospCol
                mciKey = CStr(s) & "|" & CStr(h)
                Set mciCol = mciDict(mciKey)
                Call SortMciDescending(mciCol, mciArr)

                For mi = LBound(mciArr) To UBound(mciArr)
                    mciVal = mciArr(mi)
                    qKey = mciKey & "|" & mciVal
                    qty = qtyDict(qKey)

                    wsPH.Cells(r, 2).Value = monthLabel
                    wsPH.Cells(r, 3).Value = wkLabel
                    wsPH.Cells(r, 4).Value = s
                    wsPH.Cells(r, 5).Value = nameDict(CStr(h))
                    wsPH.Cells(r, 6).Value = FormatMci(mciVal)
                    wsPH.Cells(r, 7).Value = qty
                    sumWritten = sumWritten + qty
                    r = r + 1
                Next mi
            Next h

            groupEnd = r - 1
            wsPH.Cells(groupStart, 8).Formula = "=SUM(G" & groupStart & ":G" & groupEnd & ")"
            totalRowsAdded = totalRowsAdded + (groupEnd - groupStart + 1)
        End If
    Next s

    expected = ExpectedTotal(True, True)
    If sumWritten <> expected Then
        msg = msg & "[합계 불일치] 주문데이터 기준 " & expected & " / Per hospital 반영 " & sumWritten & vbCrLf
    End If

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    If errMsg <> "" Then
        Report STEP_NAME, errMsg, vbCritical
    ElseIf skipped Then
        Report STEP_NAME, msg, vbExclamation
    ElseIf totalRowsAdded = 0 Then
        Report STEP_NAME, wkLabel & "에 해당하는 주문데이터가 없어 추가된 행이 없습니다.", vbExclamation
    Else
        pivotMsg = RefreshPivot()
        If pivotMsg <> "" Then msg = msg & pivotMsg & vbCrLf
        If msg = "" Then
            Report STEP_NAME, wkLabel & " (" & monthLabel & ") " & totalRowsAdded & "행 추가 완료 (" & startRow & "~" & (r - 1) & "행), 피벗 갱신 완료", vbInformation
        Else
            Report STEP_NAME, wkLabel & " (" & monthLabel & ") " & totalRowsAdded & "행 추가 완료 (" & startRow & "~" & (r - 1) & "행)." & vbCrLf & msg, vbExclamation
        End If
    End If
    Exit Sub

ErrH:
    errMsg = "오류 " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Sub

' 피벗 소스 범위를 Per hospital 마지막 행까지 늘리고 갱신 (오류 시 메시지 반환, 정상이면 "")
Private Function RefreshPivot() As String
    Dim wsPH As Worksheet, pt As PivotTable, lastRow As Long
    On Error GoTo ErrH
    Set wsPH = Sheets(SH_PERHOSP)
    lastRow = wsPH.Cells(wsPH.rows.Count, 2).End(xlUp).Row
    Set pt = Sheets(SH_PIVOT).PivotTables(1)
    pt.ChangePivotCache ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:="'" & SH_PERHOSP & "'!B3:H" & lastRow)
    pt.RefreshTable
    RefreshPivot = ""
    Exit Function
ErrH:
    RefreshPivot = "[피벗 갱신 실패] " & Err.Description & " (Step3b를 따로 실행해 확인해주세요)"
End Function

Sub Step3b_PerHospital_Pivot_갱신()
    Dim m As String
    m = RefreshPivot()
    If m = "" Then
        Report "3b. 피벗", "Summary Pivot 갱신 완료!", vbInformation
    Else
        Report "3b. 피벗", m, vbCritical
    End If
End Sub

Sub Accumulate(hospOrder As Object, mciDict As Object, qtyDict As Object, _
               ByVal shipKey As String, ByVal hospID As String, ByVal mciVal As Double, ByVal qty As Double)
    Dim hospCol As Collection: Set hospCol = hospOrder(shipKey)
    Dim found As Boolean: found = False
    Dim v As Variant
    For Each v In hospCol
        If CStr(v) = hospID Then
            found = True
            Exit For
        End If
    Next v
    If Not found Then hospCol.Add hospID

    Dim mciKey As String: mciKey = shipKey & "|" & hospID
    If Not mciDict.exists(mciKey) Then Set mciDict(mciKey) = New Collection
    Dim mciCol As Collection: Set mciCol = mciDict(mciKey)
    Dim mciFound As Boolean: mciFound = False
    For Each v In mciCol
        If CDbl(v) = mciVal Then
            mciFound = True
            Exit For
        End If
    Next v
    If Not mciFound Then mciCol.Add mciVal

    Dim qKey As String: qKey = mciKey & "|" & mciVal
    AddQty qtyDict, qKey, qty
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

'==============================================================
' Step 4: Container Log 추가
'   시트 구조(2026 Containers+면장) 기준:
'   - 주차 라벨은 "Week 37" 형식, 월 라벨은 영문 대문자(첫 주차에만, 병합)
'   - 선적 행은 최근 방식대로 6종으로 분리
'       GroupB-01- high dose caps / GroupB-01- diagnosis caps / GroupB-01-sol /
'       GroupA-01- high dose caps / GroupA-01- diagnosis caps / GroupB-02- caps
'   - 마지막 데이터 행은 C열 기준, "TOTAL" 행 위에 쓰고 모자라면 행을 삽입
'==============================================================
Private Sub Step4_Run(ByVal wkNum As Long)

    Const STEP_NAME As String = "4. Container Log"

    Dim wsCL As Worksheet
    Dim wkLabel As String, monthLabel As String
    Dim weekSun As Date
    Dim dupCheck As Range, ans As VbMsgBoxResult
    Dim labels As Variant
    Dim qtys(1 To 6) As Double
    Dim i As Long, k As Long, r As Long, kind As Long
    Dim totalRow As Long, lastDataRow As Long, startRow As Long, endRow As Long
    Dim nRows As Long, avail As Long, need As Long, c As Long
    Dim sumQty As Double, expected As Double, unclass As Double
    Dim isH As Boolean
    Dim msg As String, errMsg As String, skipped As Boolean
    Dim overwrite As Boolean, blkR1 As Long, blkR2 As Long, oldRows As Long, extra As Long
    Dim doneWord As String

    On Error GoTo ErrH
    Application.ScreenUpdating = False
    Set wsCL = Sheets(SH_CONTAINER)
    LoadOrders

    wkLabel = "Week " & wkNum

    ' 중복 확인: "Week 37" 형식과 예전 "01" 형식 모두 확인
    Set dupCheck = wsCL.Columns(2).Find(What:=wkLabel, LookAt:=xlWhole, LookIn:=xlValues)
    If dupCheck Is Nothing Then
        Set dupCheck = wsCL.Columns(2).Find(What:=Format(wkNum, "00"), LookAt:=xlWhole, LookIn:=xlValues)
    End If
    If Not dupCheck Is Nothing Then
        ans = MsgBox(wkLabel & " 데이터가 이미 있습니다. 선적 행과 수량을 새로 계산한 값으로 바꿀까요?" & vbCrLf & _
                     "(E열 컨테이너 수, G열 면장은 그대로 남습니다)", vbYesNo + vbQuestion, STEP_NAME)
        If ans = vbNo Then
            skipped = True
            msg = wkLabel & " 데이터가 이미 있어 아무것도 바꾸지 않았습니다."
            GoTo CleanUp
        End If
        overwrite = True
        blkR1 = dupCheck.Row
    End If

    weekSun = WeekMonday(wkNum) + 6
    monthLabel = UCase(MonthEn(Month(weekSun)))

    labels = Array("GroupB-01- high dose caps", "GroupB-01- diagnosis caps", "GroupB-01-sol", _
                   "GroupA-01- high dose caps", "GroupA-01- diagnosis caps", "GroupB-02- caps")

    If gSolNote <> "" Then msg = msg & gSolNote & vbCrLf
    ' ---- 분류 ----
    For i = 1 To gOrdN
        kind = 0
        If gOrd(i).is2nd Then
            kind = 6
        ElseIf gOrd(i).isSol Then
            kind = 3
        Else
            isH = InList(GROUP_A_LIST, gOrd(i).id)
            Select Case gOrd(i).mci
                Case 25, 30, 50, 100, 200
                    kind = IIf(isH, 4, 1)
                Case 2, 5, 10
                    kind = IIf(isH, 5, 2)
            End Select
        End If
        If kind = 0 Then
            msg = msg & "[미분류 mCi " & gOrd(i).mci & "] 제외됨: " & gOrd(i).id & " / " & gOrd(i).qty & "ea" & vbCrLf
            unclass = unclass + gOrd(i).qty
        Else
            qtys(kind) = qtys(kind) + gOrd(i).qty
        End If
    Next i

    For k = 1 To 6
        If qtys(k) > 0 Then nRows = nRows + 1
    Next k
    If nRows = 0 Then GoTo CleanUp

    ' ---- TOTAL 행과 마지막 데이터 행 찾기 ----
    totalRow = 0
    For r = wsCL.Cells(wsCL.rows.Count, 3).End(xlUp).Row To 1 Step -1
        If UCase(Trim(CStr(wsCL.Cells(r, 3).Value))) = "TOTAL" Then
            totalRow = r
            Exit For
        End If
    Next r
    If totalRow = 0 Then
        errMsg = "C열에서 'TOTAL' 행을 찾지 못했습니다. 시트 구조를 확인해주세요."
        GoTo CleanUp
    End If

    lastDataRow = 5
    For r = totalRow - 1 To 6 Step -1
        If Trim(CStr(wsCL.Cells(r, 3).Value)) <> "" Then
            lastDataRow = r
            Exit For
        End If
    Next r
    startRow = lastDataRow + 1

    If overwrite Then
        ' ---- 덮어쓰기: 기존 주차 블록(blkR1~blkR2)을 찾아 C:D를 비우고 행 수를 맞춘다 ----
        blkR2 = blkR1
        Do While blkR2 + 1 < totalRow
            If Trim(CStr(wsCL.Cells(blkR2 + 1, 2).Value)) = "" And Trim(CStr(wsCL.Cells(blkR2 + 1, 3).Value)) <> "" Then
                blkR2 = blkR2 + 1
            Else
                Exit Do
            End If
        Loop
        oldRows = blkR2 - blkR1 + 1
        wsCL.Range(wsCL.Cells(blkR1, 3), wsCL.Cells(blkR2, 4)).ClearContents
        If nRows < oldRows Then
            wsCL.rows(blkR1 + nRows & ":" & blkR2).Delete
            totalRow = totalRow - (oldRows - nRows)
        ElseIf nRows > oldRows Then
            extra = nRows - oldRows
            wsCL.rows(blkR2 + 1).Resize(extra).Insert Shift:=xlDown
            totalRow = totalRow + extra
        End If
        If nRows <> oldRows Then
            For c = 4 To 7
                wsCL.Cells(totalRow, c).Formula = "=SUM(" & ColL(CLng(c)) & "6:" & ColL(CLng(c)) & (totalRow - 1) & ")"
            Next c
        End If
        startRow = blkR1
    Else
        ' ---- 빈 행이 모자라면 TOTAL 위에 행 삽입 후 TOTAL 합계 범위 재설정 ----
        avail = totalRow - startRow
        If nRows > avail Then
            need = nRows - avail
            wsCL.rows(totalRow).Resize(need).Insert Shift:=xlDown
            totalRow = totalRow + need
            For c = 4 To 7
                wsCL.Cells(totalRow, c).Formula = "=SUM(" & ColL(CLng(c)) & "6:" & ColL(CLng(c)) & (totalRow - 1) & ")"
            Next c
        End If
    End If

    ' ---- 값 쓰기 ----
    r = startRow
    For k = 1 To 6
        If qtys(k) > 0 Then
            wsCL.Cells(r, 3).Value = labels(k - 1)
            wsCL.Cells(r, 4).Value = qtys(k)
            sumQty = sumQty + qtys(k)
            r = r + 1
        End If
    Next k
    endRow = r - 1

    wsCL.Cells(startRow, 2).Value = wkLabel
    wsCL.Cells(startRow, 6).Formula = "=SUM(E" & startRow & ":E" & endRow & ")"

    ' ---- 병합/테두리/월 라벨 (실패해도 값은 이미 들어간 상태) ----
    Call Container_Layout(wsCL, startRow, endRow, monthLabel, overwrite)

    expected = ExpectedTotal(True, True) - unclass
    If sumQty <> expected Then
        msg = msg & "[합계 불일치] 주문데이터 기준 " & expected & " / Container Log 반영 " & sumQty & vbCrLf
    End If

CleanUp:
    Application.ScreenUpdating = True

    If errMsg <> "" Then
        Report STEP_NAME, errMsg, vbCritical
    ElseIf skipped Then
        Report STEP_NAME, msg, vbExclamation
    ElseIf nRows = 0 Then
        Report STEP_NAME, wkLabel & "에 해당하는 주문데이터가 없어 추가된 행이 없습니다.", vbExclamation
    Else
        If overwrite Then doneWord = "덮어쓰기" Else doneWord = "추가"
        Report STEP_NAME, wkLabel & " (" & monthLabel & ") " & nRows & "개 선적 행 " & doneWord & " 완료 (" & startRow & "~" & endRow & "행)." & vbCrLf & _
               "E열(컨테이너 수)와 G열(면장)은 직접 입력해주세요. 병합/테두리가 맞는지 눈으로 확인해주세요." & vbCrLf & msg, _
               IIf(msg = "", vbInformation, vbExclamation)
    End If
    Exit Sub

ErrH:
    errMsg = "오류 " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Sub

' 주차 블록의 병합(B, F열), 테두리, 월 라벨(A열 병합)을 기존 모양에 맞춰 적용
Private Sub Container_Layout(ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, ByVal monthLabel As String, ByVal keepMonth As Boolean)
    Dim r As Long, anchorRow As Long, topRow As Long, bottomRow As Long
    Dim lastMonth As String, ma As Range, startR As Long

    On Error Resume Next
    Application.DisplayAlerts = False

    ' 주차(B), 주간 합계(F) 병합
    If r2 > r1 Then
        ws.Range(ws.Cells(r1, 2), ws.Cells(r2, 2)).Merge
        ws.Range(ws.Cells(r1, 6), ws.Cells(r2, 6)).Merge
    End If
    ws.Range(ws.Cells(r1, 2), ws.Cells(r2, 2)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(r1, 2), ws.Cells(r2, 2)).VerticalAlignment = xlCenter
    ws.Range(ws.Cells(r1, 6), ws.Cells(r2, 6)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(r1, 6), ws.Cells(r2, 6)).VerticalAlignment = xlCenter

    ' 테두리: 안쪽 얇게, 주차 블록 바깥 굵게
    With ws.Range(ws.Cells(r1, 2), ws.Cells(r2, 7))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .BorderAround xlContinuous, xlMedium
    End With

    ' 월 라벨(A열): 마지막으로 라벨이 있는 행을 찾아 같은 달이면 병합 범위를 늘리고, 달이 바뀌면 새로 씀
    anchorRow = 0
    If keepMonth Then startR = r1 Else startR = r1 - 1
    For r = startR To 6 Step -1
        If Trim(CStr(ws.Cells(r, 1).Value)) <> "" Then
            anchorRow = r
            Exit For
        End If
    Next r
    If anchorRow > 0 Then lastMonth = UCase(Trim(CStr(ws.Cells(anchorRow, 1).Value)))

    If anchorRow > 0 And (keepMonth Or lastMonth = monthLabel) Then
        Set ma = ws.Cells(anchorRow, 1).MergeArea
        topRow = ma.Row
        bottomRow = ma.Row + ma.rows.Count - 1
        If bottomRow < r2 Then bottomRow = r2
        ma.UnMerge
        ws.Range(ws.Cells(topRow, 1), ws.Cells(bottomRow, 1)).Merge
    Else
        ' 새 달: 이전 달 병합이 새 행까지 걸쳐 있으면 r1-1까지로 줄이고 새 라벨을 씀
        If anchorRow > 0 Then
            Set ma = ws.Cells(anchorRow, 1).MergeArea
            topRow = ma.Row
            If ma.Row + ma.rows.Count - 1 >= r1 Then
                ma.UnMerge
                If r1 - 1 > topRow Then ws.Range(ws.Cells(topRow, 1), ws.Cells(r1 - 1, 1)).Merge
            End If
        End If
        ws.Cells(r1, 1).Value = monthLabel
        If r2 > r1 Then ws.Range(ws.Cells(r1, 1), ws.Cells(r2, 1)).Merge
        ws.Range(ws.Cells(r1, 1), ws.Cells(r2, 1)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(r1, 1), ws.Cells(r2, 1)).VerticalAlignment = xlCenter
    End If

    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

'==============================================================
' Step 5: Summary 추가
'   - 직전 블록의 서식 복사 (직전 블록 헤더가 "Week "로 시작할 때만)
'   - 10mCi 등 미분류 수량은 경고로 표시
'==============================================================
Private Sub Step5_Run(ByVal wkNum As Long)

    Const STEP_NAME As String = "5. Summary"

    Dim wsSum As Worksheet
    Dim wkShort As String, monthLabel As String, wkLabelFull As String
    Dim weekSun As Date
    Dim dupCheck As Range
    Dim mci25 As Double, mciHighOthers As Double, mci5 As Double, mci2 As Double, solutionQty As Double
    Dim i As Long, r As Long
    Dim lastAllShipRow As Long, hdrRow As Long, r25 As Long, r5 As Long, r2 As Long, rSol As Long, rAll As Long
    Dim prevHdr As Long, fmtCopied As Boolean
    Dim msg As String, errMsg As String, skipped As Boolean, written As Boolean
    Dim ans As VbMsgBoxResult, overwrite As Boolean

    On Error GoTo ErrH
    Application.ScreenUpdating = False
    Set wsSum = Sheets(SH_SUMMARY)
    LoadOrders

    wkShort = "Week " & wkNum
    weekSun = WeekMonday(wkNum) + 6
    monthLabel = UCase(MonthText(Month(weekSun), UseKoreanMonth(wsSum, 3)))
    wkLabelFull = wkShort & "- " & monthLabel

    ' "Week 40- ..." 로 시작하는 블록이 있는지 확인 (월 라벨 표기가 달라도 찾도록 앞부분만 비교)
    Set dupCheck = wsSum.Columns(3).Find(What:=wkShort & "- ", LookAt:=xlPart, LookIn:=xlValues)
    If Not dupCheck Is Nothing Then
        ans = MsgBox(Trim(CStr(dupCheck.Value)) & " 블록이 이미 있습니다. 새로 계산한 값으로 덮어쓸까요?", vbYesNo + vbQuestion, STEP_NAME)
        If ans = vbNo Then
            skipped = True
            msg = Trim(CStr(dupCheck.Value)) & " 블록이 이미 있어 아무것도 바꾸지 않았습니다."
            GoTo CleanUp
        End If
        overwrite = True
    End If

    If gSolNote <> "" Then msg = msg & gSolNote & vbCrLf
    For i = 1 To gOrdN
        If gOrd(i).isSol Then
            solutionQty = solutionQty + gOrd(i).qty
        Else
            Select Case gOrd(i).mci
                Case 25
                    mci25 = mci25 + gOrd(i).qty
                Case 30, 50, 100, 200
                    mciHighOthers = mciHighOthers + gOrd(i).qty
                Case 5
                    mci5 = mci5 + gOrd(i).qty
                Case 2
                    mci2 = mci2 + gOrd(i).qty
                Case 10
                    msg = msg & "[10mCi] Summary 블록에 칸이 없어 제외됨: " & gOrd(i).id & " / " & gOrd(i).qty & "ea" & vbCrLf
                Case Else
                    msg = msg & "[미분류 mCi " & gOrd(i).mci & "] 제외됨: " & gOrd(i).id & " / " & gOrd(i).qty & "ea" & vbCrLf
            End Select
        End If
    Next i

    If overwrite Then
        ' 기존 블록 위치를 그대로 사용 (서식은 건드리지 않음)
        hdrRow = dupCheck.Row
        wkLabelFull = Trim(CStr(dupCheck.Value))
        If Trim(CStr(wsSum.Cells(hdrRow + 5, 3).Value)) <> "Capsules- All shipments" Then
            errMsg = "기존 " & wkLabelFull & " 블록이 6행 구조가 아니라서 덮어쓸 수 없습니다. 직접 지우고 다시 실행해주세요."
            GoTo CleanUp
        End If
    Else
        lastAllShipRow = 0
        For r = wsSum.Cells(wsSum.rows.Count, 3).End(xlUp).Row To 1 Step -1
            If Trim(CStr(wsSum.Cells(r, 3).Value)) = "Capsules- All shipments" Then
                lastAllShipRow = r
                Exit For
            End If
        Next r
        If lastAllShipRow = 0 Then
            errMsg = "'Capsules- All shipments' 마지막 줄을 찾지 못했습니다. 시트 구조를 확인해주세요."
            GoTo CleanUp
        End If
        hdrRow = lastAllShipRow + 4
    End If

    r25 = hdrRow + 1
    r5 = hdrRow + 2
    r2 = hdrRow + 3
    rSol = hdrRow + 4
    rAll = hdrRow + 5

    ' ---- 새 블록일 때만 직전 블록 서식 복사 (6행 블록 구조 + 헤더가 "Week "로 시작할 때만) ----
    If Not overwrite Then
        prevHdr = lastAllShipRow - 5
        If prevHdr >= 1 Then
            If Left(Trim(CStr(wsSum.Cells(prevHdr, 3).Value)), 5) = "Week " Then
                wsSum.Range(wsSum.Cells(prevHdr, 2), wsSum.Cells(lastAllShipRow, 8)).Copy
                wsSum.Cells(hdrRow, 2).PasteSpecial xlPasteFormats
                Application.CutCopyMode = False
                fmtCopied = True
            End If
        End If
        If Not fmtCopied Then msg = msg & "[서식] 직전 블록 위치를 확인하지 못해 서식은 복사하지 않았습니다." & vbCrLf
    End If

    wsSum.Cells(hdrRow, 3).Value = wkLabelFull
    wsSum.Cells(hdrRow, 6).Value = wkShort

    wsSum.Cells(r25, 3).Value = "25mCi"
    wsSum.Cells(r25, 4).Value = mci25
    wsSum.Cells(r25, 6).Value = "High dose"
    wsSum.Cells(r25, 7).Formula = "=D" & rAll & "+D" & r25

    wsSum.Cells(r5, 3).Value = "5mCi"
    wsSum.Cells(r5, 4).Value = mci5
    wsSum.Cells(r5, 6).Value = "Low dose"
    wsSum.Cells(r5, 7).Formula = "=D" & r5 & "+D" & r2

    wsSum.Cells(r2, 3).Value = "2mCi"
    wsSum.Cells(r2, 4).Value = mci2
    wsSum.Cells(r2, 6).Value = "Solution"
    wsSum.Cells(r2, 7).Formula = "=D" & rSol

    wsSum.Cells(rSol, 3).Value = "Solution"
    wsSum.Cells(rSol, 4).Value = solutionQty

    wsSum.Cells(rAll, 3).Value = "Capsules- All shipments"
    wsSum.Cells(rAll, 4).Value = mciHighOthers
    written = True

CleanUp:
    Application.CutCopyMode = False
    Application.ScreenUpdating = True

    If errMsg <> "" Then
        Report STEP_NAME, errMsg, vbCritical
    ElseIf skipped Then
        Report STEP_NAME, msg, vbExclamation
    ElseIf written Then
        If msg = "" Then
            Report STEP_NAME, wkLabelFull & IIf(overwrite, " 블록 덮어쓰기 완료! (", " 블록 추가 완료! (") & hdrRow & "~" & rAll & "행)", vbInformation
        Else
            Report STEP_NAME, wkLabelFull & IIf(overwrite, " 블록 덮어쓰기 완료 (", " 블록 추가 완료 (") & hdrRow & "~" & rAll & "행)." & vbCrLf & msg, vbExclamation
        End If
    End If
    Exit Sub

ErrH:
    errMsg = "오류 " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Sub
