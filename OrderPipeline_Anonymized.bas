Option Explicit

' ============================================================
' Pharma Order Processing Pipeline (Portfolio / Anonymized Version)
' Original client names, agency codes, and partner names have been
' replaced with generic placeholders (PARTNER_CO, CLIENT01-31, etc.)
' Sheet names translated to English for readability.
' ============================================================

Sub ToggleEvents(ByVal State As Boolean)
    Application.ScreenUpdating = State
    Application.DisplayAlerts = State
    Application.EnableEvents = State
    If State Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
End Sub

Sub Step0_Clear_Columns()
    ' Clears staging columns before a new run
    Dim ws As Worksheet: Set ws = Sheets("Distribution")
    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row

    If lastRow < 15 Then Exit Sub

    Application.ScreenUpdating = False

    ws.Range("C15:C" & lastRow).ClearContents
    ws.Range("E15:E" & lastRow).ClearContents
    ws.Range("F15:F" & lastRow).ClearContents
    ws.Range("J15:J" & lastRow).ClearContents

    Application.ScreenUpdating = True

    MsgBox "Columns C, E, F, J cleared."
End Sub

Sub Step1_AutoFill_Smart_Optimized()
    ' Expands one order-summary row per unit quantity into individual rows
    ' (e.g. qty=4 -> 4 separate rows), for the "international partner" order stream.
    Dim wsOrder As Worksheet, wsTarget As Worksheet
    Dim i As Long, k As Long, targetRow As Long, lastRowOrder As Long
    Dim currentAgency As String, currentDay As String
    Dim mCiValue As Double, qty As Long
    Dim rowValue As String, lineData As Variant

    Set wsOrder = Sheets("OrderData_Intl")
    Set wsTarget = Sheets("Distribution")

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    wsTarget.Range("B15:C981, E15:G981, L15:Q981").ClearContents
    On Error Resume Next
    wsTarget.Range("L15:P981").UnMerge
    On Error GoTo 0

    lastRowOrder = wsOrder.Cells(wsOrder.Rows.Count, 1).End(xlUp).Row
    If lastRowOrder < 2 Then
        MsgBox "No data found in OrderData_Intl sheet.", vbExclamation
        GoTo CleanUp
    End If

    targetRow = 15

    For i = 2 To lastRowOrder
        rowValue = Application.Trim(wsOrder.Cells(i, 1).Value)
        If rowValue <> "" Then
            lineData = Split(rowValue, " ")

            If UBound(lineData) < 3 Then
                currentAgency = CStr(wsOrder.Cells(i, 1).Value)
                currentDay = wsOrder.Cells(i, 2).Value
                mCiValue = Val(wsOrder.Cells(i, 3).Value)
                qty = Val(wsOrder.Cells(i, 4).Value)
            Else
                currentAgency = CStr(lineData(0))
                currentDay = lineData(1)
                mCiValue = Val(lineData(2))
                qty = Val(lineData(3))
            End If

            If qty > 0 Then
                For k = 1 To qty
                    wsTarget.Cells(targetRow, "C").Value = currentAgency
                    wsTarget.Cells(targetRow, "E").Value = currentDay
                    wsTarget.Cells(targetRow, "F").Value = mCiValue
                    wsTarget.Cells(targetRow, "J").Value = "PARTNER_CO"
                    targetRow = targetRow + 1
                Next k
            End If
        End If
    Next i

    If targetRow > 15 Then
        With wsTarget.Range("D15:D" & targetRow - 1)
            .FillDown
        End With
    End If

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Step 1 complete. Rows generated: " & (targetRow - 15)
End Sub

Sub Step2_Final_Merge_Resolved()
    ' Sorts rows by dose size / day / client priority, then merges adjacent
    ' rows that belong to the same "package" (a package can span 1, 2 or 4
    ' physical units depending on the dose denomination).
    Dim ws As Worksheet: Set ws = Sheets("Distribution")
    Dim i As Long, k As Long, lastRow As Long, r As Long
    Dim actualMerge As Integer, mergeCount As Integer
    Dim noCounter As Long: noCounter = 1

    lastRow = 0
    For r = 981 To 15 Step -1
        If ws.Cells(r, "J").Value = "PARTNER_CO" Then
            lastRow = r
            Exit For
        End If
    Next r
    If lastRow < 15 Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    On Error Resume Next
    ws.Range("A15:Q" & lastRow).UnMerge
    On Error GoTo 0

    With ws.Sort
        .SortFields.Clear

        .SortFields.Add Key:=ws.Range("F15:F" & lastRow), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal

        .SortFields.Add Key:=ws.Range("E15:E" & lastRow), _
            SortOn:=xlSortOnValues, Order:=xlAscending, _
            CustomOrder:="Sun,Sat,Fri,Thu,Wed,Tue,Mon", DataOption:=xlSortNormal

        ' Client priority order — replace CLIENT01..CLIENT31 with your own
        ' client/agency code list when reusing this template.
        .SortFields.Add Key:=ws.Range("C15:C" & lastRow), _
            SortOn:=xlSortOnValues, Order:=xlAscending, _
            CustomOrder:="'CLIENT01','CLIENT02','CLIENT03','CLIENT04','CLIENT05','CLIENT06','CLIENT07','CLIENT08','CLIENT09','CLIENT10','CLIENT11','CLIENT12','CLIENT13','CLIENT14','CLIENT15','CLIENT16','CLIENT17','CLIENT18','CLIENT19','CLIENT20','CLIENT21','CLIENT22','CLIENT23','CLIENT24','CLIENT25','CLIENT26','CLIENT27','CLIENT28','CLIENT29','CLIENT30','CLIENT31'", DataOption:=xlSortNormal
        .SetRange ws.Range("A15:Q" & lastRow)
        .Header = xlNo
        .Apply
    End With

    ws.Range("Q15:Q" & lastRow).Interior.ColorIndex = xlNone

    i = 15
    Do While i <= lastRow
        Dim individualValue As Double: individualValue = Val(ws.Cells(i, "F").Value)
        Dim currentDay As String: currentDay = Trim(ws.Cells(i, "E").Value)
        Dim currentAgency As String: currentAgency = Trim(ws.Cells(i, "C").Value)

        ' Merge count depends on dose denomination — adjust to your own SKUs.
        Select Case individualValue
            Case 50, 30, 25, 5, 2: mergeCount = 4
            Case 100: mergeCount = 2
            Case Else: mergeCount = 1
        End Select

        actualMerge = 0
        For k = 0 To mergeCount - 1
            If (i + k) > lastRow Then Exit For
            If Trim(ws.Cells(i + k, "E").Value) <> currentDay Or _
               Trim(ws.Cells(i + k, "C").Value) <> currentAgency Then Exit For
            actualMerge = actualMerge + 1
        Next k

        If actualMerge = 0 Then actualMerge = 1

        If actualMerge = 1 Then
            ws.Cells(i, "L").Value = "No."
            ws.Cells(i, "M").Value = "=A" & i

            With ws.Range(ws.Cells(i, "N"), ws.Cells(i, "P"))
                .Merge
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Value = "1 Package (" & individualValue & "mCi)"
            End With
        Else
            ws.Cells(i, "L").Value = "No."
            For k = 0 To actualMerge - 1
                ws.Cells(i, 13 + k).Value = "=A" & (i + k)
            Next k

            With ws.Range(ws.Cells(i + 1, "L"), ws.Cells(i + actualMerge - 1, "P"))
                .Merge
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Value = "1 Package (" & (individualValue * actualMerge) & "mCi)"
            End With
        End If

        With ws.Range(ws.Cells(i, "L"), ws.Cells(i + actualMerge - 1, "P"))
            .Borders.LineStyle = xlNone
            .BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        End With

        ' Highlight priority client group (e.g. expedited region)
        Select Case currentAgency
            Case "CLIENT26", "CLIENT30", "CLIENT18", "CLIENT27", "CLIENT17", "CLIENT_SPECIAL"
                ws.Cells(i, "Q").Resize(actualMerge, 1).Interior.Color = RGB(255, 192, 203)
        End Select

        noCounter = noCounter + actualMerge
        i = i + actualMerge
    Loop

    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Merge complete."
End Sub

Sub Step3_Domestic_Only_Final_Fixed()
    ' Same expand-and-sort logic as Step1/2, applied to the domestic order stream,
    ' appended below the international rows on the same sheet.
    Dim wsOrder As Worksheet, wsTarget As Worksheet
    Dim i As Long, k As Long, targetRow As Long, startRow As Long
    Dim lastRowOrder As Long, qty As Long, lastR As Long
    Dim lastPackageRow As Long, searchRow As Long
    Dim r As Long

    Set wsOrder = Sheets("OrderData_Domestic")
    Set wsTarget = Sheets("Distribution")

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    lastPackageRow = 0
    For searchRow = 981 To 15 Step -1
        If wsTarget.Cells(searchRow, "J").Value = "PARTNER_CO" Then
            lastPackageRow = searchRow
            Exit For
        End If
    Next searchRow

    If lastPackageRow = 0 Then
        lastPackageRow = 14
    End If

    startRow = lastPackageRow + 1
    targetRow = startRow

    On Error Resume Next
    wsTarget.Range("A" & startRow & ":Q" & 1000).UnMerge
    On Error GoTo 0

    With wsTarget
        .Range("B" & startRow & ":C" & 976).ClearContents
        .Range("E" & startRow & ":G" & 976).ClearContents
        .Range("J" & startRow & ":J" & 976).ClearContents
        .Range("L" & startRow & ":Q" & 976).ClearContents
    End With

    lastRowOrder = wsOrder.Cells(wsOrder.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRowOrder
        qty = Val(wsOrder.Cells(i, 4).Value)
        If qty > 0 And Trim(wsOrder.Cells(i, 1).Value) <> "" Then
            For k = 1 To qty
                wsTarget.Cells(targetRow, "C").Value = wsOrder.Cells(i, 1).Value
                wsTarget.Cells(targetRow, "E").Value = wsOrder.Cells(i, 2).Value
                wsTarget.Cells(targetRow, "F").Value = Val(wsOrder.Cells(i, 3).Value)
                wsTarget.Cells(targetRow, "J").Value = wsOrder.Cells(i, 5).Value
                targetRow = targetRow + 1
            Next k
        End If
    Next i

    If targetRow > startRow Then
        lastR = targetRow - 1
        With wsTarget
            With .Sort
                .SortFields.Clear

                .SortFields.Add Key:=wsTarget.Range("F" & startRow & ":F" & lastR), _
                    SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal

                .SortFields.Add Key:=wsTarget.Range("E" & startRow & ":E" & lastR), _
                    SortOn:=xlSortOnValues, Order:=xlAscending, _
                    CustomOrder:="Sat,Fri,Thu,Wed,Tue,Mon", DataOption:=xlSortNormal

                .SetRange wsTarget.Range("A" & startRow & ":Q" & lastR)
                .Header = xlNo
                .Apply
            End With

            For r = startRow To lastR
                With .Range("L" & r & ":P" & r)
                    .Merge
                    .Borders.LineStyle = xlContinuous
                End With
                .Range("B" & r & ":C" & r).Borders.LineStyle = xlContinuous
                .Range("E" & r & ":G" & r).Borders.LineStyle = xlContinuous
                .Range("J" & r).Borders.LineStyle = xlContinuous
                .Range("L" & r & ":Q" & r).Borders.LineStyle = xlContinuous
            Next r
        End With
    End If

    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Domestic data sorted. Rows generated: " & (targetRow - startRow)
End Sub

Sub Step4_Solution_Mapping()
    ' Maps raw client-entered names to standardized display names via a
    ' lookup table, then expands and sorts a third order stream ("solution"
    ' orders received via a separate reconciliation form).
    Dim wsOrder As Worksheet, wsTarget As Worksheet, wsMap As Worksheet
    Dim vData As Variant
    Dim i As Long, k As Long, r As Long
    Dim lastRowOrder As Long, totalQty As Long
    Dim targetRow As Long, lastR As Long
    Dim clientDisplayName As String
    Dim startRow As Long

    Set wsOrder = Sheets("SolutionData")
    Set wsTarget = Sheets("Distribution")
    Set wsMap = Sheets("ClientNameMap")

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    lastRowOrder = wsOrder.Cells(wsOrder.Rows.Count, 1).End(xlUp).Row
    If lastRowOrder < 2 Then
        MsgBox "No data found in SolutionData sheet.", vbExclamation
        GoTo CleanUp
    End If

    vData = wsOrder.Range("A2:E" & lastRowOrder).Value

    totalQty = 0
    For i = 1 To UBound(vData, 1)
        totalQty = totalQty + Val(vData(i, 4))
    Next i

    If totalQty = 0 Then
        MsgBox "No quantity data found.", vbExclamation
        GoTo CleanUp
    End If

    startRow = 987
    targetRow = startRow
    lastR = startRow + totalQty - 1

    On Error Resume Next
    wsTarget.Range("A" & startRow & ":Q" & startRow + 500).UnMerge
    On Error GoTo 0

    wsTarget.Range("C" & startRow & ":C" & lastR).ClearContents
    wsTarget.Range("E" & startRow & ":F" & lastR).ClearContents
    wsTarget.Range("J" & startRow & ":J" & lastR).ClearContents
    wsTarget.Range("L" & startRow & ":Q" & lastR).ClearContents

    r = 0
    For i = 1 To UBound(vData, 1)
        If Val(vData(i, 4)) > 0 Then
            clientDisplayName = ""
            On Error Resume Next
            clientDisplayName = Application.WorksheetFunction.Index(wsMap.Range("A:A"), _
                           Application.WorksheetFunction.Match("*" & vData(i, 1) & "*", wsMap.Range("A:A"), 0))
            On Error GoTo 0

            If clientDisplayName = "" Then clientDisplayName = vData(i, 1)

            For k = 1 To Val(vData(i, 4))
                wsTarget.Cells(startRow + r, "C").Value = clientDisplayName
                wsTarget.Cells(startRow + r, "E").Value = vData(i, 2)
                wsTarget.Cells(startRow + r, "F").Value = vData(i, 3)
                wsTarget.Cells(startRow + r, "J").Value = vData(i, 5)
                r = r + 1
            Next k
        End If
    Next i

    If lastR >= startRow Then
        With wsTarget.Sort
            .SortFields.Clear
            .SortFields.Add Key:=wsTarget.Range("F" & startRow & ":F" & lastR), _
                SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
            .SortFields.Add Key:=wsTarget.Range("E" & startRow & ":E" & lastR), _
                SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:="Sat,Fri,Thu,Wed,Tue,Mon", DataOption:=xlSortNormal
            .SetRange wsTarget.Range("A" & startRow & ":Q" & lastR)
            .Header = xlNo
            .Apply
        End With
    End If

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Solution data sorted. Rows generated: " & totalQty
End Sub
