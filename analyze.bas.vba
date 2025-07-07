Attribute VB_Name = "analyze"
Sub GO()
    Dim FirstReferenceRow As Integer
    Dim LastReferenceRow As Integer
    Dim result() As t_BOM3
    Dim lastRow As Long
    Dim i, j, k, Qty As Integer
    Dim BOQ() As t_BOM1
    Dim QtyIssue As Boolean
    On Error Resume Next
    
    ' Getting the references's last row
    FirstReferenceRow = 6
    LastReferenceRow = FirstReferenceRow
    Dim EmptyFound As Boolean: EmptyFound = False
    While (Not EmptyFound)
        If Worksheets("Main").Cells(LastReferenceRow + 1, 3) = "" Then EmptyFound = True
        If (Not EmptyFound) Then LastReferenceRow = LastReferenceRow + 1
    Wend
    k = 0
    For i = FirstReferenceRow To LastReferenceRow
        j = 0
        Qty = 0
        QtyIssue = False
        Qty = CInt(Worksheets("main").Cells(i, 4).value)
        If Qty = 0 Then
            Qty = 1
            QtyIssue = True
        End If
        j = addItem(Worksheets("main").Cells(i, 3).value, Qty, result)
        k = k + 1
        If j = -1 Then
            ' Not recognized
            Worksheets("main").Cells(i, 2).value = CStr(k)
            Worksheets("main").Cells(i, 3).Interior.Color = vbRed
            Worksheets("main").Cells(i, 3).Font.Color = vbBlack
            Worksheets("main").Cells(i, 5).value = "Not recognized"
        Else
            Worksheets("main").Cells(i, 2).value = CStr(k)
            Worksheets("main").Cells(i, 3).Interior.Color = vbWhite
            Worksheets("main").Cells(i, 3).Font.Color = vbBlack
            If Not QtyIssue Then
                Worksheets("main").Cells(i, 4).Interior.Color = vbWhite
                Worksheets("main").Cells(i, 5).value = "OK"
            Else
                Worksheets("main").Cells(i, 4).Interior.Color = vbYellow
                Worksheets("main").Cells(i, 5).value = "Processed as Qty = 1"
            End If
            
            '''''''''''''''''''''''''''
            Dim b As Integer
            b = UBound(result(UBound(result)).Items) - LBound(result(UBound(result)).Items) + 1
            If b = result(UBound(result)).Qty_Driver Then Worksheets("main").Cells(i, 5).value = "OK"
            If b < result(UBound(result)).Qty_Driver Then Worksheets("main").Cells(i, 5).value = "Too many drivers"
            If b > result(UBound(result)).Qty_Driver Then Worksheets("main").Cells(i, 5).value = "Missing driver"
            '''''''''''''''''''''''''''
            Dim Missing_PCB As Boolean
            Missing_PCB = False
            Dim cc As Integer
            For cc = LBound(result(UBound(result)).Items) To UBound(result(UBound(result)).Items)
                Dim h As Integer
                h = 1
                Dim para As t_Parameters
                para = getAllParameters(result(UBound(result)).Items(cc).reference)
                If (result(UBound(result)).Items(cc).HasHalfFoot And para.length > 6) Or (para.length = 60) Then h = 2
                If result(UBound(result)).Items(cc).Qty_PCB <> h Then
                    Missing_PCB = True
                    Exit For
                End If
            Next
            If Missing_PCB Then Worksheets("main").Cells(i, 5).value = "Missing PCB"
            '''''''''''''''''''''''''''''
            Dim op As Integer
            Dim Missing_Optic_Diffuser As Boolean
            Missing_Optic_Diffuser = False
            Dim Missing_Optic_Grazer As Boolean
            Missing_Optic_Grazer = False
            Dim Missing_Optic_Washer As Boolean
            Missing_Optic_Washer = False
            Dim Missing_Optic_Sym As Boolean
            Missing_Optic_Sym = False
            Dim Missing_Optic_HC_TIR As Boolean
            Missing_Optic_HC_TIR = False
            Dim Missing_Optic_HC_Opal As Boolean
            Missing_Optic_HC_Opal = False
            For op = LBound(result(UBound(result)).Items) To UBound(result(UBound(result)).Items)
                Dim s As String
                s = UCase(Mid(result(UBound(result)).Items(op).reference, 1, 3))
                'Dim para As t_Parameters
                para = getAllParameters(result(UBound(result)).Items(op).reference)
                ' Opal case
                If s = "BIO" Or s = "BOO" Or s = "BJO" Then
                    If result(UBound(result)).Items(op).Qty_Optic_Diffuser = 0 Then Missing_Optic_Diffuser = True
                End If
                If Missing_Optic_Diffuser Then Exit For
                '''''''''''
                'Grazer case
                If (s = "BIW" Or s = "BOW" Or s = "BJW") And para.BeamAngle = "G" Then
                    If result(UBound(result)).Items(op).Qty_Optic_Lens = 0 Then Missing_Optic_Grazer = True
                    If result(UBound(result)).Items(op).Qty_Optic_Reflector = 0 Then Missing_Optic_Grazer = True
                    If result(UBound(result)).Items(op).Qty_Optic_Kick_Reflector = 0 Then Missing_Optic_Grazer = True
                End If
                If Missing_Optic_Grazer Then Exit For
                ''''''''''''
                'Washer case
                If (s = "BIW" Or s = "BOW" Or s = "BJW") And para.BeamAngle = "W" Then
                    If result(UBound(result)).Items(op).Qty_Optic_Lens = 0 Then Missing_Optic_Washer = True
                    If result(UBound(result)).Items(op).Qty_Optic_Reflector = 0 Then Missing_Optic_Washer = True
                    If result(UBound(result)).Items(op).Qty_Optic_Kick_Reflector = 0 Then Missing_Optic_Washer = True
                    If result(UBound(result)).Items(op).Qty_Optic_Fresnel = 0 Then Missing_Optic_Washer = True
                End If
                If Missing_Optic_Washer Then Exit For
                '''''''''''''
                'Symmetrical case
                If s = "BIS" Or s = "BOS" Or s = "BJS" Then
                    If result(UBound(result)).Items(op).Qty_Optic_Lens = 0 Then Missing_Optic_Sym = True
                    If result(UBound(result)).Items(op).Qty_Optic_Reflector = 0 Then Missing_Optic_Sym = True
                End If
                If Missing_Optic_Sym Then Exit For
                ''''''''''''''
                 'Grazer Honeycomb TIR
                If s = "BXH" Or s = "BKH" Or s = "BHS" Then
                    If result(UBound(result)).Items(op).Qty_Optic_Lens = 0 Then Missing_Optic_HC_TIR = True
                    If result(UBound(result)).Items(op).Qty_Optic_Reflector = 0 Then Missing_Optic_HC_TIR = True
                End If
                If Missing_Optic_HC_TIR Then Exit For
                ''''''''''''''''
                'Grazer Honeycomb Opal
                If s = "BXO" Or s = "BKO" Or s = "BHO" Then
                    If result(UBound(result)).Items(op).Qty_Optic_Reflector = 0 Then Missing_Optic_HC_Opal = True
                End If
                If Missing_Optic_HC_Opal Then Exit For
                ''''''''''''''''
            Next
            If Missing_Optic_Diffuser Then Worksheets("main").Cells(i, 5).value = "Missing Optic"
            If Missing_Optic_Grazer Then Worksheets("main").Cells(i, 5).value = "Missing Optic"
            If Missing_Optic_Washer Then Worksheets("main").Cells(i, 5).value = "Missing Optic"
            If Missing_Optic_Sym Then Worksheets("main").Cells(i, 5).value = "Missing Optic"
            If Missing_Optic_HC_TIR Then Worksheets("main").Cells(i, 5).value = "Missing Optic"
            If Missing_Optic_HC_Opal Then Worksheets("main").Cells(i, 5).value = "Missing Optic"
            
        End If
    Next i
    Call displayBOM.displayBOM(result, "---")
    Call displayBOM.displayBOM_ERPLayout(result, "---")
    Call displayBOM.displayBOM_Pricelist(result, "---")

    ' Getting the BOQ
    Dim np As Integer
    For i = LBound(result) To UBound(result)
        On Error Resume Next
        For j = LBound(result(i).Items) To UBound(result(i).Items)
            For k = LBound(result(i).Items(j).Items) To UBound(result(i).Items(j).Items)
                On Error Resume Next
                np = UBound(BOQ) + 1
                If Err.Number <> 0 Then np = 1
                ReDim Preserve BOQ(1 To np)
                BOQ(np) = result(i).Items(j).Items(k)
            Next k
        Next j
    Next i
    
    ' Filtering the BOQ
    For i = LBound(BOQ) To UBound(BOQ) - 1
        For j = i + 1 To UBound(BOQ)
            If BOQ(i).ERP <> "-" And BOQ(i).ERP = BOQ(j).ERP And BOQ(i).Description = BOQ(j).Description And BOQ(i).length = BOQ(j).length Then
                BOQ(i).TQty = BOQ(i).TQty + BOQ(j).TQty
                BOQ(j).ERP = "-"
            End If
        Next j
    Next i

    ' Displaying the BOQ
    Worksheets("BOQ").Activate
    Application.ScreenUpdating = False
    Cells.Clear
    Dim ResultRow As Integer
    ResultRow = 1
    Cells(ResultRow, 1).value = "Category"
    Cells(ResultRow, 2).value = "Item"
    Cells(ResultRow, 3).value = "ERP"
    Cells(ResultRow, 4).value = "Length"
    Cells(ResultRow, 5).value = "Qty"
    Cells(ResultRow, 6).value = "Description"
    
    For i = 1 To 6
        Cells(ResultRow, i).Interior.Color = RGB(198, 239, 208)
        Cells(ResultRow, i).Font.ForeColor = RGB(0, 97, 0)
        Cells(ResultRow, i).Font.Bold = True
        Cells(ResultRow, i).Borders.LineStyle = xlContinuous
    Next
    ActiveSheet.Columns("A").HorizontalAlignment = xlCenter
    ActiveSheet.Columns("B").HorizontalAlignment = xlCenter
    ActiveSheet.Columns("C").HorizontalAlignment = xlCenter
    ActiveSheet.Columns("D").HorizontalAlignment = xlCenter
    ActiveSheet.Columns("E").HorizontalAlignment = xlCenter
    ActiveSheet.Columns("F").HorizontalAlignment = xlLeft

    ResultRow = ResultRow + 1

    For i = LBound(BOQ) To UBound(BOQ)
        If BOQ(i).ERP <> "-" Then
            Cells(ResultRow, 1).value = BOQ(i).Category
            Cells(ResultRow, 2).value = BOQ(i).Item
            Cells(ResultRow, 3).value = BOQ(i).ERP
            Cells(ResultRow, 4).value = LengthToDisplay(BOQ(i).length)
            Cells(ResultRow, 5).value = ValueToDisplay(BOQ(i).TQty, "pc")
            Cells(ResultRow, 6).value = BOQ(i).Description
            For j = 1 To 6
                Cells(ResultRow, j).Borders.LineStyle = xlContinuous
            Next j
            ResultRow = ResultRow + 1
        End If
    Next i
    
    Columns("A:F").EntireColumn.AutoFit
    Application.ScreenUpdating = True
    Columns("A").EntireColumn.ColumnWidth = Columns("A").EntireColumn.ColumnWidth + 1
    Columns("B").EntireColumn.ColumnWidth = Columns("B").EntireColumn.ColumnWidth + 1
    Columns("C").EntireColumn.ColumnWidth = Columns("C").EntireColumn.ColumnWidth + 1
    Columns("D").EntireColumn.ColumnWidth = Columns("D").EntireColumn.ColumnWidth + 1
    Columns("E").EntireColumn.ColumnWidth = Columns("E").EntireColumn.ColumnWidth + 1
    Columns("F").EntireColumn.ColumnWidth = Columns("F").EntireColumn.ColumnWidth + 1
    
    Call sortBOQ(ResultRow)
    
    Worksheets("main").Select
        
End Sub

Private Sub sortBOQ(lr As Integer)
    Dim ws As Worksheet
    Dim rng As Range
    Dim sortRange As Range

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("BOQ")

    ' Set the range to be sorted
    Set rng = ws.Range("A1:F" & CStr(lr))

    ' Define the sort range
    Set sortRange = rng.Columns("A:C")

    ' Sort the range
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=sortRange.Columns(1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=sortRange.Columns(2), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=sortRange.Columns(3), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange rng
        .Header = xlYes ' Assuming the first row contains headers
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub



