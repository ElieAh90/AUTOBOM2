Attribute VB_Name = "displayBOM"
Dim Margin As Single
Dim Mark_Up As Single


Public Sub displayBOM(t() As t_BOM3, ByVal refNb As String)
    Const StartingResultRow As Integer = 3
    ' Filling "Result" sheet
    Dim colRange As Range
    Worksheets("Detailed BOM").Activate
    Application.ScreenUpdating = False
    Cells.Clear
    Sheets("Result Format").Cells.Copy Destination:=Sheets("Detailed BOM").Range("A1")
    'Cells(1, 11).value = refNb
    Dim ResultRow As Integer
    ResultRow = StartingResultRow
    Dim p As Integer
    For p = LBound(t) To UBound(t)
    
        Dim itemRecognized As Boolean
        itemRecognized = False
        On Error Resume Next
        Dim q As Integer
        q = LBound(t(p).Items)
        If Err.Number = 0 Then itemRecognized = True
    
        Cells(ResultRow, 2).value = "Reference"
        Cells(ResultRow, 3).value = "Qty"
        Cells(ResultRow, 10).value = "EXW RM"
        Cells(ResultRow, 11).value = "Landed RM"
        Cells(ResultRow, 12).value = "EXW Driver"
        Cells(ResultRow, 13).value = "Landed Driver"
        Set colRange = Range(Cells(ResultRow, 4), Cells(ResultRow, 9))
        colRange.Merge
        Cells(ResultRow, 4).HorizontalAlignment = xlCenter
        Cells(ResultRow, 1).Interior.Color = RGB(235, 235, 235)
        Cells(ResultRow, 2).Interior.Color = RGB(235, 235, 235)
        Cells(ResultRow, 3).Interior.Color = RGB(235, 235, 235)
        Cells(ResultRow, 4).Interior.Color = RGB(235, 235, 235)
        Cells(ResultRow, 10).Interior.Color = RGB(235, 235, 235)
        Cells(ResultRow, 11).Interior.Color = RGB(235, 235, 235)
        Cells(ResultRow, 12).Interior.Color = RGB(235, 235, 235)
        Cells(ResultRow, 13).Interior.Color = RGB(235, 235, 235)
        Cells(ResultRow, 4).Font.Bold = False
        Cells(ResultRow, 4).Font.Size = Cells(ResultRow, 1).Font.Size - 1
            
        If t(p).ProvidedLength <> t(p).RequiredLength And itemRecognized Then
            Dim s As String
            s = "The item "
            Dim p1 As Integer
            p1 = Len(s)
            s = s & t(p).reference & " will be replaced by "
            Dim p2 As Integer
            p2 = Len(s)
            If resultIndex = 0 Then
                s = s & t(p).ProvidedReference & " (custom solution)"
            Else
                s = s & t(p).ProvidedReference & " (closest feasible length)"
            End If
            Cells(ResultRow, 4).value = s
            Cells(ResultRow, 4).Font.Color = vbRed
            Cells(ResultRow, 4).Characters(p1 + 1, Len(t(p).reference)).Font.Bold = True
            Cells(ResultRow, 4).Characters(p2 + 1, Len(t(p).ProvidedReference)).Font.Bold = True
        End If
        ResultRow = ResultRow + 1
            
        Cells(ResultRow, 1).value = "# " & p
        Cells(ResultRow, 1).Font.Bold = True
        If t(p).ProvidedLength <> t(p).RequiredLength And itemRecognized Then
            Cells(ResultRow, 2).value = t(p).ProvidedReference
            Cells(ResultRow, 2).Characters(5, Len(CStr(getAllParameters(t(p).ProvidedReference).length))).Font.Color = vbRed
        Else
            Cells(ResultRow, 2).value = t(p).reference
        End If
        Cells(ResultRow, 2).Font.Bold = True
        
        Cells(ResultRow, 2).HorizontalAlignment = xlLeft
        Cells(ResultRow, 3).value = t(p).Qty
        Cells(ResultRow, 3).Font.Bold = True
        Set colRange = Range(Cells(ResultRow, 4), Cells(ResultRow, 9))
        colRange.Merge
        
        If t(p).ProvidedLength <> t(p).RequiredLength And itemRecognized Then
            Cells(ResultRow, 4).value = t(p).ProvidedDescription
        Else
            Cells(ResultRow, 4).value = t(p).Description
        End If
        Cells(ResultRow, 4).WrapText = True
        Cells(ResultRow, 4).Font.Bold = True
        Cells(ResultRow, 4).RowHeight = 2 * Cells(ResultRow, 4).RowHeight
        '''''''''
        'Cells(ResultRow, 10).value = t(p).CostEach ' calculated below instead
        
        'Cells(ResultRow, 11).value = t(p).CostEach * t(p).Qty
        Cells(ResultRow, 11).value = "Todo"
                
        Cells(ResultRow, 10).Font.Bold = True
        Cells(ResultRow, 11).Font.Bold = True
        Cells(ResultRow, 12).Font.Bold = True
        Cells(ResultRow, 13).Font.Bold = True
        'Cells(ResultRow, 10).Interior.Color = vbYellow
        'If t(p).CostEach > 0 Then
            'Cells(ResultRow, 11).Interior.Color = vbGreen
        'Else
            'Cells(ResultRow, 11).Interior.Color = vbRed
        'End If
        Cells(ResultRow, 4).Borders(xlEdgeLeft).LineStyle = xlDouble
        ResultRow = ResultRow + 1
        ' Displaying the real composition
        If itemRecognized Then
            For q = LBound(t(p).Items) To UBound(t(p).Items)
                Cells(ResultRow, 2).value = t(p).Items(q).dashedReference
                Cells(ResultRow, 2).Font.Bold = True
                Cells(ResultRow, 2).HorizontalAlignment = xlRight
                Cells(ResultRow, 3).value = t(p).Items(q).Qty
                Cells(ResultRow, 2).Font.Size = Cells(ResultRow, 2).Font.Size - 1
                Cells(ResultRow, 3).Font.Size = Cells(ResultRow, 3).Font.Size - 1
                Set colRange = Range(Cells(ResultRow, 4), Cells(ResultRow, 9))
                colRange.Merge
                Cells(ResultRow, 4).value = t(p).Items(q).Description
                Cells(ResultRow, 4).WrapText = True
                Cells(ResultRow, 4).Font.Bold = True
                Cells(ResultRow, 4).Font.Italic = True
                Cells(ResultRow, 4).Font.Size = Cells(ResultRow, 4).Font.Size - 1
                Cells(ResultRow, 4).RowHeight = 2 * Cells(ResultRow, 4).RowHeight
                '''''''
                'Cells(ResultRow, 10).value = t(p).Items(q).CostEach
                'Cells(ResultRow, 11).value = t(p).Items(q).CostEach * t(p).Items(q).Qty
                Cells(ResultRow, 10).value = "-"
                'Cells(ResultRow, 11).value = t(p).Items(q).CostEach
                
                Dim TotalLanded_Row As Integer
                TotalLanded_Row = ResultRow
                Cells(ResultRow, 10).Font.Bold = True
                Cells(ResultRow, 11).Font.Bold = True
                Cells(ResultRow, 12).Font.Bold = True
                Cells(ResultRow, 14).Font.Bold = True
                Cells(ResultRow, 10).Font.Italic = True
                Cells(ResultRow, 11).Font.Italic = True
                Cells(ResultRow, 12).Font.Italic = True
                Cells(ResultRow, 14).Font.Italic = True
                Cells(ResultRow, 10).Font.Size = Cells(ResultRow, 10).Font.Size - 1
                Cells(ResultRow, 11).Font.Size = Cells(ResultRow, 11).Font.Size - 1
                Cells(ResultRow, 12).Font.Size = Cells(ResultRow, 12).Font.Size - 1
                Cells(ResultRow, 14).Font.Size = Cells(ResultRow, 14).Font.Size - 1
                Cells(ResultRow, 4).Borders(xlEdgeLeft).LineStyle = xlDouble
                Cells(ResultRow - 1 - 1, 2).Borders(xlEdgeBottom).LineStyle = xlDot
                Cells(ResultRow - 1 - 1, 3).Borders(xlEdgeBottom).LineStyle = xlDot
                Cells(ResultRow - 1 - 1, 4).Borders(xlEdgeBottom).LineStyle = xlDot
                Cells(ResultRow - 1 - 1, 5).Borders(xlEdgeBottom).LineStyle = xlDot
                Cells(ResultRow - 1 - 1, 6).Borders(xlEdgeBottom).LineStyle = xlDot
                Cells(ResultRow - 1 - 1, 7).Borders(xlEdgeBottom).LineStyle = xlDot
                Cells(ResultRow - 1 - 1, 8).Borders(xlEdgeBottom).LineStyle = xlDot
                Cells(ResultRow - 1 - 1, 9).Borders(xlEdgeBottom).LineStyle = xlDot
                Cells(ResultRow - 1 - 1, 10).Borders(xlEdgeBottom).LineStyle = xlDot
                Cells(ResultRow - 1 - 1, 11).Borders(xlEdgeBottom).LineStyle = xlDot
                ResultRow = ResultRow + 1
                
                Cells(ResultRow, 4).value = "Item"
                Cells(ResultRow, 5).value = "ERP"
                Cells(ResultRow, 6).value = "Length"
                Cells(ResultRow, 7).value = "Qty"
                Cells(ResultRow, 8).value = "T.Qty"
                Cells(ResultRow, 9).value = "Description"
                
                Cells(ResultRow, 10).value = "EXW Unit"
                Cells(ResultRow, 11).value = "EXW Total"
                Cells(ResultRow, 12).value = "Landed Total GA"
                Cells(ResultRow, 13).value = "Multiplier GA"
                Cells(ResultRow, 14).value = "Landed Total LB"
                Cells(ResultRow, 15).value = "Multiplier LB"
                
                Cells(ResultRow, 4).Borders(xlEdgeLeft).LineStyle = xlDouble
                Cells(ResultRow, 4).Font.Italic = True
                Cells(ResultRow, 4).Font.Size = Cells(ResultRow, 4).Font.Size - 2
                Cells(ResultRow, 4).HorizontalAlignment = xlCenter
                Cells(ResultRow, 4).Font.Underline = True
                For f = 5 To 15
                    Cells(ResultRow, f).Font.Italic = True
                    Cells(ResultRow, f).Font.Size = Cells(ResultRow, f).Font.Size - 2
                    Cells(ResultRow, f).HorizontalAlignment = xlCenter
                    Cells(ResultRow, f).Font.Underline = True
                Next
                
                
                ResultRow = ResultRow + 1
                Dim elecPart As Byte
                Dim manlPart As Byte
                elecPart = t(p).Items(q).ElectricPart / t(p).Items(q).CostEach * 100
                manlPart = t(p).Items(q).ManlaborPart / t(p).Items(q).CostEach * 100
                Cells(ResultRow + 1, 2).value = "Inset part: " + Format(elecPart, "0") + "%"
                Cells(ResultRow + 2, 2).value = "Control part: " + Format(manlPart, "0") + "%"
                Cells(ResultRow, 2).value = "Body part: " + Format((100 - elecPart - manlPart), "0") + "%"

                Cells(ResultRow, 2).HorizontalAlignment = xlRight
                Cells(ResultRow + 1, 2).HorizontalAlignment = xlRight
                Cells(ResultRow + 2, 2).HorizontalAlignment = xlRight
                
                Dim r As Integer
                Dim TotalStartingRow As Integer
                Dim DriverRow As Integer
                TotalStartingRow = ResultRow
                For r = LBound(t(p).Items(q).Items) To UBound(t(p).Items(q).Items)
                    If UCase(t(p).Items(q).Items(r).Category) = "DRIVERS" Then DriverRow = ResultRow
                    Cells(ResultRow, 4).value = t(p).Items(q).Items(r).Item & " / " & t(p).Items(q).Items(r).Category
                    Cells(ResultRow, 5).value = t(p).Items(q).Items(r).ERP
                    Cells(ResultRow, 6).value = LengthToDisplay(t(p).Items(q).Items(r).length)
                    Cells(ResultRow, 7).value = Format3(t(p).Items(q).Items(r).Qty)
                    If (t(p).Items(q).Items(r).Qty = 0) Then
                        Cells(ResultRow, 7).EntireRow.Font.Color = vbRed
                        Cells(ResultRow, 7).EntireRow.Font.Bold = True
                    End If
                    Cells(ResultRow, 8).value = Format3(t(p).Items(q).Items(r).TQty)
                    Cells(ResultRow, 9).value = t(p).Items(q).Items(r).Description
                    Cells(ResultRow, 10).value = t(p).Items(q).Items(r).CostEach
                    'Cells(ResultRow, 11).value = t(p).Items(q).Items(r).CostEach * t(p).Items(q).Items(r).Qty * t(p).Items(q).Qty
                    Cells(ResultRow, 11).value = "=J" & Format(ResultRow, "0") & "*G" + Format(ResultRow, "0")
                    Cells(ResultRow, 10).Font.Italic = True
                    Cells(ResultRow, 11).Font.Italic = True
                    Cells(ResultRow, 10).Font.Size = Cells(ResultRow, 10).Font.Size - 2
                    Cells(ResultRow, 11).Font.Size = Cells(ResultRow, 11).Font.Size - 2
                    
                    Cells(ResultRow, 13).value = t(p).Items(q).Items(r).MultiplierGA
                    'Cells(ResultRow, 12).value = t(p).Items(q).Items(r).MultiplierGA * Cells(ResultRow, 11).value
                    Cells(ResultRow, 12).value = "=K" & Format(ResultRow, "0") & "*M" + Format(ResultRow, "0")
                    
                    'TotalLandedGA = TotalLandedGA + Cells(ResultRow, 12).value
                    Cells(ResultRow, 15).value = t(p).Items(q).Items(r).MultiplierLB
                    'Cells(ResultRow, 14).value = t(p).Items(q).Items(r).MultiplierLB * Cells(ResultRow, 11).value
                    Cells(ResultRow, 14).value = "=K" & Format(ResultRow, "0") & "*O" + Format(ResultRow, "0")
                    
                    'TotalLandedLB = TotalLandedLB + Cells(ResultRow, 14).value
                    Cells(ResultRow, 4).Borders(xlEdgeLeft).LineStyle = xlDouble
                    Cells(ResultRow, 4).Font.Italic = True
                    Cells(ResultRow, 4).Font.Size = Cells(ResultRow, 4).Font.Size - 2
                    Cells(ResultRow, 5).Font.Italic = True
                    Cells(ResultRow, 5).Font.Size = Cells(ResultRow, 5).Font.Size - 2
                    Cells(ResultRow, 6).Font.Italic = True
                    Cells(ResultRow, 6).Font.Size = Cells(ResultRow, 6).Font.Size - 2
                    Cells(ResultRow, 7).Font.Italic = True
                    Cells(ResultRow, 7).Font.Size = Cells(ResultRow, 7).Font.Size - 2
                    Cells(ResultRow, 8).Font.Italic = True
                    Cells(ResultRow, 8).Font.Size = Cells(ResultRow, 8).Font.Size - 2
                    Cells(ResultRow, 9).Font.Italic = True
                    Cells(ResultRow, 9).Font.Size = Cells(ResultRow, 9).Font.Size - 2
                    Cells(ResultRow, 10).Font.Italic = True
                    Cells(ResultRow, 10).Font.Size = Cells(ResultRow, 10).Font.Size - 2
                    Cells(ResultRow, 11).Font.Italic = True
                    Cells(ResultRow, 11).Font.Size = Cells(ResultRow, 11).Font.Size - 2
                    Cells(ResultRow, 12).Font.Italic = True
                    Cells(ResultRow, 12).Font.Size = Cells(ResultRow, 12).Font.Size - 2
                    Cells(ResultRow, 13).Font.Italic = True
                    Cells(ResultRow, 13).Font.Size = Cells(ResultRow, 13).Font.Size - 2
                    Cells(ResultRow, 14).Font.Italic = True
                    Cells(ResultRow, 14).Font.Size = Cells(ResultRow, 14).Font.Size - 2
                    Cells(ResultRow, 15).Font.Italic = True
                    Cells(ResultRow, 15).Font.Size = Cells(ResultRow, 15).Font.Size - 2
                    
                    
                    
                    ResultRow = ResultRow + 1
                Next
                ' Display the landed of GA and LB &ExW Total
                Cells(TotalLanded_Row, 12).value = "=SUM(L" & Format(TotalStartingRow, "0") & ":L" & Format(ResultRow - 1, "0") & ")"
                Cells(TotalLanded_Row, 14).value = "=SUM(N" & Format(TotalStartingRow, "0") & ":N" & Format(ResultRow - 1, "0") & ")"
                Cells(TotalLanded_Row, 11).value = "=SUM(K" & Format(TotalStartingRow, "0") & ":K" & Format(ResultRow - 1, "0") & ")"
                Cells(TotalLanded_Row - 1, 10).value = "=K" & Format(TotalLanded_Row, "0") & "-K" + Format(DriverRow, "0")
                Cells(TotalLanded_Row - 1, 11).value = "=N" & Format(TotalLanded_Row, "0") & "-N" + Format(DriverRow, "0")
                Cells(TotalLanded_Row - 1, 12).value = "=K" + Format(DriverRow, "0")
                Cells(TotalLanded_Row - 1, 13).value = "=N" + Format(DriverRow, "0")
                
                ResultRow = ResultRow + 1
                
            Next
        End If
        
        Cells(ResultRow - 1, 1).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 2).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 3).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 4).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 5).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 6).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 7).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 8).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 9).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 10).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 11).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 12).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 13).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 14).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 15).Borders(xlEdgeBottom).LineStyle = xlDouble
        ResultRow = ResultRow + 1
    Next
    Columns("A:O").EntireColumn.AutoFit
    Application.ScreenUpdating = True
    Columns("D").EntireColumn.ColumnWidth = Columns("D").EntireColumn.ColumnWidth + 1
    Columns("E").EntireColumn.ColumnWidth = Columns("E").EntireColumn.ColumnWidth + 1
    Columns("F").EntireColumn.ColumnWidth = Columns("F").EntireColumn.ColumnWidth + 1
    Columns("G").EntireColumn.ColumnWidth = Columns("G").EntireColumn.ColumnWidth + 1
    Columns("H").EntireColumn.ColumnWidth = Columns("H").EntireColumn.ColumnWidth + 1
    Columns("I").EntireColumn.ColumnWidth = Columns("I").EntireColumn.ColumnWidth + 1
End Sub

Function LengthToDisplay(ByVal length As Single) As String
    If length = 0 Then
        LengthToDisplay = "-"
    Else
        LengthToDisplay = length & " mm"
    End If
End Function

Function ValueToDisplay(ByVal value As Single, ByVal unit As String) As String
    If value = Int(value) Then
        ValueToDisplay = CStr(value)
    Else
        ValueToDisplay = Format(value, "0.00") & " " & unit
    End If
End Function

Public Sub displayBOM_ERPLayout(t() As t_BOM3, ByVal refNb As String)
    Const StartingResultRow As Integer = 3
    ' Filling "Result" sheet
    Dim colRange As Range
    Worksheets("ERP Layout").Activate
    Application.ScreenUpdating = False
    Cells.Clear
    Sheets("Result Format").Cells.Copy Destination:=Sheets("ERP Layout").Range("A1")
    'Cells(1, 11).value = refNb
    Dim ResultRow As Integer
    ResultRow = StartingResultRow
    Dim p As Integer
    For p = LBound(t) To UBound(t)

        Dim itemRecognized As Boolean
        itemRecognized = False
        On Error Resume Next
        Dim q As Integer
        q = LBound(t(p).Items)
        If Err.Number = 0 Then itemRecognized = True
    
        Cells(ResultRow, 2).value = "Reference"
        Cells(ResultRow, 3).value = "Qty"
        Cells(ResultRow, 11).value = "Unit Cost"
        Cells(ResultRow, 12).value = "Total Cost"
        Set colRange = Range(Cells(ResultRow, 4), Cells(ResultRow, 10))
        colRange.Merge
        Cells(ResultRow, 4).HorizontalAlignment = xlCenter
        Cells(ResultRow, 1).Interior.Color = RGB(235, 235, 235)
        Cells(ResultRow, 2).Interior.Color = RGB(235, 235, 235)
        Cells(ResultRow, 3).Interior.Color = RGB(235, 235, 235)
        Cells(ResultRow, 4).Interior.Color = RGB(235, 235, 235)
        Cells(ResultRow, 11).Interior.Color = RGB(235, 235, 235)
        Cells(ResultRow, 12).Interior.Color = RGB(235, 235, 235)
        Cells(ResultRow, 4).Font.Bold = False
        Cells(ResultRow, 4).Font.Size = Cells(ResultRow, 1).Font.Size - 1
            
        If t(p).ProvidedLength <> t(p).RequiredLength And itemRecognized Then
            Dim s As String
            s = "The item "
            Dim p1 As Integer
            p1 = Len(s)
            s = s & t(p).reference & " will be replaced by "
            Dim p2 As Integer
            p2 = Len(s)
            If resultIndex = 0 Then
                s = s & t(p).ProvidedReference & " (custom solution)"
            Else
                s = s & t(p).ProvidedReference & " (closest feasible length)"
            End If
            Cells(ResultRow, 4).value = s
            Cells(ResultRow, 4).Font.Color = vbRed
            Cells(ResultRow, 4).Characters(p1 + 1, Len(t(p).reference)).Font.Bold = True
            Cells(ResultRow, 4).Characters(p2 + 1, Len(t(p).ProvidedReference)).Font.Bold = True
        End If
        ResultRow = ResultRow + 1
            
        Cells(ResultRow, 1).value = "# " & p
        Cells(ResultRow, 1).Font.Bold = True
        If t(p).ProvidedLength <> t(p).RequiredLength And itemRecognized Then
            Cells(ResultRow, 2).value = t(p).ProvidedReference
            Cells(ResultRow, 2).Characters(5, Len(CStr(getAllParameters(t(p).ProvidedReference).length))).Font.Color = vbRed
        Else
            Cells(ResultRow, 2).value = t(p).reference
        End If
        Cells(ResultRow, 2).Font.Bold = True
        
        Cells(ResultRow, 2).HorizontalAlignment = xlLeft
        Cells(ResultRow, 3).value = t(p).Qty
        Cells(ResultRow, 3).Font.Bold = True
        Set colRange = Range(Cells(ResultRow, 4), Cells(ResultRow, 10))
        colRange.Merge
        
        If t(p).ProvidedLength <> t(p).RequiredLength And itemRecognized Then
            Cells(ResultRow, 4).value = t(p).ProvidedDescription
        Else
            Cells(ResultRow, 4).value = t(p).Description
        End If
        Cells(ResultRow, 4).WrapText = True
        Cells(ResultRow, 4).Font.Bold = True
        Cells(ResultRow, 4).RowHeight = 2 * Cells(ResultRow, 4).RowHeight
        '''''''''
        Cells(ResultRow, 11).value = t(p).CostEach
        Cells(ResultRow, 12).value = t(p).CostEach * t(p).Qty
        Cells(ResultRow, 11).Font.Bold = True
        Cells(ResultRow, 12).Font.Bold = True
        Cells(ResultRow, 11).Interior.Color = vbYellow
        If t(p).CostEach > 0 Then
            Cells(ResultRow, 12).Interior.Color = vbGreen
        Else
            Cells(ResultRow, 12).Interior.Color = vbRed
        End If
        Cells(ResultRow, 4).Borders(xlEdgeLeft).LineStyle = xlDouble
        ResultRow = ResultRow + 1
        ' Displaying the real composition
            If itemRecognized Then
            For q = LBound(t(p).Items) To UBound(t(p).Items)
                Cells(ResultRow, 2).value = t(p).Items(q).reference
                Cells(ResultRow, 2).Font.Bold = True
                Cells(ResultRow, 2).HorizontalAlignment = xlRight
                Cells(ResultRow, 3).value = t(p).Items(q).Qty
                Cells(ResultRow, 2).Font.Size = Cells(ResultRow, 2).Font.Size - 1
                Cells(ResultRow, 3).Font.Size = Cells(ResultRow, 3).Font.Size - 1
                Set colRange = Range(Cells(ResultRow, 4), Cells(ResultRow, 10))
                colRange.Merge
                Cells(ResultRow, 4).value = t(p).Items(q).Description
                Cells(ResultRow, 4).WrapText = True
                Cells(ResultRow, 4).Font.Bold = True
                Cells(ResultRow, 4).Font.Italic = True
                Cells(ResultRow, 4).Font.Size = Cells(ResultRow, 4).Font.Size - 1
                Cells(ResultRow, 4).RowHeight = 2 * Cells(ResultRow, 4).RowHeight
                '''''''
                Cells(ResultRow, 11).value = t(p).Items(q).CostEach
                Cells(ResultRow, 12).value = t(p).Items(q).CostEach * t(p).Items(q).Qty
                Cells(ResultRow, 11).Font.Bold = True
                Cells(ResultRow, 12).Font.Bold = True
                Cells(ResultRow, 11).Font.Italic = True
                Cells(ResultRow, 12).Font.Italic = True
                Cells(ResultRow, 11).Font.Size = Cells(ResultRow, 10).Font.Size - 1
                Cells(ResultRow, 12).Font.Size = Cells(ResultRow, 11).Font.Size - 1
                Cells(ResultRow, 4).Borders(xlEdgeLeft).LineStyle = xlDouble
                Cells(ResultRow - 1, 2).Borders(xlEdgeBottom).LineStyle = xlDot
                Cells(ResultRow - 1, 3).Borders(xlEdgeBottom).LineStyle = xlDot
                Cells(ResultRow - 1, 4).Borders(xlEdgeBottom).LineStyle = xlDot
                Cells(ResultRow - 1, 5).Borders(xlEdgeBottom).LineStyle = xlDot
                Cells(ResultRow - 1, 6).Borders(xlEdgeBottom).LineStyle = xlDot
                Cells(ResultRow - 1, 7).Borders(xlEdgeBottom).LineStyle = xlDot
                Cells(ResultRow - 1, 8).Borders(xlEdgeBottom).LineStyle = xlDot
                Cells(ResultRow - 1, 9).Borders(xlEdgeBottom).LineStyle = xlDot
                Cells(ResultRow - 1, 11).Borders(xlEdgeBottom).LineStyle = xlDot
                Cells(ResultRow - 1, 12).Borders(xlEdgeBottom).LineStyle = xlDot
                ResultRow = ResultRow + 1
                
                Cells(ResultRow, 4).value = "Item"
                Cells(ResultRow, 5).value = "ERP cat"
                Cells(ResultRow, 6).value = "ERP"
                Cells(ResultRow, 7).value = "Length"
                Cells(ResultRow, 8).value = "Qty"
                Cells(ResultRow, 9).value = "T.Qty"
                Cells(ResultRow, 10).value = "Description"
                Cells(ResultRow, 4).Borders(xlEdgeLeft).LineStyle = xlDouble
                Cells(ResultRow, 4).Font.Italic = True
                Cells(ResultRow, 4).Font.Size = Cells(ResultRow, 4).Font.Size - 2
                Cells(ResultRow, 4).HorizontalAlignment = xlCenter
                Cells(ResultRow, 4).Font.Underline = True
                Cells(ResultRow, 5).Font.Italic = True
                Cells(ResultRow, 5).Font.Size = Cells(ResultRow, 5).Font.Size - 2
                Cells(ResultRow, 5).HorizontalAlignment = xlCenter
                Cells(ResultRow, 5).Font.Underline = True
                Cells(ResultRow, 6).Font.Italic = True
                Cells(ResultRow, 6).Font.Size = Cells(ResultRow, 6).Font.Size - 2
                Cells(ResultRow, 6).HorizontalAlignment = xlCenter
                Cells(ResultRow, 6).Font.Underline = True
                Cells(ResultRow, 7).Font.Italic = True
                Cells(ResultRow, 7).Font.Size = Cells(ResultRow, 7).Font.Size - 2
                Cells(ResultRow, 7).HorizontalAlignment = xlCenter
                Cells(ResultRow, 7).Font.Underline = True
                Cells(ResultRow, 8).Font.Italic = True
                Cells(ResultRow, 8).Font.Size = Cells(ResultRow, 8).Font.Size - 2
                Cells(ResultRow, 8).HorizontalAlignment = xlCenter
                Cells(ResultRow, 8).Font.Underline = True
                Cells(ResultRow, 9).Font.Italic = True
                Cells(ResultRow, 9).Font.Size = Cells(ResultRow, 9).Font.Size - 2
                Cells(ResultRow, 9).HorizontalAlignment = xlCenter
                Cells(ResultRow, 9).Font.Underline = True
                Cells(ResultRow, 10).Font.Italic = True
                Cells(ResultRow, 10).Font.Size = Cells(ResultRow, 10).Font.Size - 2
                Cells(ResultRow, 10).HorizontalAlignment = xlCenter
                Cells(ResultRow, 10).Font.Underline = True
                
                ResultRow = ResultRow + 1
                Dim elecPart As Byte
                Dim manlPart As Byte
                elecPart = t(p).Items(q).ElectricPart / t(p).Items(q).CostEach * 100
                manlPart = t(p).Items(q).ManlaborPart / t(p).Items(q).CostEach * 100
                Cells(ResultRow + 1, 2).value = "Inset part: " + Format(elecPart, "0") + "%"
                Cells(ResultRow + 2, 2).value = "Control part: " + Format(manlPart, "0") + "%"
                Cells(ResultRow, 2).value = "Body part: " + Format((100 - elecPart - manlPart), "0") + "%"

                Cells(ResultRow, 2).HorizontalAlignment = xlRight
                Cells(ResultRow + 1, 2).HorizontalAlignment = xlRight
                Cells(ResultRow + 2, 2).HorizontalAlignment = xlRight
                
                Dim r As Integer
                Dim InternalLoop As Integer
                InternalLoop = 0
                Dim looped As Boolean
                For r = LBound(t(p).Items(q).Items) To UBound(t(p).Items(q).Items)
                    looped = False
IL:
                    Cells(ResultRow, 4).value = t(p).Items(q).Items(r).Item & " / " & t(p).Items(q).Items(r).Category
                    Cells(ResultRow, 5).value = "item"
                    Cells(ResultRow, 6).value = t(p).Items(q).Items(r).ERP

                    Dim b As String
                    If t(p).Items(q).Items(r).length = 0 Then
                        b = t(p).Items(q).Items(r).Qty
                    Else
                        b = Format(t(p).Items(q).Items(r).length / 1000, "0.00")
                        If t(p).Items(q).Items(r).Qty > 1 And InternalLoop = 0 And Not looped Then
                            InternalLoop = t(p).Items(q).Items(r).Qty
                        End If
                    End If
                    
                    If InternalLoop > 0 Then
                        InternalLoop = InternalLoop - 1
                        If InternalLoop = 0 Then looped = True
                    End If
                    
                    Cells(ResultRow, 7).value = b
                    Cells(ResultRow, 8).value = Format3(t(p).Items(q).Items(r).Qty)
                    
                    If (t(p).Items(q).Items(r).Qty = 0) Then
                        Cells(ResultRow, 8).EntireRow.Font.Color = vbRed
                        Cells(ResultRow, 8).EntireRow.Font.Bold = True
                    End If
                    
                    Cells(ResultRow, 9).value = Format3(t(p).Items(q).Items(r).TQty)
                    Cells(ResultRow, 10).value = t(p).Items(q).Items(r).Description
                    Cells(ResultRow, 11).value = t(p).Items(q).Items(r).CostEach
                    Cells(ResultRow, 12).value = t(p).Items(q).Items(r).CostEach * t(p).Items(q).Items(r).Qty * t(p).Items(q).Qty
                    Cells(ResultRow, 11).Font.Italic = True
                    Cells(ResultRow, 12).Font.Italic = True
                    Cells(ResultRow, 11).Font.Size = Cells(ResultRow, 11).Font.Size - 2
                    Cells(ResultRow, 12).Font.Size = Cells(ResultRow, 12).Font.Size - 2
                    Cells(ResultRow, 4).Borders(xlEdgeLeft).LineStyle = xlDouble
                    Cells(ResultRow, 4).Font.Italic = True
                    Cells(ResultRow, 4).Font.Size = Cells(ResultRow, 4).Font.Size - 2
                    Cells(ResultRow, 5).Font.Italic = True
                    Cells(ResultRow, 5).Font.Size = Cells(ResultRow, 5).Font.Size - 2
                    Cells(ResultRow, 6).Font.Italic = True
                    Cells(ResultRow, 6).Font.Size = Cells(ResultRow, 6).Font.Size - 2
                    Cells(ResultRow, 7).Font.Italic = True
                    Cells(ResultRow, 7).Font.Size = Cells(ResultRow, 7).Font.Size - 2
                    Cells(ResultRow, 8).Font.Italic = True
                    Cells(ResultRow, 8).Font.Size = Cells(ResultRow, 8).Font.Size - 2
                    Cells(ResultRow, 9).Font.Italic = True
                    Cells(ResultRow, 9).Font.Size = Cells(ResultRow, 9).Font.Size - 2
                    Cells(ResultRow, 10).Font.Italic = True
                    Cells(ResultRow, 10).Font.Size = Cells(ResultRow, 10).Font.Size - 2
                    ResultRow = ResultRow + 1
                    If InternalLoop > 0 Then GoTo IL
                Next
            Next
        End If
        
        Cells(ResultRow - 1, 1).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 2).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 3).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 4).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 5).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 6).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 7).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 8).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 9).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 10).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 11).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 12).Borders(xlEdgeBottom).LineStyle = xlDouble
        
        ResultRow = ResultRow + 1
    Next
    Columns("A:K").EntireColumn.AutoFit
    Application.ScreenUpdating = True
    Columns("D").EntireColumn.ColumnWidth = Columns("D").EntireColumn.ColumnWidth + 1
    Columns("E").EntireColumn.ColumnWidth = Columns("E").EntireColumn.ColumnWidth + 1
    Columns("F").EntireColumn.ColumnWidth = Columns("F").EntireColumn.ColumnWidth + 1
    Columns("G").EntireColumn.ColumnWidth = Columns("G").EntireColumn.ColumnWidth + 1
    Columns("H").EntireColumn.ColumnWidth = Columns("H").EntireColumn.ColumnWidth + 1
    Columns("I").EntireColumn.ColumnWidth = Columns("I").EntireColumn.ColumnWidth + 1
    Columns("J").EntireColumn.ColumnWidth = Columns("J").EntireColumn.ColumnWidth + 1
End Sub

Public Sub displayBOM_Pricelist(t() As t_BOM3, ByVal refNb As String)

    Margin = Sheets("Coefficient").Cells(13, 6).value
    Mark_Up = Sheets("Coefficient").Cells(13, 17).value

    Const StartingResultRow As Integer = 3
    ' Filling "Result" sheet
    Dim colRange As Range
    Worksheets("Pricelist").Activate
    Application.ScreenUpdating = False
    Cells.Clear
    Sheets("Result Format").Cells.Copy Destination:=Sheets("Pricelist").Range("A1")
    'Cells(1, 11).value = refNb
    Dim ResultRow As Integer
    ResultRow = StartingResultRow
    Dim p As Integer
    For p = LBound(t) To UBound(t)
    
        Dim itemRecognized As Boolean
        itemRecognized = False
        On Error Resume Next
        Dim q As Integer
        q = LBound(t(p).Items)
        If Err.Number = 0 Then itemRecognized = True
    
        Cells(ResultRow, 2).value = "Reference"
        Cells(ResultRow, 3).value = "Qty"
        Cells(ResultRow, 10).value = "Unit Price"
        Cells(ResultRow, 11).value = "Total Price"
        Set colRange = Range(Cells(ResultRow, 4), Cells(ResultRow, 9))
        colRange.Merge
        Cells(ResultRow, 4).HorizontalAlignment = xlCenter
        Cells(ResultRow, 1).Interior.Color = RGB(235, 235, 235)
        Cells(ResultRow, 2).Interior.Color = RGB(235, 235, 235)
        Cells(ResultRow, 3).Interior.Color = RGB(235, 235, 235)
        Cells(ResultRow, 4).Interior.Color = RGB(235, 235, 235)
        Cells(ResultRow, 10).Interior.Color = RGB(235, 235, 235)
        Cells(ResultRow, 11).Interior.Color = RGB(235, 235, 235)
        Cells(ResultRow, 4).Font.Bold = False
        Cells(ResultRow, 4).Font.Size = Cells(ResultRow, 1).Font.Size - 1
            
        If t(p).ProvidedLength <> t(p).RequiredLength And itemRecognized Then
            Dim s As String
            s = "The item "
            Dim p1 As Integer
            p1 = Len(s)
            s = s & t(p).reference & " will be replaced by "
            Dim p2 As Integer
            p2 = Len(s)
            If resultIndex = 0 Then
                s = s & t(p).ProvidedReference & " (custom solution)"
            Else
                s = s & t(p).ProvidedReference & " (closest feasible length)"
            End If
            Cells(ResultRow, 4).value = s
            Cells(ResultRow, 4).Font.Color = vbRed
            Cells(ResultRow, 4).Characters(p1 + 1, Len(t(p).reference)).Font.Bold = True
            Cells(ResultRow, 4).Characters(p2 + 1, Len(t(p).ProvidedReference)).Font.Bold = True
        End If
        ResultRow = ResultRow + 1
            
        Cells(ResultRow, 1).value = "# " & p
        Cells(ResultRow, 1).Font.Bold = True
        If t(p).ProvidedLength <> t(p).RequiredLength And itemRecognized Then
            Cells(ResultRow, 2).value = t(p).ProvidedReference
            Cells(ResultRow, 2).Characters(5, Len(CStr(getAllParameters(t(p).ProvidedReference).length))).Font.Color = vbRed
        Else
            Cells(ResultRow, 2).value = t(p).reference
        End If
        Cells(ResultRow, 2).Font.Bold = True
        
        Cells(ResultRow, 2).HorizontalAlignment = xlLeft
        Cells(ResultRow, 3).value = t(p).Qty
        Cells(ResultRow, 3).Font.Bold = True
        Set colRange = Range(Cells(ResultRow, 4), Cells(ResultRow, 9))
        colRange.Merge
        
        If t(p).ProvidedLength <> t(p).RequiredLength And itemRecognized Then
            Cells(ResultRow, 4).value = t(p).ProvidedDescription
        Else
            Cells(ResultRow, 4).value = t(p).Description
        End If
        Cells(ResultRow, 4).WrapText = True
        Cells(ResultRow, 4).Font.Bold = True
        Cells(ResultRow, 4).RowHeight = 2 * Cells(ResultRow, 4).RowHeight
        '''''''''
        Cells(ResultRow, 10).value = t(p).CostEach * Mark_Up / Margin
        Cells(ResultRow, 11).value = t(p).CostEach * t(p).Qty * Mark_Up / Margin
        Cells(ResultRow, 10).Font.Bold = True
        Cells(ResultRow, 11).Font.Bold = True
        Cells(ResultRow, 10).Interior.Color = vbYellow
        If t(p).CostEach > 0 Then
            Cells(ResultRow, 11).Interior.Color = vbGreen
        Else
            Cells(ResultRow, 11).Interior.Color = vbRed
        End If
        Cells(ResultRow, 4).Borders(xlEdgeLeft).LineStyle = xlDouble
        ResultRow = ResultRow + 1
        ' Displaying the real composition
        If itemRecognized Then
            For q = LBound(t(p).Items) To UBound(t(p).Items)
                Cells(ResultRow, 2).value = t(p).Items(q).dashedReference
                Cells(ResultRow, 2).Font.Bold = True
                Cells(ResultRow, 2).HorizontalAlignment = xlRight
                Cells(ResultRow, 3).value = t(p).Items(q).Qty
                Cells(ResultRow, 2).Font.Size = Cells(ResultRow, 2).Font.Size - 1
                Cells(ResultRow, 3).Font.Size = Cells(ResultRow, 3).Font.Size - 1
                Set colRange = Range(Cells(ResultRow, 4), Cells(ResultRow, 9))
                colRange.Merge
                Cells(ResultRow, 4).value = t(p).Items(q).Description
                Cells(ResultRow, 4).WrapText = True
                Cells(ResultRow, 4).Font.Bold = True
                Cells(ResultRow, 4).Font.Italic = True
                Cells(ResultRow, 4).Font.Size = Cells(ResultRow, 4).Font.Size - 1
                Cells(ResultRow, 4).RowHeight = 2 * Cells(ResultRow, 4).RowHeight
                '''''''
                Cells(ResultRow, 10).value = t(p).Items(q).CostEach * Mark_Up / Margin
                Cells(ResultRow, 11).value = t(p).Items(q).CostEach * t(p).Items(q).Qty * Mark_Up / Margin
                Cells(ResultRow, 10).Font.Bold = True
                Cells(ResultRow, 11).Font.Bold = True
                Cells(ResultRow, 10).Font.Italic = True
                Cells(ResultRow, 11).Font.Italic = True
                Cells(ResultRow, 10).Font.Size = Cells(ResultRow, 10).Font.Size - 1
                Cells(ResultRow, 11).Font.Size = Cells(ResultRow, 11).Font.Size - 1
                Cells(ResultRow, 4).Borders(xlEdgeLeft).LineStyle = xlDouble
                Cells(ResultRow - 1, 2).Borders(xlEdgeBottom).LineStyle = xlDot
                Cells(ResultRow - 1, 3).Borders(xlEdgeBottom).LineStyle = xlDot
                Cells(ResultRow - 1, 4).Borders(xlEdgeBottom).LineStyle = xlDot
                Cells(ResultRow - 1, 5).Borders(xlEdgeBottom).LineStyle = xlDot
                Cells(ResultRow - 1, 6).Borders(xlEdgeBottom).LineStyle = xlDot
                Cells(ResultRow - 1, 7).Borders(xlEdgeBottom).LineStyle = xlDot
                Cells(ResultRow - 1, 8).Borders(xlEdgeBottom).LineStyle = xlDot
                Cells(ResultRow - 1, 9).Borders(xlEdgeBottom).LineStyle = xlDot
                Cells(ResultRow - 1, 10).Borders(xlEdgeBottom).LineStyle = xlDot
                Cells(ResultRow - 1, 11).Borders(xlEdgeBottom).LineStyle = xlDot
                ResultRow = ResultRow + 1
                
'                Cells(ResultRow, 4).value = "Item"
'                Cells(ResultRow, 5).value = "ERP"
'                Cells(ResultRow, 6).value = "Length"
'                Cells(ResultRow, 7).value = "Qty"
'                Cells(ResultRow, 8).value = "T.Qty"
'                Cells(ResultRow, 9).value = "Description"
'                Cells(ResultRow, 4).Borders(xlEdgeLeft).LineStyle = xlDouble
'                Cells(ResultRow, 4).Font.Italic = True
'                Cells(ResultRow, 4).Font.Size = Cells(ResultRow, 4).Font.Size - 2
'                Cells(ResultRow, 4).HorizontalAlignment = xlCenter
'                Cells(ResultRow, 4).Font.Underline = True
'                Cells(ResultRow, 5).Font.Italic = True
'                Cells(ResultRow, 5).Font.Size = Cells(ResultRow, 5).Font.Size - 2
'                Cells(ResultRow, 5).HorizontalAlignment = xlCenter
'                Cells(ResultRow, 5).Font.Underline = True
'                Cells(ResultRow, 6).Font.Italic = True
'                Cells(ResultRow, 6).Font.Size = Cells(ResultRow, 6).Font.Size - 2
'                Cells(ResultRow, 6).HorizontalAlignment = xlCenter
'                Cells(ResultRow, 6).Font.Underline = True
'                Cells(ResultRow, 7).Font.Italic = True
'                Cells(ResultRow, 7).Font.Size = Cells(ResultRow, 7).Font.Size - 2
'                Cells(ResultRow, 7).HorizontalAlignment = xlCenter
'                Cells(ResultRow, 7).Font.Underline = True
'                Cells(ResultRow, 8).Font.Italic = True
'                Cells(ResultRow, 8).Font.Size = Cells(ResultRow, 8).Font.Size - 2
'                Cells(ResultRow, 8).HorizontalAlignment = xlCenter
'                Cells(ResultRow, 8).Font.Underline = True
'                Cells(ResultRow, 9).Font.Italic = True
'                Cells(ResultRow, 9).Font.Size = Cells(ResultRow, 9).Font.Size - 2
'                Cells(ResultRow, 9).HorizontalAlignment = xlCenter
'                Cells(ResultRow, 9).Font.Underline = True
                
'                ResultRow = ResultRow + 1
'                Dim elecPart As Byte
'                Dim manlPart As Byte
'                elecPart = t(p).Items(q).ElectricPart / t(p).Items(q).CostEach * 100
'                manlPart = t(p).Items(q).ManlaborPart / t(p).Items(q).CostEach * 100
'                Cells(ResultRow + 1, 2).value = "Electrical part: " + Format(elecPart, "0") + "%"
'                Cells(ResultRow + 2, 2).value = "Manlabor part: " + Format(manlPart, "0") + "%"
'                Cells(ResultRow, 2).value = "Mechanical part: " + Format((100 - elecPart - manlPart), "0") + "%"

'                Cells(ResultRow, 2).HorizontalAlignment = xlRight
'                Cells(ResultRow + 1, 2).HorizontalAlignment = xlRight
'                Cells(ResultRow + 2, 2).HorizontalAlignment = xlRight
                
'                Dim r As Integer
'                For r = LBound(t(p).Items(q).Items) To UBound(t(p).Items(q).Items)
'                    Cells(ResultRow, 4).value = t(p).Items(q).Items(r).Item & " / " & t(p).Items(q).Items(r).Category
'                    Cells(ResultRow, 5).value = t(p).Items(q).Items(r).ERP
'                    Cells(ResultRow, 6).value = LengthToDisplay(t(p).Items(q).Items(r).length)
'                    Cells(ResultRow, 7).value = Format3(t(p).Items(q).Items(r).Qty)
'                    If (t(p).Items(q).Items(r).Qty = 0) Then
'                        Cells(ResultRow, 7).EntireRow.Font.Color = vbRed
'                        Cells(ResultRow, 7).EntireRow.Font.Bold = True
'                    End If
'                    Cells(ResultRow, 8).value = Format3(t(p).Items(q).Items(r).TQty)
'                    Cells(ResultRow, 9).value = t(p).Items(q).Items(r).Description
'                    Cells(ResultRow, 10).value = t(p).Items(q).Items(r).CostEach
'                    Cells(ResultRow, 11).value = t(p).Items(q).Items(r).CostEach * t(p).Items(q).Items(r).Qty * t(p).Items(q).Qty
'                    Cells(ResultRow, 10).Font.Italic = True
'                    Cells(ResultRow, 11).Font.Italic = True
'                    Cells(ResultRow, 10).Font.Size = Cells(ResultRow, 10).Font.Size - 2
'                    Cells(ResultRow, 11).Font.Size = Cells(ResultRow, 11).Font.Size - 2
'                    Cells(ResultRow, 4).Borders(xlEdgeLeft).LineStyle = xlDouble
'                    Cells(ResultRow, 4).Font.Italic = True
'                    Cells(ResultRow, 4).Font.Size = Cells(ResultRow, 4).Font.Size - 2
'                    Cells(ResultRow, 5).Font.Italic = True
'                    Cells(ResultRow, 5).Font.Size = Cells(ResultRow, 5).Font.Size - 2
'                    Cells(ResultRow, 6).Font.Italic = True
'                    Cells(ResultRow, 6).Font.Size = Cells(ResultRow, 6).Font.Size - 2
'                    Cells(ResultRow, 7).Font.Italic = True
'                    Cells(ResultRow, 7).Font.Size = Cells(ResultRow, 7).Font.Size - 2
'                    Cells(ResultRow, 8).Font.Italic = True
'                    Cells(ResultRow, 8).Font.Size = Cells(ResultRow, 8).Font.Size - 2
'                    Cells(ResultRow, 9).Font.Italic = True
'                    Cells(ResultRow, 9).Font.Size = Cells(ResultRow, 9).Font.Size - 2
'                    ResultRow = ResultRow + 1
'                Next
            Next
        End If
        
        Cells(ResultRow - 1, 1).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 2).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 3).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 4).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 5).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 6).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 7).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 8).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 9).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 10).Borders(xlEdgeBottom).LineStyle = xlDouble
        Cells(ResultRow - 1, 11).Borders(xlEdgeBottom).LineStyle = xlDouble
        ResultRow = ResultRow + 1
    Next
    Columns("A:K").EntireColumn.AutoFit
    Application.ScreenUpdating = True
    Columns("D").EntireColumn.ColumnWidth = Columns("D").EntireColumn.ColumnWidth + 1
    Columns("E").EntireColumn.ColumnWidth = Columns("E").EntireColumn.ColumnWidth + 1
    Columns("F").EntireColumn.ColumnWidth = Columns("F").EntireColumn.ColumnWidth + 1
    Columns("G").EntireColumn.ColumnWidth = Columns("G").EntireColumn.ColumnWidth + 1
    Columns("H").EntireColumn.ColumnWidth = Columns("H").EntireColumn.ColumnWidth + 1
    Columns("I").EntireColumn.ColumnWidth = Columns("I").EntireColumn.ColumnWidth + 1
End Sub

