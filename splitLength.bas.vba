Attribute VB_Name = "splitLength"

Type t_SplitResult
    length As Integer
    Qty As Integer
    wiring As String
End Type

    Public lengthsCount As Integer
    Public usedStandards() As Integer
    Public deltaLength() As Integer
    Public fixturesCount() As Integer
    Public usedLengthNumber() As Integer
    Public resultIndex As Integer
    Public RequiredLength As Integer
    Public customResult() As t_SplitResult

Public Function splitLengthOf(partNumber As String) As t_BOM2()
    Dim res() As t_SplitResult
    Dim param As t_Parameters
    param = getAllParameters(partNumber)
    res = splitLength(param.length, param.wiring)
    Dim bom() As t_BOM2
    ReDim bom(1 To 1)
    Dim i As Integer
    For i = LBound(res) To UBound(res)
        ReDim Preserve bom(1 To i)
        bom(i).Qty = res(i).Qty
        bom(i).reference = assembleReference(param, res(i).length, res(i).wiring, False)
        bom(i).dashedReference = assembleReference(param, res(i).length, res(i).wiring, True)
        bom(i).Description = getDescriptionOf(bom(i).reference, True)
    Next i
    
    splitLengthOf = bom
    
End Function

Public Function assembleReference(param As t_Parameters, length As Integer, wiring As String, WithDashes As Boolean) As String
    Dim s As String
    s = param.FType
    If WithDashes Then s = s + "-"
    s = s + param.Mounting
    If length <= 0 Then
        s = s + CStr(param.length)
    Else
        s = s + CStr(length)
    End If
    s = s + param.BodyFinish
    If WithDashes Then s = s + "-"
    s = s + param.OutputPower
    s = s + param.Voltage
    s = s + param.Dimming
    If WithDashes Then s = s + "-"
    s = s + param.Baffles_Diffuser
    If WithDashes Then s = s + "-"
    s = s + param.BeamAngle
    s = s + param.CRI
    s = s + param.CCT
    If WithDashes Then s = s + "-"
    s = s + param.Emergency
    If wiring <> "" Then
        s = s + wiring
    Else
        s = s + param.wiring
    End If
    assembleReference = s
End Function

Public Function changeLengthTo(reference As String, length As Integer) As String
    changeLengthTo = Mid(reference, 1, 4) & CStr(length)
    Dim m As Integer
    m = 4
    Do
        m = m + 1
    Loop Until Not IsNumeric(Mid(reference, m, 1))
    changeLengthTo = changeLengthTo & Mid(reference, m)
End Function

Public Function splitLength(runLength As Integer, wiring As String) As t_SplitResult()
    
    RequiredLength = runLength
    
    Dim standards(1 To 10) As Integer
    standards(1) = 6
    standards(2) = 12
    standards(3) = 18
    standards(4) = 24
    standards(5) = 30
    standards(6) = 36
    standards(7) = 42
    standards(8) = 48
    standards(9) = 54
    standards(10) = 60
    
    If runLength < standards(1) Then runLength = standards(1)
    
'    Dim lengthsCount As Integer
    lengthsCount = UBound(standards) - LBound(standards) + 1
    
    Dim minPieces() As Single
    ReDim minPieces(1 To lengthsCount)
    
    Dim i, j, k As Integer
    For i = 1 To lengthsCount
        minPieces(i) = runLength / standards(i)
    Next
    
'    Dim usedStandards() As Integer
    ReDim usedStandards(1 To lengthsCount, 1 To lengthsCount)

    'filling the diagonal values
    For i = 1 To lengthsCount
        usedStandards(i, lengthsCount - i + 1) = Int(minPieces(lengthsCount - i + 1))
    Next
'    Dim deltaLength() As Integer
'    Dim fixturesCount() As Integer
'    Dim usedLengthNumber() As Integer
    ReDim deltaLength(1 To lengthsCount) As Integer
    ReDim fixturesCount(1 To lengthsCount) As Integer
    ReDim usedLengthNumber(1 To lengthsCount) As Integer
    
    'filling the remaining values (at left in the table)
    For i = 1 To lengthsCount
        ' calculating deltaLength(i)
        deltaLength(i) = runLength
        For j = 1 To lengthsCount
            deltaLength(i) = deltaLength(i) - usedStandards(i, j) * standards(j)
        Next j
        ' selecting the other lengths
        While deltaLength(i) >= standards(1)
            For k = lengthsCount - i To 1 Step -1
                If (deltaLength(i) >= standards(k)) Then
                    usedStandards(i, k) = usedStandards(i, k) + 1
                    deltaLength(i) = deltaLength(i) - standards(k)
                    Exit For
                End If
            Next k
        Wend
        ' calculating the fixturesCount array
        fixturesCount(i) = 0
        For j = 1 To lengthsCount
            fixturesCount(i) = fixturesCount(i) + usedStandards(i, j)
        Next j
        ' calculating the usedLengthNumber array
        usedLengthNumber(i) = 0
        For j = 1 To lengthsCount
            If usedStandards(i, j) > 0 Then usedLengthNumber(i) = usedLengthNumber(i) + 1
        Next j
    Next i
    
    ' searching for the optimum solution
'    Dim resultIndex As Integer
    
    Dim smallestDeltaLength As Integer
    smallestDeltaLength = deltaLength(1)
    For i = 2 To lengthsCount
        If deltaLength(i) < smallestDeltaLength Then smallestDeltaLength = deltaLength(i)
    Next i
    
    Dim smallestFixturesCount As Integer
    smallestFixturesCount = 30000  ' initialize it with a big number
    For i = 1 To lengthsCount
        If fixturesCount(i) < smallestFixturesCount And deltaLength(i) = smallestDeltaLength Then smallestFixturesCount = fixturesCount(i)
    Next i
    
    Dim smallestUsedLengthNumber As Integer
    smallestUsedLengthNumber = usedLengthNumber(1)
    For i = 2 To lengthsCount
        If usedLengthNumber(i) < smallestUsedLengthNumber Then smallestUsedLengthNumber = usedLengthNumber(i)
    Next i
      
    For i = 1 To lengthsCount
        If deltaLength(i) = smallestDeltaLength Then
            resultIndex = i
            Exit For
        End If
    Next i
    
    ' a better split (equal length) might be available
    If usedLengthNumber(resultIndex) > smallestUsedLengthNumber Then
        j = resultIndex + 1
        k = resultIndex + 4 ' max lines after the already selected one
        If j > lengthsCount Then j = lengthsCount
        If k > lengthsCount Then k = lengthsCount
        For i = j To k
            If deltaLength(i) = smallestDeltaLength And usedLengthNumber(i) = smallestUsedLengthNumber And fixturesCount(i) <= (fixturesCount(resultIndex) + 5) Then
                resultIndex = i
                Exit For
            End If
        Next i
    End If

    If Sheets("Main").IsAutoSelectSingleStandard And Sheets("Main").FixturesUsedIn(resultIndex) = 1 Then GoTo SkipSplit
    If Sheets("Main").IsAutoSelectAllReferences Then GoTo SkipSplit

    frm_RunSplit.Show vbModal
        
SkipSplit:
    ' prepare the result in a detailed table
    Dim result() As t_SplitResult
    If resultIndex = 0 Then
        ' Custom solution
        result = customResult
    Else
        ' One of the pre-generated solutions
        ReDim result(1 To usedLengthNumber(resultIndex))
        j = 1
        For i = lengthsCount To 1 Step -1
            If usedStandards(resultIndex, i) > 0 Then
                result(j).length = standards(i)
                result(j).Qty = usedStandards(resultIndex, i)
                j = j + 1
            End If
        Next
    End If
    
    Call splitForInstallation(result, wiring)
    
    splitLength = result

End Function

Public Function splitForInstallation(ByRef g() As t_SplitResult, Optional wiring As String = "")

    Dim fixturesCount As Integer
    Dim i, j As Integer

    fixturesCount = 0
    For i = LBound(g) To UBound(g)
        fixturesCount = fixturesCount + g(i).Qty
    Next i
    
    If fixturesCount = 0 Then
        ReDim g(0)
        Exit Function
    End If
    
    If fixturesCount = 1 Then
        If wiring <> "" Then
            g(LBound(g)).wiring = wiring
        Else
            g(LBound(g)).wiring = "S"
        End If
        Exit Function
    End If

    If fixturesCount = 2 Then
        If LBound(g) = UBound(g) Then
            ReDim Preserve g(LBound(g) To LBound(g) + 1)
            g(LBound(g)).Qty = 1
            g(LBound(g)).wiring = "B"
            g(UBound(g)).Qty = 1
            g(UBound(g)).length = g(LBound(g)).length
            g(UBound(g)).wiring = "E"
        Else
            g(LBound(g)).wiring = "B"
            g(UBound(g)).wiring = "E"
        End If
        Exit Function
    End If
    
    Dim fixtures() As Integer
    ReDim fixtures(1 To fixturesCount)
    
    k = 1
    For i = LBound(g) To UBound(g)
        For j = 1 To g(i).Qty
            fixtures(k) = g(i).length
            k = k + 1
        Next j
    Next i
    
    ReDim g(1 To 1)
    Dim g_Index As Integer
    
    ' starting with the beginning
    g_Index = 1
    g(g_Index).Qty = 1
    g(g_Index).wiring = "B"
    g(g_Index).length = fixtures(1)
    
    ' filling the middle
    For i = 2 To fixturesCount - 1
        If fixtures(i) <> g(g_Index).length Then
            g_Index = g_Index + 1
            ReDim Preserve g(1 To g_Index)
            g(g_Index).Qty = 1
            g(g_Index).length = fixtures(i)
            g(g_Index).wiring = "M"
        Else
            If g_Index = 1 Then
                g_Index = g_Index + 1
                ReDim Preserve g(1 To g_Index)
                g(g_Index).Qty = 1
                g(g_Index).length = fixtures(i)
                g(g_Index).wiring = "M"
            Else
                g(g_Index).Qty = g(g_Index).Qty + 1
            End If
        End If
    Next i
    
    ' fixing the end
    g_Index = g_Index + 1
    ReDim Preserve g(1 To g_Index)
    g(g_Index).Qty = 1
    g(g_Index).length = fixtures(i)
    g(g_Index).wiring = "E"

End Function
