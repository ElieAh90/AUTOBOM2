Attribute VB_Name = "getBOM"
    
Type t_BOM1 ' BOM
    Category As String
    Item As String
    ERP As String
    Description As String
    length As Single
    Qty As Single
    TQty As Single
    QtyRelatedToDriver As Boolean
    CostEach As Single
    MultiplierGA As Single
    MultiplierLB As Single
End Type

Type t_BOM2 ' Fixture
    reference As String
    dashedReference As String
    Description As String
    Qty As Integer
    CostEach As Single
    Items() As t_BOM1
    MechanicalPart As Single
    ElectricPart As Single
    ManlaborPart As Single
    Qty_Driver As Integer
    Qty_PCB As Integer
    Qty_Optic_Lens As Integer
    Qty_Optic_Diffuser As Integer
    Qty_Optic_Reflector As Integer
    Qty_Optic_Kick_Reflector As Integer
    Qty_Optic_Fresnel As Integer
    HasHalfFoot As Boolean
End Type

Type t_BOM3     ' Run
    reference As String
    Description As String
    RequiredLength As Integer
    ProvidedLength As Integer
    ProvidedReference As String
    'ProvidedDashedReference As String
    ProvidedDescription As String
    Qty As Integer
    CostEach As Single
    Items() As t_BOM2
    Qty_Driver As Integer
    Qty_PCB As Integer
    Qty_Optic_Lens As Integer
    Qty_Optic_Diffuser As Integer
    Qty_Optic_Reflector As Integer
    Qty_Optic_Kick_Reflector As Integer
    Qty_Optic_Fresnel As Integer
    
End Type

Type t_Parameters
    Family As String
    FType As String
    Mounting As String
    length As Integer
    BodyFinish As String
    OutputPower As String
    Voltage As String
    Dimming As String
    Baffles_Diffuser As String
    BeamAngle As String
    CRI As String
    CCT As String
    Emergency As String
    wiring As String
End Type

Public Function addItem(runReference As String, Qty As Integer, ByRef list() As t_BOM3) As Integer
    Dim np As Integer
    Dim param As t_Parameters
    On Error Resume Next
    np = UBound(list) + 1
    If Err.Number <> 0 Then np = 1
    ReDim Preserve list(1 To np)
    runReference = cleanPartNumber(runReference)
    runReference = UCase(runReference)
    param = getAllParameters(runReference)
    list(np).reference = runReference
    list(np).Qty = Qty
    list(np).Description = getDescriptionOf(runReference, False)
    list(np).RequiredLength = param.length
    list(np).ProvidedLength = 0
    If list(np).Description <> "" Then
        list(np).Items = splitLength.splitLengthOf(runReference)
        Dim i As Integer
        list(np).Qty_Driver = 0
        list(np).Qty_PCB = 0
        list(np).Qty_Optic_Diffuser = 0
        list(np).Qty_Optic_Fresnel = 0
        list(np).Qty_Optic_Kick_Reflector = 0
        list(np).Qty_Optic_Lens = 0
        list(np).Qty_Optic_Reflector = 0
        
        For i = LBound(list(np).Items) To UBound(list(np).Items)
            Call getBOM(list(np).Items(i), list(np).Qty)
            list(np).ProvidedLength = list(np).ProvidedLength + getAllParameters(list(np).Items(i).reference).length * list(np).Items(i).Qty
            list(np).Qty_Driver = list(np).Qty_Driver + list(np).Items(i).Qty_Driver
            list(np).Qty_PCB = list(np).Qty_PCB + list(np).Items(i).Qty_PCB
            list(np).Qty_Optic_Diffuser = list(np).Qty_Optic_Diffuser + list(np).Items(i).Qty_Optic_Diffuser
            list(np).Qty_Optic_Fresnel = list(np).Qty_Optic_Fresnel + list(np).Items(i).Qty_Optic_Fresnel
            list(np).Qty_Optic_Kick_Reflector = list(np).Qty_Optic_Kick_Reflector + list(np).Items(i).Qty_Optic_Kick_Reflector
            list(np).Qty_Optic_Lens = list(np).Qty_Optic_Lens + list(np).Items(i).Qty_Optic_Lens
            list(np).Qty_Optic_Reflector = list(np).Qty_Optic_Reflector + list(np).Items(i).Qty_Optic_Reflector
            ''''
            list(np).Items(i).HasHalfFoot = False
            If getAllParameters(list(np).Items(i).reference).length Mod 12 <> 0 Then list(np).Items(i).HasHalfFoot = True
            ''''
        Next i
        

        
        list(np).ProvidedReference = changeLengthTo(list(np).reference, list(np).ProvidedLength)
        'list(np).ProvidedDashedReference = assembleReference(getAllParameters(list(np).ProvidedReference), -1, "", True)
        list(np).ProvidedDescription = getDescriptionOf(list(np).ProvidedReference, False)
        Call updateCostOfRun(list(np))
        addItem = 1
    Else
        list(np).Description = "Not recognized"
        addItem = -1
    End If
End Function

Public Function cleanPartNumber(reference As String) As String
    ' Remove all undesired characters
    reference = Replace(reference, " ", "")
    reference = Replace(reference, "-", "")
    reference = Replace(reference, "_", "")
    reference = Replace(reference, "/", "")
    reference = Replace(reference, ".", "")
    reference = Replace(reference, ",", "")
    reference = Replace(reference, ";", "")
    reference = Replace(reference, "'", "")
    reference = Replace(reference, "\", "")
    reference = Replace(reference, "|", "")
    reference = Replace(reference, "*", "")
    reference = Replace(reference, "+", "")
    reference = Replace(reference, "(", "")
    reference = Replace(reference, ")", "")
    reference = Replace(reference, "=", "")
    reference = Replace(reference, "&", "")
    reference = Replace(reference, "^", "")
    reference = Replace(reference, "%", "")
    reference = Replace(reference, "$", "")
    reference = Replace(reference, "#", "")
    reference = Replace(reference, "@", "")
    reference = Replace(reference, "!", "")
    reference = Replace(reference, "'", "")
    reference = Replace(reference, """", "")
    cleanPartNumber = reference
End Function

Public Function getAllParameters(reference As String) As t_Parameters
    Dim res As t_Parameters
    res.Family = Mid(reference, 1, 1)
    res.FType = Mid(reference, 1, 3)
    res.Mounting = Mid(reference, 4, 1)
    Dim m As Integer
    m = 4
    Do
        m = m + 1
    Loop Until Not IsNumeric(Mid(reference, m, 1))
    res.length = Mid(reference, 5, m - 5)
    res.BodyFinish = Mid(reference, m, 1)
    res.OutputPower = Mid(reference, m + 1, 1)
    res.Voltage = Mid(reference, m + 2, 1)
    res.Dimming = Mid(reference, m + 3, 1)
    res.Baffles_Diffuser = Mid(reference, m + 4, 1)
    res.BeamAngle = Mid(reference, m + 5, 1)
    res.CRI = Mid(reference, m + 6, 1)
    res.CCT = Mid(reference, m + 7, 2)
 
    Select Case Len(Mid(reference, m))
        Case 11:
            res.Emergency = Mid(reference, m + 9, 1)
            res.wiring = Mid(reference, m + 10, 1)
        Case 10:
            Dim temp As String
            temp = Mid(reference, m + 9, 1)
            If IsNumeric(temp) Then
                res.Emergency = temp
                res.wiring = "S"        ' Single
            Else
                res.Emergency = "0"     ' No emergency
                res.wiring = temp
            End If
        Case 9:
            res.Emergency = "0"         ' No Emergency
            res.wiring = "S"            ' Single
    End Select
    getAllParameters = res
    
End Function

Public Function getBOM(ByRef ref As t_BOM2, oq As Integer)
    On Error Resume Next
    Dim p As t_Parameters
    p = getAllParameters(ref.reference)
    Dim row As Integer
    Dim LastDataRow As Integer
    Dim np As Integer
    Dim component As String
    Dim rQty As Single
    Dim v16 As String
    Dim v19 As String
    Worksheets("Database").Activate
    LastDataRow = Cells(Rows.Count, 1).End(xlUp).row
    np = 0
    ref.Qty_Driver = 0
    ref.Qty_PCB = 0
    ref.Qty_Optic_Diffuser = 0
    ref.Qty_Optic_Fresnel = 0
    ref.Qty_Optic_Kick_Reflector = 0
    ref.Qty_Optic_Lens = 0
    ref.Qty_Optic_Reflector = 0
    
    For row = 4 To LastDataRow
        
        If isStated(p.FType, Cells(row, 5)) And isStated(p.Mounting, Cells(row, 6)) And isStated(p.wiring, Cells(row, 7)) _
        And isStated(CStr(p.length), Cells(row, 8)) And isStated(p.OutputPower, Cells(row, 9)) And isStated(p.Voltage, Cells(row, 10)) _
        And isStated(p.Dimming, Cells(row, 11)) And isStated(p.Baffles_Diffuser, Cells(row, 12)) And isStated(p.BeamAngle, Cells(row, 13)) _
        And isStated(p.CRI, Cells(row, 14)) And isStated(p.CCT, Cells(row, 15)) And isStated(p.BodyFinish, Cells(row, 16)) And Cells(row, 2) <> "" And Cells(row, 4) <> "" Then
            
            Dim d As Integer
            d = ItemAlreadyInBOM(ref.Items, Cells(row, 3))
            If d > 0 Then
                np = d
            Else
                np = np + 1
                ReDim Preserve ref.Items(1 To np)
                ref.Items(np).Qty = 0
            End If

            ref.Items(np).ERP = Cells(row, 3).value
            ref.Items(np).Item = Cells(row, 2).value
            ref.Items(np).Category = Cells(row, 4).value
            
            If UCase(ref.Items(np).Category) = "DRIVERS" Then ref.Qty_Driver = ref.Qty_Driver + 1
            If UCase(ref.Items(np).Category) = "PCB" Then ref.Qty_PCB = ref.Qty_PCB + 1
            If UCase(ref.Items(np).Category) = "OPTIC_DIFFUSER" Then ref.Qty_Optic_Diffuser = ref.Qty_Optic_Diffuser + 1
            If UCase(ref.Items(np).Category) = "OPTIC_FRESNEL" Then ref.Qty_Optic_Fresnel = ref.Qty_Optic_Fresnel + 1
            If UCase(ref.Items(np).Category) = "OPTIC_KICK_REFLECTOR" Then ref.Qty_Optic_Kick_Reflector = ref.Qty_Optic_Kick_Reflector + 1
            If UCase(ref.Items(np).Category) = "OPTIC_LENS" Then ref.Qty_Optic_Lens = ref.Qty_Optic_Lens + 1
            If UCase(ref.Items(np).Category) = "OPTIC_REFLECTOR" Then ref.Qty_Optic_Reflector = ref.Qty_Optic_Reflector + 1
            
            Dim m As Double
            Dim temp_CE As Single
            m = CDbl(p.length / 12)
            ref.Items(np).QtyRelatedToDriver = False
            v19 = Replace(Cells(row, 20).value, " ", "")
            
            If v19 = "" Then v19 = "0"
            
            rQty = CDbl(v19)
            v16 = Replace(Cells(row, 17).value, " ", "")
            
            If UCase(v16) = "PC" Then
                ' checking if the quantity is driver related
                Dim sp As String
                If Len(v19) >= 3 Then
                    sp = Mid(v19, Len(v19) - 1, 2)
                    If UCase(sp) = "/D" Then
                        ref.Items(np).QtyRelatedToDriver = True
                        rQty = CDbl(Mid(v19, 1, Len(v19) - 2))
                    End If
                End If
                '''
                ref.Items(np).Qty = ref.Items(np).Qty + CDbl(Cells(row, 19).value) * Int(p.length / 12) + rQty
                ref.Items(np).CostEach = Cells(row, 21).value
            Else
                ref.Items(np).length = Cells(row, 19).value * m + v19
                ref.Items(np).Qty = Cells(row, 18).value
                If ref.Items(np).Qty = 0 Then ref.Items(np).Qty = 1
                If Cells(row, 19).value = 0 Then                        ' calculation here
                     ref.Items(np).CostEach = Cells(row, 21).value
                Else
                    ref.Items(np).CostEach = Cells(row, 21).value * (m * Cells(row, 19).value + Cells(row, 20).value) / 300
                End If
                
            
            End If
            
            ref.Items(np).MultiplierGA = Cells(row, 26).value
            ref.Items(np).MultiplierLB = Cells(row, 27).value
            
            temp_CE = ref.Items(np).CostEach * ref.Items(np).Qty
            ref.CostEach = ref.CostEach + temp_CE
            If InStr(1, ref.Items(np).Category, "Aluminum, 1.5mm", vbTextCompare) Then ref.ElectricPart = ref.ElectricPart + temp_CE
            If InStr(1, ref.Items(np).Category, "Accessories", vbTextCompare) Then ref.ElectricPart = ref.ElectricPart + temp_CE
            If InStr(1, ref.Items(np).Category, "Aluminum, 2mm", vbTextCompare) Then ref.ElectricPart = ref.ElectricPart + temp_CE
            If InStr(1, ref.Items(np).Category, "Cables with Connectors", vbTextCompare) Then ref.ElectricPart = ref.ElectricPart + temp_CE
            If InStr(1, ref.Items(np).Category, "Connectors", vbTextCompare) Then ref.ElectricPart = ref.ElectricPart + temp_CE
            If InStr(1, ref.Items(np).Category, "Diffusers", vbTextCompare) Then ref.ElectricPart = ref.ElectricPart + temp_CE
            If InStr(1, ref.Items(np).Category, "Optic_Diffuser", vbTextCompare) Then ref.ElectricPart = ref.ElectricPart + temp_CE
            If InStr(1, ref.Items(np).Category, "Optic_Fresnel", vbTextCompare) Then ref.ElectricPart = ref.ElectricPart + temp_CE
            If InStr(1, ref.Items(np).Category, "Optic_Kick_Reflector", vbTextCompare) Then ref.ElectricPart = ref.ElectricPart + temp_CE
            If InStr(1, ref.Items(np).Category, "Optic_Lens", vbTextCompare) Then ref.ElectricPart = ref.ElectricPart + temp_CE
            If InStr(1, ref.Items(np).Category, "Optic_Reflector", vbTextCompare) Then ref.ElectricPart = ref.ElectricPart + temp_CE
            If InStr(1, ref.Items(np).Category, "PCB", vbTextCompare) Then ref.ElectricPart = ref.ElectricPart + temp_CE
            If InStr(1, ref.Items(np).Category, "Wires", vbTextCompare) Then ref.ElectricPart = ref.ElectricPart + temp_CE
            If InStr(1, ref.Items(np).Category, "Cables", vbTextCompare) Then ref.ElectricPart = ref.ElectricPart + temp_CE
            
            If InStr(1, ref.Items(np).Category, "Drivers", vbTextCompare) Then ref.ManlaborPart = ref.ManlaborPart + temp_CE
            ref.Items(np).Description = Cells(row, 22).value

        End If
    Next row

    ' Checking if any item is driver related to update its quantity
    Dim i As Integer
    Dim drivers As Integer
    drivers = 0
    ' Calculating the number of drivers
    For i = LBound(ref.Items) To UBound(ref.Items)
        'If InStr(1, ref.Items(i).Category, "Elec", vbTextCompare) And InStr(1, ref.Items(i).Description, "Drivers", vbTextCompare) Then
         If InStr(1, ref.Items(i).Category, "Drivers", vbTextCompare) Then
            drivers = drivers + ref.Items(i).Qty
        End If
    Next i
    ' updating the quantity if necessary
    For i = LBound(ref.Items) To UBound(ref.Items)
        If ref.Items(i).QtyRelatedToDriver Then
            ref.Items(i).Qty = ref.Items(i).Qty * drivers
        End If
    Next i
    
    ' loading the T.Qty
    For i = LBound(ref.Items) To UBound(ref.Items)
        ref.Items(i).TQty = ref.Items(i).Qty * ref.Qty * oq
    Next i
    
End Function

Public Function ItemAlreadyInBOM(ref() As t_BOM1, s As String)
    ItemAlreadyInBOM = 0
    Dim i As Integer
    On Error GoTo out
    '''
    If s = "" Then GoTo out
    '''
    For i = LBound(ref) To UBound(ref)
        If ref(i).ERP = s Then
            ItemAlreadyInBOM = i
            Exit Function
        End If
    Next i
out:
    ItemAlreadyInBOM = 0
End Function

Public Sub updateCostOfRun(ByRef ref As t_BOM3)
    Dim i As Integer
    For i = LBound(ref.Items) To UBound(ref.Items)
        ref.CostEach = ref.CostEach + ref.Items(i).CostEach * ref.Items(i).Qty
    Next i
End Sub

Public Function isStated(data As String, source As String) As Boolean
    Dim sp() As String
    Dim o As Variant
    isStated = False
    If source = "" Then
        isStated = True
        Exit Function
    End If
    source = Replace(source, ",", " ")
    source = Replace(source, "/", " ")
    source = Replace(source, "|", " ")
    source = Replace(source, ";", " ")
    sp = Split(source)
    For Each o In sp
        If UCase(data) = UCase(o) Then
            isStated = True
            Exit Function
        End If
    Next o
End Function
