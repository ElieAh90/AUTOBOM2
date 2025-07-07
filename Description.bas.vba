Attribute VB_Name = "Description"

Public Function Format2(value As Double) As String
    If Int(value) = value Then
        Format2 = CStr(value)
    Else
        Format2 = Format(value, "#.#")
    End If
End Function

Public Function Format3(value As Single) As String
    If Int(value) = value Then
        Format3 = CStr(value)
    Else
        Format3 = Format(value, "#.##")
    End If
End Function

Public Function getDescriptionOf(partNumber As String, w As Boolean) As String
    
    Dim param As t_Parameters
    param = getAllParameters(partNumber)

    Dim FType As String
    Dim FEnvironment As String
    Dim FOptic As String
    Dim Mounting As String
    Dim BodyFinish As String
    Dim OutputPower As String
    Dim Voltage As String
    Dim Dimming As String
    Dim Baffles_Diffuser As String
    Dim BeamAngle As String
    Dim CRI As String
    Dim CCT As String
    Dim Emergency As String
    Dim wiring As String
    Dim LengthFoot As Double

    LengthFoot = CDbl(param.length) / 12
    
    If param.FType = "BOP" And LengthFoot > 0.5 Then LengthFoot = 0
    If param.FType <> "BOP" And LengthFoot = 0.5 Then LengthFoot = 0
    
    FType = getTypeDescription(param.FType)
    FEnvironment = getEnvironmentDescription(param.FType)
    FOptic = getOpticDescription(param.FType)
    Mounting = GetMountingDescription(param.Mounting, param.FType)
    BodyFinish = GetBodyfinishDescription(param.BodyFinish)
    OutputPower = GetOutputPowerDescription(param.OutputPower, param.FType)
    Dimming = GetDimmingDescription(param.Dimming)
    Baffles_Diffuser = GetDiffuserDescription(param.Baffles_Diffuser)
    BeamAngle = GetBeamAngleDescription(param.BeamAngle, param.FType)
    CRI = GetCRIDescription(param.CRI)
    CCT = GetCCTDescription(param.CCT)
    Emergency = GetEmergencyDescription(param.Emergency)
    Voltage = GetVoltageDescription(param.Voltage)
    wiring = GetWiringDescription(param.wiring)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ' Billet Outdoor, Recessed with Trim, Flat Black Symmetrical Baffles 50°,
    ' 6 inches, 0.5ft, Black Body Finish, 120/277 VAC, 4W/ft, Dali to 1% - Remote Driver, CRI90, 2700K-5000K, No emergency.
        
    getDescriptionOf = "Billet " & FEnvironment & ", " & Mounting
    If InStr(1, getDescriptionOf, "opal", vbTextCompare) = 0 Then getDescriptionOf = getDescriptionOf & ", " & Baffles_Diffuser
    If InStr(1, getDescriptionOf, "opal", vbTextCompare) = 0 Then getDescriptionOf = getDescriptionOf & " " & FOptic
    If InStr(1, getDescriptionOf, "opal", vbTextCompare) = 0 Then
        getDescriptionOf = getDescriptionOf & ", " & BeamAngle & ", "
    Else
        getDescriptionOf = getDescriptionOf & ", "
    End If
    getDescriptionOf = getDescriptionOf & param.length & " inch, " & LengthFoot & " ft, "
    'getDescriptionOf = getDescriptionOf & param.length & " inch, " & Format2(LengthFoot) & " ft, "
    If UCase(Baffles_Diffuser) <> "OPAL" Then getDescriptionOf = getDescriptionOf '& vbCrLf
    getDescriptionOf = getDescriptionOf & BodyFinish & " Body Finish, "
    If UCase(Baffles_Diffuser) = "OPAL" Then getDescriptionOf = getDescriptionOf '& vbCrLf
    getDescriptionOf = getDescriptionOf & Voltage & ", "
    getDescriptionOf = getDescriptionOf & OutputPower & ", " & Dimming
    getDescriptionOf = getDescriptionOf & ", CRI" & CRI & ", " & CCT & ", " & Emergency
    If w Then getDescriptionOf = getDescriptionOf & ", " & wiring
    
    'If param.length = 0 Then getDescriptionOf = ""
    If LengthFoot = 0 Then getDescriptionOf = ""
    If FType = "" Then getDescriptionOf = ""
    If Mounting = "" Then getDescriptionOf = ""
    If BodyFinish = "" Then getDescriptionOf = ""
    If OutputPower = "" Then getDescriptionOf = ""
    If Dimming = "" Then getDescriptionOf = ""
    If Baffles_Diffuser = "" Then getDescriptionOf = ""
    If BeamAngle = "" Then getDescriptionOf = ""
    If CRI = "" Then getDescriptionOf = ""
    If CCT = "" Then getDescriptionOf = ""
    If Emergency = "" Then getDescriptionOf = ""
    If Voltage = "" Then getDescriptionOf = ""
    If wiring = "" Then getDescriptionOf = ""
    
    
End Function

Public Function getTypeDescription(t As String) As String
    Dim lastRow As Long
    Dim i As Integer
    lastRow = Worksheets("Billet Nomenclature").Cells(Rows.Count, "A").End(xlUp).row
    getTypeDescription = ""
    For i = 3 To lastRow
        If Worksheets("Billet Nomenclature").Cells(i, 1).value = t Then
            getTypeDescription = Worksheets("Billet Nomenclature").Cells(i, 2).value
            Exit For
        End If
    Next i
End Function

Public Function getEnvironmentDescription(t As String) As String
    Dim lastRow As Long
    Dim i As Integer
    lastRow = Worksheets("Billet Nomenclature").Cells(Rows.Count, "A").End(xlUp).row
    getEnvironmentDescription = ""
    For i = 3 To lastRow
        If Worksheets("Billet Nomenclature").Cells(i, 1).value = t Then
            getEnvironmentDescription = Worksheets("Billet Nomenclature").Cells(i, 4).value
            Exit For
        End If
    Next i
End Function

Public Function getOpticDescription(t As String) As String
    Dim lastRow As Long
    Dim i As Integer
    lastRow = Worksheets("Billet Nomenclature").Cells(Rows.Count, "A").End(xlUp).row
    getOpticDescription = ""
    For i = 3 To lastRow
        If Worksheets("Billet Nomenclature").Cells(i, 1).value = t Then
            getOpticDescription = Worksheets("Billet Nomenclature").Cells(i, 6).value
            Exit For
        End If
    Next i
End Function

Public Function GetMountingDescription(t As String, d As String) As String
    Dim lastRow As Long
    Dim i As Integer
    lastRow = Worksheets("Billet Nomenclature").Cells(Rows.Count, "H").End(xlUp).row
    GetMountingDescription = ""
    For i = 3 To lastRow
        If Worksheets("Billet Nomenclature").Cells(i, 8).value = t Then
            GetMountingDescription = Worksheets("Billet Nomenclature").Cells(i, 9).value
            If d = "BOP" And (t <> "E" And t <> "F") Then GetMountingDescription = ""
            If d <> "BOP" And (t = "E" Or t = "F") Then GetMountingDescription = ""
            Exit For
        End If
    Next i
End Function

Public Function GetBodyfinishDescription(t As String) As String
    Dim lastRow As Long
    Dim i As Integer
    lastRow = Worksheets("Billet Nomenclature").Cells(Rows.Count, "N").End(xlUp).row
    GetBodyfinishDescription = ""
    For i = 3 To lastRow
        If Worksheets("Billet Nomenclature").Cells(i, 14).value = t Then
            GetBodyfinishDescription = Worksheets("Billet Nomenclature").Cells(i, 15).value
            Exit For
        End If
    Next i
End Function

Public Function GetOutputPowerDescription(t As String, d As String) As String
    Dim lastRow As Long
    Dim i As Integer
    Dim c As Integer
    c = 19
    If Mid(d, 3, 1) = "O" Then c = 20
    If d = "BOP" Then c = 18
    lastRow = Worksheets("Billet Nomenclature").Cells(Rows.Count, "Q").End(xlUp).row
    GetOutputPowerDescription = ""
    For i = 3 To lastRow
        If Worksheets("Billet Nomenclature").Cells(i, 17).value = t Then
            GetOutputPowerDescription = Worksheets("Billet Nomenclature").Cells(i, c).value
            Exit For
        End If
    Next i
End Function

Public Function GetDimmingDescription(t As String) As String
    Dim lastRow As Long
    Dim i As Integer
    lastRow = Worksheets("Billet Nomenclature").Cells(Rows.Count, "Y").End(xlUp).row
    GetDimmingDescription = ""
    For i = 3 To lastRow
        If Worksheets("Billet Nomenclature").Cells(i, 25).value = t Then
            GetDimmingDescription = Worksheets("Billet Nomenclature").Cells(i, 26).value
            Exit For
        End If
    Next i
End Function

Public Function GetDiffuserDescription(t As String) As String
    Dim lastRow As Long
    Dim i As Integer
    lastRow = Worksheets("Billet Nomenclature").Cells(Rows.Count, "AB").End(xlUp).row
    GetDiffuserDescription = ""
    For i = 3 To lastRow
        If Worksheets("Billet Nomenclature").Cells(i, 28).value = t Then
            GetDiffuserDescription = Worksheets("Billet Nomenclature").Cells(i, 29).value
            Exit For
        End If
    Next i
End Function

Public Function GetBeamAngleDescription(t As String, d As String) As String
    Dim lastRow As Long
    Dim i As Integer
    Dim c As Integer
    c = 37  'BXO | BIO | BHO | BOO | BKO | BJO
    If d = "BXH" Or d = "BHS" Or d = "BKH" Then c = 36
    If d = "BIS" Or d = "BOS" Or d = "BJS" Then c = 35
    If d = "BIW" Or d = "BOW" Or d = "BJW" Then c = 34
    If d = "BIK" Or d = "BOK" Then c = 33
    If d = "BOP" Then c = 32
    lastRow = Worksheets("Billet Nomenclature").Cells(Rows.Count, "AE").End(xlUp).row
    GetBeamAngleDescription = ""
    For i = 3 To lastRow
        If Worksheets("Billet Nomenclature").Cells(i, 31).value = t Then
            GetBeamAngleDescription = Worksheets("Billet Nomenclature").Cells(i, c).value
            Exit For
        End If
    Next i
End Function

Public Function GetCRIDescription(t As String) As String
    Dim lastRow As Long
    Dim i As Integer
    lastRow = Worksheets("Billet Nomenclature").Cells(Rows.Count, "AM").End(xlUp).row
    GetCRIDescription = ""
    For i = 3 To lastRow
        If Worksheets("Billet Nomenclature").Cells(i, 39).value = t Then
            GetCRIDescription = Worksheets("Billet Nomenclature").Cells(i, 40).value
            Exit For
        End If
    Next i
End Function

Public Function GetCCTDescription(t As String) As String
    Dim lastRow As Long
    Dim i As Integer
    lastRow = Worksheets("Billet Nomenclature").Cells(Rows.Count, "AP").End(xlUp).row
    GetCCTDescription = ""
    For i = 3 To lastRow
        If Worksheets("Billet Nomenclature").Cells(i, 42).value = t Then
            GetCCTDescription = Worksheets("Billet Nomenclature").Cells(i, 43).value
            Exit For
        End If
    Next i
End Function

Public Function GetEmergencyDescription(t As String) As String
    Dim lastRow As Long
    Dim i As Integer
    lastRow = Worksheets("Billet Nomenclature").Cells(Rows.Count, "AS").End(xlUp).row
    GetEmergencyDescription = ""
    For i = 3 To lastRow
        If Worksheets("Billet Nomenclature").Cells(i, 45).value = t Then
            GetEmergencyDescription = Worksheets("Billet Nomenclature").Cells(i, 46).value
            Exit For
        End If
    Next i
End Function

Public Function GetVoltageDescription(t As String) As String
    Dim lastRow As Long
    Dim i As Integer
    lastRow = Worksheets("Billet Nomenclature").Cells(Rows.Count, "V").End(xlUp).row
    GetVoltageDescription = ""
    For i = 3 To lastRow
        If Worksheets("Billet Nomenclature").Cells(i, 22).value = t Then
            GetVoltageDescription = Worksheets("Billet Nomenclature").Cells(i, 23).value
            Exit For
        End If
    Next i
End Function

Public Function GetWiringDescription(t As String) As String
    Dim lastRow As Long
    Dim i As Integer
    lastRow = Worksheets("Billet Nomenclature").Cells(Rows.Count, "AV").End(xlUp).row
    GetWiringDescription = ""
    For i = 3 To lastRow
        If Worksheets("Billet Nomenclature").Cells(i, 48).value = t Then
            GetWiringDescription = Worksheets("Billet Nomenclature").Cells(i, 49).value
            Exit For
        End If
    Next i
End Function


