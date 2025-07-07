Attribute VB_Name = "frm_RunSplit"
Attribute VB_Base = "0{0FD83E84-98FB-4046-B430-45C04D2B7DC4}{51041F15-E113-4322-A026-955381007CC1}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Dim customLength As Integer
Dim ProposedLength As Integer

Private Sub selectIndex(i As Integer)
    frame_Custom.ForeColor = vbButtonText
    l1.BackColor = vbButtonFace
    l2.BackColor = vbButtonFace
    l3.BackColor = vbButtonFace
    l4.BackColor = vbButtonFace
    l5.BackColor = vbButtonFace
    l6.BackColor = vbButtonFace
    l7.BackColor = vbButtonFace
    l8.BackColor = vbButtonFace
    l9.BackColor = vbButtonFace
    cmd_12.BackColor = vbButtonFace
    cmd_18.BackColor = vbButtonFace
    cmd_24.BackColor = vbButtonFace
    cmd_30.BackColor = vbButtonFace
    cmd_36.BackColor = vbButtonFace
    cmd_42.BackColor = vbButtonFace
    cmd_48.BackColor = vbButtonFace
    cmd_54.BackColor = vbButtonFace
    cmd_60.BackColor = vbButtonFace
    If i = 1 Then l1.BackColor = vbGreen
    If i = 2 Then l2.BackColor = vbGreen
    If i = 3 Then l3.BackColor = vbGreen
    If i = 4 Then l4.BackColor = vbGreen
    If i = 5 Then l5.BackColor = vbGreen
    If i = 6 Then l6.BackColor = vbGreen
    If i = 7 Then l7.BackColor = vbGreen
    If i = 8 Then l8.BackColor = vbGreen
    If i = 9 Then l9.BackColor = vbGreen
    If i = 0 Then
        frame_Custom.ForeColor = vbGreen
        cmd_12.BackColor = vbGreen
        cmd_18.BackColor = vbGreen
        cmd_24.BackColor = vbGreen
        cmd_30.BackColor = vbGreen
        cmd_36.BackColor = vbGreen
        cmd_42.BackColor = vbGreen
        cmd_48.BackColor = vbGreen
        cmd_54.BackColor = vbGreen
        cmd_60.BackColor = vbGreen
    End If
End Sub

Private Function displayCustomResult(ByRef g() As t_SplitResult)
    l_Custom.Caption = ""
    Dim i As Integer
    Dim s As String
    s = ""
    For i = LBound(g) To UBound(g)
        s = s & g(i).Qty & " x " & CStr(g(i).length) & """ " & g(i).wiring & vbCrLf
    Next i
    l_Custom.Caption = s
End Function

Private Sub updateProposedLength()
    ProposedLength = 0
    Dim i As Integer
    For i = 1 To customLength
        ProposedLength = ProposedLength + customResult(i).length
    Next i
    l_ProposedLength.Caption = CStr(ProposedLength) & """"
    l_DifferenceLength.Caption = CStr(RequiredLength - ProposedLength) & """"
End Sub

Private Sub cmd_12_Click()
    customLength = customLength + 1
    ReDim Preserve customResult(1 To customLength)
    customResult(customLength).length = 12
    customResult(customLength).Qty = 1
    updateProposedLength
    Dim c() As t_SplitResult
    c = customResult
    Call splitForInstallation(c)
    Call displayCustomResult(c)
    resultIndex = 0
    Call selectIndex(resultIndex)
End Sub

Private Sub cmd_18_Click()
    customLength = customLength + 1
    ReDim Preserve customResult(1 To customLength)
    customResult(customLength).length = 18
    customResult(customLength).Qty = 1
    updateProposedLength
    Dim c() As t_SplitResult
    c = customResult
    Call splitForInstallation(c)
    Call displayCustomResult(c)
    resultIndex = 0
    Call selectIndex(resultIndex)
End Sub

Private Sub cmd_24_Click()
    customLength = customLength + 1
    ReDim Preserve customResult(1 To customLength)
    customResult(customLength).length = 24
    customResult(customLength).Qty = 1
    updateProposedLength
    Dim c() As t_SplitResult
    c = customResult
    Call splitForInstallation(c)
    Call displayCustomResult(c)
    resultIndex = 0
    Call selectIndex(resultIndex)
End Sub

Private Sub cmd_30_Click()
    customLength = customLength + 1
    ReDim Preserve customResult(1 To customLength)
    customResult(customLength).length = 30
    customResult(customLength).Qty = 1
    updateProposedLength
    Dim c() As t_SplitResult
    c = customResult
    Call splitForInstallation(c)
    Call displayCustomResult(c)
    resultIndex = 0
    Call selectIndex(resultIndex)
End Sub

Private Sub cmd_36_Click()
    customLength = customLength + 1
    ReDim Preserve customResult(1 To customLength)
    customResult(customLength).length = 36
    customResult(customLength).Qty = 1
    updateProposedLength
    Dim c() As t_SplitResult
    c = customResult
    Call splitForInstallation(c)
    Call displayCustomResult(c)
    resultIndex = 0
    Call selectIndex(resultIndex)
End Sub

Private Sub cmd_42_Click()
    customLength = customLength + 1
    ReDim Preserve customResult(1 To customLength)
    customResult(customLength).length = 42
    customResult(customLength).Qty = 1
    updateProposedLength
    Dim c() As t_SplitResult
    c = customResult
    Call splitForInstallation(c)
    Call displayCustomResult(c)
    resultIndex = 0
    Call selectIndex(resultIndex)
End Sub

Private Sub cmd_48_Click()
    customLength = customLength + 1
    ReDim Preserve customResult(1 To customLength)
    customResult(customLength).length = 48
    customResult(customLength).Qty = 1
    updateProposedLength
    Dim c() As t_SplitResult
    c = customResult
    Call splitForInstallation(c)
    Call displayCustomResult(c)
    resultIndex = 0
    Call selectIndex(resultIndex)
End Sub

Private Sub cmd_54_Click()
    customLength = customLength + 1
    ReDim Preserve customResult(1 To customLength)
    customResult(customLength).length = 54
    customResult(customLength).Qty = 1
    updateProposedLength
    Dim c() As t_SplitResult
    c = customResult
    Call splitForInstallation(c)
    Call displayCustomResult(c)
    resultIndex = 0
    Call selectIndex(resultIndex)
End Sub

Private Sub cmd_60_Click()
    customLength = customLength + 1
    ReDim Preserve customResult(1 To customLength)
    customResult(customLength).length = 60
    customResult(customLength).Qty = 1
    updateProposedLength
    Dim c() As t_SplitResult
    c = customResult
    Call splitForInstallation(c)
    Call displayCustomResult(c)
    resultIndex = 0
    Call selectIndex(resultIndex)
End Sub

Private Sub cmd_Confirmed_Click()
        
    Unload Me
End Sub

Private Sub cmd_Undo_Click()
    If customLength > 0 Then
        customLength = customLength - 1
        If customLength = 0 Then
            ReDim customResult(1 To 1)
        Else
            ReDim Preserve customResult(1 To customLength)
        End If
        Dim c() As t_SplitResult
        c = customResult
        Call splitForInstallation(c)
        Call displayCustomResult(c)
        resultIndex = 0
        Call selectIndex(resultIndex)
        updateProposedLength
    End If
End Sub

Private Sub l1_Click()
    resultIndex = 1
    Call selectIndex(resultIndex)
End Sub

Private Sub l2_Click()
    resultIndex = 2
    Call selectIndex(resultIndex)
End Sub

Private Sub l3_Click()
    resultIndex = 3
    Call selectIndex(resultIndex)
End Sub

Private Sub l4_Click()
    resultIndex = 4
    Call selectIndex(resultIndex)
End Sub

Private Sub l5_Click()
    resultIndex = 5
    Call selectIndex(resultIndex)
End Sub

Private Sub l6_Click()
    resultIndex = 6
    Call selectIndex(resultIndex)
End Sub

Private Sub l7_Click()
    resultIndex = 7
    Call selectIndex(resultIndex)
End Sub

Private Sub l8_Click()
    resultIndex = 8
    Call selectIndex(resultIndex)
End Sub

Private Sub l9_Click()
    resultIndex = 9
    Call selectIndex(resultIndex)
End Sub

Private Function getDelta(d As Integer)
    If d = 0 Then getDelta = "Same length"
    If d < 0 Then getDelta = CStr(d) & """" & " longer"
    If d > 0 Then getDelta = CStr(d) & """" & " shorter"
End Function

Private Sub UserForm_Initialize()
    Dim s As String
    For i = 1 To lengthsCount
        s = ""
        If usedStandards(i, 9) > 0 Then s = s & CStr(usedStandards(i, 9)) & " pc of 60""" & vbCrLf
        If usedStandards(i, 8) > 0 Then s = s & CStr(usedStandards(i, 8)) & " pc of 54""" & vbCrLf
        If usedStandards(i, 7) > 0 Then s = s & CStr(usedStandards(i, 7)) & " pc of 48""" & vbCrLf
        If usedStandards(i, 6) > 0 Then s = s & CStr(usedStandards(i, 6)) & " pc of 42""" & vbCrLf
        If usedStandards(i, 5) > 0 Then s = s & CStr(usedStandards(i, 5)) & " pc of 36""" & vbCrLf
        If usedStandards(i, 4) > 0 Then s = s & CStr(usedStandards(i, 4)) & " pc of 30""" & vbCrLf
        If usedStandards(i, 3) > 0 Then s = s & CStr(usedStandards(i, 3)) & " pc of 24""" & vbCrLf
        If usedStandards(i, 2) > 0 Then s = s & CStr(usedStandards(i, 2)) & " pc of 18""" & vbCrLf
        If usedStandards(i, 1) > 0 Then s = s & CStr(usedStandards(i, 1)) & " pc of 12""" & vbCrLf
        
        If i = 1 Then l_Suggestion1.Caption = s
        If i = 2 Then l_Suggestion2.Caption = s
        If i = 3 Then l_Suggestion3.Caption = s
        If i = 4 Then l_Suggestion4.Caption = s
        If i = 5 Then l_Suggestion5.Caption = s
        If i = 6 Then l_Suggestion6.Caption = s
        If i = 7 Then l_Suggestion7.Caption = s
        If i = 8 Then l_Suggestion8.Caption = s
        If i = 9 Then l_Suggestion9.Caption = s
    
        If i = 1 Then
            l_Delta1.Caption = getDelta(deltaLength(i))
            If deltaLength(i) <> 0 Then l_Delta1.ForeColor = vbRed
        End If
        If i = 2 Then
            l_Delta2.Caption = getDelta(deltaLength(i))
            If deltaLength(i) <> 0 Then l_Delta2.ForeColor = vbRed
        End If
        If i = 3 Then
            l_Delta3.Caption = getDelta(deltaLength(i))
            If deltaLength(i) <> 0 Then l_Delta3.ForeColor = vbRed
        End If
        If i = 4 Then
            l_Delta4.Caption = getDelta(deltaLength(i))
            If deltaLength(i) <> 0 Then l_Delta4.ForeColor = vbRed
        End If
        If i = 5 Then
            l_Delta5.Caption = getDelta(deltaLength(i))
            If deltaLength(i) <> 0 Then l_Delta5.ForeColor = vbRed
        End If
        If i = 6 Then
            l_Delta6.Caption = getDelta(deltaLength(i))
            If deltaLength(i) <> 0 Then l_Delta6.ForeColor = vbRed
        End If
        If i = 7 Then
            l_Delta7.Caption = getDelta(deltaLength(i))
            If deltaLength(i) <> 0 Then l_Delta7.ForeColor = vbRed
        End If
        If i = 8 Then
            l_Delta8.Caption = getDelta(deltaLength(i))
            If deltaLength(i) <> 0 Then l_Delta8.ForeColor = vbRed
        End If
        If i = 9 Then
            l_Delta9.Caption = getDelta(deltaLength(i))
            If deltaLength(i) <> 0 Then l_Delta9.ForeColor = vbRed
        End If
        
        
    Next i

    Call selectIndex(resultIndex)
    
    l1.Top = 25
    l1.Left = 12
    l1.Height = 17
    l1.Width = 54
    
    l2.Top = 25
    l2.Left = l1.Left + l1.Width
    l2.Height = 17
    l2.Width = 54
    
    l3.Top = 25
    l3.Left = l2.Left + l2.Width
    l3.Height = 17
    l3.Width = 54
    
    l4.Top = 25
    l4.Left = l3.Left + l3.Width
    l4.Height = 17
    l4.Width = 54
    
    l5.Top = 25
    l5.Left = l4.Left + l4.Width
    l5.Height = 17
    l5.Width = 54

    l6.Top = 25
    l6.Left = l5.Left + l5.Width + 1
    l6.Height = 17
    l6.Width = 54 - 1

    l7.Top = 25
    l7.Left = l6.Left + l6.Width + 1
    l7.Height = 17
    l7.Width = 54 - 1

    l8.Top = 25
    l8.Left = l7.Left + l7.Width + 1
    l8.Height = 17
    l8.Width = 54 - 1

    l9.Top = 25
    l9.Left = l8.Left + l8.Width
    l9.Height = 17
    l9.Width = 54
    
    ' prepare the custom result
    ReDim customResult(1 To 1)
    customLength = 0
    
    l_RequiredLength.Caption = CStr(RequiredLength) & """"
    l_ProposedLength.Caption = CStr(ProposedLength) & """"
    

End Sub
