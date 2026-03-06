Attribute VB_Name = "modFormHelpers"
Public Function GetFractionSurdFromControls(frm As Object, _
                                            eqNum As Integer, _
                                            prefix As String) As FractionSurd
    
    Dim fs As FractionSurd
    Dim signVal As Long
    Dim toggleName As String

    Select Case prefix
        Case "A": toggleName = "tglSignFT" & eqNum
        Case "B": toggleName = "tglSignST" & eqNum
        Case "C": toggleName = "tglSignCT" & eqNum
    End Select
    ' -----------------------------
    ' Read Sign
    ' -----------------------------
    If frm.Controls(toggleName).Caption = "-" Then
        signVal = -1
    Else
        signVal = 1
    End If
    
    ' -----------------------------
    ' Read Numerator
    ' -----------------------------
    fs.num.coeff = CLng(frm.Controls("txt" & prefix & "NC" & eqNum).value)
    fs.num.radicand = CLng(frm.Controls("txt" & prefix & "NR" & eqNum).value)
    
    ' -----------------------------
    ' Read Denominator
    ' -----------------------------
    fs.den.coeff = CLng(frm.Controls("txt" & prefix & "DC" & eqNum).value)
    fs.den.radicand = CLng(frm.Controls("txt" & prefix & "DR" & eqNum).value)
    
    ' -----------------------------
    ' Default Handling
    ' -----------------------------
    
    If fs.num.radicand <= 0 Then fs.num.radicand = 1
    If fs.den.coeff = 0 Then fs.den.coeff = 1
    If fs.den.radicand <= 0 Then fs.den.radicand = 1
    
    ' -----------------------------
    ' Apply Sign to Numerator
    ' -----------------------------
    fs.num.coeff = fs.num.coeff * signVal
    
    GetFractionSurdFromControls = fs
    
End Function


Public Function GetFractionTermFromControls(frm As Object, _
                                            eqNum As Integer, _
                                            prefix As String) As FractionTerm
                                            
    Dim ft As FractionTerm
    
    ft.coeff = GetFractionSurdFromControls(frm, eqNum, prefix)
    ft.variableID = GetVariableID(frm, eqNum, prefix)
    
    GetFractionTermFromControls = ft
    
End Function
Public Function GetVariableID(frm As Object, _
                              eqNum As Integer, _
                              prefix As String) As Integer

    Dim userVar As String
    Dim pVar As String
    Dim sVar As String
    
    ' Variable entered for this term (from term control)
    userVar = Trim(frm.Controls("cmbVar" & prefix & eqNum).value)
    
    ' Primary & Secondary variables defined by user
    pVar = Trim(frm.txtPVar.value)
    sVar = Trim(frm.txtSVar.value)
    
    If userVar = pVar Then
        GetVariableID = 1
        
    ElseIf userVar = sVar Then
        GetVariableID = 2
        
    Else
        GetVariableID = 0   ' constant term
        
    End If

End Function
Public Sub UpdateEquationPreview(frm As Object, eqNum As Integer)
    If frm.IsUpdating Then Exit Sub
    On Error GoTo ErrHandler

    Dim parts(1 To 4) As String
    Dim eqPos As Integer
    eqPos = val(frm.Controls("txtETP" & eqNum).value)
    Dim i As Integer
    Dim j As Integer
    Dim termCtrlNames As Variant
    termCtrlNames = Array("PVTP", "SVTP", "CTP")

    Dim varNames As Variant
    varNames = Array(frm.txtPVar.value, frm.txtSVar.value, "")

    Dim termPrefix As Variant
    termPrefix = Array("FT", "ST", "CT")
    
    Dim coeffPrefix As Variant
    coeffPrefix = Array("A", "B", "C")
    
    For i = 0 To 2
    
        Dim tPos As Integer
        tPos = val(frm.Controls("txt" & Array("PVTP", "SVTP", "CTP")(i) & eqNum).value)
    
        If tPos >= 1 And tPos <= 4 Then
    
            Dim isLeft As Boolean
            Dim isLead As Boolean
            
            isLeft = (tPos < eqPos)
            isLead = False   ' we decide later
                
            parts(tPos) = GetLaTeXString( _
                frm.Controls("tglSign" & termPrefix(i) & eqNum).Caption, _
                frm.Controls("txt" & coeffPrefix(i) & "NC" & eqNum).value, _
                frm.Controls("txt" & coeffPrefix(i) & "NR" & eqNum).value, _
                frm.Controls("txt" & coeffPrefix(i) & "DC" & eqNum).value, _
                frm.Controls("txt" & coeffPrefix(i) & "DR" & eqNum).value, _
                varNames(i), _
                False)
    
        End If
    
    Next i

    Dim leftPart As String
    Dim rightPart As String
    Dim firstLeft As Boolean
    Dim firstRight As Boolean
    
    firstLeft = True
    firstRight = True
    
    For j = 1 To 4
    
        If parts(j) <> "" Then
            
            If j < eqPos Then
                
                If firstLeft Then
                    parts(j) = RemoveLeadingPlus(parts(j))
                    firstLeft = False
                End If
                
                leftPart = leftPart & " " & parts(j)
                
            ElseIf j > eqPos Then
                
                If firstRight Then
                    parts(j) = RemoveLeadingPlus(parts(j))
                    firstRight = False
                End If
                
                rightPart = rightPart & " " & parts(j)
                
            End If
            
        End If
    
    Next j

    Dim fullLatex As String
    Select Case eqPos
    
        Case 1
            fullLatex = "0 = " & Trim(rightPart)
    
        Case 4
            fullLatex = Trim(leftPart) & " = 0"
    
        Case Else
            fullLatex = Trim(leftPart) & " = " & Trim(rightPart)

    End Select
    ' Save to form property
    If eqNum = 1 Then
        frm.LastLtx1 = fullLatex
    Else
        frm.LastLtx2 = fullLatex
    End If
    Debug.Print "Eq"; eqNum; "=", fullLatex
    ' Render preview
    RenderMathJax fullLatex, frm, eqNum

    ' Update combined question
    RenderMathJax frm.LastLtx1 & " ; " & frm.LastLtx2, frm, "Question"

    Exit Sub

ErrHandler:
    MsgBox "UpdateEquationPreview crashed: " & Err.Description

End Sub
Public Sub UpdateQuestionPreview(frm As Object)

    Dim eq1 As String
    Dim eq2 As String
    
    eq1 = frm.LastLtx1
    eq2 = frm.LastLtx2
    
    ' Prevent one equation from clearing the other
    If eq1 = "" And eq2 = "" Then Exit Sub
    
    Dim combined As String
    
    If eq1 <> "" And eq2 <> "" Then
        combined = eq1 & " \; ; \; " & eq2
    ElseIf eq1 <> "" Then
        combined = eq1
    Else
        combined = eq2
    End If
    
    RenderMathJax combined, frm, "Question"

End Sub
Public Sub RearrangeFrames(frm As Object, eqNum As Integer)
    Dim startX As Integer: startX = 10: Dim margin As Integer: margin = 10
    Dim pos As Integer: Dim currentLeft As Integer: Dim ctrl As Object
    
    Dim pv As Integer: pv = val(frm.Controls("txtPVTP" & eqNum).value)
    Dim sv As Integer: sv = val(frm.Controls("txtSVTP" & eqNum).value)
    Dim et As Integer: et = val(frm.Controls("txtETP" & eqNum).value)
    Dim ct As Integer: ct = val(frm.Controls("txtCTP" & eqNum).value)

    currentLeft = startX
    For pos = 1 To 4
        Set ctrl = Nothing
        If pv = pos Then
            Set ctrl = frm.Controls("fraFT" & eqNum)
            HandleSignVisibility frm.Controls("tglSignFT" & eqNum), pos, et
        ElseIf sv = pos Then
            Set ctrl = frm.Controls("fraST" & eqNum)
            HandleSignVisibility frm.Controls("tglSignST" & eqNum), pos, et
        ElseIf ct = pos Then
            Set ctrl = frm.Controls("fraCT" & eqNum)
            HandleSignVisibility frm.Controls("tglSignCT" & eqNum), pos, et
        ElseIf et = pos Then
            Set ctrl = frm.Controls("lblET" & eqNum)
            ' CORRECTED NESTED IF
            If pos = 1 Then
                frm.Controls("lblET" & eqNum).Caption = "0 ="
            ElseIf pos = 4 Then
                frm.Controls("lblET" & eqNum).Caption = "= 0"
            Else
                frm.Controls("lblET" & eqNum).Caption = "="
            End If
        End If
        
        If Not ctrl Is Nothing Then
            ctrl.left = currentLeft
            currentLeft = currentLeft + ctrl.Width + margin
        End If
    Next pos
End Sub
Public Sub HandleSignVisibility(tgl As Object, p As Integer, e As Integer)
    If tgl.Caption = "+" And (p = 1 Or p = e + 1) Then tgl.ForeColor = tgl.BackColor Else tgl.ForeColor = vbBlack
End Sub
Public Function GetSafeValue(frm As Object, ctrlName As String) As Variant

    On Error GoTo SafeExit
    
    If ControlExists(frm, ctrlName) Then
        GetSafeValue = frm.Controls(ctrlName).value
    Else
        GetSafeValue = ""
    End If
    
    Exit Function

SafeExit:
    GetSafeValue = ""

End Function
Public Function ControlExists(frm As Object, ctrlName As String) As Boolean
    On Error Resume Next
    ControlExists = Not frm.Controls(ctrlName) Is Nothing
    On Error GoTo 0
End Function
Public Function FindControlRecursive(obj As Object, ctrlName As String) As String
    Dim ctrl As Object
    On Error Resume Next
    
    ' Look at all children of the current object
    For Each ctrl In obj.Controls
        If ctrl.Name = ctrlName Then
            FindControlRecursive = ctrl.value
            Exit Function
        End If
        
        ' If this control is a container (Frame/MultiPage), look inside it
        If TypeName(ctrl) = "Frame" Or TypeName(ctrl) = "MultiPage" Then
            Dim result As String
            result = FindControlRecursive(ctrl, ctrlName)
            If result <> "" Then
                FindControlRecursive = result
                Exit Function
            End If
        End If
    Next ctrl
End Function

