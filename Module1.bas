Attribute VB_Name = "Module1"
Option Explicit

Private Const KSP As String = "\kern-5pt"
Public EqCounter As Long
Public Function NextEq() As String

    EqCounter = EqCounter + 1
    NextEq = "\dots \text{(" & EqCounter & ")}"

End Function


Public Function GetGlobalLCM(fs1 As FractionSurd, _
                             fs2 As FractionSurd, _
                             fs3 As FractionSurd) As Long
    
    Dim d1 As Long, d2 As Long, d3 As Long
    
    d1 = GetSquaredDenominator(fs1)
    d2 = GetSquaredDenominator(fs2)
    d3 = GetSquaredDenominator(fs3)
    
    GetGlobalLCM = lcm(lcm(d1, d2), d3)
    
End Function

Public Function GetLCMSurdMultiplier(globalLCM As Long) As Surd
    
    GetLCMSurdMultiplier = SquareRootToSurd(globalLCM)
    
End Function
Public Function ConvertToFinalSurd(fs As FractionSurd, _
                                   globalLCM As Long, _
                                   ByRef result As Surd) As Boolean

    Dim numeratorSquared As Long
    Dim denominatorSquared As Long
    Dim tempResult As Long
    
    ' (mvn)^2 = m^2 * n
    numeratorSquared = (fs.num.coeff * fs.num.coeff) * fs.num.radicand
    denominatorSquared = (fs.den.coeff * fs.den.coeff) * fs.den.radicand
    
    ' Multiply numerator by global LCM (squared domain)
    numeratorSquared = numeratorSquared * globalLCM
    
    ' Safe exact division check
    If Not SafeDivide(numeratorSquared, denominatorSquared, tempResult) Then
        ConvertToFinalSurd = False
        Exit Function
    End If
    
    ' Convert back from squared domain
    result = SquareRootToSurd(tempResult)
    
    ConvertToFinalSurd = True

End Function

Public Function GetFinalGCD(s1 As Surd, _
                            s2 As Surd, _
                            s3 As Surd) As Long
    
    Dim v1 As Long, v2 As Long, v3 As Long
    
    v1 = GetSurdSquaredValue(s1)
    v2 = GetSurdSquaredValue(s2)
    v3 = GetSurdSquaredValue(s3)
    
    GetFinalGCD = GCD(GCD(v1, v2), v3)
    
End Function


Public Function ProcessEquation(frm As Object, prefix As String, _
                                ByRef final1 As Surd, _
                                ByRef final2 As Surd, _
                                ByRef final3 As Surd) As Boolean
    
    Dim fs1 As FractionSurd
    Dim fs2 As FractionSurd
    Dim fs3 As FractionSurd
    
    Dim globalLCM As Long
    Dim multiplier As Surd
    
    fs1 = GetFractionSurdFromControls(frm, 1, prefix)
    fs2 = GetFractionSurdFromControls(frm, 2, prefix)
    fs3 = GetFractionSurdFromControls(frm, 3, prefix)
    
    If Not ValidateFractionSurd(fs1) _
    Or Not ValidateFractionSurd(fs2) _
    Or Not ValidateFractionSurd(fs3) Then
        
        MsgBox "Beyond the scope of current Grade"
        Exit Function
    End If
    
    globalLCM = GetGlobalLCM(fs1, fs2, fs3)
    
    If globalLCM <= 0 Then
        MsgBox "Beyond the scope of current Grade"
        Exit Function
    End If
    
    multiplier = GetLCMSurdMultiplier(globalLCM)
    
    If Not ConvertToFinalSurd(fs1, globalLCM, final1) Then
        MsgBox "Beyond the scope of current Grade"
        Exit Function
    End If
    
    If Not ConvertToFinalSurd(fs2, globalLCM, final2) Then
        MsgBox "Beyond the scope of current Grade"
        Exit Function
    End If
    
    If Not ConvertToFinalSurd(fs3, globalLCM, final3) Then
        MsgBox "Beyond the scope of current Grade"
        Exit Function
    End If
    
    If Not ReduceFinalEquation(final1, final2, final3) Then
        MsgBox "Beyond the scope of current Grade"
        Exit Function
    End If
    
    ProcessEquation = True
    
End Function


Public Function SafeDivide(dividend As Long, _
                           divisor As Long, _
                           ByRef result As Long) As Boolean
    
    If divisor = 0 Then
        SafeDivide = False
        Exit Function
    End If
    
    If dividend Mod divisor <> 0 Then
        SafeDivide = False
        Exit Function
    End If
    
    result = dividend \ divisor
    SafeDivide = True
    
End Function
Public Function ValidateFractionSurd(fs As FractionSurd) As Boolean

    ' Denominator coefficient cannot be zero
    If fs.den.coeff = 0 Then Exit Function
    
    ' Radicands must be positive
    If fs.den.radicand <= 0 Then Exit Function
    If fs.num.radicand <= 0 Then Exit Function
    
    ' Radicands should not be zero
    If fs.den.radicand = 0 Then Exit Function
    If fs.num.radicand = 0 Then Exit Function
    
    ValidateFractionSurd = True

End Function


' Recursive search helper







Public Sub SolveSystem(frm As Object)

    Dim webSolution As String
    Dim eq1Final As StandardForm
    Dim eq2Final As StandardForm
    Dim result As SystemType
    EqCounter = 0
    webSolution = "\begin{aligned}"
    
    ' Process Equation 1
    eq1Final = ProcessSingleEquation(frm, 1, webSolution)
    
    webSolution = webSolution & "\\[12pt]"
    
    ' Process Equation 2
    eq2Final = ProcessSingleEquation(frm, 2, webSolution)
    
    ' Validate System
    result = ValidateSystem(eq1Final, eq2Final)
    
    Select Case result
    
        Case SystemIdentity
            webSolution = webSolution & WrapStep( _
                "System Classification", _
                "\text{Both equations reduce to } 0 = 0. Infinite solutions.")
            
            RenderStepByStep webSolution, frm
            Exit Sub
            
        Case SystemContradictory
            webSolution = webSolution & WrapStep( _
                "System Classification", _
                "\text{System is inconsistent. No solution.}")
            
            RenderStepByStep webSolution, frm
            Exit Sub
            
        Case SystemDependent
            webSolution = webSolution & WrapStep( _
                "System Classification", _
                "\text{Equations are proportional. Infinite solutions.}")
            
            RenderStepByStep webSolution, frm
            Exit Sub
            
        Case SystemIndependent
            webSolution = webSolution & WrapStep( _
                "System Classification", _
                "\text{Determinant } \neq 0. Unique solution exists.")
            
            ' Next step: call elimination
            ' SolveByElimination eq1Final, eq2Final, webSolution
            
    End Select
    
    webSolution = webSolution & "\end{aligned}"
    
    RenderStepByStep webSolution, frm

End Sub






Public Function RearrangeToStandard(frm As Object, _
                                    eqNum As Integer) _
                                    As StandardForm

    Dim sf As StandardForm
    
    ' Build FractionSurds from controls
    Dim fsA As FractionSurd
    Dim fsB As FractionSurd
    Dim fsC As FractionSurd
    
    fsA = GetFractionSurdFromControls(frm, eqNum, "A")
    fsB = GetFractionSurdFromControls(frm, eqNum, "B")
    fsC = GetFractionSurdFromControls(frm, eqNum, "C")
    
    ' Move terms according to positions
    ' (Your existing side-switch logic goes here)
    
    ' Convert to StandardForm
    sf.aCoeff = fsA
    sf.bCoeff = fsB
    sf.constCoeff = fsC
    
    RearrangeToStandard = sf

End Function







Public Function NegateFractionSurd(fs As FractionSurd) As FractionSurd

    fs.num.coeff = -fs.num.coeff
    NegateFractionSurd = fs

End Function








Public Sub RunCommonPipeline(frm As Object)

    Dim webSolution As String
    Dim eq1Final As StandardForm
    Dim eq2Final As StandardForm
    Dim result As SystemType

    webSolution = "\begin{aligned}" & vbCrLf

    ' ===============================
    ' GIVEN EQUATIONS BLOCK
    ' ===============================
    
    Dim eq1Raw As String
    Dim eq2Raw As String
    
    eq1Raw = GetStep1Latex(frm, 1)
    eq2Raw = GetStep1Latex(frm, 2)

    webSolution = webSolution & _
        "& \left." & _
        "\begin{matrix}" & _
        eq1Raw & " \\" & _
        eq2Raw & _
        "\end{matrix}" & _
        "\right\}" & _
        "\quad \text{Given equations} \\[12pt]" & vbCrLf

    ' ===============================
    ' Equation 1
    ' ===============================
    
    webSolution = webSolution & LtxTextLine("Equation (1)")
    eq1Final = ProcessSingleEquation(frm, 1, webSolution)

    

    ' ===============================
    ' Equation 2
    ' ===============================
    
    webSolution = webSolution & LtxTextLine("Equation (2)")
    eq2Final = ProcessSingleEquation(frm, 2, webSolution)

    

    ' ===============================
    ' SYSTEM VALIDATION
    ' ===============================
    
    result = ValidateSystem(eq1Final, eq2Final)

    webSolution = webSolution & LtxTextLine("System classification")

    Select Case result

        Case SystemIdentity
            webSolution = webSolution & _
                LtxEquationLine("Both equations reduce to 0 = 0. Infinite solutions.")

        Case SystemContradictory
            webSolution = webSolution & _
                LtxEquationLine("System is inconsistent. No solution exists.")

        Case SystemDependent
            webSolution = webSolution & _
                LtxEquationLine("Equations are proportional. Infinite solutions.")

        Case SystemIndependent
            webSolution = webSolution & _
                LtxEquationLine("\text{Determinant} \neq 0. \text{Unique solution exists.}")

    End Select

    webSolution = webSolution & "\end{aligned}"

    ' ===============================
    ' SAVE TO DATABASE
    ' ===============================
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Database")

    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(r, 1).value = webSolution

    ' ===============================
    ' DISPLAY
    ' ===============================
    
    RenderStepByStep webSolution, frm

End Sub







Public Function ProcessSingleEquation(frm As Object, _
                                      eqNum As Integer, _
                                      ByRef webSolution As String) _
                                      As StandardForm

    Dim lastEquationLatex As String
    Dim sf As StandardForm
    Dim status As EquationStatus
    
    On Error GoTo DebugError

    ' =====================================
    ' ALWAYS DISPLAY EQUATION FIRST
    ' =====================================
    
    Dim step1Str As String
    step1Str = GetStep1Latex(frm, eqNum)
    
    lastEquationLatex = step1Str
    
    webSolution = webSolution & _
        "& \quad " & step1Str & " \\[5pt]" & vbCrLf

    ' =====================================
    ' STEP 2 – Rearrange
    ' =====================================
    
    sf = RearrangeToStandard(frm, eqNum)
    
    Dim step2Str As String
    step2Str = StandardFormToLatex(sf, frm.txtPVar.value, frm.txtSVar.value)
    
    If Replace(step1Str, " ", "") <> Replace(step2Str, " ", "") Then
        
        lastEquationLatex = step2Str
        
        webSolution = webSolution & _
            LtxTextLine("Rearranged to ax + by = c")
        
        webSolution = webSolution & _
            LtxEquationLine(step2Str)
            
    End If

    ' =====================================
    ' VALIDATE EQUATION
    ' =====================================
    
    status = ValidateEquation(sf)

    If status = Identity Or status = Contradiction Then
        
        lastEquationLatex = StandardFormToLatex(sf, _
                            frm.txtPVar.value, frm.txtSVar.value)
        
        webSolution = webSolution & _
            LtxTextLine("Equation cannot proceed")
        
        webSolution = webSolution & _
            LtxEquationLine(lastEquationLatex)
        
        GoTo ApplyNumbering
        
    End If

    ' =====================================
    ' STEP 3 – Normalize Sign
    ' =====================================
    
    If IsNegativeLeading(sf) Then
    
        sf = NormalizeLeadingSign(sf)
        
        lastEquationLatex = StandardFormToLatex(sf, _
                            frm.txtPVar.value, frm.txtSVar.value)
        
        webSolution = webSolution & _
            LtxTextLine("Multiply entire equation by -1 to make leading coefficient positive")
        
        webSolution = webSolution & _
            LtxEquationLine(lastEquationLatex)
        
    End If

    ' =====================================
    ' STEP 4 – Remove Denominators
    ' =====================================
    
    Dim lcmSurd As Surd
    lcmSurd = GetLCMMultiplier(sf)
    
    If Not (lcmSurd.coeff = 1 And lcmSurd.radicand = 1) Then
    
        Dim expandedLine As String
        expandedLine = ExpandLCMLine(sf, lcmSurd, _
                        frm.txtPVar.value, frm.txtSVar.value)
        
        webSolution = webSolution & _
            LtxTextLine("Multiplying the equation by " & _
                        FormatSurdToLatex(lcmSurd) & " we get")
        
        webSolution = webSolution & _
            LtxEquationLine(expandedLine)
    
        sf = ApplyLCM(sf, lcmSurd)
        
        lastEquationLatex = StandardFormToLatex(sf, _
                            frm.txtPVar.value, frm.txtSVar.value)
        
        webSolution = webSolution & _
            LtxEquationLine(lastEquationLatex)
    
    End If

    ' =====================================
    ' STEP 5 – GCD Reduction
    ' =====================================
    
    Dim gcdSurd As Surd
    gcdSurd = GetGCDFactor(sf)
        
    If Not (gcdSurd.coeff = 1 And gcdSurd.radicand = 1) Then
        
        Dim gcdValue As Long
        gcdValue = gcdSurd.coeff
        
        Dim divisionLine As String
        divisionLine = ExpandGCDDivisionLine(sf, gcdValue, _
                        frm.txtPVar.value, frm.txtSVar.value)
        
        webSolution = webSolution & _
            LtxTextLine("Dividing the equation by " & gcdValue)
        
        webSolution = webSolution & _
            LtxEquationLine(divisionLine)
        
        sf = ApplyGCDReduction(sf, gcdSurd)
        
        lastEquationLatex = StandardFormToLatex(sf, _
                            frm.txtPVar.value, frm.txtSVar.value)
        
        webSolution = webSolution & _
            LtxEquationLine(lastEquationLatex)
        
    End If

ApplyNumbering:

    ' =====================================
    ' NUMBER ONLY THE FINAL EQUATION
    ' =====================================
    
    Dim numberedLatex As String
    numberedLatex = lastEquationLatex & " \dots (" & eqNum & ")"
    
    webSolution = left(webSolution, _
        InStrRev(webSolution, lastEquationLatex) - 1) _
        & numberedLatex _
        & Mid(webSolution, _
        InStrRev(webSolution, lastEquationLatex) + Len(lastEquationLatex))

    ProcessSingleEquation = sf
    Exit Function

DebugError:
    MsgBox "Error inside ProcessSingleEquation: " & Err.Description

End Function







Public Function RemoveLeadingPlus(txt As String) As String
    If left(Trim(txt), 1) = "+" Then
        RemoveLeadingPlus = Mid(Trim(txt), 2)
    Else
        RemoveLeadingPlus = txt
    End If
End Function




Public Function BuildSurd(c As String, r As String) As String
    Dim vC As String: vC = IIf(c = "", "1", c): Dim vr As String: vr = IIf(r = "", "1", r)
    If vr = "1" Then BuildSurd = vC Else BuildSurd = IIf(vC = "1", "", vC) & "\sqrt{" & vr & "}"
End Function







' 3. Simplifies Sqr(Val) into m * Sqr(n)
' e.g., SimplifySquareRoot(300, m, n) results in m=10, n=3




' ... [Keep GCD, LCM, SimplifySquareRoot as they are] ...



'============================First half ends========================================================================================================================================


Private Function CalculateTermAfterMult(frm As Object, eqNum As Integer, prefix As String, multM As Long, multN As Long) As String
    Dim nc As Long, nr As Long, dc As Long, dr As Long
    Dim resM As Long, resN As Long
    Dim finalSign As String
    

    ' 1. Extract values
    nc = val(frm.Controls("txt" & prefix & "NC" & eqNum).value)
    nr = val(frm.Controls("txt" & prefix & "NR" & eqNum).value)
    dc = val(frm.Controls("txt" & prefix & "DC" & eqNum).value)
    dr = val(frm.Controls("txt" & prefix & "DR" & eqNum).value)
    If dc = 0 Then dc = 1: If dr = 0 Then dr = 1
    
    ' 2. The Math: (nc*m * Sqr(nr*n)) / (dc * Sqr(dr))
    ' To simplify, we square the whole thing: [(nc*m)^2 * (nr*n)] / [dc^2 * dr]
    Dim squaredVal As Long
    squaredVal = ((nc * multM) ^ 2 * (nr * multN)) / (dc ^ 2 * dr)
    
    ' 3. Square root the result back into m*Sqr(n)
    SimplifySquareRoot squaredVal, resM, resN
    
    ' 4. Handle Sign and Formatting
    ' (Logic to determine if term is negative based on Step 2/3 flips)
    ' ...
    
    If resN = 1 Then
        CalculateTermAfterMult = CStr(resM)
    Else
        CalculateTermAfterMult = resM & "\sqrt{" & resN & "}"
    End If
    
    ' Append variable name
    If prefix = "A" Then CalculateTermAfterMult = CalculateTermAfterMult & frm.txtPVar.value
    If prefix = "B" Then CalculateTermAfterMult = " + " & CalculateTermAfterMult & frm.txtSVar.value
End Function






' Helper to store surd as "Coefficient|Radicand"
Private Function SimplifyToPipeFormat(ByVal val As Long, ByVal Sign As Double) As String
    Dim m As Long, n As Long
    SimplifySquareRoot Abs(val), m, n
    ' Combine sign with m
    SimplifyToPipeFormat = (m * Sign) & "|" & n
End Function

Public Function ProcessSystem(frm As Object) As String
    
    Dim webSolution As String
    Dim sf1 As StandardForm
    Dim sf2 As StandardForm
    Dim result As SystemType
    
    webSolution = "\begin{aligned}" & vbCrLf
    
    Dim eq1Raw As String
    Dim eq2Raw As String
    
    eq1Raw = GetStep1Latex(frm, 1)
    eq2Raw = GetStep1Latex(frm, 2)
    
    ' -------------------------
    ' Given Equations Block
    ' -------------------------
    
    webSolution = webSolution & _
        "& \left." & _
        "\begin{matrix}" & _
        eq1Raw & " \\" & _
        eq2Raw & _
        "\end{matrix}" & _
        "\right\}" & _
        "\quad \text{Given equations} \\[12pt]" & vbCrLf
    
    ' -------------------------
    ' Equation 1
    ' -------------------------
    
    webSolution = webSolution & _
        LtxTextLine("Equation (1)")
    
    sf1 = ProcessSingleEquation(frm, 1, webSolution)
    
    
    
    ' -------------------------
    ' Equation 2
    ' -------------------------
    
    webSolution = webSolution & _
        LtxTextLine("Equation (2)")
    
    sf2 = ProcessSingleEquation(frm, 2, webSolution)
    'webSolution = webSolution & _
    'LtxTextLine("Applying Cramer's Rule")

    webSolution = webSolution & _
    AppendCramerStep1(sf1, sf2, frm.txtPVar.value, frm.txtSVar.value)
    webSolution = webSolution & AppendDeterminantD(sf1, sf2)
    webSolution = webSolution & AppendDeterminantDx(sf1, sf2, frm.txtPVar.value)
    webSolution = webSolution & AppendDeterminantDy(sf1, sf2, frm.txtSVar.value)
    Dim D As FractionSurd
    Dim Dx As FractionSurd
    Dim Dy As FractionSurd
    
    D = ComputeDeterminant(sf1, sf2)
    Dx = ComputeDeterminantDx(sf1, sf2)
    Dy = ComputeDeterminantDy(sf1, sf2)
    
    webSolution = webSolution & _
        AppendFinalCramerStep(D, Dx, Dy, _
                              frm.txtPVar.value, _
                              frm.txtSVar.value)
'
'    ' -------------------------
'    ' System Classification
'    ' -------------------------
'
'    result = ValidateSystem(sf1, sf2)
'
'    webSolution = webSolution & _
'        LtxTextLine("System classification")
'
'    Select Case result
'
'        Case SystemIdentity
'            webSolution = webSolution & _
'                LtxEquationLine("Both equations reduce to identities. Infinite solutions.")
'
'        Case SystemContradictory
'            webSolution = webSolution & _
'                LtxEquationLine("System is inconsistent. No solution.")
'
'        Case SystemDependent
'            webSolution = webSolution & _
'                LtxEquationLine("Equations are proportional. Infinite solutions.")
'
'        Case SystemIndependent
'            webSolution = webSolution & _
'                LtxEquationLine("\text{Determinant} \neq 0. \text{Unique solution exists.}")
'
'    End Select
'
    webSolution = webSolution & "\end{aligned}"
    
    ProcessSystem = webSolution
    
End Function



'=================================================Module1 first half==========================================================================================================================
'=================================================Module1 second half==========================================================================================================================





' Formats a coefficient and radicand into clean LaTeX (e.g., 5, \sqrt{3}, or 5\sqrt{3})





' New Helper to keep math exact
Private Sub GetSquaredParts(frm As Object, eqNum As Integer, prefix As String, _
                            ByRef outSqNum As Long, ByRef outSqDen As Long)
    Dim nc As Long, nr As Long, dc As Long, dr As Long
    nc = val(frm.Controls("txt" & prefix & "NC" & eqNum).value): If nc = 0 Then nc = 1
    nr = val(frm.Controls("txt" & prefix & "NR" & eqNum).value): If nr = 0 Then nr = 1
    dc = val(frm.Controls("txt" & prefix & "DC" & eqNum).value): If dc = 0 Then dc = 1
    dr = val(frm.Controls("txt" & prefix & "DR" & eqNum).value): If dr = 0 Then dr = 1
    
    outSqNum = (nc ^ 2) * nr
    outSqDen = (dc ^ 2) * dr
End Sub
Public Sub CopyToClipboard(txt As String)
    Dim objData As Object
    ' Use late binding for MSForms.DataObject to avoid reference issues
    Set objData = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    
    objData.SetText txt
    objData.PutInClipboard
    
    MsgBox "Solution LaTeX copied to clipboard!", vbInformation
End Sub


Public Function CleanLeadingPlus(ByVal s As String, _
                                 ByVal isFirst As Boolean) As String

    s = Trim(s)

    If isFirst Then
        If left(s, 2) = "+ " Then
            CleanLeadingPlus = Mid(s, 3)
            Exit Function
        End If
    End If

    CleanLeadingPlus = s

End Function

' 3. Formatting: Converts string coefficient to LaTeX term


' 4. Sign Manipulation



' 1. Essential GCD Helper








Public Function ParseStringToSurd(s As String) As Surd

    Dim res As Surd
    
    s = Trim(s)
    
    ' Default case
    If s = "" Or s = "+" Then
        res.coeff = 1
        res.radicand = 1
        ParseStringToSurd = res
        Exit Function
    End If
    
    ' Handle pure integer
    If InStr(s, "\sqrt") = 0 Then
        res.coeff = CLng(s)
        res.radicand = 1
        ParseStringToSurd = res
        Exit Function
    End If
    
    ' Handle surd format like 3\sqrt{5}
    Dim coeffPart As String
    Dim radPart As String
    
    coeffPart = left(s, InStr(s, "\sqrt") - 1)
    
    If coeffPart = "" Or coeffPart = "+" Then
        res.coeff = 1
    ElseIf coeffPart = "-" Then
        res.coeff = -1
    Else
        res.coeff = CLng(coeffPart)
    End If
    
    radPart = Mid(s, InStr(s, "{") + 1)
    radPart = left(radPart, InStr(radPart, "}") - 1)
    
    res.radicand = CLng(radPart)
    
    ParseStringToSurd = res

End Function

' 2. Reverse the squaring: Takes a number and returns p\sqrt{q}


' 1. The Main Function your code is currently calling





'======================================================Cramer's Rule=======================================================================================================================





Public Sub GenerateAllMethodSolutions(frm As Object, r As Long)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Database")
    
    ' -----------------------------------
    ' Generate Each Method Independently
    ' -----------------------------------
    
    'ws.Cells(r, 46).Value = ProcessSubstitutionMethod(frm)   ' AT
    ws.Cells(r, 47).value = ProcessEliminationMethod(frm)    ' AU
    'ws.Cells(r, 48).Value = ProcessGraphicalMethod(frm)      ' AV
    ws.Cells(r, 49).value = ProcessCramerMethod(frm)         ' AW
    
    ' Optional: Keep standardized only (AX)
    ws.Cells(r, 50).value = ProcessSystem(frm)

End Sub
Public Function GenerateInitialStandardSteps(frm As Object, _
                                             ByRef sf1 As StandardForm, _
                                             ByRef sf2 As StandardForm) As String

    Dim latex As String
    
    latex = "\begin{aligned}" & vbCrLf
    
    ' Given equations block
    latex = latex & _
        "& \left.\begin{matrix}" & _
        GetStep1Latex(frm, 1) & " \\" & _
        GetStep1Latex(frm, 2) & _
        "\end{matrix}\right\}\quad \text{Given equations} \\[12pt]"
    
    ' Equation (1)
    latex = latex & LtxTextLine("Equation (1)")
    sf1 = ProcessSingleEquation(frm, 1, latex)
    
    ' Equation (2)
    latex = latex & LtxTextLine("Equation (2)")
    sf2 = ProcessSingleEquation(frm, 2, latex)
    
    GenerateInitialStandardSteps = latex

End Function












Public Function SolveSingleVariable(a As FractionSurd, _
                                   c As FractionSurd) As FractionSurd

    Dim result As FractionSurd
    
    result = DivideFractionSurd(c, a)
    
    SimplifyFractionSurd result
    
    SolveSingleVariable = result

End Function







Public Function CreateStandardForm(a As FractionSurd, _
                                   b As FractionSurd, _
                                   c As FractionSurd) As StandardForm

    Dim sf As StandardForm
    
    sf.aCoeff = a
    sf.bCoeff = b
    sf.constCoeff = c
    
    CreateStandardForm = sf

End Function



Public Function MultiplyEquation(sf As StandardForm, _
                                 k As Long) As StandardForm

    Dim result As StandardForm

    result.aCoeff = MultiplyFractionByInteger(sf.aCoeff, k)
    result.bCoeff = MultiplyFractionByInteger(sf.bCoeff, k)
    result.constCoeff = MultiplyFractionByInteger(sf.constCoeff, k)

    MultiplyEquation = result

End Function







Public Function ZeroFraction() As FractionSurd

    Dim z As FractionSurd
    
    z.num.coeff = 0
    z.num.radicand = 1
    
    z.den.coeff = 1
    z.den.radicand = 1
    
    ZeroFraction = z

End Function




Public Function ReadStandardForm(ws As Worksheet, _
                                 r As Long, _
                                 eqIndex As Long) As StandardForm

    Dim sf As StandardForm

    Dim a As FractionSurd
    Dim b As FractionSurd
    Dim c As FractionSurd

    ' Columns depend on which equation
    Dim baseCol As Long

    If eqIndex = 1 Then
        baseCol = 8
    Else
        baseCol = 27
    End If

    ' Read coefficients from worksheet
    a = ReadFractionSurd(ws, r, baseCol)
    b = ReadFractionSurd(ws, r, baseCol + 5)
    c = ReadFractionSurd(ws, r, baseCol + 10)

    sf.aCoeff = a
    sf.bCoeff = b
    sf.constCoeff = c

    ReadStandardForm = sf

End Function
Public Function BuildCommonInitialSteps(sf1 As StandardForm, _
                                        sf2 As StandardForm, _
                                        var1 As String, _
                                        var2 As String) As String

    Dim latex As String

    latex = ""

    latex = latex & _
        "& " & BuildLinearTermLatex(sf1.aCoeff, var1) & _
        " + " & BuildLinearTermLatex(sf1.bCoeff, var2) & _
        " = " & FormatFractionSurdToLatex(sf1.constCoeff) & " \\[6pt]"

    latex = latex & _
        "& " & BuildLinearTermLatex(sf2.aCoeff, var1) & _
        " + " & BuildLinearTermLatex(sf2.bCoeff, var2) & _
        " = " & FormatFractionSurdToLatex(sf2.constCoeff) & " \\[10pt]"

    BuildCommonInitialSteps = latex

End Function

Public Function ReadFractionSurd(ws As Worksheet, _
                                 r As Long, _
                                 startCol As Long) As FractionSurd

    Dim fs As FractionSurd
    Dim signVal As Long

    ' Sign column
    signVal = ws.Cells(r, startCol).value
    
    ' Numerator
    fs.num.coeff = ws.Cells(r, startCol + 1).value
    fs.num.radicand = ws.Cells(r, startCol + 2).value

    ' Denominator
    fs.den.coeff = ws.Cells(r, startCol + 3).value
    fs.den.radicand = ws.Cells(r, startCol + 4).value

    ' Apply sign
    fs.num.coeff = fs.num.coeff * signVal

    ReadFractionSurd = fs

End Function
