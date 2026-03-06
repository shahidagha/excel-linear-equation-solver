Attribute VB_Name = "modElimination"
Public Function ProcessEliminationMethod(frm As Object) As String

    Dim latex As String
    Dim sf1 As StandardForm
    Dim sf2 As StandardForm
    
    ' Initial standard steps
    latex = GenerateInitialStandardSteps(frm, sf1, sf2)
    
    ' Add elimination steps
    latex = latex & AppendEliminationSteps(sf1, sf2, _
                                           frm.txtPVar.value, _
                                           frm.txtSVar.value)
    
    latex = latex & "\end{aligned}"
    
    ProcessEliminationMethod = latex

End Function
Public Function DirectEliminationSequence(sf1 As StandardForm, _
                                          sf2 As StandardForm, _
                                          var1 As String, _
                                          var2 As String) As String

    Dim latex As String
    latex = ""

    Dim a1 As Long
    Dim a2 As Long
    Dim b1 As Long
    Dim b2 As Long
    
    a1 = sf1.aCoeff.num.coeff
    a2 = sf2.aCoeff.num.coeff
    
    b1 = sf1.bCoeff.num.coeff
    b2 = sf2.bCoeff.num.coeff


    ' --------------------------------
    ' Decide which variable eliminates
    ' --------------------------------

    Dim eliminateVar As String
    
    If Abs(b1) = Abs(b2) And b1 <> b2 Then
        eliminateVar = var2
    If Abs(b1) = Abs(b2) Then
        eliminateVar = var2
    End If
    Else
        eliminateVar = var1
    End If


    ' --------------------------------
    ' Decide ADD or SUBTRACT
    ' --------------------------------

    Dim useAdd As Boolean
    
    If eliminateVar = var2 Then
        
        If Sgn(b1) <> Sgn(b2) Then
            useAdd = True
        Else
            useAdd = False
        End If
        
    Else
        
        If Sgn(a1) <> Sgn(a2) Then
            useAdd = True
        Else
            useAdd = False
        End If
        
    End If


    ' --------------------------------
    ' Display operation
    ' --------------------------------

    If useAdd Then
    
        latex = latex & _
        "& \text{Adding equations (1) and (2)} \\[6pt]"
        
        latex = latex & _
        BuildVerticalAddLayout(sf1, sf2, var1, var2)
        
    Else
    
        latex = latex & _
        "& \text{Subtracting equations (1) and (2)} \\[6pt]"
        
        latex = latex & _
        BuildVerticalSubtractLayout(sf1, sf2, var1, var2)
        
    End If


    ' --------------------------------
    ' Compute result equation
    ' --------------------------------

    Dim newA As FractionSurd
    Dim newB As FractionSurd
    Dim newC As FractionSurd

    If useAdd Then
    
        newA = AddFractionSurd(sf1.aCoeff, sf2.aCoeff)
        newB = AddFractionSurd(sf1.bCoeff, sf2.bCoeff)
        newC = AddFractionSurd(sf1.constCoeff, sf2.constCoeff)
        
    Else
    
        newA = SubtractFractionSurd(sf1.aCoeff, sf2.aCoeff)
        newB = SubtractFractionSurd(sf1.bCoeff, sf2.bCoeff)
        newC = SubtractFractionSurd(sf1.constCoeff, sf2.constCoeff)
        
    End If
    latex = latex & _
    FinalizeEliminationStep(newA, newB, newC, sf1, var1, var2)

   
    DirectEliminationSequence = latex

End Function
Public Function ChooseEliminationVariable(sf1 As StandardForm, _
                                          sf2 As StandardForm, _
                                          var1 As String, _
                                          var2 As String) As String

    Dim a1 As Long
    Dim a2 As Long
    Dim b1 As Long
    Dim b2 As Long
    
    a1 = Abs(sf1.aCoeff.num.coeff)
    a2 = Abs(sf2.aCoeff.num.coeff)
    
    b1 = Abs(sf1.bCoeff.num.coeff)
    b2 = Abs(sf2.bCoeff.num.coeff)

    Dim lcmX As Long
    Dim lcmY As Long
    
    lcmX = Application.WorksheetFunction.lcm(a1, a2)
    lcmY = Application.WorksheetFunction.lcm(b1, b2)

    Dim scoreX As Long
    Dim scoreY As Long

    scoreX = Abs((lcmX / a1) * b1) + Abs((lcmX / a2) * b2)
    scoreY = Abs((lcmY / b1) * a1) + Abs((lcmY / b2) * a2)

    If scoreX < scoreY Then
        ChooseEliminationVariable = var1
    Else
        ChooseEliminationVariable = var2
    End If

End Function
Public Function CrossEliminationSequence(sf1 As StandardForm, _
                                         sf2 As StandardForm, _
                                         var1 As String, _
                                         var2 As String) As String

    Dim latex As String
    latex = ""

    Dim add_a As FractionSurd
    Dim add_b As FractionSurd
    Dim add_c As FractionSurd
    
    Dim sub_a As FractionSurd
    Dim sub_b As FractionSurd
    Dim sub_c As FractionSurd


    ' ---------------------------------
    ' ADD equations
    ' ---------------------------------

    add_a = AddFractionSurd(sf1.aCoeff, sf2.aCoeff)
    add_b = AddFractionSurd(sf1.bCoeff, sf2.bCoeff)
    add_c = AddFractionSurd(sf1.constCoeff, sf2.constCoeff)

    latex = latex & _
        "& \text{Adding equations:} \\[6pt]"

    latex = latex & _
        BuildVerticalAddLayout(sf1, sf2, var1, var2)


    ' ---------------------------------
    ' Print equation (3)
    ' ---------------------------------

    latex = latex & _
        "& " & BuildEquationLatex(add_a, add_b, add_c, var1, var2) & _
        " \dots \text{(3)} \\[10pt]"



    ' ---------------------------------
    ' SUBTRACT equations
    ' ---------------------------------

    Dim absA1 As Long
    Dim absA2 As Long
    
    Dim firstEq As StandardForm
    Dim secondEq As StandardForm

    absA1 = Abs(sf1.aCoeff.num.coeff)
    absA2 = Abs(sf2.aCoeff.num.coeff)

    If absA1 >= absA2 Then
    
        firstEq = sf1
        secondEq = sf2
        
    Else
    
        firstEq = sf2
        secondEq = sf1
        
    End If


    sub_a = SubtractFractionSurd(firstEq.aCoeff, secondEq.aCoeff)
    sub_b = SubtractFractionSurd(firstEq.bCoeff, secondEq.bCoeff)
    sub_c = SubtractFractionSurd(firstEq.constCoeff, secondEq.constCoeff)

    latex = latex & _
        "& \text{Subtracting equations:} \\[6pt]"

    latex = latex & _
        BuildVerticalSubtractLayout(firstEq, secondEq, var1, var2)


    ' ---------------------------------
    ' Print equation (4)
    ' ---------------------------------

    latex = latex & _
        "& " & BuildEquationLatex(sub_a, sub_b, sub_c, var1, var2) & _
        " \dots \text{(4)} \\[10pt]"



    ' ---------------------------------
    ' Stage 2 : eliminate y
    ' ---------------------------------

    latex = latex & _
        "& \text{Adding equations (3) and (4):} \\[6pt]"

    latex = latex & _
        BuildVerticalAddLayout( _
            CreateStandardForm(add_a, add_b, add_c), _
            CreateStandardForm(sub_a, sub_b, sub_c), _
            var1, var2)


    Dim final_a As FractionSurd
    Dim final_c As FractionSurd

    final_a = AddFractionSurd(add_a, sub_a)
    final_c = AddFractionSurd(add_c, sub_c)


    latex = latex & _
        "& " & BuildEquationLatex(final_a, ZeroFraction(), final_c, var1, var2) & " \\[10pt]"


    ' ---------------------------------
    ' Solve remaining variable
    ' ---------------------------------

    'latex = latex & _
     '   SolveAfterElimination(newA, newB, newC, sf1, var1, var2)
   latex = latex & FinalizeEliminationStep(final_a, ZeroFraction(), final_c, sf1, var1, var2)
    CrossEliminationSequence = latex

End Function
Public Function CrossEliminationSequenceCore(sf1 As StandardForm, _
                                         sf2 As StandardForm, _
                                         var1 As String, _
                                         var2 As String) As String

    Dim latex As String
    latex = ""

    Dim add_a As FractionSurd
    Dim add_b As FractionSurd
    Dim add_c As FractionSurd
    
    Dim sub_a As FractionSurd
    Dim sub_b As FractionSurd
    Dim sub_c As FractionSurd
    
    ' ---------------------------------
    ' ADD equations
    ' ---------------------------------
    
    add_a = AddFractionSurd(sf1.aCoeff, sf2.aCoeff)
    add_b = AddFractionSurd(sf1.bCoeff, sf2.bCoeff)
    add_c = AddFractionSurd(sf1.constCoeff, sf2.constCoeff)
    
    latex = latex & _
        "& \text{Adding equations:} \\[6pt]"
    
    latex = latex & _
        BuildVerticalAddLayout(sf1, sf2, var1, var2) & _
        "\\[8pt]"
    
    
    ' ---------- Divide if possible ----------
    
    Dim gAdd As Long
    gAdd = Application.WorksheetFunction.GCD( _
                Abs(add_a.num.coeff), _
                Abs(add_b.num.coeff), _
                Abs(add_c.num.coeff))



    If gAdd > 1 Then
        
    latex = latex & _
        "& \text{Dividing the equation by } " & gAdd & _
        "\text{ we get} \\[6pt]"
        
    add_a.num.coeff = add_a.num.coeff / gAdd
    add_b.num.coeff = add_b.num.coeff / gAdd
    add_c.num.coeff = add_c.num.coeff / gAdd
    
    Dim eqLine As String
    eqLine = ""
    
    ' ---- Build Linear Expression Cleanly ----
    
    If Not IsZeroFraction(add_a) Then
        eqLine = BuildLinearTermLatex(add_a, var1)
    End If
    
    If Not IsZeroFraction(add_b) Then
        
        If eqLine <> "" Then
            If add_b.num.coeff > 0 Then
                eqLine = eqLine & " + " & BuildLinearTermLatex(add_b, var2)
            Else
                Dim tempB As FractionSurd
                tempB = add_b
                tempB.num.coeff = Abs(tempB.num.coeff)
                eqLine = eqLine & " - " & BuildLinearTermLatex(tempB, var2)
            End If
        Else
            eqLine = BuildLinearTermLatex(add_b, var2)
        End If
        
    End If
    
    latex = latex & _
        "& " & eqLine & _
        " = " & FormatFractionSurdToLatex(add_c) & _
        " \dots \text{(3)} \\[10pt]"
        
End If
    
    
    ' ---------------------------------
    ' SUBTRACT equations (Stage-1 corrected logic)
    ' ---------------------------------
    
    Dim absA1 As Long
    Dim absA2 As Long
    
    Dim firstEq As StandardForm
    Dim secondEq As StandardForm
    
    absA1 = Abs(sf1.aCoeff.num.coeff)
    absA2 = Abs(sf2.aCoeff.num.coeff)
    
    If absA1 >= absA2 Then
        
        ' sf1 - sf2
        firstEq = sf1
        secondEq = sf2
        
    Else
        
        ' sf2 - sf1
        firstEq = sf2
        secondEq = sf1
        
    End If
    
    ' Compute subtraction using selected order
    sub_a = SubtractFractionSurd(firstEq.aCoeff, secondEq.aCoeff)
    sub_b = SubtractFractionSurd(firstEq.bCoeff, secondEq.bCoeff)
    sub_c = SubtractFractionSurd(firstEq.constCoeff, secondEq.constCoeff)
    
    latex = latex & _
        "& \text{Subtracting equations:} \\[6pt]"
    
    latex = latex & _
        BuildVerticalSubtractLayout(firstEq, secondEq, var1, var2) & _
        "\\[8pt]"
    
    
    ' ---------- Divide if possible ----------
    
    Dim gSub As Long
    gSub = Application.WorksheetFunction.GCD( _
                Abs(sub_a.num.coeff), _
                Abs(sub_b.num.coeff), _
                Abs(sub_c.num.coeff))
    
    If gSub > 1 Then
    
        latex = latex & _
            "& \text{Dividing the equation by } " & gSub & _
            "\text{ we get} \\[6pt]"
        
        sub_a.num.coeff = sub_a.num.coeff / gSub
        sub_b.num.coeff = sub_b.num.coeff / gSub
        sub_c.num.coeff = sub_c.num.coeff / gSub
        
    End If
    
    ' --- Now print equation (4) properly ---
    
    Dim signY As String
    Dim absY As FractionSurd
    
    absY = sub_b
    absY.num.coeff = Abs(absY.num.coeff)
    
    If sub_b.num.coeff < 0 Then
        signY = " - "
    Else
        signY = " + "
    End If

    latex = latex & _
        "& " & BuildLinearTermLatex(sub_a, var1) & _
        signY & BuildLinearTermLatex(absY, var2) & _
        " = " & FormatFractionSurdToLatex(sub_c) & _
        " \dots \text{(4)} \\[10pt]"
        
    ' ---- Build Linear Expression Cleanly ----
    ' ---------------------------------
    ' STAGE-2 : Eliminate y from (3) and (4)
    ' ---------------------------------
    
    Dim eq3_a As FractionSurd
    Dim eq3_b As FractionSurd
    Dim eq3_c As FractionSurd
    
    Dim eq4_a As FractionSurd
    Dim eq4_b As FractionSurd
    Dim eq4_c As FractionSurd
    
    ' (3)  x + y = A
    eq3_a = add_a
    eq3_b = add_b
    eq3_c = add_c
    
    ' (4)  x - y = B
    eq4_a = sub_a
    eq4_b = sub_b
    eq4_c = sub_c
    
    latex = latex & _
    "& \text{Adding equations (3) and (4):} \\[6pt]"
    
    latex = latex & _
    BuildVerticalAddLayout( _
        CreateStandardForm(eq3_a, eq3_b, eq3_c), _
        CreateStandardForm(eq4_a, eq4_b, eq4_c), _
        var1, var2) & _
    "\\[8pt]"
    
    
    ' ---------------------------------
    ' Compute result : 2x = A + B
    ' ---------------------------------
    
    Dim final_a As FractionSurd
    Dim final_c As FractionSurd
    
    final_a = AddFractionSurd(eq3_a, eq4_a)
    final_c = AddFractionSurd(eq3_c, eq4_c)
    
    latex = latex & _
    "& " & FormatFractionSurdToLatex(final_a) & var1 & _
    " = " & FormatFractionSurdToLatex(final_c) & " \\[8pt]"
    
    
    ' ---------------------------------
    ' x = constant / coefficient
    ' ---------------------------------
    
    Dim xVal As FractionSurd
    
    latex = latex & _
    "& " & var1 & " = \frac{" & _
    FormatFractionSurdToLatex(final_c) & "}{" & _
    FormatFractionSurdToLatex(final_a) & "} \\[6pt]"
    
    xVal = DivideFractionSurd(final_c, final_a)
    
    SimplifyFractionSurd xVal
    
    latex = latex & _
    "& " & var1 & " = " & _
    FormatFractionSurdToLatex(xVal) & " \\[10pt]"
    ' ---------------------------------
    ' STAGE-3 : Substitute to find y
    ' ---------------------------------
    Dim yVal As FractionSurd
    
    latex = latex & _
    "& \text{Substitute } " & var1 & " = " & _
    FormatFractionSurdToLatex(xVal) & _
    "\text{ in equation (3)} \\[8pt]"
    
    ' x + y = A
    latex = latex & _
    "& " & BuildLinearTermLatex(eq3_a, var1) & _
    " + " & BuildLinearTermLatex(eq3_b, var2) & _
    " = " & FormatFractionSurdToLatex(eq3_c) & " \\[6pt]"
    
    ' substitute value
    latex = latex & _
    "& " & FormatFractionSurdToLatex(xVal) & _
    " + " & var2 & _
    " = " & FormatFractionSurdToLatex(eq3_c) & " \\[6pt]"
    
    ' y = A - x
    latex = latex & _
    "& " & var2 & " = " & _
    FormatFractionSurdToLatex(eq3_c) & _
    " - " & FormatFractionSurdToLatex(xVal) & " \\[6pt]"
    
    yVal = SubtractFractionSurd(eq3_c, xVal)
    
    SimplifyFractionSurd yVal
    
    latex = latex & _
    "& " & var2 & " = " & _
    FormatFractionSurdToLatex(yVal) & " \\[10pt]"
    
    ' ---------------------------------
    ' FINAL SOLUTION
    ' ---------------------------------
    
    latex = latex & _
    "& \therefore \: (" & var1 & "," & var2 & ") = (" & _
    FormatFractionSurdToLatex(xVal) & "," & _
    FormatFractionSurdToLatex(yVal) & ")"
            
    
    CrossEliminationSequenceCore = latex

End Function
Public Function LCMEliminationSequence(sf1 As StandardForm, _
                                       sf2 As StandardForm, _
                                       var1 As String, _
                                       var2 As String) As String

    Dim latex As String
    latex = ""

    Dim b1 As Long
    Dim b2 As Long
    Dim lcmVal As Long
    
    b1 = Abs(sf1.bCoeff.num.coeff)
    b2 = Abs(sf2.bCoeff.num.coeff)

    lcmVal = Application.WorksheetFunction.lcm(b1, b2)

    Dim m1 As Long
    Dim m2 As Long

    m1 = lcmVal / b1
    m2 = lcmVal / b2


    latex = latex & _
        "& \text{Applying LCM elimination on } " & var2 & " \\[6pt]"


    ' ----------------------------------
    ' Display multiplication step
    ' ----------------------------------

    Dim multText As String
multText = ""

If m1 <> 1 And m2 <> 1 Then

    multText = "\text{Multiplying equation (1) by } " & m1 & _
               "\text{ and equation (2) by } " & m2

ElseIf m1 <> 1 Then

    multText = "\text{Multiplying equation (1) by } " & m1

ElseIf m2 <> 1 Then

    multText = "\text{Multiplying equation (2) by } " & m2

End If

If multText <> "" Then

    latex = latex & _
        "& " & multText & " \\[8pt]"

End If

    ' ----------------------------------
    ' Create multiplied equations
    ' ----------------------------------

    Dim nsf1 As StandardForm
    Dim nsf2 As StandardForm

    nsf1 = MultiplyEquation(sf1, m1)
    nsf2 = MultiplyEquation(sf2, m2)


    ' ----------------------------------
    ' Print multiplied equations
    ' ----------------------------------

    latex = latex & _
        "& " & BuildEquationLatex(nsf1.aCoeff, nsf1.bCoeff, nsf1.constCoeff, var1, var2) & _
        " \dots \text{(3)} \\[6pt]"

    latex = latex & _
        "& " & BuildEquationLatex(nsf2.aCoeff, nsf2.bCoeff, nsf2.constCoeff, var1, var2) & _
        " \dots \text{(4)} \\[10pt]"



    ' ----------------------------------
    ' Decide ADD or SUBTRACT automatically
    ' ----------------------------------

    Dim useAdd As Boolean

    If Sgn(nsf1.bCoeff.num.coeff) <> Sgn(nsf2.bCoeff.num.coeff) Then
        useAdd = True
    Else
        useAdd = False
    End If


    ' ----------------------------------
    ' Decide equation order (larger |a|)
    ' ----------------------------------

    Dim absA1 As Long
    Dim absA2 As Long

    absA1 = Abs(nsf1.aCoeff.num.coeff)
    absA2 = Abs(nsf2.aCoeff.num.coeff)

    Dim topEq As StandardForm
    Dim bottomEq As StandardForm

    If absA1 >= absA2 Then
    
        topEq = nsf1
        bottomEq = nsf2
        
    Else
    
        topEq = nsf2
        bottomEq = nsf1
        
    End If


    ' ----------------------------------
    ' Vertical elimination layout
    ' ----------------------------------

    If useAdd Then
    
        latex = latex & _
            "& \text{Adding equations (4) and (3)} \\[6pt]"
            
        latex = latex & _
            BuildVerticalAddLayout(topEq, bottomEq, var1, var2)
        
    Else
    
        latex = latex & _
            "& \text{Subtracting equations (4) and (3)} \\[6pt]"
            
        latex = latex & _
            BuildVerticalSubtractLayout(topEq, bottomEq, var1, var2)
        
    End If


    ' ----------------------------------
    ' Compute result equation
    ' ----------------------------------

    Dim newA As FractionSurd
    Dim newB As FractionSurd
    Dim newC As FractionSurd

    If useAdd Then
    
        newA = AddFractionSurd(topEq.aCoeff, bottomEq.aCoeff)
        newB = AddFractionSurd(topEq.bCoeff, bottomEq.bCoeff)
        newC = AddFractionSurd(topEq.constCoeff, bottomEq.constCoeff)
        
    Else
    
        newA = SubtractFractionSurd(topEq.aCoeff, bottomEq.aCoeff)
        newB = SubtractFractionSurd(topEq.bCoeff, bottomEq.bCoeff)
        newC = SubtractFractionSurd(topEq.constCoeff, bottomEq.constCoeff)
        
    End If


   ' ----------------------------------
    ' Print result equation
    ' ----------------------------------
    
    latex = latex & _
        PrintEquation(newA, newB, newC, var1, var2)
    
    
    
    ' ----------------------------------
    ' Choose easier original equation
    ' ----------------------------------
    
    Dim useEq As StandardForm
    
    Dim score1 As Long
    Dim score2 As Long
    
    score1 = Abs(sf1.aCoeff.num.coeff) + Abs(sf1.bCoeff.num.coeff)
    score2 = Abs(sf2.aCoeff.num.coeff) + Abs(sf2.bCoeff.num.coeff)
    
    If score1 <= score2 Then
        useEq = sf1
    Else
        useEq = sf2
    End If
    
    
    ' ----------------------------------
    ' Solve remaining variable
    ' ----------------------------------
    
    'latex = latex & _
    '    SolveAfterElimination(newA, newB, newC, useEq, var1, var2)
    latex = latex & _
    FinalizeEliminationStep(newA, newB, newC, sf1, var1, var2)
    
    LCMEliminationSequence = latex

End Function
Public Function LCMEliminationSequenceCore(sf1 As StandardForm, _
                                       sf2 As StandardForm, _
                                       var1 As String, _
                                       var2 As String) As String

    Dim latex As String
    latex = ""

    Dim b1 As Long
    Dim b2 As Long
    Dim lcmVal As Long
    
    b1 = Abs(sf1.bCoeff.num.coeff)
    b2 = Abs(sf2.bCoeff.num.coeff)
    
    lcmVal = Application.WorksheetFunction.lcm(b1, b2)
    
    Dim m1 As Long
    Dim m2 As Long
    
    m1 = lcmVal / b1
    m2 = lcmVal / b2


    ' -----------------------------
    ' LCM step text
    ' -----------------------------
    
    latex = latex & _
        "& \text{Applying LCM elimination on } " & var2 & " \\[6pt]"


    If m1 = 1 And m2 <> 1 Then
    
        latex = latex & _
            "& \text{Multiplying equation (2) by } " & m2 & " \\[8pt]"
            
    ElseIf m2 = 1 And m1 <> 1 Then
    
        latex = latex & _
            "& \text{Multiplying equation (1) by } " & m1 & " \\[8pt]"
            
    Else
    
        latex = latex & _
            "& \text{Multiplying equation (1) by } " & m1 & _
            "\text{ and equation (2) by } " & m2 & " \\[8pt]"
            
    End If


    ' -----------------------------
    ' Build multiplied equations
    ' -----------------------------
    
    Dim nsf1 As StandardForm
    Dim nsf2 As StandardForm

    nsf1 = MultiplyEquation(sf1, m1)
    nsf2 = MultiplyEquation(sf2, m2)


    ' -----------------------------
    ' Equation labels
    ' -----------------------------
    
    Dim eqA As StandardForm
    Dim eqB As StandardForm
    
    Dim labelA As String
    Dim labelB As String
    
    Dim nextLabel As Integer
    nextLabel = 3


    ' Equation (1)
    If m1 = 1 Then
    
        eqA = sf1
        labelA = "1"
        
    Else
    
        eqA = nsf1
        labelA = CStr(nextLabel)
        
        latex = latex & _
            "& " & BuildLinearTermLatex(eqA.aCoeff, var1) & _
            " + " & BuildLinearTermLatex(eqA.bCoeff, var2) & _
            " = " & FormatFractionSurdToLatex(eqA.constCoeff) & _
            " \dots \text{(" & labelA & ")} \\[6pt]"
            
        nextLabel = nextLabel + 1
        
    End If


    ' Equation (2)
    If m2 = 1 Then
    
        eqB = sf2
        labelB = "2"
        
    Else
    
        eqB = nsf2
        labelB = CStr(nextLabel)
        
        latex = latex & _
            "& " & BuildLinearTermLatex(eqB.aCoeff, var1) & _
            " + " & BuildLinearTermLatex(eqB.bCoeff, var2) & _
            " = " & FormatFractionSurdToLatex(eqB.constCoeff) & _
            " \dots \text{(" & labelB & ")} \\[10pt]"
            
    End If


    ' -----------------------------
    ' Larger |a| rule
    ' -----------------------------
    
    Dim topEq As StandardForm
    Dim bottomEq As StandardForm
    
    Dim topLabel As String
    Dim bottomLabel As String
    
    If Abs(eqA.aCoeff.num.coeff) >= Abs(eqB.aCoeff.num.coeff) Then
    
        topEq = eqA
        bottomEq = eqB
        
        topLabel = labelA
        bottomLabel = labelB
        
    Else
    
        topEq = eqB
        bottomEq = eqA
        
        topLabel = labelB
        bottomLabel = labelA
        
    End If


    ' -----------------------------
    ' Subtraction step
    ' -----------------------------
    
    Dim bTop As Long
    Dim bBottom As Long
    
    bTop = topEq.bCoeff.num.coeff
    bBottom = bottomEq.bCoeff.num.coeff
    
    If Sgn(bTop) <> Sgn(bBottom) Then
    
        latex = latex & _
            "& \text{Adding equations (" & topLabel & ") and (" & bottomLabel & ")} \\[6pt]"
        
        latex = latex & _
            BuildVerticalAddLayout(topEq, bottomEq, var1, var2)
    
    Else
    
        latex = latex & _
            "& \text{Subtracting equations (" & topLabel & ") and (" & bottomLabel & ")} \\[6pt]"
        
        latex = latex & _
            BuildVerticalSubtractLayout(topEq, bottomEq, var1, var2)
    
    End If
    ' -----------------------------
' Result equation
' -----------------------------

Dim newA As FractionSurd
Dim newB As FractionSurd
Dim newC As FractionSurd

If Sgn(bTop) <> Sgn(bBottom) Then

    ' ADD equations
    newA = AddFractionSurd(topEq.aCoeff, bottomEq.aCoeff)
    newB = AddFractionSurd(topEq.bCoeff, bottomEq.bCoeff)
    newC = AddFractionSurd(topEq.constCoeff, bottomEq.constCoeff)

Else

    ' SUBTRACT equations
    newA = SubtractFractionSurd(topEq.aCoeff, bottomEq.aCoeff)
    newB = SubtractFractionSurd(topEq.bCoeff, bottomEq.bCoeff)
    newC = SubtractFractionSurd(topEq.constCoeff, bottomEq.constCoeff)

End If


    ' Clean printing of result row
    Dim eqLine As String
    eqLine = ""
    
    If Not IsZeroFraction(newA) Then
        eqLine = BuildLinearTermLatex(newA, var1)
    End If
    
    If Not IsZeroFraction(newB) Then
        
        If eqLine <> "" Then
        
            If newB.num.coeff > 0 Then
            
                eqLine = eqLine & " + " & BuildLinearTermLatex(newB, var2)
                
            Else
            
                Dim tempB As FractionSurd
                tempB = newB
                tempB.num.coeff = Abs(tempB.num.coeff)
                
                eqLine = eqLine & " - " & BuildLinearTermLatex(tempB, var2)
                
            End If
            
        Else
        
            eqLine = BuildLinearTermLatex(newB, var2)
            
        End If
        
    End If
    
    
    latex = latex & _
        "& " & eqLine & _
        " = " & FormatFractionSurdToLatex(newC) & " \\[10pt]"

    ' -----------------------------
    ' Choose easier original equation
    ' -----------------------------
    
    Dim useEq As StandardForm
    
    Dim score1 As Long
    Dim score2 As Long
    
    score1 = Abs(sf1.aCoeff.num.coeff) + Abs(sf1.bCoeff.num.coeff)
    score2 = Abs(sf2.aCoeff.num.coeff) + Abs(sf2.bCoeff.num.coeff)
    
    If score1 <= score2 Then
        useEq = sf1
    Else
        useEq = sf2
    End If


    latex = latex & _
        AppendEliminationSolve(newA, newC, useEq, var1, var2)
    
    LCMEliminationSequenceCore = latex

End Function
Public Function AppendEliminationSteps(sf1 As StandardForm, _
                                       sf2 As StandardForm, _
                                       var1 As String, _
                                       var2 As String) As String

    Dim latex As String
    Dim methodType As String

    latex = ""

    ' ---------------------------------
    ' Select elimination strategy
    ' ---------------------------------

    methodType = SelectEliminationMethod(sf1, sf2)

    If methodType = "DIRECT" Then

        latex = latex & _
            DirectEliminationSequence(sf1, sf2, var1, var2)

    ElseIf methodType = "CROSS" Then

        latex = latex & _
            CrossEliminationSequence(sf1, sf2, var1, var2)

    Else

        latex = latex & _
            LCMEliminationSequence(sf1, sf2, var1, var2)

    End If

    AppendEliminationSteps = latex

End Function
Public Function SelectEliminationMethod(sf1 As StandardForm, _
                                        sf2 As StandardForm) As String

    Dim a1 As Long
    Dim a2 As Long
    Dim b1 As Long
    Dim b2 As Long
    
    a1 = Abs(sf1.aCoeff.num.coeff)
    a2 = Abs(sf2.aCoeff.num.coeff)
    
    b1 = Abs(sf1.bCoeff.num.coeff)
    b2 = Abs(sf2.bCoeff.num.coeff)

    
    ' --------------------------------
    ' 1. Direct Elimination
    ' --------------------------------
    
    If a1 = a2 Or b1 = b2 Then
        SelectEliminationMethod = "DIRECT"
        Exit Function
    End If
    
    
    ' --------------------------------
    ' 2. Cross Elimination
    ' --------------------------------
    
    If a1 = b2 And a2 = b1 Then
        SelectEliminationMethod = "CROSS"
        Exit Function
    End If
    
    
    ' --------------------------------
    ' 3. LCM Elimination
    ' --------------------------------
    
    SelectEliminationMethod = "LCM"

End Function
Public Function FinalizeEliminationStep(newA As FractionSurd, _
                                        newB As FractionSurd, _
                                        newC As FractionSurd, _
                                        useEq As StandardForm, _
                                        var1 As String, _
                                        var2 As String) As String

    Dim latex As String

    latex = PrintEquation(newA, newB, newC, var1, var2) & _
            SolveAfterElimination(newA, newB, newC, useEq, var1, var2)

    FinalizeEliminationStep = latex

End Function
Public Function SolveAfterElimination(a As FractionSurd, _
                                      b As FractionSurd, _
                                      c As FractionSurd, _
                                      useEq As StandardForm, _
                                      var1 As String, _
                                      var2 As String) As String

    Dim latex As String
    Dim xVal As FractionSurd
    Dim yVal As FractionSurd
    
    latex = ""

    ' -----------------------------
    ' Solve first variable
    ' -----------------------------
    
    xVal = SolveSingleVariable(a, c)

    latex = latex & PrintVariableResult(var1, xVal)

    ' -----------------------------
    ' Substitute into equation
    ' -----------------------------
    
    latex = latex & _
        "& \text{Substitute } " & var1 & " = " & _
        FormatFractionSurdToLatex(xVal) & _
        "\text{ in equation (1)} \\[6pt]"
    
    latex = latex & _
        PrintEquation(useEq.aCoeff, useEq.bCoeff, useEq.constCoeff, _
                      var1, var2)

    latex = latex & _
        BuildSubstitutionSteps(useEq, xVal, var1, var2)

    ' -----------------------------
    ' Compute y
    ' -----------------------------
    
    Dim axPart As FractionSurd
    Dim tempConst As FractionSurd
    
    axPart = MultiplyTwoFractionSurds(useEq.aCoeff, xVal)
    SimplifyFractionSurd axPart
    
    tempConst = SubtractFractionSurd(useEq.constCoeff, axPart)
    SimplifyFractionSurd tempConst
    
    If Abs(useEq.bCoeff.num.coeff) = 1 Then
    
        ' y = tempConst already printed earlier
        yVal = tempConst
    
    Else
    
        latex = latex & _
            "& " & var2 & " = \frac{" & _
            FormatFractionSurdToLatex(tempConst) & "}{" & _
            FormatFractionSurdToLatex(useEq.bCoeff) & "} \\"
    
        yVal = DivideFractionSurd(tempConst, useEq.bCoeff)
        SimplifyFractionSurd yVal
    
        latex = latex & _
            "& " & var2 & " = " & _
            FormatFractionSurdToLatex(yVal) & " \\[10pt]"
    
    End If
    ' -----------------------------
    ' Final answer
    ' -----------------------------

    latex = latex & _
        PrintFinalSolution(var1, var2, xVal, yVal)

    SolveAfterElimination = latex

End Function
Public Function EliminateVariable(sf1 As StandardForm, _
                                  sf2 As StandardForm, _
                                  var1 As String, _
                                  var2 As String, _
                                  eliminateVar1 As Boolean) As String

    Dim latex As String
    
    Dim newCoeff As FractionSurd
    Dim newConst As FractionSurd
    
    If eliminateVar1 Then
        
        ' Eliminate var1 (x)
        
        newCoeff = AddFractionSurd(sf1.bCoeff, sf2.bCoeff)
        newConst = AddFractionSurd(sf1.constCoeff, sf2.constCoeff)
        
        latex = latex & _
            "& " & FormatFractionSurdToLatex(newCoeff) & var2 & _
            " = " & FormatFractionSurdToLatex(newConst) & " \\[8pt]"
        
        Dim yVal As FractionSurd
        yVal = DivideFractionSurd(newConst, newCoeff)
        
        latex = latex & _
            "& " & var2 & " = " & _
            FormatFractionSurdToLatex(yVal) & " \\[8pt]"
        
    Else
        
        ' Eliminate var2 (y)
        
        newCoeff = AddFractionSurd(sf1.aCoeff, sf2.aCoeff)
        newConst = AddFractionSurd(sf1.constCoeff, sf2.constCoeff)
        
        latex = latex & _
            "& " & FormatFractionSurdToLatex(newCoeff) & var1 & _
            " = " & FormatFractionSurdToLatex(newConst) & " \\[8pt]"
        
        Dim xVal As FractionSurd
        xVal = DivideFractionSurd(newConst, newCoeff)
        
        latex = latex & _
            "& " & var1 & " = " & _
            FormatFractionSurdToLatex(xVal) & " \\[8pt]"
        
    End If
    
    EliminateVariable = latex

End Function
Public Function SolveLinearSingleVariable(a As FractionSurd, _
                                          b As FractionSurd, _
                                          c As FractionSurd, _
                                          varName As String) As String

    Dim latex As String
    latex = ""
    
    Dim rhs1 As FractionSurd
    Dim rhs2 As FractionSurd
    Dim result As FractionSurd
    
    ' ax - b = c
    
    latex = latex & _
        "& " & BuildLinearTermLatex(a, varName) & _
        " - " & FormatFractionSurdToLatex(b) & _
        " = " & FormatFractionSurdToLatex(c) & " \\[6pt]"
    
    
    ' ax = c + b
    
    latex = latex & _
        "& " & BuildLinearTermLatex(a, varName) & _
        " = " & FormatFractionSurdToLatex(c) & _
        " + " & FormatFractionSurdToLatex(b) & " \\[6pt]"
    
    
    ' compute c + b
    
    rhs1 = AddFractionSurd(c, b)
    
    latex = latex & _
        "& " & BuildLinearTermLatex(a, varName) & _
        " = " & FormatFractionSurdToLatex(rhs1) & " \\[6pt]"
    
    
    ' x = rhs / a
    
    latex = latex & _
        "& " & varName & " = \frac{" & _
        FormatFractionSurdToLatex(rhs1) & "}{" & _
        FormatFractionSurdToLatex(a) & "} \\[6pt]"
    
    
    result = DivideFractionSurd(rhs1, a)
    
    SimplifyFractionSurd result
    
    
    latex = latex & _
        "& " & varName & " = " & _
        FormatFractionSurdToLatex(result) & " \\[10pt]"
    
    
    SolveLinearSingleVariable = latex

End Function
Public Function SolveCrossEquations(eq3 As StandardForm, _
                                    eq4 As StandardForm, _
                                    var1 As String, _
                                    var2 As String) As String

    Dim latex As String
    latex = ""
    
    Dim newA As FractionSurd
    Dim newC As FractionSurd
    
    ' Add equations
    
    newA = AddFractionSurd(eq3.aCoeff, eq4.aCoeff)
    newC = AddFractionSurd(eq3.constCoeff, eq4.constCoeff)
    
    
    latex = latex & _
        "& \text{Adding equation (3) and (4)} \\[6pt]"
    
    
    latex = latex & _
        "& " & BuildLinearTermLatex(eq3.aCoeff, var1) & _
        " + " & BuildLinearTermLatex(eq3.bCoeff, var2) & _
        " = " & FormatFractionSurdToLatex(eq3.constCoeff) & " \\"
    
    
    latex = latex & _
        "& " & BuildLinearTermLatex(eq4.aCoeff, var1) & _
        " - " & BuildLinearTermLatex(eq4.bCoeff, var2) & _
        " = " & FormatFractionSurdToLatex(eq4.constCoeff) & " \\"
    
    
    latex = latex & _
        "& " & BuildLinearTermLatex(newA, var1) & _
        " = " & FormatFractionSurdToLatex(newC) & " \\[8pt]"
    
    
    SolveCrossEquations = latex

End Function
Public Function SolveSingleCoefficient(a As FractionSurd, _
                                       c As FractionSurd, _
                                       varName As String) As String

    Dim latex As String
    latex = ""
    
    Dim result As FractionSurd
    
    
    ' ax = c
    
    latex = latex & _
        "& " & BuildLinearTermLatex(a, varName) & _
        " = " & FormatFractionSurdToLatex(c) & " \\[6pt]"
    
    
    ' x = c / a
    
    latex = latex & _
        "& " & varName & " = \frac{" & _
        FormatFractionSurdToLatex(c) & "}{" & _
        FormatFractionSurdToLatex(a) & "} \\[6pt]"
    
    
    result = DivideFractionSurd(c, a)
    
    SimplifyFractionSurd result
    
    
    latex = latex & _
        "& " & varName & " = " & _
        FormatFractionSurdToLatex(result) & " \\[10pt]"
    
    
    SolveSingleCoefficient = latex

End Function
