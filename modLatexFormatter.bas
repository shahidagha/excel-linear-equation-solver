Attribute VB_Name = "modLatexFormatter"
Public Function FormatFractionSurdToLatex(fr As FractionSurd) As String

    Dim numLatex As String
    Dim denLatex As String
    Dim signStr As String
    
    numLatex = FormatSurdToLatex(fr.num)
    denLatex = FormatSurdToLatex(fr.den)
    
    ' Detect sign from numerator only
    If left(Trim(numLatex), 1) = "-" Then
        signStr = "-"
        numLatex = Mid(Trim(numLatex), 2)
    Else
        signStr = ""
    End If
    
    If denLatex = "1" Then
        FormatFractionSurdToLatex = signStr & numLatex
    Else
        FormatFractionSurdToLatex = signStr & "\frac{" & numLatex & "}{" & denLatex & "}"
    End If

End Function
Public Function AppendBackSubstitution(a As FractionSurd, _
                                       c As FractionSurd, _
                                       xVal As FractionSurd, _
                                       var1 As String, _
                                       var2 As String) As String

    Dim latex As String
    latex = ""
    
    Dim ax As FractionSurd
    Dim yVal As FractionSurd
    
    ' y = c - ax
    
    latex = latex & _
        "& " & var2 & " = " & _
        FormatFractionSurdToLatex(c) & _
        " - " & BuildLinearTermLatex(a, var1) & " \\[6pt]"
    
    
    ' y = c - a(value)
    
    latex = latex & _
        "& " & var2 & " = " & _
        FormatFractionSurdToLatex(c) & _
        " - " & FormatFractionSurdToLatex(a) & _
        "(" & FormatFractionSurdToLatex(xVal) & ") \\[6pt]"
    
    
    ' compute ax
    
    ax = MultiplyFractionSurd(a, xVal)
    
    
    latex = latex & _
        "& " & var2 & " = " & _
        FormatFractionSurdToLatex(c) & _
        " - " & FormatFractionSurdToLatex(ax) & " \\[6pt]"
    
    
    ' compute final y
    
    yVal = SubtractFractionSurd(c, ax)
    
    SimplifyFractionSurd yVal
    
    
    latex = latex & _
        "& " & var2 & " = " & _
        FormatFractionSurdToLatex(yVal) & " \\[10pt]"
    
    
    latex = latex & _
        "& \therefore (" & var1 & "," & var2 & ") = (" & _
        FormatFractionSurdToLatex(xVal) & "," & _
        FormatFractionSurdToLatex(yVal) & ")"
    
    
    AppendBackSubstitution = latex

End Function
Public Function AppendCramerStep1(sf1 As StandardForm, _
                                  sf2 As StandardForm, _
                                  pVar As String, _
                                  sVar As String) As String

    Dim latex As String

    latex = ""

    latex = latex & _
        "& \; \text{Comparing eq(1) with } a_1" & pVar & _
        " + b_1" & sVar & " = c_1 \;\text{, we get }\\[4pt]"

    latex = latex & _
        "& \qquad a_1 = " & FormatFractionSurdToLatex(sf1.aCoeff) & _
        ",\;b_1 = " & FormatFractionSurdToLatex(sf1.bCoeff) & _
        ",\;c_1 = " & FormatFractionSurdToLatex(sf1.constCoeff) & _
        " \\[8pt]"

    latex = latex & _
        "& \; \text{Comparing eq(2) with } a_2" & pVar & _
        " + b_2" & sVar & " = c_2 \; \text{, we get}\\[4pt]"

    latex = latex & _
        "& \qquad a_2 = " & FormatFractionSurdToLatex(sf2.aCoeff) & _
        ",\;b_2 = " & FormatFractionSurdToLatex(sf2.bCoeff) & _
        ",\;c_2 = " & FormatFractionSurdToLatex(sf2.constCoeff) & _
        " \\[8pt]"

    AppendCramerStep1 = latex

End Function
Public Function AppendDeterminantD(sf1 As StandardForm, _
                                   sf2 As StandardForm) As String

    Dim latex As String
    
    Dim a1 As String, b1 As String
    Dim a2 As String, b2 As String
    
    Dim prod1 As FractionSurd
    Dim prod2 As FractionSurd
    Dim D As FractionSurd
    
    Dim prod1Disp As String
    Dim prod2Disp As String
    Dim prod2Abs As String
    Dim finalD As String
    
    ' -----------------------------
    ' Format coefficients
    ' -----------------------------
    
    a1 = FormatFractionSurdToLatex(sf1.aCoeff)
    b1 = FormatFractionSurdToLatex(sf1.bCoeff)
    a2 = FormatFractionSurdToLatex(sf2.aCoeff)
    b2 = FormatFractionSurdToLatex(sf2.bCoeff)
    
    ' Wrap individual coefficients ONLY if negative
    If left(a1, 1) = "-" Then a1 = "(" & a1 & ")"
    If left(b1, 1) = "-" Then b1 = "(" & b1 & ")"
    If left(a2, 1) = "-" Then a2 = "(" & a2 & ")"
    If left(b2, 1) = "-" Then b2 = "(" & b2 & ")"
    
    ' -----------------------------
    ' Compute products
    ' -----------------------------
    
    prod1 = MultiplyFractionBySurd(sf1.aCoeff, sf2.bCoeff.num)
    prod2 = MultiplyFractionBySurd(sf2.aCoeff, sf1.bCoeff.num)
    
    ' Subtract
    D = prod1
    D.num = SubtractSurds(prod1.num, prod2.num)
    
    ' Display versions
    prod1Disp = FormatFractionSurdToLatex(prod1)
    prod2Disp = FormatFractionSurdToLatex(prod2)
    finalD = FormatFractionSurdToLatex(D)
    
    ' -----------------------------
    ' Build LaTeX
    ' -----------------------------
    
    latex = ""
    
    ' Matrix line
    latex = latex & _
        "& D=\begin{vmatrix} " & _
        FormatFractionSurdToLatex(sf1.aCoeff) & " & " & _
        FormatFractionSurdToLatex(sf1.bCoeff) & " \\ " & _
        FormatFractionSurdToLatex(sf2.aCoeff) & " & " & _
        FormatFractionSurdToLatex(sf2.bCoeff) & _
        " \end{vmatrix} \\[6pt]"
    
    ' Line 1: product form
    latex = latex & _
        "& \qquad = " & a1 & " \times " & b2 & _
        " - " & a2 & " \times " & b1 & " \\[6pt]"
    
    ' Line 2: substituted products
    If left(prod2Disp, 1) = "-" Then
        prod2Disp = "(" & prod2Disp & ")"
    End If
    
    latex = latex & _
        "& \qquad = " & prod1Disp & _
        " - " & prod2Disp & " \\[6pt]"
    
    ' Line 3: sign resolution (only if second product negative)
    If left(FormatFractionSurdToLatex(prod2), 1) = "-" Then
        
        prod2Abs = Mid(FormatFractionSurdToLatex(prod2), 2)
        
        latex = latex & _
            "& \qquad = " & prod1Disp & _
            " + " & prod2Abs & " \\[6pt]"
    End If
    
    ' Final result
    latex = latex & _
        "& \qquad = " & finalD & " \\[10pt]"
    
    AppendDeterminantD = latex

End Function
Public Function AppendDeterminantDx(sf1 As StandardForm, _
                                    sf2 As StandardForm, _
                                    pVar As String) As String

    Dim latex As String
    
    Dim c1 As String, b1 As String
    Dim c2 As String, b2 As String
    
    Dim prod1 As FractionSurd
    Dim prod2 As FractionSurd
    Dim Dx As FractionSurd
    
    Dim prod1Disp As String
    Dim prod2Disp As String
    Dim prod2Abs As String
    Dim finalDx As String
    
    ' Format coefficients
    c1 = FormatFractionSurdToLatex(sf1.constCoeff)
    b1 = FormatFractionSurdToLatex(sf1.bCoeff)
    c2 = FormatFractionSurdToLatex(sf2.constCoeff)
    b2 = FormatFractionSurdToLatex(sf2.bCoeff)
    
    ' Wrap negatives (line 1 only)
    If left(c1, 1) = "-" Then c1 = "(" & c1 & ")"
    If left(b1, 1) = "-" Then b1 = "(" & b1 & ")"
    If left(c2, 1) = "-" Then c2 = "(" & c2 & ")"
    If left(b2, 1) = "-" Then b2 = "(" & b2 & ")"
    
    ' Compute products
    prod1 = MultiplyFractionBySurd(sf1.constCoeff, sf2.bCoeff.num)
    prod2 = MultiplyFractionBySurd(sf2.constCoeff, sf1.bCoeff.num)
    
    Dx = prod1
    Dx.num = SubtractSurds(prod1.num, prod2.num)
    
    prod1Disp = FormatFractionSurdToLatex(prod1)
    prod2Disp = FormatFractionSurdToLatex(prod2)
    finalDx = FormatFractionSurdToLatex(Dx)
    
    ' Matrix line
    latex = "& D_" & pVar & "=\begin{vmatrix} " & _
            FormatFractionSurdToLatex(sf1.constCoeff) & " & " & _
            FormatFractionSurdToLatex(sf1.bCoeff) & " \\ " & _
            FormatFractionSurdToLatex(sf2.constCoeff) & " & " & _
            FormatFractionSurdToLatex(sf2.bCoeff) & _
            " \end{vmatrix} \\[6pt]"
    
    ' Line 1
    latex = latex & _
        "& \qquad = " & c1 & " \times " & b2 & _
        " - " & c2 & " \times " & b1 & " \\[6pt]"
    
    ' Line 2
    If left(prod2Disp, 1) = "-" Then
        prod2Disp = "(" & prod2Disp & ")"
    End If
    
    latex = latex & _
        "& \qquad = " & prod1Disp & _
        " - " & prod2Disp & " \\[6pt]"
    
    ' Sign resolution
    If left(FormatFractionSurdToLatex(prod2), 1) = "-" Then
        prod2Abs = Mid(FormatFractionSurdToLatex(prod2), 2)
        latex = latex & _
            "& \qquad = " & prod1Disp & _
            " + " & prod2Abs & " \\[6pt]"
    End If
    
    ' Final
    latex = latex & _
        "& \qquad = " & finalDx & " \\[10pt]"
    
    AppendDeterminantDx = latex

End Function
Public Function AppendDeterminantDy(sf1 As StandardForm, _
                                    sf2 As StandardForm, _
                                    sVar As String) As String
    Dim latex As String
    
    Dim a1 As String, c1 As String
    Dim a2 As String, c2 As String
    
    Dim prod1 As FractionSurd
    Dim prod2 As FractionSurd
    Dim Dy As FractionSurd
    
    Dim prod1Disp As String
    Dim prod2Disp As String
    Dim prod2Abs As String
    Dim finalDy As String
    
    a1 = FormatFractionSurdToLatex(sf1.aCoeff)
    c1 = FormatFractionSurdToLatex(sf1.constCoeff)
    a2 = FormatFractionSurdToLatex(sf2.aCoeff)
    c2 = FormatFractionSurdToLatex(sf2.constCoeff)
    
    If left(a1, 1) = "-" Then a1 = "(" & a1 & ")"
    If left(c1, 1) = "-" Then c1 = "(" & c1 & ")"
    If left(a2, 1) = "-" Then a2 = "(" & a2 & ")"
    If left(c2, 1) = "-" Then c2 = "(" & c2 & ")"
    
    prod1 = MultiplyFractionBySurd(sf1.aCoeff, sf2.constCoeff.num)
    prod2 = MultiplyFractionBySurd(sf2.aCoeff, sf1.constCoeff.num)
    
    Dy = prod1
    Dy.num = SubtractSurds(prod1.num, prod2.num)
    
    prod1Disp = FormatFractionSurdToLatex(prod1)
    prod2Disp = FormatFractionSurdToLatex(prod2)
    finalDy = FormatFractionSurdToLatex(Dy)
    
    latex = "& D_" & sVar & "=\begin{vmatrix}" & _
            FormatFractionSurdToLatex(sf1.aCoeff) & " & " & _
            FormatFractionSurdToLatex(sf1.constCoeff) & " \\ " & _
            FormatFractionSurdToLatex(sf2.aCoeff) & " & " & _
            FormatFractionSurdToLatex(sf2.constCoeff) & _
            " \end{vmatrix} \\[6pt]"
    
    latex = latex & _
        "& \qquad = " & a1 & " \times " & c2 & _
        " - " & a2 & " \times " & c1 & " \\[6pt]"
    
    If left(prod2Disp, 1) = "-" Then
        prod2Disp = "(" & prod2Disp & ")"
    End If
    
    latex = latex & _
        "& \qquad = " & prod1Disp & _
        " - " & prod2Disp & " \\[6pt]"
    
    If left(FormatFractionSurdToLatex(prod2), 1) = "-" Then
        prod2Abs = Mid(FormatFractionSurdToLatex(prod2), 2)
        latex = latex & _
            "& \qquad = " & prod1Disp & _
            " + " & prod2Abs & " \\[6pt]"
    End If
    
    latex = latex & _
        "& \qquad = " & finalDy & " \\[10pt]"
    
    AppendDeterminantDy = latex

End Function
Public Function AppendEliminationSolve(newA As FractionSurd, _
                                       newC As FractionSurd, _
                                       useEq As StandardForm, _
                                       var1 As String, _
                                       var2 As String) As String

    Dim latex As String
    latex = ""

    Dim xVal As FractionSurd
    Dim axPart As FractionSurd
    Dim tempConst As FractionSurd
    Dim yVal As FractionSurd


    ' ----------------------------------------
    ' Solve first variable
    ' ----------------------------------------

    xVal = DivideFractionSurd(newC, newA)
    SimplifyFractionSurd xVal

    latex = latex & _
        "& " & var1 & " = " & _
        FormatFractionSurdToLatex(xVal) & " \\[10pt]"


    ' ----------------------------------------
    ' Substitute into chosen equation
    ' ----------------------------------------

    latex = latex & _
        "& \text{Substitute } " & var1 & " = " & _
        FormatFractionSurdToLatex(xVal) & _
        "\text{ in equation (1)} \\[6pt]"

    latex = latex & _
        "& " & BuildEquationLatex(useEq.aCoeff, useEq.bCoeff, useEq.constCoeff, var1, var2) & " \\"


    ' ----------------------------------------
    ' Substitute value
    ' ----------------------------------------

    latex = latex & _
        "& " & FormatFractionSurdToLatex(useEq.aCoeff) & _
        "(" & FormatFractionSurdToLatex(xVal) & ")" & _
        " + " & BuildLinearTermLatex(useEq.bCoeff, var2) & _
        " = " & FormatFractionSurdToLatex(useEq.constCoeff) & " \\"


    ' ----------------------------------------
    ' Compute a*x
    ' ----------------------------------------

    axPart = MultiplyTwoFractionSurds(useEq.aCoeff, xVal)
    SimplifyFractionSurd axPart

    latex = latex & _
        "& " & FormatFractionSurdToLatex(axPart) & _
        " + " & BuildLinearTermLatex(useEq.bCoeff, var2) & _
        " = " & FormatFractionSurdToLatex(useEq.constCoeff) & " \\"


    ' ----------------------------------------
    ' Move constant to RHS
    ' ----------------------------------------

    latex = latex & _
        "& " & BuildLinearTermLatex(useEq.bCoeff, var2) & _
        " = " & FormatFractionSurdToLatex(useEq.constCoeff) & _
        " - " & FormatFractionSurdToLatex(axPart) & " \\"


    tempConst = SubtractFractionSurd(useEq.constCoeff, axPart)
    SimplifyFractionSurd tempConst

    latex = latex & _
        "& " & BuildLinearTermLatex(useEq.bCoeff, var2) & _
        " = " & FormatFractionSurdToLatex(tempConst) & " \\"


    
   ' ----------------------------------------
    ' Divide by coefficient of y (only if needed)
    ' ----------------------------------------
    
    If useEq.bCoeff.num.coeff <> 1 Then
    
        yVal = DivideFractionSurd(tempConst, useEq.bCoeff)
        SimplifyFractionSurd yVal
    
        latex = latex & _
            "& " & var2 & " = " & _
            FormatFractionSurdToLatex(yVal) & " \\[10pt]"
    
    Else
    
        yVal = tempConst
        SimplifyFractionSurd yVal
    
    End If
    ' ----------------------------------------
    ' Final Answer
    ' ----------------------------------------

    latex = latex & _
        "& \therefore (" & var1 & "," & var2 & ") = (" & _
        FormatFractionSurdToLatex(xVal) & "," & _
        FormatFractionSurdToLatex(yVal) & ")"

    AppendEliminationSolve = latex

End Function
Public Function BuildDivisionStep(a As FractionSurd, _
                                  b As FractionSurd, _
                                  divisor As Long, _
                                  var1 As String, _
                                  eqLabel As String) As String

    Dim latex As String
    Dim newA As FractionSurd
    Dim newB As FractionSurd

    newA = a
    newB = b

    newA.num.coeff = newA.num.coeff / divisor
    newB.num.coeff = newB.num.coeff / divisor

    latex = ""
    latex = latex & "& \text{Dividing the equation by } " & divisor & _
            "\text{ we get} \\[6pt]"

    latex = latex & _
        "& " & FormatFractionSurdToLatex(newA) & var1 & _
        " = " & FormatFractionSurdToLatex(newB) & _
        " \dots \text{(" & eqLabel & ")} \\[10pt]"

    BuildDivisionStep = latex

End Function
Public Function FormatSignedTerm(fs As FractionSurd, varName As String) As String

    If fs.num.coeff < 0 Then
        FormatSignedTerm = "- " & AbsCoeff(fs, varName)
    Else
        FormatSignedTerm = "+ " & AbsCoeff(fs, varName)
    End If

End Function

Public Function BuildLinearExpression(a As FractionSurd, _
                                      b As FractionSurd, _
                                      var1 As String, _
                                      var2 As String) As String

    Dim expr As String
    expr = ""

    ' -------------------------
    ' First term (ax)
    ' -------------------------

    If Not IsZeroFraction(a) Then
        expr = BuildLinearTermLatex(a, var1)
    End If


    ' -------------------------
    ' Second term (by)
    ' -------------------------

    If Not IsZeroFraction(b) Then

        Dim absB As FractionSurd
        absB = b
        absB.num.coeff = Abs(absB.num.coeff)

        If expr <> "" Then

            If b.num.coeff > 0 Then
                expr = expr & " + " & BuildLinearTermLatex(absB, var2)
            Else
                expr = expr & " - " & BuildLinearTermLatex(absB, var2)
            End If

        Else

            If b.num.coeff < 0 Then
                expr = "-" & BuildLinearTermLatex(absB, var2)
            Else
                expr = BuildLinearTermLatex(absB, var2)
            End If

        End If

    End If

    BuildLinearExpression = expr

End Function
Public Function AppendCramerSteps(sf1 As StandardForm, _
                                  sf2 As StandardForm, _
                                  var1 As String, _
                                  var2 As String) As String

    Dim latex As String

    latex = ""

    latex = latex & _
        "& \\text{Cramer's rule solution not implemented yet.}"

    AppendCramerSteps = latex

End Function

Public Function AppendGraphicalSteps(sf1 As StandardForm, _
                                     sf2 As StandardForm, _
                                     var1 As String, _
                                     var2 As String) As String

    Dim latex As String
    
    latex = ""
    
    latex = latex & _
        "& \\text{Graphical method not implemented yet.}"
    
    AppendGraphicalSteps = latex

End Function
Public Function AppendFinalCramerStep(D As FractionSurd, _
                                      Dx As FractionSurd, _
                                      Dy As FractionSurd, _
                                      var1 As String, _
                                      var2 As String) As String

    Dim latex As String
    Dim xVal As FractionSurd
    Dim yVal As FractionSurd
    
    latex = ""
    
    ' ----------------------------------
    ' CASE 1 : D ? 0 ? Unique Solution
    ' ----------------------------------
    
    If Not IsZeroFraction(D) Then
        
        xVal = DivideFractionSurd(Dx, D)
        yVal = DivideFractionSurd(Dy, D)
        
        ' IMPORTANT: Simplify in-place (Sub, not Function)
        SimplifyFractionSurd xVal
        SimplifyFractionSurd yVal
        
        latex = latex & _
            "& " & var1 & "= \frac{D_" & var1 & "}{D} = " & _
            "\frac{" & FormatFractionSurdToLatex(Dx) & "}{" & _
            FormatFractionSurdToLatex(D) & "} = " & _
            FormatFractionSurdToLatex(xVal) & " \\[8pt]"
        
        latex = latex & _
            "& " & var2 & "= \frac{D_" & var2 & "}{D} = " & _
            "\frac{" & FormatFractionSurdToLatex(Dy) & "}{" & _
            FormatFractionSurdToLatex(D) & "} = " & _
            FormatFractionSurdToLatex(yVal) & " \\[10pt]"
        
        latex = latex & _
            "& \therefore \: (" & var1 & "," & var2 & ") = (" & _
            FormatFractionSurdToLatex(xVal) & "," & _
            FormatFractionSurdToLatex(yVal) & ")"
        
        AppendFinalCramerStep = latex
        Exit Function
    End If
    
    
    ' ----------------------------------
    ' CASE 2 : D = 0
    ' ----------------------------------
    
    latex = latex & _
        "& \text{Since } D = 0 \\[6pt]"
    
    
    ' Subcase A : Dx = 0 AND Dy = 0 ? Infinite solutions
    
    If IsZeroFraction(Dx) And IsZeroFraction(Dy) Then
        
        latex = latex & _
            "& \text{and } D_" & var1 & " = 0,\; D_" & var2 & " = 0 \\[6pt]"
        
        latex = latex & _
            "& \therefore \text{System has infinitely many solutions.}"
        
    Else
        
        ' Subcase B : Inconsistent
        
        latex = latex & _
            "& \text{but at least one of } D_" & var1 & ", D_" & var2 & " \neq 0 \\[6pt]"
        
        latex = latex & _
            "& \therefore \text{System has no solution.}"
        
    End If
    
    AppendFinalCramerStep = latex

End Function
Public Function AppendSubstitutionSolve(a2 As FractionSurd, _
                                        b2 As FractionSurd, _
                                        a1 As FractionSurd, _
                                        c1 As FractionSurd, _
                                        c2 As FractionSurd, _
                                        var1 As String, _
                                        var2 As String) As String

    Dim latex As String
    latex = ""

    Dim term1 As FractionSurd
    Dim term2 As FractionSurd
    Dim lhsCoeff As FractionSurd
    Dim rhsConst As FractionSurd
    Dim xVal As FractionSurd
    Dim yVal As FractionSurd
    
    Dim negA1 As FractionSurd
    
    ' -----------------------------
    ' Expand b2(c1 - a1x)
    ' -----------------------------
    
    term1 = MultiplyFractionSurd(c1, b2)
    
    negA1 = a1
    negA1.num.coeff = -negA1.num.coeff
    
    term2 = MultiplyFractionSurd(negA1, b2)
    
    
    ' -----------------------------
    ' Combine x terms
    ' -----------------------------
    
    lhsCoeff = AddFractionSurd(a2, term2)
    
    
    ' -----------------------------
    ' Move constants
    ' -----------------------------
    
    rhsConst = SubtractFractionSurd(c2, term1)
    
    
    latex = latex & _
    "& " & BuildLinearTermLatex(a2, var1) & _
    " + " & FormatFractionSurdToLatex(term1) & _
    " + " & BuildLinearTermLatex(term2, var1) & _
    " = " & FormatFractionSurdToLatex(c2) & " \\[6pt]"
    
    
    latex = latex & _
    "& " & BuildLinearTermLatex(lhsCoeff, var1) & _
    " + " & FormatFractionSurdToLatex(term1) & _
    " = " & FormatFractionSurdToLatex(c2) & " \\[6pt]"
    
    
    latex = latex & _
    "& " & BuildLinearTermLatex(lhsCoeff, var1) & _
    " = " & FormatFractionSurdToLatex(rhsConst) & " \\[8pt]"
    
    
    ' -----------------------------
    ' Solve x
    ' -----------------------------
    
    xVal = DivideFractionSurd(rhsConst, lhsCoeff)
    SimplifyFractionSurd xVal
    
    latex = latex & _
    "& " & var1 & " = \frac{" & _
    FormatFractionSurdToLatex(rhsConst) & "}{" & _
    FormatFractionSurdToLatex(lhsCoeff) & "} \\[6pt]"
    
    
    latex = latex & _
    "& " & var1 & " = " & _
    FormatFractionSurdToLatex(xVal) & " \\[10pt]"
    
    
    ' -----------------------------
    ' Back substitute
    ' -----------------------------
    
    latex = latex & _
    "& " & var2 & " = " & _
    FormatFractionSurdToLatex(c1) & _
    " - " & BuildLinearTermLatex(a1, var1) & " \\[6pt]"
    
    
    latex = latex & _
    "& " & var2 & " = " & _
    FormatFractionSurdToLatex(c1) & _
    " - " & FormatFractionSurdToLatex(a1) & _
    "(" & FormatFractionSurdToLatex(xVal) & ") \\[6pt]"
    
    
    Dim axPart As FractionSurd
    
    axPart = MultiplyFractionSurd(a1, xVal)
    
    yVal = SubtractFractionSurd(c1, axPart)
    
    SimplifyFractionSurd yVal
    
    
    latex = latex & _
    "& " & var2 & " = " & _
    FormatFractionSurdToLatex(c1) & _
    " - " & FormatFractionSurdToLatex(axPart) & " \\[6pt]"
    
    
    latex = latex & _
    "& " & var2 & " = " & _
    FormatFractionSurdToLatex(yVal) & " \\[10pt]"
    
    
    latex = latex & _
    "& \therefore (" & var1 & "," & var2 & ") = (" & _
    FormatFractionSurdToLatex(xVal) & "," & _
    FormatFractionSurdToLatex(yVal) & ")"
    
    
    AppendSubstitutionSolve = latex

End Function
Public Function BuildSignedTermLatex(f As FractionSurd, _
                                     varName As String) As String

    Dim term As String
    term = BuildLinearTermLatex(f, varName)
    
    If term = "" Then
        BuildSignedTermLatex = ""
        Exit Function
    End If
    
    If left(term, 1) = "-" Then
        BuildSignedTermLatex = " - " & Mid(term, 2)
    Else
        BuildSignedTermLatex = " + " & term
    End If

End Function
Public Function AppendSubstitutionSteps(sf1 As StandardForm, _
                                        sf2 As StandardForm, _
                                        var1 As String, _
                                        var2 As String) As String

    Dim latex As String
    latex = ""

    Dim a1 As FractionSurd, b1 As FractionSurd, c1 As FractionSurd
    Dim a2 As FractionSurd, b2 As FractionSurd, c2 As FractionSurd

    a1 = sf1.aCoeff
    b1 = sf1.bCoeff
    c1 = sf1.constCoeff

    a2 = sf2.aCoeff
    b2 = sf2.bCoeff
    c2 = sf2.constCoeff

    latex = latex & "& \text{Applying Substitution Method} \\[10pt]"

    ' -------------------------------------------------
    ' Stage-1 : Choose variable to isolate
    ' -------------------------------------------------

    Dim isolateY As Boolean

    If Abs(b1.num.coeff) = 1 Then
        isolateY = True
    ElseIf Abs(a1.num.coeff) = 1 Then
        isolateY = False
    ElseIf Abs(b1.num.coeff) <= Abs(a1.num.coeff) Then
        isolateY = True
    Else
        isolateY = False
    End If


    ' -------------------------------------------------
    ' Stage-2 : Isolation
    ' -------------------------------------------------

    If isolateY Then

        latex = latex & _
        "& \text{From equation (1)} \\[6pt]"

        latex = latex & _
        "& " & BuildLinearTermLatex(a1, var1) & _
        " + " & BuildLinearTermLatex(b1, var2) & _
        " = " & FormatFractionSurdToLatex(c1) & " \\[6pt]"

        latex = latex & _
        "& " & var2 & " = " & _
        FormatFractionSurdToLatex(c1) & " - " & _
        BuildLinearTermLatex(a1, var1) & " \\[10pt]"


        ' -------------------------------------------------
        ' Stage-3 : Substitution
        ' -------------------------------------------------

        latex = latex & _
        "& \text{Substitute in equation (2)} \\[6pt]"

        latex = latex & _
        "& " & BuildLinearTermLatex(a2, var1) & _
        " + " & BuildLinearTermLatex(b2, var2) & _
        " = " & FormatFractionSurdToLatex(c2) & " \\[6pt]"

        latex = latex & _
        "& " & BuildLinearTermLatex(a2, var1) & _
        " + " & FormatFractionSurdToLatex(b2) & _
        "(" & FormatFractionSurdToLatex(c1) & _
        " - " & BuildLinearTermLatex(a1, var1) & _
        ") = " & FormatFractionSurdToLatex(c2) & " \\[10pt]"

    Else

        latex = latex & _
        "& \text{From equation (1)} \\[6pt]"

        latex = latex & _
        "& " & BuildLinearTermLatex(a1, var1) & _
        " + " & BuildLinearTermLatex(b1, var2) & _
        " = " & FormatFractionSurdToLatex(c1) & " \\[6pt]"

        latex = latex & _
        "& " & var1 & " = " & _
        FormatFractionSurdToLatex(c1) & " - " & _
        BuildLinearTermLatex(b1, var2) & " \\[10pt]"


        latex = latex & _
        "& \text{Substitute in equation (2)} \\[6pt]"

        latex = latex & _
        "& (" & FormatFractionSurdToLatex(c1) & _
        " - " & BuildLinearTermLatex(b1, var2) & _
        ") + " & BuildLinearTermLatex(b2, var2) & _
        " = " & FormatFractionSurdToLatex(c2) & " \\[10pt]"

    End If


    AppendSubstitutionSteps = latex

End Function

Public Function ExpandSubstitutionStep(a2 As FractionSurd, _
                                       b2 As FractionSurd, _
                                       a1 As FractionSurd, _
                                       c1 As FractionSurd, _
                                       c2 As FractionSurd, _
                                       var1 As String) As String

    Dim latex As String
    latex = ""

    Dim termConst As FractionSurd
    Dim termX As FractionSurd
    
    Dim negA1 As FractionSurd
    
    ' b2 * c1
    termConst = MultiplyFractionSurd(c1, b2)
    
    ' -a1
    negA1 = a1
    negA1.num.coeff = -negA1.num.coeff
    
    ' b2 * (-a1)
    termX = MultiplyFractionSurd(negA1, b2)
    
    
    latex = latex & _
    "& " & BuildLinearTermLatex(a2, var1) & _
    BuildSignedTermLatex(termConst, "") & _
    BuildSignedTermLatex(termX, var1) & _
    " = " & FormatFractionSurdToLatex(c2) & " \\[6pt]"
    
    
    ExpandSubstitutionStep = latex

End Function
Public Function AppendWithSign(baseStr As String, _
                                newTerm As String) As String
                                
    If baseStr = "" Then
        AppendWithSign = newTerm
    Else
        AppendWithSign = baseStr & " + " & newTerm
    End If
    
End Function
Public Function BuildEquationLatex(a As FractionSurd, _
                                   b As FractionSurd, _
                                   c As FractionSurd, _
                                   var1 As String, _
                                   var2 As String, _
                                   Optional showEqNumber As Boolean = False, _
                                   Optional eqLabel As String = "") As String

    Dim eqLine As String
    eqLine = ""

    ' -------------------------
    ' First term (ax)
    ' -------------------------

    If Not IsZeroFraction(a) Then
        eqLine = BuildLinearTermLatex(a, var1)
    End If


    ' -------------------------
    ' Second term (by)
    ' -------------------------

    If Not IsZeroFraction(b) Then

        Dim absB As FractionSurd
        absB = b
        absB.num.coeff = Abs(absB.num.coeff)

        If eqLine <> "" Then

            If b.num.coeff > 0 Then
                eqLine = eqLine & " + " & BuildLinearTermLatex(absB, var2)
            Else
                eqLine = eqLine & " - " & BuildLinearTermLatex(absB, var2)
            End If

        Else

            If b.num.coeff < 0 Then
                eqLine = "-" & BuildLinearTermLatex(absB, var2)
            Else
                eqLine = BuildLinearTermLatex(absB, var2)
            End If

        End If

    End If


    ' -------------------------
    ' RHS
    ' -------------------------

    eqLine = eqLine & " = " & FormatFractionSurdToLatex(c)


    ' -------------------------
    ' Equation numbering
    ' -------------------------

    If showEqNumber Then

        If eqLabel <> "" Then
            eqLine = eqLine & " \dots \text{(" & eqLabel & ")}"
        Else
            EqCounter = EqCounter + 1
            eqLine = eqLine & " \dots \text{(" & EqCounter & ")}"
        End If

    End If


    BuildEquationLatex = eqLine

End Function
Public Function BuildLinearTermLatex(f As FractionSurd, _
                                     varName As String) As String

    Dim c As Long
    c = f.num.coeff   ' denominator is 1 after Step-5
    
    If c = 0 Then
        BuildLinearTermLatex = ""
        Exit Function
    End If
    
    If c = 1 Then
        BuildLinearTermLatex = varName
    ElseIf c = -1 Then
        BuildLinearTermLatex = "-" & varName
    Else
        BuildLinearTermLatex = CStr(c) & varName
    End If

End Function
Public Function BuildSideLatex(frm As Object, _
                                eqNum As Integer, _
                                isLHS As Boolean) As String
    
    Dim output As String
    Dim eqPos As Integer
    
    eqPos = val(frm.Controls("txtETP" & eqNum).value)
    
    Dim posA As Integer
    Dim posB As Integer
    Dim posC As Integer
    
    posA = val(frm.Controls("txtPVTP" & eqNum).value)
    posB = val(frm.Controls("txtSVTP" & eqNum).value)
    posC = val(frm.Controls("txtCTP" & eqNum).value)
    
    ' Process A
    If (isLHS And posA < eqPos) Or _
       (Not isLHS And posA > eqPos) Then
       
        output = output & FormatFractionSurdToLatex( _
                    GetFractionSurdFromControls(frm, eqNum, "A")) _
                    & frm.txtPVar.value
    End If
    
    ' Process B
    If (isLHS And posB < eqPos) Or _
       (Not isLHS And posB > eqPos) Then
       
        output = AppendWithSign(output, _
                    FormatFractionSurdToLatex( _
                    GetFractionSurdFromControls(frm, eqNum, "B")) _
                    & frm.txtSVar.value)
    End If
    
    ' Process Constant
    If (isLHS And posC < eqPos) Or _
       (Not isLHS And posC > eqPos) Then
       
        output = AppendWithSign(output, _
                    FormatFractionSurdToLatex( _
                    GetFractionSurdFromControls(frm, eqNum, "C")))
    End If
    
    BuildSideLatex = output
    
End Function
Public Function BuildSubstitutionSteps(eq As StandardForm, _
                                       xVal As FractionSurd, _
                                       var1 As String, _
                                       var2 As String) As String

    Dim latex As String
    Dim axPart As FractionSurd
    Dim tempConst As FractionSurd

    latex = ""

    ' Substitute value
    latex = latex & _
        "& " & FormatFractionSurdToLatex(eq.aCoeff) & _
        "(" & FormatFractionSurdToLatex(xVal) & ")" & _
        " + " & BuildLinearTermLatex(eq.bCoeff, var2) & _
        " = " & FormatFractionSurdToLatex(eq.constCoeff) & " \\"

    ' Compute ax
    axPart = MultiplyTwoFractionSurds(eq.aCoeff, xVal)
    SimplifyFractionSurd axPart

    latex = latex & _
        "& " & FormatFractionSurdToLatex(axPart) & _
        " + " & BuildLinearTermLatex(eq.bCoeff, var2) & _
        " = " & FormatFractionSurdToLatex(eq.constCoeff) & " \\"

    ' Move constant
    latex = latex & _
        "& " & BuildLinearTermLatex(eq.bCoeff, var2) & _
        " = " & FormatFractionSurdToLatex(eq.constCoeff) & _
        " - " & FormatFractionSurdToLatex(axPart) & " \\"

    ' Compute RHS
    tempConst = SubtractFractionSurd(eq.constCoeff, axPart)
    SimplifyFractionSurd tempConst

    latex = latex & _
        "& " & BuildLinearTermLatex(eq.bCoeff, var2) & _
        " = " & FormatFractionSurdToLatex(tempConst) & " \\"

    BuildSubstitutionSteps = latex

End Function
Public Function BuildTermLatex(coeff As FractionSurd, _
                               varName As String, _
                               firstTerm As Boolean) As String

    Dim valueLatex As String
    Dim absLatex As String
    Dim finalSign As String
    
    valueLatex = FormatFractionSurdToLatex(coeff)
    
    If valueLatex = "" Or valueLatex = "0" Then
        BuildTermLatex = ""
        Exit Function
    End If
    
    ' Detect negative
    If left(Trim(valueLatex), 1) = "-" Then
        finalSign = "-"
        absLatex = Trim(Mid(Trim(valueLatex), 2))
    Else
        finalSign = "+"
        absLatex = valueLatex
    End If
    
    ' Remove 1 coefficient for variables
    If absLatex = "1" And varName <> "" Then absLatex = ""
    
    ' FIRST TERM: no leading +
    If firstTerm Then
        If finalSign = "-" Then
            BuildTermLatex = "-" & absLatex & varName
        Else
            BuildTermLatex = absLatex & varName
        End If
    Else
        BuildTermLatex = " " & finalSign & " " & absLatex & varName
    End If

End Function
Public Function BuildTextTerm(ByVal s As String, _
                              ByVal nc As Long, ByVal nr As Long, _
                              ByVal dc As Long, ByVal dr As Long, _
                              ByVal v As String, ByVal lead As Boolean) As String

    Dim num As String, den As String, combined As String
    Dim finalSign As String
    
    ' ---------------------------
    ' NORMALIZE NUMERIC SIGN
    ' ---------------------------
    If nc < 0 Then
        finalSign = "-"
        nc = Abs(nc)
    Else
        finalSign = "+"
    End If
    
    ' Apply toggle sign
    If s = "-" Then
        If finalSign = "+" Then
            finalSign = "-"
        Else
            finalSign = "+"
        End If
    End If
    
    ' ---------------------------
    ' BUILD NUMERATOR / DENOM
    ' ---------------------------
    num = IIf(nc = 0 Or nc = 1, "", CStr(nc)) & _
          IIf(nr = 0 Or nr = 1, "", "v" & CStr(nr))
    
    If num = "" Then num = "1"
    
    den = IIf(dc = 0 Or dc = 1, "", CStr(dc)) & _
          IIf(dr = 0 Or dr = 1, "", "v" & CStr(dr))
    
    If den = "" Or den = "1" Then
        combined = num
    Else
        combined = "(" & num & "/" & den & ")"
    End If
    
    ' Remove coefficient 1 before variable
    If combined = "1" And v <> "" Then combined = ""
    
    ' Hide leading +
    If finalSign = "+" And lead Then finalSign = ""
    
    BuildTextTerm = finalSign & combined & v

End Function
Public Function BuildUserInputLatex( _
        ft1 As FractionTerm, _
        ft2 As FractionTerm, _
        ft3 As FractionTerm, _
        pVar As String, _
        sVar As String) As String

    Dim arr(1 To 3) As FractionTerm
    Dim i As Integer
    
    Dim lhs As String
    Dim rhs As String
    
    arr(1) = ft1
    arr(2) = ft2
    arr(3) = ft3
    
    lhs = ""
    rhs = ""
    
    For i = 1 To 3
        
        If Not IsZeroFractionSurd(arr(i).coeff) Then
            
            Dim termLatex As String
            termLatex = FormatFractionSurdToLatex(arr(i).coeff)
            
            Select Case arr(i).variableID
                Case 1
                    termLatex = termLatex & pVar
                Case 2
                    termLatex = termLatex & sVar
                Case 0
                    ' constant term
            End Select
            
            ' Constant goes to RHS for Step 1 display
            If arr(i).variableID = 0 Then
                rhs = termLatex
            Else
                If lhs = "" Then
                    lhs = termLatex
                Else
                    If left(termLatex, 1) = "-" Then
                        lhs = lhs & " " & termLatex
                    Else
                        lhs = lhs & " + " & termLatex
                    End If
                End If
            End If
            
        End If
        
    Next i
    
    If lhs = "" Then lhs = "0"
    If rhs = "" Then rhs = "0"
    
    BuildUserInputLatex = lhs & " = " & rhs

End Function

Public Function FormatSurdToLatex(s As Surd) As String

    If s.coeff = 0 Then
        FormatSurdToLatex = "0"
        Exit Function
    End If
    
    If s.radicand = 1 Then
        FormatSurdToLatex = CStr(s.coeff)
    ElseIf Abs(s.coeff) = 1 Then
        If s.coeff < 0 Then
            FormatSurdToLatex = "-\sqrt{" & s.radicand & "}"
        Else
            FormatSurdToLatex = "\sqrt{" & s.radicand & "}"
        End If
    Else
        FormatSurdToLatex = s.coeff & "\sqrt{" & s.radicand & "}"
    End If

End Function
Public Function BuildVerticalAddLayout(sf1 As StandardForm, _
                                       sf2 As StandardForm, _
                                       var1 As String, _
                                       var2 As String) As String

    BuildVerticalAddLayout = _
        BuildVerticalLayoutCore(sf1, sf2, var1, var2, "+")

End Function
Public Function BuildVerticalSubtractLayout(sf1 As StandardForm, _
                                            sf2 As StandardForm, _
                                            var1 As String, _
                                            var2 As String) As String

    BuildVerticalSubtractLayout = _
        BuildVerticalLayoutCore(sf1, sf2, var1, var2, "-")

End Function
Public Function BuildVerticalLayoutCore(sfTop As StandardForm, _
                                        sfBottom As StandardForm, _
                                        var1 As String, _
                                        var2 As String, _
                                        opSymbol As String) As String

    Dim latex As String
    
    Dim newA As FractionSurd
    Dim newB As FractionSurd
    Dim newC As FractionSurd
    Dim undersetSignA As String
    Dim undersetSignB As String
    Dim undersetSignC As String
    ' --------------------------------
    ' Compute result equation
    ' --------------------------------
    
    If opSymbol = "+" Then
        
        newA = AddFractionSurd(sfTop.aCoeff, sfBottom.aCoeff)
        newB = AddFractionSurd(sfTop.bCoeff, sfBottom.bCoeff)
        newC = AddFractionSurd(sfTop.constCoeff, sfBottom.constCoeff)
        
    Else
        
        newA = SubtractFractionSurd(sfTop.aCoeff, sfBottom.aCoeff)
        newB = SubtractFractionSurd(sfTop.bCoeff, sfBottom.bCoeff)
        newC = SubtractFractionSurd(sfTop.constCoeff, sfBottom.constCoeff)
        
    End If
    If SignColumn(sfBottom.aCoeff) = "+" Then
    undersetSignA = "\underset{(-)}"
    Else
    undersetSignA = "\underset{(+)}"
    End If
    If SignColumn(sfBottom.bCoeff) = "+" Then
    undersetSignB = "\underset{(-)}"
    Else
    undersetSignB = "\underset{(+)}"
    End If
    If SignColumn(sfBottom.constCoeff) = "+" Then
    undersetSignC = "\underset{(-)}"
    Else
    undersetSignC = "\underset{(+)}"
    End If

    latex = "& \begin{array}{cccccccc}"


    ' --------------------------------
    ' TOP ROW
    ' --------------------------------

    latex = latex & _
        " &" & KSP & "{}" & _
        " &" & KSP & "{}" & _
        " &" & KSP & CleanArrayCell(AbsCoeff(sfTop.aCoeff, var1)) & _
        " &" & KSP & SignColumn(sfTop.bCoeff) & _
        " &" & KSP & CleanArrayCell(AbsCoeff(sfTop.bCoeff, var2)) & _
        " &" & KSP & "=" & _
        " &" & KSP & "{}" & _
        " &" & KSP & Replace(FormatFractionSurdToLatex(sfTop.constCoeff), "-", "") & " \\"


    ' --------------------------------
    ' SECOND ROW
    ' --------------------------------
    If opSymbol = "+" Then
        latex = latex & _
            " &" & KSP & "\boldsymbol{" & opSymbol & "}" & _
            " &" & KSP & "{}" & _
            " &" & KSP & CleanArrayCell(AbsCoeff(sfBottom.aCoeff, var1)) & _
            " &" & KSP & SignColumn(sfBottom.bCoeff) & _
            " &" & KSP & CleanArrayCell(AbsCoeff(sfBottom.bCoeff, var2)) & _
            " &" & KSP & "=" & _
            " &" & KSP & "{}" & _
            " &" & KSP & Replace(FormatFractionSurdToLatex(sfBottom.constCoeff), "-", "") & " \\"
    Else
            latex = latex & _
            " &" & KSP & "\boldsymbol{" & opSymbol & "}" & _
            " &" & KSP & undersetSignA & "{}" & _
            " &" & KSP & CleanArrayCell(AbsCoeff(sfBottom.aCoeff, var1)) & _
            " &" & KSP & undersetSignB & SignColumn(sfBottom.bCoeff) & _
            " &" & KSP & CleanArrayCell(AbsCoeff(sfBottom.bCoeff, var2)) & _
            " &" & KSP & "=" & _
            " &" & KSP & undersetSignC & "{}" & _
            " &" & KSP & Replace(FormatFractionSurdToLatex(sfBottom.constCoeff), "-", "") & " \\"
    End If
    ' --------------------------------
    ' RESULT ROW
    ' --------------------------------

    latex = latex & "\hline"
    
        Dim signB As String
        Dim termB As String
    
    If newB.num.coeff = 0 Then
        signB = "{}"
        termB = "{}"
    Else
        signB = SignColumn(newB)
        termB = CleanArrayCell(AbsCoeff(newB, var2))
    End If
    
    latex = latex & _
        " &" & KSP & "{}" & _
        " &" & KSP & FirstSign(newA) & _
        " &" & KSP & CleanArrayCell(AbsCoeff(newA, var1)) & _
        " &" & KSP & signB & _
        " &" & KSP & termB & _
        " &" & KSP & "=" & _
        " &" & KSP & FirstSign(newC) & _
        " &" & KSP & Replace(FormatFractionSurdToLatex(newC), "-", "")
    
        latex = latex & "\end{array} \\[10pt]"
    BuildVerticalLayoutCore = latex

End Function
Public Function ExpandGCDDivisionLine(sf As StandardForm, _
                                       gcdValue As Long, _
                                       pVar As String, _
                                       sVar As String) As String

    Dim result As String
    
    ' A term
    result = "\frac{" & _
             FormatFractionSurdToLatex(sf.aCoeff) & pVar & "}{" & gcdValue & "}"
             
    ' B term
    If sf.bCoeff.num.coeff < 0 Then
        result = result & " - "
        result = result & "\frac{" & _
                 Replace(FormatFractionSurdToLatex(sf.bCoeff), "-", "", 1, 1) & _
                 sVar & "}{" & gcdValue & "}"
    Else
        result = result & " + "
        result = result & "\frac{" & _
                 FormatFractionSurdToLatex(sf.bCoeff) & sVar & "}{" & gcdValue & "}"
    End If
    
    ' Constant
    result = result & " = \frac{" & _
             FormatFractionSurdToLatex(sf.constCoeff) & _
             "}{" & gcdValue & "}"
             
    ExpandGCDDivisionLine = result

End Function
Public Function ExpandLCMLine(sf As StandardForm, _
                              lcm As Surd, _
                              pVar As String, _
                              sVar As String) As String

    Dim lcmLatex As String
    lcmLatex = FormatSurdToLatex(lcm)
    
    Dim result As String
    
    '========================
    ' A TERM
    '========================
    
    Dim aSign As Long
    aSign = Sgn(sf.aCoeff.num.coeff)
    
    Dim aMag As FractionSurd
    aMag = sf.aCoeff
    aMag.num.coeff = Abs(aMag.num.coeff)
    
    If aSign < 0 Then
        result = "-"
    End If
    
    result = result & _
             FormatFractionSurdToLatex(aMag) & _
             pVar & " \times " & lcmLatex
             
    '========================
    ' B TERM
    '========================
    
    Dim bSign As Long
    bSign = Sgn(sf.bCoeff.num.coeff)
    
    Dim bMag As FractionSurd
    bMag = sf.bCoeff
    bMag.num.coeff = Abs(bMag.num.coeff)
    
    If bSign < 0 Then
        result = result & " - "
    Else
        result = result & " + "
    End If
    
    result = result & _
             FormatFractionSurdToLatex(bMag) & _
             sVar & " \times " & lcmLatex
    
    '========================
    ' CONSTANT
    '========================
    
    result = result & " = (" & _
             FormatFractionSurdToLatex(sf.constCoeff) & _
             ") \times " & lcmLatex

    ExpandLCMLine = result

End Function
Public Function ExtractInnerLtx(ByVal txt As String) As String
    ' Removes $$ if they exist so the string can sit inside \begin{aligned}
    ExtractInnerLtx = Replace(txt, "$$", "")
End Function
Public Function FlipSignString(ByVal ltx As String) As String
    Dim s As String: s = Trim(ltx)
    If left(s, 2) = "+ " Then
        FlipSignString = "- " & Mid(s, 3)
    ElseIf left(s, 1) = "+" Then
        FlipSignString = "-" & Mid(s, 2)
    ElseIf left(s, 2) = "- " Then
        FlipSignString = "+ " & Mid(s, 3)
    ElseIf left(s, 1) = "-" Then
        FlipSignString = "+" & Mid(s, 2)
    Else
        ' If no sign was present (leading term), it was implicitly positive
        FlipSignString = "- " & s
    End If
End Function
Public Function FlipStringSign(s As String) As String
    If left(s, 1) = "-" Then FlipStringSign = Mid(s, 2) Else FlipStringSign = "-" & s
End Function
Public Function FormatDeterminant2x2(a As String, _
                                     b As String, _
                                     c As String, _
                                     D As String) As String

    FormatDeterminant2x2 = "\begin{vmatrix} " & _
                           a & " & " & b & " \\ " & _
                           c & " & " & D & _
                           " \end{vmatrix}"

End Function
Public Function FormatSurd(m As Long, n As Long) As String
    If n = 1 Then
        FormatSurd = CStr(m)
    ElseIf m = 1 Then
        FormatSurd = "\sqrt{" & n & "}"
    Else
        FormatSurd = m & "\sqrt{" & n & "}"
    End If
End Function
Public Function FormatTerm_Symbolic(coeffStr As String, varChar As String) As String
    If coeffStr = "0" Then FormatTerm_Symbolic = "": Exit Function
    
    Dim s As String, val As String
    If left(coeffStr, 1) = "-" Then
        s = "- "
        val = Mid(coeffStr, 2)
    Else
        s = "+ "
        val = coeffStr
    End If
    
    ' Handle Fraction format
    If InStr(val, "/") > 0 Then
        Dim parts() As String
        parts = Split(val, "/")
        val = "\frac{" & parts(0) & "}{" & parts(1) & "}"
    End If
    
    FormatTerm_Symbolic = s & val & varChar
End Function
Public Function GetEqSideLatex(frm As Object, eqNum As Integer, side As String) As String
    Dim ltx As String: ltx = ""
    Dim isFirstOnSide As Boolean: isFirstOnSide = True
    Dim termLatex As String
    
    If side = "LHS" Then
        ' --- FORCE Ax + By FORM ---
        ' 1. Process Primary Variable (A)
        termLatex = GetTermLatex(frm, eqNum, "A", True)
        If termLatex <> "" Then
            ltx = ltx & CleanLeadingPlus(termLatex, isFirstOnSide)
            isFirstOnSide = False
        End If
        
        ' 2. Process Secondary Variable (B)
        termLatex = GetTermLatex(frm, eqNum, "B", True)
        If termLatex <> "" Then
            ltx = ltx & CleanLeadingPlus(termLatex, isFirstOnSide)
            isFirstOnSide = False
        End If
        
    Else
        ' --- FORCE CONSTANT (C) ON RHS ---
        termLatex = GetTermLatex(frm, eqNum, "C", True)
        If termLatex <> "" Then
            ltx = ltx & CleanLeadingPlus(termLatex, isFirstOnSide)
        Else
            ltx = "0" ' Show 0 if RHS is empty
        End If
    End If
    
    GetEqSideLatex = ltx
End Function


Private Function GetStep2Latex(frm As Object, eqNum As Integer) As String
    ' This shows the "Raw" rearrangement before any signs are flipped
    Dim ltx As String
    ltx = GetEqSideLatex(frm, eqNum, "LHS") & " &= " & GetEqSideLatex(frm, eqNum, "RHS") & _
          " && \text{...Arranging terms (Step 2)} \\[12pt]"
    GetStep2Latex = ltx
End Function
'Public Function GetStep2Latex_FromState(cA As Double, cB As Double, cC As Double) As String
'    Dim lhs As String
 '   lhs = CleanLeadingPlus(FormatTerm(cA, "x") & FormatTerm(cB, "y"))
'    GetStep2Latex_FromState = lhs & " &= " & cC & " && \text{...Arranging terms (Step 2)} \\[12pt]"
'End Function
Public Function GetStep2Latex_FromState(cA As String, cB As String, cC As String, Optional factor As String = "") As String
    Dim lhs As String
    lhs = CleanLeadingPlus(FormatTerm_Symbolic(cA, "x") & FormatTerm_Symbolic(cB, "y"), True)
    GetStep2Latex_FromState = lhs & " &= " & cC & " && \text{...Arranging terms} \\[12pt]"
End Function

Public Function GetStep3Latex(frm As Object, eqNum As Integer) As String
    Dim ltx As String
    Dim termA As String, termB As String, termC As String
    Dim flippedA As String, flippedB As String, flippedC As String
    
    ' 1. Get the rearranged terms (as they appear at the end of Step 2)
    ' We use CheckSide = True to ensure we see the signs after moving to LHS/RHS
    termA = GetTermLatex(frm, eqNum, "A", True)
    termB = GetTermLatex(frm, eqNum, "B", True)
    termC = GetTermLatex(frm, eqNum, "C", True)
    
    ' 2. Clean the leading plus for the display of the current state
    Dim currentLHS As String
    currentLHS = CleanLeadingPlus(termA, True) & termB
    
    ' 3. Generate the "Flipped" versions for the result line
    ' We manually flip the signs to show the result of multiplying by -1
    flippedA = CleanLeadingPlus(FlipSignString(termA), True)
    flippedB = FlipSignString(termB)
    flippedC = FlipSignString(termC)
    
    ' 4. Build the LaTeX
    ' Line 1: Show the multiplication by (-1)
    ltx = "-1 \left( " & currentLHS & " \right) &= -1 \left( " & termC & " \right) " & _
          " && \text{...Multiplying by -1 (Step 3)} \\[8pt]"
    
    ' Line 2: Show the result with signs normalized
    ltx = ltx & flippedA & " " & flippedB & " &= " & flippedC & " && \text{...Lead sign normalized} \\[12pt]"
    
    GetStep3Latex = ltx
End Function
Public Function GetGivenEquationsBlock(frm As Object) As String
    Dim eq1 As String: eq1 = Replace(frm.LastLtx1, "$$", "")
    Dim eq2 As String: eq2 = Replace(frm.LastLtx2, "$$", "")
    
    GetGivenEquationsBlock = "\begin{array}{l} " & eq1 & " \\[10pt] " & eq2 & _
                             " \end{array} \left. \text{\rule{0pt}{3em}} \right\} \text{... Given equations}"
End Function
Public Function GetRearrangementLatex(frm As Object, eqNum As Integer) As String
    Dim ltx As String
    Dim eqPos As Integer: eqPos = val(frm.Controls("txtETP" & eqNum).value)
    
    ' 1. Show the Arranged form (Variables Left, Constant Right)
    ' We build the strings for Term A, B, and C based on the logic we refined
    Dim strA As String, strB As String, strC As String
    strA = GetTermLatex(frm, eqNum, "A", True) ' True means check if flip needed
    strB = GetTermLatex(frm, eqNum, "B", True)
    strC = GetTermLatex(frm, eqNum, "C", True)
    
    ltx = strA & " " & strB & " &= " & strC & " && \text{...Arranging terms (Step 2)} \\[10pt]"
    
    ' 2. Step 3: Check if Lead Term is Negative
    ' (This logic assumes GetTermLatex handles the sign flip internally)
    If left(Trim(strA), 1) = "-" Then
        strA = Mid(strA, 2) ' Remove leading minus
        strB = FlipSignString(strB)
        strC = FlipSignString(strC)
        ltx = ltx & strA & " " & strB & " &= " & strC & " && \text{...Multiplying by -1 (Step 3)} \\[10pt]"
    End If
    
    GetRearrangementLatex = ltx
End Function
Public Function GetLaTeXString(s As String, nc As String, nr As String, _
                               dc As String, dr As String, _
                               ByVal v As String, lead As Boolean) As String

    Dim n As String, D As String, t As String
    Dim finalSign As String
    Dim numVal As Long
    
    ' ----------------------------------
    ' NORMALIZE SIGN
    ' ----------------------------------
    numVal = CLng(val(nc))
    
    If numVal < 0 Then
        finalSign = "-"
        nc = CStr(Abs(numVal))
    Else
        finalSign = "+"
    End If
    
    ' Apply toggle sign
    If s = "-" Then
        If finalSign = "+" Then
            finalSign = "-"
        Else
            finalSign = "+"
        End If
    End If
    
    ' ----------------------------------
    ' BUILD NUMERATOR / DENOMINATOR
    ' ----------------------------------
    n = BuildSurd(nc, nr)
    D = BuildSurd(dc, dr)
    
    If D = "1" Then
        t = n
    Else
        t = "\frac{" & n & "}{" & D & "}"
    End If
    
    ' Remove coefficient 1 before variable
    If t = "1" And v <> "" Then t = ""
    
    ' Hide leading +
    If finalSign = "+" And lead Then finalSign = ""
    
    GetLaTeXString = finalSign & t & v

End Function



Public Sub RenderStepByStep(ltxBody As String, frm As Object)
    Dim html As String
    ' Note the use of "align" environment for multi-line steps
    html = "<html><head><script src='js/mathjax/MathJax.js?config=TeX-AMS_HTML'></script>" & _
           "<style>body { font-size: 1.2em; background-color: #FFFFFF; padding: 20px; font-family: sans-serif; }" & _
           ".step { color: #2c3e50; font-weight: bold; margin-bottom: 5px; }</style>" & _
           "</head><body>" & ltxBody & "</body></html>"
    
    On Error Resume Next
    With frm.webSolution.Document
        .Open
        .Write html
        .Close
    End With
End Sub
Public Function GetSafeCaption(frm As Object, ctrlName As String) As String

    On Error GoTo SafeExit
    
    If ControlExists(frm, ctrlName) Then
        GetSafeCaption = frm.Controls(ctrlName).Caption
    Else
        GetSafeCaption = "+"
    End If
    
    Exit Function

SafeExit:
    GetSafeCaption = "+"

End Function
Public Function GetStep1Latex(frm As Object, eqNum As Integer) As String

    ' Simply use the already-correct preview LaTeX
    If eqNum = 1 Then
        GetStep1Latex = frm.LastLtx1
    Else
        GetStep1Latex = frm.LastLtx2
    End If

End Function
Public Function GetStep3Latex_FromSurds(sA As Surd, sB As Surd, sC As Surd) As String
    ' Formats the equation for the display
    Dim termA As String: termA = FormatSurdToLatex(sA)
    Dim termB As String: termB = FormatSurdToLatex(sB)
    Dim termC As String: termC = FormatSurdToLatex(sC)
    
    ' Assuming Step 3 is "Normalization", you typically show the result
    GetStep3Latex_FromSurds = "\text{Normalized: } & " & termA & "x + " & termB & "y + " & termC & " = 0 \\"
End Function

Public Function GetStep4Latex_FromSurds(sA As Surd, sB As Surd, sC As Surd, lcm As Surd) As String
    ' Multiply terms by the LCM using your existing multiplication logic
    ' We keep the terms as new variables to display the "Result" of the multiplication
    Dim termA As Surd: termA = MultiplySurds(sA, lcm)
    Dim termB As Surd: termB = MultiplySurds(sB, lcm)
    Dim termC As Surd: termC = MultiplySurds(sC, lcm)
    
    ' Format using your existing LaTeX helper
    Dim strA As String: strA = FormatSurdToLatex(termA)
    Dim strB As String: strB = FormatSurdToLatex(termB)
    Dim strC As String: strC = FormatSurdToLatex(termC)
    
    ' Create the display string
    GetStep4Latex_FromSurds = "\text{Multiply by LCM: } & " & strA & "x + " & strB & "y + " & strC & " = 0 \\"
End Function
Public Function GetStep5Latex_FromSurds(sA As Surd, sB As Surd, sC As Surd, gSurd As Surd) As String
    ' This converts your reduced Surd variables into the final LaTeX string
    Dim strA As String: strA = FormatSurdToLatex(sA)
    Dim strB As String: strB = FormatSurdToLatex(sB)
    Dim strC As String: strC = FormatSurdToLatex(sC)
    
    ' Combine them into the final equation format
    ' Note: We assume the user wants to see the final simplified form
    GetStep5Latex_FromSurds = "\text{Simplified: } & " & strA & "x + " & strB & "y + " & strC & " = 0 \\"
End Function
Public Function GetSymbolicEq(frm As Object, eqNum As Integer) As String

    On Error GoTo SafeFail

    Dim parts(1 To 4) As String
    Dim eqPos As Integer
    eqPos = val(frm.Controls("txtETP" & eqNum).value)

    ' ?? Separate prefixes properly
    Dim coefPrefix As Variant
    coefPrefix = Array("A", "B", "C")   ' For txtANC, txtBNC, txtCNC

    Dim togglePrefix As Variant
    togglePrefix = Array("F", "S", "C") ' For tglSignFT, tglSignST, tglSignCT

    Dim posPrefix As Variant
    posPrefix = Array("PVTP", "SVTP", "CTP")

    Dim varNames As Variant
    varNames = Array(frm.txtPVar.value, frm.txtSVar.value, "")

    Dim i As Integer
    Dim pos As Integer
    Dim Sign As String
    Dim m As Long, n As Long, D As Long, r As Long
    Dim vName As String
    Dim ncName As String

    For i = 0 To 2

        ' -------- Coefficient Controls --------
        ncName = "txt" & coefPrefix(i) & "NC" & eqNum
        If Not ControlExists(frm, ncName) Then GoTo NextTerm

        m = val(frm.Controls(ncName).value)
        n = val(frm.Controls("txt" & coefPrefix(i) & "NR" & eqNum).value)
        D = val(frm.Controls("txt" & coefPrefix(i) & "DC" & eqNum).value)
        r = val(frm.Controls("txt" & coefPrefix(i) & "DR" & eqNum).value)

        ' -------- Toggle Controls --------
        If ControlExists(frm, "tglSign" & togglePrefix(i) & "T" & eqNum) Then
            Sign = frm.Controls("tglSign" & togglePrefix(i) & "T" & eqNum).Caption
        Else
            Sign = "+"
        End If

        ' -------- Position --------
        pos = val(frm.Controls("txt" & posPrefix(i) & eqNum).value)
        If pos < 1 Or pos > 4 Then GoTo NextTerm

        vName = varNames(i)

        parts(pos) = BuildTextTerm(Sign, m, n, D, r, vName, (pos = 1 Or pos = eqPos + 1))

NextTerm:
    Next i

    Dim leftSide As String, rightSide As String
    Dim j As Integer

    For j = 1 To 4
        If parts(j) <> "" Then
            If j < eqPos Then
                leftSide = leftSide & " " & parts(j)
            ElseIf j > eqPos Then
                rightSide = rightSide & " " & parts(j)
            End If
        End If
    Next j

    If eqPos = 1 Then
        GetSymbolicEq = "0 =" & rightSide
    ElseIf eqPos = 4 Then
        GetSymbolicEq = leftSide & " = 0"
    Else
        GetSymbolicEq = Trim(leftSide) & " =" & rightSide
    End If

    Exit Function

SafeFail:
    Debug.Print "Symbolic generation failed:", Err.Description
    GetSymbolicEq = ""
    Err.Clear

End Function
Public Function GetTermLatex(frm As Object, eqNum As Integer, prefix As String, CheckSide As Boolean) As String
    Dim nc As Long, nr As Long, dc As Long, dr As Long
    Dim pos As Integer, eqPos As Integer
    Dim finalSign As String
    Dim strNum As String, strDen As String
    Dim mathContent As String
    Dim varName As String
    
    ' 1. Load Numeric Values
    nc = val(frm.Controls("txt" & prefix & "NC" & eqNum).value): If nc = 0 Then nc = 1
    nr = val(frm.Controls("txt" & prefix & "NR" & eqNum).value): If nr = 0 Then nr = 1
    dc = val(frm.Controls("txt" & prefix & "DC" & eqNum).value): If dc = 0 Then dc = 1
    dr = val(frm.Controls("txt" & prefix & "DR" & eqNum).value): If dr = 0 Then dr = 1
    
    ' 2. Load Position Metadata
    eqPos = val(frm.Controls("txtETP" & eqNum).value)
    If prefix = "A" Then pos = val(frm.Controls("txtPVTP" & eqNum).value)
    If prefix = "B" Then pos = val(frm.Controls("txtSVTP" & eqNum).value)
    If prefix = "C" Then pos = val(frm.Controls("txtCTP" & eqNum).value)
    
    ' Get the user-selected sign from the toggle button
    finalSign = frm.Controls("tglSign" & IIf(prefix = "A", "FT", IIf(prefix = "B", "ST", "CT")) & eqNum).Caption
    
    ' 3. ENFORCED STANDARDIZATION LOGIC (Step 2)
    ' This logic moves variables to LHS and constants to RHS
    If CheckSide Then
        If prefix = "A" Or prefix = "B" Then
            ' Variables (A, B) belong on LHS (pos <= eqPos)
            ' If found on RHS (pos > eqPos), flip the sign
            If pos > eqPos Then finalSign = IIf(finalSign = "+", "-", "+")
        ElseIf prefix = "C" Then
            ' Constant (C) belongs on RHS (pos > eqPos)
            ' If found on LHS (pos <= eqPos), flip the sign
            If pos <= eqPos Then finalSign = IIf(finalSign = "+", "-", "+")
        End If
    End If

    ' 4. Build LaTeX Math Content
    strNum = IIf(nr = 1, CStr(nc), IIf(nc = 1, "", nc) & "\sqrt{" & nr & "}")
    strDen = IIf(dc = 1 And dr = 1, "", IIf(dr = 1, CStr(dc), IIf(dc = 1, "", dc) & "\sqrt{" & dr & "}"))
    mathContent = IIf(strDen = "", strNum, "\frac{" & strNum & "}{" & strDen & "}")
    
    ' 5. Variable Attachment
    varName = IIf(prefix = "A", frm.txtPVar.value, IIf(prefix = "B", frm.txtSVar.value, ""))
    If mathContent = "1" And varName <> "" Then mathContent = ""
    
    ' 6. Final Combined String
    ' Note: We do NOT strip the leading plus here anymore;
    ' GetEqSideLatex handles that via CleanLeadingPlus for better control.
    GetTermLatex = IIf(finalSign = "+", "+ ", "- ") & mathContent & varName
End Function
' --- IMPROVED RENDERER WITH MOTW ---
Public Sub RenderMathJax(ltx As String, frm As Object, n As Variant)

    Dim targetBrowser As Object
    
    On Error Resume Next
    
    ' ------------------------------------
    ' Decide Target Safely
    ' ------------------------------------
    
    If n = "Question" Then
        
        Set targetBrowser = frm.webQuestion
        
    ElseIf n = "Solution" Then
        
        Set targetBrowser = frm.webSolution
        
    ElseIf n = 1 Then
        
        Set targetBrowser = frm.webPreview1
        
    ElseIf n = 2 Then
        
        Set targetBrowser = frm.webPreview2
        
    Else
        Exit Sub   ' Prevent crash
    End If
    
    On Error GoTo 0
    
    ' ------------------------------------
    ' Safety Check
    ' ------------------------------------
    
    If targetBrowser Is Nothing Then Exit Sub
    If targetBrowser.Document Is Nothing Then Exit Sub
    
    ' ------------------------------------
    ' Update MathJax
    ' ------------------------------------
    
    targetBrowser.Document.body.innerHTML = "$$ " & ltx & " $$"
    
    targetBrowser.Document.parentWindow.execScript _
    "if(window.MathJax){MathJax.Hub.Queue(['Typeset',MathJax.Hub,document.body]);}"

End Sub
Public Function LtxEquationLine(eq As String) As String
    If InStr(eq, "\frac") > 0 Then
        LtxEquationLine = "& \quad " & eq & " \\[5pt]" & vbCrLf
    Else
        LtxEquationLine = "& \quad " & eq & " \\" & vbCrLf
    End If
End Function
Public Function LtxTextLine(txt As String) As String
    LtxTextLine = "& \text{" & txt & "} \\" & vbCrLf
End Function
Public Function PrintEquation(a As FractionSurd, _
                              b As FractionSurd, _
                              c As FractionSurd, _
                              var1 As String, _
                              var2 As String, _
                              Optional showNumber As Boolean = False, _
                              Optional eqLabel As String = "") As String

    Dim latex As String
    Dim expr As String

    ' Build LHS expression
    expr = BuildLinearExpression(a, b, var1, var2)

    latex = "& " & expr & " = " & FormatFractionSurdToLatex(c)

    ' Optional numbering
    If showNumber Then
        If eqLabel <> "" Then
            latex = latex & " \dots \text{(" & eqLabel & ")}"
        End If
    End If

    latex = latex & " \\"

    PrintEquation = latex

End Function
Public Function PrintFinalSolution(var1 As String, _
                                   var2 As String, _
                                   xVal As FractionSurd, _
                                   yVal As FractionSurd) As String

    Dim latex As String
    
    latex = "& \therefore (" & var1 & "," & var2 & ") = (" & _
            FormatFractionSurdToLatex(xVal) & "," & _
            FormatFractionSurdToLatex(yVal) & ")"
    
    PrintFinalSolution = latex

End Function
Public Function PrintVariableResult(varName As String, _
                                    value As FractionSurd) As String

    Dim latex As String
    
    latex = "& " & varName & " = " & _
            FormatFractionSurdToLatex(value) & " \\[10pt]"
    
    PrintVariableResult = latex

End Function
Public Sub RenderHTML(wb As Object, htmlContent As String)

    wb.Navigate "about:blank"
    Do While wb.ReadyState <> 4
        DoEvents
    Loop

    wb.Document.Open
    wb.Document.Write htmlContent
    wb.Document.Close

End Sub
Public Function StandardFormToLatex(sf As StandardForm, _
                                   pVar As String, _
                                   sVar As String) As String

    Dim result As String
    Dim firstTerm As Boolean
    firstTerm = True
    
    Dim termStr As String
    
    ' -------------------------
    ' A TERM (Primary Variable)
    ' -------------------------
    termStr = BuildTermLatex(sf.aCoeff, pVar, firstTerm)
    
    If termStr <> "" Then
        result = result & termStr
        firstTerm = False
    End If
    
    ' -------------------------
    ' B TERM (Secondary Variable)
    ' -------------------------
    termStr = BuildTermLatex(sf.bCoeff, sVar, firstTerm)
    
    If termStr <> "" Then
        result = result & termStr
        firstTerm = False
    End If
    
    ' -------------------------
    ' EQUALS + CONSTANT
    ' -------------------------
    result = result & " = " & FormatFractionSurdToLatex(sf.constCoeff)
    
    StandardFormToLatex = result

End Function
Public Function WrapStep(title As String, content As String) As String

    WrapStep = _
        "\textit{" & title & "}\\ " & vbCrLf & _
        content & "\\[8pt]" & vbCrLf

End Function
Public Function WrapIfNegative(valueLatex As String) As String
    
    If left(Trim(valueLatex), 1) = "-" Then
        WrapIfNegative = "(" & valueLatex & ")"
    Else
        WrapIfNegative = valueLatex
    End If

End Function
Public Function GetLCMMultiplierLtx(sf As StandardForm) As String

    Dim lcmSurd As Surd
    
    lcmSurd = GetLCMMultiplier(sf)
    
    GetLCMMultiplierLtx = FormatSurdToLatex(lcmSurd)

End Function
Public Function AbsCoeff(f As FractionSurd, varName As String) As String


    Dim c As Long
    c = Abs(f.num.coeff)

    If c = 0 Then
        AbsCoeff = " &  & "
        Exit Function
    End If

    If c = 1 Then
        AbsCoeff = " & " & varName & " & "
    Else
        AbsCoeff = " & " & c & varName & " & "
    End If
End Function
Public Function SignColumn(f As FractionSurd) As String

    If f.num.coeff < 0 Then
        SignColumn = "-"
    Else
        SignColumn = "+"
    End If

End Function
Public Function CleanArrayCell(rawPart As String) As String
    
    Dim temp As String
    
    temp = Trim(rawPart)
    
    ' Remove leading &
    If left(temp, 1) = "&" Then
        temp = Trim(Mid(temp, 2))
    End If
    
    ' Remove trailing &
    If right(temp, 1) = "&" Then
        temp = Trim(left(temp, Len(temp) - 1))
    End If
    
    ' If empty, return safe placeholder
    If temp = "" Then
        CleanArrayCell = "{}"
    Else
        CleanArrayCell = temp
    End If
    
End Function
Public Function FirstSign(f As FractionSurd) As String

    If f.num.coeff < 0 Then
        FirstSign = "-"
    Else
        FirstSign = "{}"
    End If

End Function


