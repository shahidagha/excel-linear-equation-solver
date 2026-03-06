Attribute VB_Name = "modCramer"
Public Function ComputeDeterminant(sf1 As StandardForm, _
                                   sf2 As StandardForm) As FractionSurd

    Dim part1 As FractionSurd
    Dim part2 As FractionSurd
    Dim result As FractionSurd
    
    ' a1*b2
    part1 = MultiplyTwoFractionSurds(sf1.aCoeff, sf2.bCoeff)
    
    ' a2*b1
    part2 = MultiplyTwoFractionSurds(sf2.aCoeff, sf1.bCoeff)
    
    ' D = part1 - part2
    result = SubtractFractionSurd(part1, part2)
    
    ComputeDeterminant = result

End Function

Public Function ComputeDeterminantDx(sf1 As StandardForm, _
                                     sf2 As StandardForm) As FractionSurd

    Dim part1 As FractionSurd
    Dim part2 As FractionSurd
    Dim result As FractionSurd
    
    ' c1*b2
    part1 = MultiplyTwoFractionSurds(sf1.constCoeff, sf2.bCoeff)
    
    ' c2*b1
    part2 = MultiplyTwoFractionSurds(sf2.constCoeff, sf1.bCoeff)
    
    result = SubtractFractionSurd(part1, part2)
    
    ComputeDeterminantDx = result

End Function

Public Function ComputeDeterminantDy(sf1 As StandardForm, _
                                     sf2 As StandardForm) As FractionSurd

    Dim part1 As FractionSurd
    Dim part2 As FractionSurd
    Dim result As FractionSurd
    
    ' a1*c2
    part1 = MultiplyTwoFractionSurds(sf1.aCoeff, sf2.constCoeff)
    
    ' a2*c1
    part2 = MultiplyTwoFractionSurds(sf2.aCoeff, sf1.constCoeff)
    
    result = SubtractFractionSurd(part1, part2)
    
    ComputeDeterminantDy = result

End Function




Public Function ProcessCramerMethod(frm As Object) As String

    Dim latex As String
    Dim sf1 As StandardForm
    Dim sf2 As StandardForm
    
    latex = GenerateInitialStandardSteps(frm, sf1, sf2)
    
    latex = latex & _
        AppendCramerStep1(sf1, sf2, _
                          frm.txtPVar.value, _
                          frm.txtSVar.value)
    
    latex = latex & AppendDeterminantD(sf1, sf2)
    latex = latex & AppendDeterminantDx(sf1, sf2, frm.txtPVar.value)
    latex = latex & AppendDeterminantDy(sf1, sf2, frm.txtSVar.value)
    
    Dim D As FractionSurd
    Dim Dx As FractionSurd
    Dim Dy As FractionSurd
    
    D = ComputeDeterminant(sf1, sf2)
    Dx = ComputeDeterminantDx(sf1, sf2)
    Dy = ComputeDeterminantDy(sf1, sf2)
    
    latex = latex & _
        AppendFinalCramerStep(D, Dx, Dy, _
                              frm.txtPVar.value, _
                              frm.txtSVar.value)
    
    latex = latex & "\end{aligned}"
    
    ProcessCramerMethod = latex

End Function

