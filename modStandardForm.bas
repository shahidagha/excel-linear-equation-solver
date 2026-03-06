Attribute VB_Name = "modStandardForm"
Public Function NormalizeLeadingSign(sf As StandardForm) As StandardForm

    ' If aCoeff is negative ? multiply whole equation by -1
    If sf.aCoeff.num.coeff < 0 Then
        
        sf.aCoeff = NegateFractionSurd(sf.aCoeff)
        sf.bCoeff = NegateFractionSurd(sf.bCoeff)
        sf.constCoeff = NegateFractionSurd(sf.constCoeff)
        
    End If
    
    NormalizeLeadingSign = sf

End Function
Public Function IsNegativeLeading(sf As StandardForm) As Boolean

    If sf.aCoeff.num.coeff < 0 Then
        IsNegativeLeading = True
        Exit Function
    End If
    
    ' Optional: handle case where aCoeff = 0
    If sf.aCoeff.num.coeff = 0 Then
        If sf.bCoeff.num.coeff < 0 Then
            IsNegativeLeading = True
            Exit Function
        End If
    End If
    
    IsNegativeLeading = False

End Function
Public Function IsIdentity(sf As StandardForm) As Boolean

    If IsZeroFractionSurd(sf.aCoeff) _
       And IsZeroFractionSurd(sf.bCoeff) _
       And IsZeroFractionSurd(sf.constCoeff) Then
       
        IsIdentity = True
    Else
        IsIdentity = False
    End If

End Function
Public Function IsContradiction(sf As StandardForm) As Boolean

    If IsZeroFractionSurd(sf.aCoeff) _
       And IsZeroFractionSurd(sf.bCoeff) _
       And Not IsZeroFractionSurd(sf.constCoeff) Then
       
        IsContradiction = True
    Else
        IsContradiction = False
    End If

End Function
Public Function DeterminantIsZero(e1 As StandardForm, _
                                  e2 As StandardForm) As Boolean
    
    Dim term1 As Surd
    Dim term2 As Surd
    Dim det As Surd
    
    ' a1*b2
    term1 = MultiplySurds(e1.aCoeff.num, e2.bCoeff.num)
    
    ' a2*b1
    term2 = MultiplySurds(e2.aCoeff.num, e1.bCoeff.num)
    
    ' determinant
    det = SubtractSurds(term1, term2)
    
    If det.coeff = 0 Then
        DeterminantIsZero = True
    Else
        DeterminantIsZero = False
    End If
    
End Function
Public Function ApplyLCM(sf As StandardForm, mult As Surd) As StandardForm

    sf.aCoeff = MultiplyFractionBySurd(sf.aCoeff, mult)
    sf.bCoeff = MultiplyFractionBySurd(sf.bCoeff, mult)
    sf.constCoeff = MultiplyFractionBySurd(sf.constCoeff, mult)

    SimplifyFractionSurd sf.aCoeff
    SimplifyFractionSurd sf.bCoeff
    SimplifyFractionSurd sf.constCoeff
    
    ApplyLCM = sf

End Function

Public Function GetGCDFactor(sf As StandardForm) As Surd

    Dim s1 As Long, s2 As Long, s3 As Long
    Dim g As Long
    
    s1 = GetSquaredSurdValue(sf.aCoeff.num)
    s2 = GetSquaredSurdValue(sf.bCoeff.num)
    s3 = GetSquaredSurdValue(sf.constCoeff.num)
    
    ' If any term is zero, ignore it in GCD
    If s1 = 0 Then s1 = s2
    If s2 = 0 Then s2 = s1
    If s3 = 0 Then s3 = s1
    
    g = GCD(s1, GCD(s2, s3))
    
    If g <= 1 Then
        GetGCDFactor.coeff = 1
        GetGCDFactor.radicand = 1
    Else
        GetGCDFactor = SquareRootToSurd(g)
    End If

End Function

Public Function ApplyGCDReduction(sf As StandardForm, g As Surd) As StandardForm

    sf.aCoeff.num = DivideSurdByFactor(sf.aCoeff.num, g)
    sf.bCoeff.num = DivideSurdByFactor(sf.bCoeff.num, g)
    sf.constCoeff.num = DivideSurdByFactor(sf.constCoeff.num, g)
    
    ApplyGCDReduction = sf

End Function
Public Function GetLCMMultiplier(sf As StandardForm) As Surd

    Dim s1 As Long, s2 As Long, s3 As Long
    Dim lcmVal As Long
    
    ' Get squared denominators
    s1 = GetSquaredDenominator(sf.aCoeff)
    s2 = GetSquaredDenominator(sf.bCoeff)
    s3 = GetSquaredDenominator(sf.constCoeff)
    
    ' Compute LCM in squared domain
    lcmVal = lcm(s1, lcm(s2, s3))
    
    ' Convert back from squared to Surd
    GetLCMMultiplier = SquareRootToSurd(lcmVal)

End Function
Public Function ReduceFinalEquation(ByRef s1 As Surd, _
                                    ByRef s2 As Surd, _
                                    ByRef s3 As Surd) As Boolean
                                    
    Dim finalGCD As Long
    Dim r1 As Surd, r2 As Surd, r3 As Surd
    
    finalGCD = GetFinalGCD(s1, s2, s3)
    
    If finalGCD <= 1 Then
        ReduceFinalEquation = True
        Exit Function
    End If
    
    If Not ReduceSurdByGCD(s1, finalGCD, r1) Then Exit Function
    If Not ReduceSurdByGCD(s2, finalGCD, r2) Then Exit Function
    If Not ReduceSurdByGCD(s3, finalGCD, r3) Then Exit Function
    
    s1 = r1
    s2 = r2
    s3 = r3
    
    ReduceFinalEquation = True
    
End Function
Public Function ReduceSurdByGCD(s As Surd, _
                                finalGCD As Long, _
                                ByRef result As Surd) As Boolean
                                
    Dim squaredVal As Long
    Dim tempResult As Long
    
    ' (mvn)^2 = m^2 * n
    squaredVal = (s.coeff * s.coeff) * s.radicand
    
    ' Safe integer division
    If Not SafeDivide(squaredVal, finalGCD, tempResult) Then
        ReduceSurdByGCD = False
        Exit Function
    End If
    
    ' Convert back from squared domain
    result = SquareRootToSurd(tempResult)
    
    ' Preserve original sign
    If s.coeff < 0 Then
        result.coeff = -Abs(result.coeff)
    End If
    
    ReduceSurdByGCD = True
    
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

