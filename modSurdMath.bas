Attribute VB_Name = "modSurdMath"
Public Function SquareRootToSurd(val As Long) As Surd

    Dim s As Surd
    Dim k As Long
    
    s.coeff = 1
    s.radicand = val
    
    k = 2
    
    ' Extract perfect square factors
    Do While k * k <= s.radicand
        
        If s.radicand Mod (k * k) = 0 Then
            
            s.coeff = s.coeff * k
            s.radicand = s.radicand / (k * k)
            k = 2
            
        Else
            k = k + 1
        End If
        
    Loop
    
    SquareRootToSurd = s

End Function
Public Sub SimplifySurd(ByRef s As Surd)

    Dim k As Long
    k = 2

    ' Extract perfect square factors from radicand
    Do While k * k <= s.radicand
        
        If s.radicand Mod (k * k) = 0 Then
            
            ' Pull square factor outside
            s.radicand = s.radicand / (k * k)
            s.coeff = s.coeff * k
            
            ' Restart factor check
            k = 2
            
        Else
            k = k + 1
        End If
        
    Loop

End Sub
Public Sub SimplifySquareRoot(ByVal val As Long, ByRef outM As Long, ByRef outN As Long)
    Dim i As Long
    outM = 1: outN = val: i = 2
    Do While i * i <= outN
        If outN Mod (i * i) = 0 Then
            outM = outM * i
            outN = outN / (i * i)
            ' Stay on same i to catch higher powers
        Else
            i = i + 1
        End If
    Loop
End Sub
Public Function SimplifyRoot(ByVal n As Long) As String
    Dim k As Long: k = 2
    Dim coeff As Long: coeff = 1
    Dim radicand As Long: radicand = n
    
    ' Extract square factors
    Do While k * k <= radicand
        If radicand Mod (k * k) = 0 Then
            coeff = coeff * k
            radicand = radicand / (k * k)
            k = 2 ' Reset to check for further factors
        Else
            k = k + 1
        End If
    Loop
    
    ' Format as string
    If radicand = 1 Then
        SimplifyRoot = CStr(coeff)
    ElseIf coeff = 1 Then
        SimplifyRoot = "\sqrt{" & radicand & "}"
    Else
        SimplifyRoot = coeff & "\sqrt{" & radicand & "}"
    End If
End Function

' 3. Standard LCM of two integers
Public Function lcm(ByVal a As Long, ByVal b As Long) As Long
    If a = 0 Or b = 0 Then lcm = 0: Exit Function
    lcm = Abs((a / GCD(a, b)) * b)
End Function
Public Function MultiplySurds(s1 As Surd, s2 As Surd) As Surd
    
    Dim result As Surd
    
    result.coeff = s1.coeff * s2.coeff
    result.radicand = s1.radicand * s2.radicand
    
    result = SquareRootToSurd(GetSurdSquaredValue(result))
    
    MultiplySurds = result
    
End Function
Public Function MultiplyFractionSurd(fs As FractionSurd, _
                                     mult As FractionSurd) As FractionSurd

    Dim result As FractionSurd

    ' Multiply numerators
    result.num.coeff = fs.num.coeff * mult.num.coeff
    result.num.radicand = fs.num.radicand * mult.num.radicand

    ' Multiply denominators
    result.den.coeff = fs.den.coeff * mult.den.coeff
    result.den.radicand = fs.den.radicand * mult.den.radicand

    MultiplyFractionSurd = result

End Function
Public Function MultiplyTwoFractionSurds(f1 As FractionSurd, _
                                         f2 As FractionSurd) As FractionSurd

    Dim result As FractionSurd
    
    ' (a/b) × (c/d) = (ac) / (bd)
    
    result.num.coeff = f1.num.coeff * f2.num.coeff
    result.num.radicand = f1.num.radicand * f2.num.radicand
    
    result.den.coeff = f1.den.coeff * f2.den.coeff
    result.den.radicand = f1.den.radicand * f2.den.radicand
    
    MultiplyTwoFractionSurds = result

End Function
Public Function AddFractionSurd(f1 As FractionSurd, _
                                f2 As FractionSurd) As FractionSurd

    Dim result As FractionSurd
    Dim leftPart As FractionSurd
    Dim rightPart As FractionSurd
    
    ' ---------------------------------
    ' Same denominator
    ' ---------------------------------
    
    If f1.den.coeff = f2.den.coeff _
       And f1.den.radicand = f2.den.radicand Then
        
        result.num.coeff = f1.num.coeff + f2.num.coeff
        result.num.radicand = f1.num.radicand
        result.den = f1.den
        
        AddFractionSurd = result
        Exit Function
        
    End If
    
    ' ---------------------------------
    ' Different denominators
    ' ---------------------------------
    Dim tempFS As FractionSurd
    
        ' Convert f2.den to FractionSurd
    tempFS.num = f2.den
    tempFS.den.coeff = 1
    tempFS.den.radicand = 1
    leftPart = MultiplyFractionSurd(f1, tempFS)
    
    ' Convert f1.den to FractionSurd
    tempFS.num = f1.den
    tempFS.den.coeff = 1
    tempFS.den.radicand = 1
    rightPart = MultiplyFractionSurd(f2, tempFS)
    
    result.num.coeff = leftPart.num.coeff + rightPart.num.coeff
    result.num.radicand = leftPart.num.radicand
    
    result.den.coeff = f1.den.coeff * f2.den.coeff
    result.den.radicand = f1.den.radicand * f2.den.radicand
    
    AddFractionSurd = result

End Function





Public Function AreEqual(f1 As FractionSurd, _
                         f2 As FractionSurd) As Boolean

    If f1.num.coeff = f2.num.coeff _
       And f1.num.radicand = f2.num.radicand _
       And f1.den.coeff = f2.den.coeff _
       And f1.den.radicand = f2.den.radicand Then
       
        AreEqual = True
        
    Else
    
        AreEqual = False
        
    End If

End Function
Public Function MultiplyFractionByInteger(fs As FractionSurd, _
                                          k As Long) As FractionSurd

    Dim result As FractionSurd
    
    result.num.coeff = fs.num.coeff * k
    result.num.radicand = fs.num.radicand
    
    result.den = fs.den
    
    MultiplyFractionByInteger = result

End Function
Public Function MultiplyFractionBySurd(fs As FractionSurd, _
                                       mult As Surd) As FractionSurd

    Dim result As FractionSurd

    ' Multiply numerator by multiplier (SIGN INCLUDED)
    result.num.coeff = fs.num.coeff * mult.coeff
    result.num.radicand = fs.num.radicand * mult.radicand

    ' Denominator remains same
    result.den.coeff = fs.den.coeff
    result.den.radicand = fs.den.radicand

    MultiplyFractionBySurd = result

End Function
Public Function SubtractFractionSurd(f1 As FractionSurd, _
                                     f2 As FractionSurd) As FractionSurd

    Dim result As FractionSurd
    Dim leftNum As Surd
    Dim rightNum As Surd
    Dim finalNum As Surd

    ' Cross multiply numerators:
    ' left = a*d
    leftNum.coeff = f1.num.coeff * f2.den.coeff
    leftNum.radicand = f1.num.radicand * f2.den.radicand

    ' right = c*b
    rightNum.coeff = f2.num.coeff * f1.den.coeff
    rightNum.radicand = f2.num.radicand * f1.den.radicand

    ' Subtract
    finalNum = SubtractSurds(leftNum, rightNum)

    ' Denominator = b*d
    result.den.coeff = f1.den.coeff * f2.den.coeff
    result.den.radicand = f1.den.radicand * f2.den.radicand

    result.num = finalNum

    SubtractFractionSurd = result

End Function
Public Function IsZeroFraction(fs As FractionSurd) As Boolean
    IsZeroFraction = (fs.num.coeff = 0)
End Function
Public Function SubtractSurds(s1 As Surd, s2 As Surd) As Surd
    
    Dim result As Surd
    
    If s1.radicand = s2.radicand Then
        result.coeff = s1.coeff - s2.coeff
        result.radicand = s1.radicand
    Else
        ' If radicals differ, determinant cannot simplify safely
        result.coeff = 1
        result.radicand = -1   ' special marker for unsupported
    End If
    
    SubtractSurds = result
    
End Function
Public Function DivideFractionSurd(num As FractionSurd, _
                                   den As FractionSurd) As FractionSurd

    Dim result As FractionSurd
    
    ' (a/b) ÷ (c/d) = (a×d)/(b×c)
    
    result.num.coeff = num.num.coeff * den.den.coeff
    result.num.radicand = num.num.radicand * den.den.radicand
    
    result.den.coeff = num.den.coeff * den.num.coeff
    result.den.radicand = num.den.radicand * den.num.radicand
    
    DivideFractionSurd = result

End Function
Public Function DivideSurdByFactor(s As Surd, g As Surd) As Surd

    Dim sSquared As Long
    Dim gSquared As Long
    Dim newSquared As Long
    
    sSquared = GetSquaredSurdValue(s)
    gSquared = GetSquaredSurdValue(g)
    
    newSquared = sSquared / gSquared
    
    DivideSurdByFactor = SquareRootToSurd(newSquared)

End Function
Public Function NormalizeAndSimplifyFractionSurd(f As FractionSurd) As FractionSurd

    Dim g As Long
    
    ' -----------------------------
    ' Move negative sign to numerator
    ' -----------------------------
    If f.den.coeff < 0 Then
        f.den.coeff = -f.den.coeff
        f.num.coeff = -f.num.coeff
    End If
    
    ' -----------------------------
    ' If radicands same ? reduce coefficients
    ' -----------------------------
    If f.num.radicand = f.den.radicand Then
        
        g = WorksheetFunction.GCD(Abs(f.num.coeff), Abs(f.den.coeff))
        
        If g > 0 Then
            f.num.coeff = f.num.coeff \ g
            f.den.coeff = f.den.coeff \ g
        End If
        
    End If
    
    NormalizeAndSimplifyFractionSurd = f

End Function
Public Sub SimplifyFractionSurd(ByRef fs As FractionSurd)

    Dim g As Long
    
    If fs.den.coeff = 0 Then Exit Sub
    
    g = Application.WorksheetFunction.GCD( _
            Abs(fs.num.coeff), _
            Abs(fs.den.coeff))
    
    If g > 1 Then
        fs.num.coeff = fs.num.coeff / g
        fs.den.coeff = fs.den.coeff / g
    End If
    
    ' Normalize sign so denominator is positive
    If fs.den.coeff < 0 Then
        fs.den.coeff = -fs.den.coeff
        fs.num.coeff = -fs.num.coeff
    End If

End Sub

Public Function GetSquaredSurdValue(s As Surd) As Long

    GetSquaredSurdValue = _
        (s.coeff * s.coeff) * s.radicand
        
End Function
Public Function AreSurdsEqual(s1 As Surd, s2 As Surd) As Boolean
    
    If s1.coeff = s2.coeff And _
       s1.radicand = s2.radicand Then
        AreSurdsEqual = True
    Else
        AreSurdsEqual = False
    End If
    
End Function

Public Function AreFractionSurdsEqual(f1 As FractionSurd, _
                                      f2 As FractionSurd) As Boolean
    
    Dim left As Surd
    Dim right As Surd
    
    ' Cross multiply
    left = MultiplySurds(f1.num, f2.den)
    right = MultiplySurds(f2.num, f1.den)
    
    AreFractionSurdsEqual = AreSurdsEqual(left, right)
    
End Function
Public Function AreOpposite(f1 As FractionSurd, _
                            f2 As FractionSurd) As Boolean

    If f1.num.radicand = f2.num.radicand Then
        If f1.num.coeff = -f2.num.coeff Then
            AreOpposite = True
            Exit Function
        End If
    End If
    
    AreOpposite = False

End Function
Public Function AreAbsEqual(f1 As FractionSurd, _
                            f2 As FractionSurd) As Boolean

    If f1.num.radicand <> f2.num.radicand Then
        AreAbsEqual = False
        Exit Function
    End If
    
    If Abs(f1.num.coeff) = Abs(f2.num.coeff) Then
        AreAbsEqual = True
    Else
        AreAbsEqual = False
    End If

End Function
Public Function AreSameSign(f1 As FractionSurd, _
                            f2 As FractionSurd) As Boolean

    If (f1.num.coeff >= 0 And f2.num.coeff >= 0) _
       Or (f1.num.coeff < 0 And f2.num.coeff < 0) Then
        AreSameSign = True
    Else
        AreSameSign = False
    End If

End Function
Public Function IsZeroSurd(s As Surd) As Boolean

    If s.coeff = 0 Then
        IsZeroSurd = True
    Else
        IsZeroSurd = False
    End If

End Function
Public Function IsZeroFractionSurd(fs As FractionSurd) As Boolean

    If fs.num.coeff = 0 Then
        IsZeroFractionSurd = True
    Else
        IsZeroFractionSurd = False
    End If

End Function
Public Function IsFractionSurdZero(fs As FractionSurd) As Boolean

    ' A fraction surd is zero if numerator coefficient is zero
    If fs.num.coeff = 0 Then
        IsFractionSurdZero = True
    Else
        IsFractionSurdZero = False
    End If

End Function
Public Function GetSurdSquaredValue(s As Surd) As Long
    GetSurdSquaredValue = (s.coeff * s.coeff) * s.radicand
End Function
Public Function GCD(ByVal a As Long, ByVal b As Long) As Long
    Do While b <> 0
        Dim temp As Long: temp = b
        b = a Mod b
        a = temp
    Loop
    GCD = a
End Function
Public Function GetSquaredDenominator(fs As FractionSurd) As Long
    
    ' (pvq)^2 = p^2 * q
    
    GetSquaredDenominator = _
        (fs.den.coeff * fs.den.coeff) * fs.den.radicand
        
End Function
