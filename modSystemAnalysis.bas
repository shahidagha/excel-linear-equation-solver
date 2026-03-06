Attribute VB_Name = "modSystemAnalysis"
Public Function ValidateSystem(e1 As StandardForm, _
                               e2 As StandardForm) As SystemType

    ' 1?? Check if either equation is contradiction
    If IsContradiction(e1) Or IsContradiction(e2) Then
        ValidateSystem = SystemContradictory
        Exit Function
    End If

    ' 2?? Check if both are identities
    If IsIdentity(e1) And IsIdentity(e2) Then
        ValidateSystem = SystemIdentity
        Exit Function
    End If

    ' 3?? Check determinant
    If DeterminantIsZero(e1, e2) Then
        ValidateSystem = SystemDependent
    Else
        ValidateSystem = SystemIndependent
    End If

End Function

Public Function ValidateEquation(sf As StandardForm) As EquationStatus
    
    Dim aZero As Boolean
    Dim bZero As Boolean
    Dim cZero As Boolean
    
    aZero = IsFractionSurdZero(sf.aCoeff)
    bZero = IsFractionSurdZero(sf.bCoeff)
    cZero = IsFractionSurdZero(sf.constCoeff)
    
    If aZero And bZero And cZero Then
        ValidateEquation = Identity
        Exit Function
    End If
    
    If aZero And bZero And Not cZero Then
        ValidateEquation = Contradiction
        Exit Function
    End If
    
    ValidateEquation = Normal

End Function
Public Function AreStandardFormsEqual(a As StandardForm, b As StandardForm) As Boolean

    If AreFractionSurdsEqual(a.aCoeff, b.aCoeff) _
       And AreFractionSurdsEqual(a.bCoeff, b.bCoeff) _
       And AreFractionSurdsEqual(a.constCoeff, b.constCoeff) Then
       
        AreStandardFormsEqual = True
    Else
        AreStandardFormsEqual = False
    End If

End Function

