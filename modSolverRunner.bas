Attribute VB_Name = "modSolverRunner"
Public Sub SolveByMethod(frm As Object, methodName As String)

    Dim webSolution As String
    Dim sf1 As StandardForm
    Dim sf2 As StandardForm
    
    webSolution = "\begin{aligned}" & vbCrLf
    
    ' Process equations
    sf1 = ProcessSingleEquation(frm, 1, webSolution)
    sf2 = ProcessSingleEquation(frm, 2, webSolution)
    
    ' Choose method
    Select Case methodName
    
        Case "ELIM"
            webSolution = webSolution & _
                CrossEliminationSequence(sf1, sf2, _
                frm.txtPVar.value, frm.txtSVar.value)
        
        Case "SUB"
            webSolution = webSolution & _
                AppendSubstitutionSolve(sf1.aCoeff, sf1.bCoeff, _
                sf2.aCoeff, sf2.constCoeff, sf1.constCoeff, _
                frm.txtPVar.value, frm.txtSVar.value)
        
        Case "CRAMER"
            webSolution = webSolution & _
                AppendCramerSteps(sf1, sf2, _
                frm.txtPVar.value, frm.txtSVar.value)
        
        Case "GRAPH"
            webSolution = webSolution & _
                AppendGraphicalSteps(sf1, sf2, _
                frm.txtPVar.value, frm.txtSVar.value)
    
    End Select
    
    webSolution = webSolution & "\end{aligned}"
    
    frm.txtLatexOutput.value = webSolution

End Sub

