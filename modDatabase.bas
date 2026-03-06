Attribute VB_Name = "modDatabase"
Public Sub SaveStandardizedData(frm As Object, r As Long)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Database")
    Dim fullSolution As String
    fullSolution = ProcessSystem(frm)
    Dim cramerSolution As String
    cramerSolution = ProcessCramerMethod(frm)
    
    ws.Cells(r, 49).value = cramerSolution   ' Column AW
    'MsgBox TypeName(frm)
    On Error GoTo ErrLine
    'On Error GoTo SaveError
    ' --- SAVE VARIABLES ---
    ws.Cells(r, 2).value = frm.txtPVar.value
    ws.Cells(r, 3).value = frm.txtSVar.value
    ' --- STEP 1: Save All Raw Input Values (Columns 4 - 41) ---
    ' Equation 1
        

10     ws.Cells(r, 4).value = GetSafeValue(frm, "txtPVTP1")
20     ws.Cells(r, 5).value = GetSafeValue(frm, "txtSVTP1")
30    ws.Cells(r, 6).value = GetSafeValue(frm, "txtETP1")
40    ws.Cells(r, 7).value = GetSafeValue(frm, "txtCTP1")
50    ws.Cells(r, 8).value = GetSafeCaption(frm, "tglSignFT1")
60    ws.Cells(r, 9).value = GetSafeValue(frm, "txtANC1")
70    ws.Cells(r, 10).value = GetSafeValue(frm, "txtANR1")
80    ws.Cells(r, 11).value = GetSafeValue(frm, "txtADC1")
90    ws.Cells(r, 12).value = GetSafeValue(frm, "txtADR1")
    
100    ws.Cells(r, 13).value = GetSafeCaption(frm, "tglSignST1")
110    ws.Cells(r, 14).value = GetSafeValue(frm, "txtBNC1")
120    ws.Cells(r, 15).value = GetSafeValue(frm, "txtBNR1")
130    ws.Cells(r, 16).value = GetSafeValue(frm, "txtBDC1")
140    ws.Cells(r, 17).value = GetSafeValue(frm, "txtBDR1")
    
150    ws.Cells(r, 18).value = GetSafeCaption(frm, "tglSignCT1")
160    ws.Cells(r, 19).value = GetSafeValue(frm, "txtCNC1")
170    ws.Cells(r, 20).value = GetSafeValue(frm, "txtCNR1")
180    ws.Cells(r, 21).value = GetSafeValue(frm, "txtCDC1")
190    ws.Cells(r, 22).value = GetSafeValue(frm, "txtCDR1")
    
    ' Equation 2
200    ws.Cells(r, 23).value = GetSafeValue(frm, "txtPVTP2")
210    ws.Cells(r, 24).value = GetSafeValue(frm, "txtSVTP2")
220    ws.Cells(r, 25).value = GetSafeValue(frm, "txtETP2")
230    ws.Cells(r, 26).value = GetSafeValue(frm, "txtCTP2")
    
240    ws.Cells(r, 27).value = GetSafeCaption(frm, "tglSignFT2")
250    ws.Cells(r, 28).value = GetSafeValue(frm, "txtANC2")
260    ws.Cells(r, 29).value = GetSafeValue(frm, "txtANR2")
270    ws.Cells(r, 30).value = GetSafeValue(frm, "txtADC2")
280    ws.Cells(r, 31).value = GetSafeValue(frm, "txtADR2")
    
290    ws.Cells(r, 32).value = GetSafeCaption(frm, "tglSignST2")
300    ws.Cells(r, 33).value = GetSafeValue(frm, "txtBNC2")
310    ws.Cells(r, 34).value = GetSafeValue(frm, "txtBNR2")
320    ws.Cells(r, 35).value = GetSafeValue(frm, "txtBDC2")
330    ws.Cells(r, 36).value = GetSafeValue(frm, "txtBDR2")
    
340    ws.Cells(r, 37).value = GetSafeCaption(frm, "tglSignCT2")
350    ws.Cells(r, 38).value = GetSafeValue(frm, "txtCNC2")
360    ws.Cells(r, 39).value = GetSafeValue(frm, "txtCNR2")
370    ws.Cells(r, 40).value = GetSafeValue(frm, "txtCDC2")
380    ws.Cells(r, 41).value = GetSafeValue(frm, "txtCDR2")

    ' --- STEP 2: Math & LaTeX Generation ---
    ' Now that we've cleared the data, re-run the calculations
390    ws.Cells(r, 42).value = GetSymbolicEq(frm, 1)
400    ws.Cells(r, 43).value = GetSymbolicEq(frm, 2)
    
410    ws.Cells(r, 44).value = frm.LastLtx1
420    ws.Cells(r, 45).value = frm.LastLtx2
    
430    ws.Cells(r, 46).value = "SUBSTITUTION"
440    ws.Cells(r, 47).value = "ELIMINATION"
450    ws.Cells(r, 48).value = "GRAPHICAL"
460    ws.Cells(r, 49).value = cramerSolution
470    ws.Cells(r, 50).value = fullSolution
Exit Sub

ErrLine:
    MsgBox "Failed at line: " & Erl
'SaveError:
 '           MsgBox "Error in SaveStandardizedData: " & Err.Description, vbCritical
End Sub
Public Function FindDuplicateRecord(eq1 As String, eq2 As String) As Long

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Database")
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    eq1 = Trim(eq1)
    eq2 = Trim(eq2)
    
    For i = 2 To lastRow
        
        Dim dbEq1 As String
        Dim dbEq2 As String
        
        dbEq1 = Trim(ws.Cells(i, 42).value)
        dbEq2 = Trim(ws.Cells(i, 43).value)
        
        ' Case 1: Same order
        If dbEq1 = eq1 And dbEq2 = eq2 Then
            FindDuplicateRecord = ws.Cells(i, 1).value
            Exit Function
        End If
        
        ' Case 2: Swapped order
        If dbEq1 = eq2 And dbEq2 = eq1 Then
            FindDuplicateRecord = ws.Cells(i, 1).value
            Exit Function
        End If
        
    Next i
    
    FindDuplicateRecord = 0

End Function
Public Function FindDuplicateRecordExcluding(eq1 As String, eq2 As String, excludeRow As Long) As Long

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Database")
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    eq1 = Trim(eq1)
    eq2 = Trim(eq2)
    
    For i = 2 To lastRow
        
        If i = excludeRow Then GoTo NextRow
        
        Dim dbEq1 As String
        Dim dbEq2 As String
        
        dbEq1 = Trim(ws.Cells(i, 42).value)
        dbEq2 = Trim(ws.Cells(i, 43).value)
        
        If (dbEq1 = eq1 And dbEq2 = eq2) Or _
           (dbEq1 = eq2 And dbEq2 = eq1) Then
           
            FindDuplicateRecordExcluding = ws.Cells(i, 1).value
            Exit Function
        End If
        
NextRow:
    Next i
    
    FindDuplicateRecordExcluding = 0

End Function
Public Sub GenerateAllSolutions(r As Long)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Database")
    
    Dim sf1 As StandardForm
    Dim sf2 As StandardForm
    
    Dim var1 As String
    Dim var2 As String
    
    Dim eliminationLatex As String
    Dim substitutionLatex As String
    Dim cramerLatex As String
    Dim graphicalLatex As String
    
    
    ' ----------------------------------
    ' Read standardized data
    ' ----------------------------------
    
    sf1 = ReadStandardForm(ws, r, 1)
    sf2 = ReadStandardForm(ws, r, 2)
    
    var1 = ws.Cells(r, 6).value
    var2 = ws.Cells(r, 7).value
    
    
    ' ----------------------------------
    ' Common Initial Steps
    ' ----------------------------------
    
    Dim baseSteps As String
    
    baseSteps = BuildCommonInitialSteps(sf1, sf2, var1, var2)
    
    
    ' ----------------------------------
    ' Elimination
    ' ----------------------------------
    
    eliminationLatex = baseSteps & _
                       AppendEliminationSteps(sf1, sf2, var1, var2)
    
    ws.Cells(r, 47).value = eliminationLatex
    
    
    ' ----------------------------------
    ' Substitution
    ' ----------------------------------
    
    substitutionLatex = baseSteps & _
                        AppendSubstitutionSteps(sf1, sf2, var1, var2)
    
    ws.Cells(r, 46).value = substitutionLatex
    
    
    ' ----------------------------------
    ' Cramer
    ' ----------------------------------
    
    cramerLatex = baseSteps & _
                  AppendCramerSteps(sf1, sf2, var1, var2)
    
    ws.Cells(r, 49).value = cramerLatex
    
    
    ' ----------------------------------
    ' Graphical
    ' ----------------------------------
    
    graphicalLatex = baseSteps & _
                     AppendGraphicalSteps(sf1, sf2, var1, var2)
    
    ws.Cells(r, 48).value = graphicalLatex

End Sub
