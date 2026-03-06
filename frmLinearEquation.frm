VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLinearEquation 
   Caption         =   "UserForm1"
   ClientHeight    =   13395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21945
   OleObjectBlob   =   "frmLinearEquation.frx":0000
End
Attribute VB_Name = "frmLinearEquation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' --- GLOBAL VARIABLES ---
Public ControlCollection As New Collection
Public OldVal As String
Public IsSwapping As Boolean
Public LastLtx1 As String
Public LastLtx2 As String
Public IsUpdating As Boolean
Private StatusClearTime As Double
Private StatusFading As Boolean
Private mThemeApplied As Boolean
Private CurrentMode As FormMode
Private IsDirty As Boolean
Private ActiveMethodColumn As Long
' --- 1. INITIALIZATION ---
Private Sub UserForm_Initialize()

    ' Base form color
    Me.BackColor = RGB(248, 249, 252)

    ' Status bar
    SetupStatusBar

    ' Build control collection
    SetupControls

    ' Setup ListView
    SetupListView

    ' Hide sections initially
    Me.fraVariable.Visible = False
    Me.fraEq1.Visible = False
    Me.fraEq2.Visible = False

    ' Apply theme LAST
    ApplyModernLightTheme Me
    RestoreImportantLabels
    AdjustSqrtLabels
    ApplyButtonIcons
    SetFormMode ModeIdle
    ApplyButtonIcons

Dim btn As Control
For Each btn In Me.Controls
    If TypeName(btn) = "CommandButton" Then
        btn.Font.Name = "Segoe UI"
        btn.Font.Size = 10
    End If
Next btn
    InitBrowser Me.webPreview1
    InitBrowser Me.webPreview2
    InitBrowser Me.webQuestion
    InitBrowser Me.webSolution
End Sub
Public Sub MarkDirty()
    
    If CurrentMode = ModeIdle Then Exit Sub
    
    IsDirty = True
    
End Sub
Private Sub StyleButton(btn As MSForms.CommandButton, isEnabled As Boolean)

    btn.Enabled = isEnabled
    
    If isEnabled Then
        btn.ForeColor = RGB(255, 255, 255)
        btn.Font.Bold = True
    Else
        btn.ForeColor = RGB(150, 150, 150)
        btn.BackColor = RGB(230, 230, 230)
        btn.Font.Bold = False
    End If

End Sub

Private Sub ApplyButtonColors()

    ' Add - Dark Gray
    Me.cmdAdd.BackColor = RGB(70, 70, 70)

    ' Save - Green
    Me.cmdSave.BackColor = RGB(40, 167, 69)

    ' Update - Blue
    Me.cmdUpdate.BackColor = RGB(0, 123, 255)

    ' Delete - Red
    Me.cmdDelete.BackColor = RGB(220, 53, 69)

    ' Reset - Neutral
    Me.cmdReset.BackColor = RGB(108, 117, 125)

End Sub
Private Sub RestoreImportantLabels()

    Dim lbl As Control

    For Each lbl In Me.Controls
        
        If TypeName(lbl) = "Label" Then
            
            ' ===============================
            ' Variable Labels (x , y)
            ' ===============================
            If lbl.Name Like "lblV*" Then
                With lbl
                    .Font.Name = "Cambria Math"
                    .Font.Size = 16
                    .Font.Bold = False
                    .Font.Italic = True
                    .ForeColor = RGB(0, 0, 0)
                    .AutoSize = False
                    .TextAlign = fmTextAlignCenter
                    .Height = 22
                End With
            End If
            
            ' ===============================
            ' Equal Labels (=)
            ' ===============================
            If lbl.Name Like "lblET*" Then
                With lbl
                    .Font.Name = "Cambria Math"
                    .Font.Size = 18
                    .Font.Bold = True
                    .Font.Italic = False
                    .ForeColor = RGB(0, 0, 0)
                    .AutoSize = False
                    .TextAlign = fmTextAlignCenter
                    .Height = 24
                End With
            End If
            
        End If
        
    Next lbl

End Sub
Private Sub AdjustSqrtLabels()

    Dim ctrl As Control
    
    For Each ctrl In Me.Controls
        
        If TypeName(ctrl) = "Label" Then
            
            If ctrl.Name Like "lblsqrt*" Then
                
                With ctrl
                    .Caption = ChrW(&H221A)
                    .Font.Name = "Cambria Math"
                    .Font.Size = 18
                    .Font.Bold = False
                    .Font.Italic = False
                    .ForeColor = RGB(0, 0, 0)
                    .TextAlign = fmTextAlignCenter
                    .Top = .Top + 5   ' adjust if needed
                End With
                
            End If
            
        End If
        
    Next ctrl

End Sub
Private Sub SetFormMode(newMode As FormMode)

    CurrentMode = newMode
    
    ApplyButtonColors

    Select Case newMode

        Case ModeIdle
            
            SetInputState False
            
            StyleButton Me.cmdAdd, True
            StyleButton Me.cmdSave, False
            StyleButton Me.cmdUpdate, False
            StyleButton Me.cmdDelete, False
            StyleButton Me.cmdReset, False

        Case ModeAdd
            
            SetInputState True
            
            StyleButton Me.cmdAdd, False
            StyleButton Me.cmdSave, True
            StyleButton Me.cmdUpdate, False
            StyleButton Me.cmdDelete, False
            StyleButton Me.cmdReset, True

        Case ModeEdit
            
            SetInputState True
            
            StyleButton Me.cmdAdd, False
            StyleButton Me.cmdSave, False
            StyleButton Me.cmdUpdate, True
            StyleButton Me.cmdDelete, True
            StyleButton Me.cmdReset, True

    End Select

End Sub
' --- 2. SOLUTION RENDERER (Fixed Security Bypass) ---
Public Sub RefreshSolutionDisplay()

    Dim r As Long
    r = val(Me.Tag)

    Dim rawLatex As String

    ' ==========================
    ' GET STORED SOLUTION (AX = 50)
    ' ==========================
    If r > 0 Then
        rawLatex = ThisWorkbook.Sheets("Database").Cells(r, 50).value
    End If

    If rawLatex = "" Then
        rawLatex = "No solution data found."
    End If

    ' ==========================
    ' RENDER MAIN SOLUTION
    ' ==========================
    Dim html As String
    Dim tempPath As String

    html = "<!-- saved from url=(0014)about:internet -->" & vbCrLf & _
           "<html><head>" & _
           "<meta http-equiv='X-UA-Compatible' content='IE=edge'>" & _
           "<script src='https://cdnjs.cloudflare.com/ajax/libs/mathjax/2.7.7/MathJax.js?config=TeX-AMS_HTML'></script>" & _
           "<style>body{margin:10px;font-family:Cambria Math;font-size:12pt;overflow:auto;}</style>" & _
           "</head><body>" & rawLatex & "</body></html>"

    tempPath = Environ("Temp") & "\vba_sln_render.html"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim ts As Object
    Set ts = fso.CreateTextFile(tempPath, True, False)
    ts.Write html
    ts.Close

    Me.webSolution.Navigate tempPath

    ' ==========================
    ' UPDATE LIVE PREVIEWS
    ' ==========================
    RenderMathJax Me.LastLtx1, Me, 1
    RenderMathJax Me.LastLtx2, Me, 2
    
    Dim qLatex As String
    qLatex = Me.LastLtx1 & " ; " & Me.LastLtx2
    RenderMathJax qLatex, Me, "Question"
    'ActiveMethodColumn = 49
    'SetActiveMethodButton cmdCramer
End Sub
Public Function WrapMath(latex As String) As String

    WrapMath = "<html><head>" & _
        "<meta http-equiv='X-UA-Compatible' content='IE=edge'>" & _
        "<script src='https://cdnjs.cloudflare.com/ajax/libs/mathjax/2.7.7/MathJax.js?config=TeX-AMS_HTML'></script>" & _
        "<style>body{margin:5px;font-family:Arial;font-size:11pt;}</style>" & _
        "</head><body> $$ " & latex & " $$ </body></html>"

End Function

' --- 3. SYNC & POSITION LOGIC ---
'Public Sub SyncAll(EqNum As Integer)
'    Debug.Print "SyncAll running for Eq"; EqNum
'   On Error GoTo 0
'    ' Always update variable labels
'    Me.Controls("lblV1" & EqNum).Caption = Me.txtPVar.Value
'    Me.Controls("lblV2" & EqNum).Caption = Me.txtSVar.Value
'
'    ' If loading record, stop here
'    If Me.IsUpdating Then
'        On Error GoTo 0
'        Exit Sub
'    End If
'
'    ' Only layout + preview update
'    RearrangeFrames Me, EqNum
'    UpdateEquationPreview Me, EqNum
'    Debug.Print "Calling UpdateEquationPreview for Eq"; EqNum
'    On Error GoTo 0
'
'End Sub
Public Sub SyncAll(eqNum As Integer)

    Me.Controls("lblV1" & eqNum).Caption = Me.txtPVar.value
    Me.Controls("lblV2" & eqNum).Caption = Me.txtSVar.value

    Dim pPos As Integer: pPos = val(Me.Controls("txtPVTP" & eqNum).value)
    Dim sPos As Integer: sPos = val(Me.Controls("txtSVTP" & eqNum).value)
    Dim cPos As Integer: cPos = val(Me.Controls("txtCTP" & eqNum).value)
    Dim ePos As Integer: ePos = val(Me.Controls("txtETP" & eqNum).value)

    RearrangeFrames Me, eqNum
    UpdateEquationPreview Me, eqNum
    If eqNum = 1 Then
    If Me.LastLtx2 = "" Then UpdateEquationPreview Me, 2
    Else
        If Me.LastLtx1 = "" Then UpdateEquationPreview Me, 1
    End If

    UpdateQuestionPreview Me
   

End Sub
Private Sub UpdateSign(tgl As MSForms.ToggleButton, pos As Integer, eqPos As Integer)
    If pos = 1 Or pos = eqPos + 1 Then
        tgl.ForeColor = IIf(tgl.Caption = "+", tgl.BackColor, vbBlack)
    Else
        tgl.ForeColor = vbBlack
    End If
End Sub

' --- 4. DATA OPERATIONS ---
Private Sub cmdAdd_Click()

    On Error GoTo CleanExit

    Application.ScreenUpdating = False

    Me.IsUpdating = True   ' ?? Lock rendering

    Me.lblLoadedID.Caption = ""
    Me.Tag = ""
    
    IsSwapping = True
    SetDefaults 1
    SetDefaults 2
    IsSwapping = False
    
    Me.fraVariable.Visible = True
    Me.fraEq1.Visible = True
    Me.fraEq2.Visible = True
    
    SyncAll 1
    SyncAll 2
    
    Me.IsUpdating = False  ' ?? Unlock rendering

    ' ? Render ONCE after everything ready
    UpdateEquationPreview Me, 1
    UpdateEquationPreview Me, 2
    UpdateQuestionPreview Me

    SetFormMode ModeAdd

CleanExit:
    Application.ScreenUpdating = True

    Me.txtANC1.SetFocus
    Me.txtANC1.SelStart = 0
    Me.txtANC1.SelLength = Len(Me.txtANC1.Text)

End Sub
    

Private Sub cmdSave_Click()

    Dim ws As Worksheet
    Dim r As Long
    Dim nextID As Long
    
    Set ws = ThisWorkbook.Sheets("Database")
    
    On Error GoTo ErrorHandler
    
    ' ======================================
    ' SAVE ONLY IN ADD MODE
    ' ======================================
    
    If CurrentMode <> ModeAdd Then Exit Sub
    
    ' ======================================
    ' DUPLICATE CHECK
    ' ======================================
    
    Dim eq1 As String
    Dim eq2 As String
    Dim duplicateSr As Long
    
    eq1 = GetSymbolicEq(Me, 1)
    eq2 = GetSymbolicEq(Me, 2)
    
    duplicateSr = FindDuplicateRecord(eq1, eq2)
    
    If duplicateSr > 0 Then
        MsgBox "Duplicate record found." & vbCrLf & _
               "Same question exists at Sr. No: " & duplicateSr, vbExclamation
        Exit Sub
    End If
    
    ' ======================================
    ' DETERMINE NEW ROW
    ' ======================================
    
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' ======================================
    ' GENERATE NEW ID
    ' ======================================
    
    If Application.WorksheetFunction.CountA(ws.Columns(1)) = 0 Then
        nextID = 1
    Else
        nextID = Application.WorksheetFunction.Max(ws.Columns(1)) + 1
    End If
    
    ws.Cells(r, 1).value = nextID
    Me.lblLoadedID.Caption = nextID
    Me.Tag = r
    
    ' ======================================
    ' SAVE RAW DATA
    ' ======================================
    
    SaveStandardizedData Me, r
    
    ' ======================================
    ' GENERATE ALL METHOD SOLUTIONS
    ' ======================================
    
    GenerateAllMethodSolutions Me, r
    
    ' ======================================
    ' RECALCULATE QUESTION NUMBERS
    ' ======================================
    
    RecalculateQuestionNumbers
    
    ' ======================================
    ' REFRESH UI
    ' ======================================
    
    UpdateListView
    RefreshSolutionDisplay
    
    UpdateStatus "Record saved successfully.", StatusSuccess
    
    ' ======================================
    ' RETURN TO IDLE MODE
    ' ======================================
    
    SetFormMode ModeIdle
    IsDirty = False
    
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & vbCrLf & _
           Err.Description & vbCrLf & _
           "Source: " & Err.Source, vbCritical

End Sub
Public Sub LoadRecordToForm(rowNum As Long)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Database")

    Me.IsUpdating = True
    Me.Tag = rowNum
    Me.lblLoadedID.Caption = ws.Cells(rowNum, 1).value

    ' Variables
    Me.txtPVar.value = ws.Cells(rowNum, 2).value
    Me.txtSVar.value = ws.Cells(rowNum, 3).value

    Dim e As Integer, colOffset As Integer

    For e = 1 To 2

        colOffset = IIf(e = 1, 0, 19)

        ' Positions
        Me.Controls("txtPVTP" & e).value = ws.Cells(rowNum, 4 + colOffset).value
        Me.Controls("txtSVTP" & e).value = ws.Cells(rowNum, 5 + colOffset).value
        Me.Controls("txtETP" & e).value = ws.Cells(rowNum, 6 + colOffset).value
        Me.Controls("txtCTP" & e).value = ws.Cells(rowNum, 7 + colOffset).value

        ' Signs
        Me.Controls("tglSignFT" & e).Caption = ws.Cells(rowNum, 8 + colOffset).value
        Me.Controls("tglSignST" & e).Caption = ws.Cells(rowNum, 13 + colOffset).value
        Me.Controls("tglSignCT" & e).Caption = ws.Cells(rowNum, 18 + colOffset).value

        ' Coefficients
        Me.Controls("txtANC" & e).value = ws.Cells(rowNum, 9 + colOffset).value
        Me.Controls("txtANR" & e).value = ws.Cells(rowNum, 10 + colOffset).value
        Me.Controls("txtADC" & e).value = ws.Cells(rowNum, 11 + colOffset).value
        Me.Controls("txtADR" & e).value = ws.Cells(rowNum, 12 + colOffset).value

        Me.Controls("txtBNC" & e).value = ws.Cells(rowNum, 14 + colOffset).value
        Me.Controls("txtBNR" & e).value = ws.Cells(rowNum, 15 + colOffset).value
        Me.Controls("txtBDC" & e).value = ws.Cells(rowNum, 16 + colOffset).value
        Me.Controls("txtBDR" & e).value = ws.Cells(rowNum, 17 + colOffset).value

        Me.Controls("txtCNC" & e).value = ws.Cells(rowNum, 19 + colOffset).value
        Me.Controls("txtCNR" & e).value = ws.Cells(rowNum, 20 + colOffset).value
        Me.Controls("txtCDC" & e).value = ws.Cells(rowNum, 21 + colOffset).value
        Me.Controls("txtCDR" & e).value = ws.Cells(rowNum, 22 + colOffset).value

    Next e

    Me.IsUpdating = False
        ' Force rebuild after loading
    UpdateEquationPreview Me, 1
    UpdateEquationPreview Me, 2

    ' Now just refresh display
    UpdateQuestionPreview Me
    RefreshSolutionDisplay
    UpdateListView
    SetFormMode ModeEdit
    Me.fraVariable.Visible = True
    Me.fraEq1.Visible = True
    Me.fraEq2.Visible = True
    Me.txtANC1.SetFocus
    Me.txtANC1.SelStart = 0
    Me.txtANC1.SelLength = Len(Me.txtANC1.Text)

End Sub
' --- 6. LISTVIEW INTERACTION ---
Public Sub UpdateListView(Optional ByVal filterText As String = "")

    Dim ws As Worksheet
    Dim li As ListItem
    Dim i As Long
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Sheets("Database")
    
    Me.ListView1.ListItems.Clear
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    For i = 2 To lastRow
        
        Dim eq1 As String
        Dim eq2 As String
        Dim qNo As String
        
        qNo = CStr(ws.Cells(i, 51).value)
        eq1 = ws.Cells(i, 42).value
        eq2 = ws.Cells(i, 43).value
        
        ' Apply filter
        If filterText = "" Or _
           InStr(1, qNo & eq1 & eq2, filterText, vbTextCompare) > 0 Then
           
            Set li = Me.ListView1.ListItems.Add(, , qNo)
            li.Tag = ws.Cells(i, 1).value
            li.SubItems(1) = eq1
            li.SubItems(2) = eq2
            
        End If
        
    Next i

End Sub
Private Sub txtSearch_Change()

    UpdateListView Me.txtSearch.Text

End Sub
Private Sub ListView1_DblClick()
    If IsDirty Then
    If MsgBox("You have unsaved changes. Continue?", _
              vbExclamation + vbYesNo) = vbNo Then Exit Sub
    End If
    If Me.ListView1.SelectedItem Is Nothing Then Exit Sub

    Dim selectedSr As Long
    selectedSr = CLng(Me.ListView1.SelectedItem.Tag)   ' ? IMPORTANT

    Dim ws As Worksheet
    Dim foundCell As Range
    
    Set ws = ThisWorkbook.Sheets("Database")
    
    Set foundCell = ws.Columns(1).Find(What:=selectedSr, LookAt:=xlWhole)
    
    If Not foundCell Is Nothing Then
        LoadRecordToForm foundCell.Row
    Else
        MsgBox "Record not found in database.", vbExclamation
    End If

End Sub


'====================================================================================================================================================================================
'====================================================================================================================================================================================
'====================================================================================================================================================================================
'====================================================================================================================================================================================



' --- 5. SWAP LOGIC ---
Public Sub HandleSwap(ByRef TargetBox As MSForms.TextBox)
    Dim currentEq As Integer: currentEq = val(right(TargetBox.Name, 1))
    Dim newPos As String: newPos = TargetBox.value
    Dim oldPos As String: oldPos = Me.OldVal
    Dim ctrl As Control
    
    If newPos = "" Or oldPos = "" Or newPos = oldPos Then Exit Sub
    
    Me.IsUpdating = True
    For Each ctrl In Me.Controls
        ' Find partner box in current equation (e.g. txtPVTP1, txtSVTP1)
        If ctrl.Name Like "*TP" & currentEq And ctrl.Name <> TargetBox.Name Then
            If ctrl.value = newPos Then
                ctrl.value = oldPos
                Exit For
            End If
        End If
    Next ctrl
    Me.IsUpdating = False
    SyncAll currentEq
End Sub

' --- 6. UTILITY FUNCTIONS (Keep as updated) ---
Private Sub cmdReset_Click()
    If IsDirty Then
    If MsgBox("You have unsaved changes. Continue?", _
              vbExclamation + vbYesNo) = vbNo Then Exit Sub
    End If
    ' ======================================
    ' CLEAR STATE
    ' ======================================
    
    Me.Tag = 0
    Me.lblLoadedID.Caption = ""
    
    ' ======================================
    ' RESET EQUATIONS
    ' ======================================
    
    IsSwapping = True
    
    SetDefaults 1
    SetDefaults 2
    
    IsSwapping = False
    
    ' ======================================
    ' HIDE UI SECTIONS
    ' ======================================
    
    Me.fraVariable.Visible = False
    Me.fraEq1.Visible = False
    Me.fraEq2.Visible = False
    
    ' ======================================
    ' CLEAR PREVIEWS
    ' ======================================
    
    Me.LastLtx1 = ""
    Me.LastLtx2 = ""
    
    ' Optional: clear rendered output
    RefreshSolutionDisplay
    
    ' ======================================
    ' RETURN TO IDLE MODE
    ' ======================================
    
    SetFormMode ModeIdle

End Sub
Private Sub SetInputState(enableInput As Boolean)

    Dim ctrl As Control
    
    For Each ctrl In Me.Controls
        
        Select Case TypeName(ctrl)
            
            Case "TextBox", "ToggleButton"
                ctrl.Enabled = enableInput
                
        End Select
        
    Next ctrl

End Sub
Public Sub RecalculateQuestionNumbers()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim qNo As Long
    
    Set ws = ThisWorkbook.Sheets("Database")
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    qNo = 1
    
    For i = 2 To lastRow
        ws.Cells(i, 51).value = qNo
        qNo = qNo + 1
    Next i

End Sub

Private Sub cmdCopySoln_Click()

    Dim ws As Worksheet
    Dim r As Long
    
    Set ws = ThisWorkbook.Sheets("Database")
    r = val(Me.Tag)
    
    If r = 0 Then
        MsgBox "No record loaded.", vbExclamation
        Exit Sub
    End If
    
    If ActiveMethodColumn = 0 Then
        MsgBox "Please select a solution method first.", vbExclamation
        Exit Sub
    End If
    
    Module1.CopyToClipboard ws.Cells(r, ActiveMethodColumn).value

End Sub

' Method triggers now stay on one page
'Private Sub cmdCrammer_Click(): RefreshSolutionDisplay: End Sub
'Private Sub cmdSubstitution_Click(): RefreshSolutionDisplay: End Sub
'Private Sub cmdElimination_Click(): RefreshSolutionDisplay: End Sub
'Private Sub cmdGraphical_Click(): RefreshSolutionDisplay: End Sub
'====================================================================================================================================================================================
'====================================================================================================================================================================================
'====================================================================================================================================================================================
'====================================================================================================================================================================================

Private Sub cmdLoad_Click()
    Call LogAllControlNames
End Sub

' --- SET DEFAULTS FOR POSITION BOXES ---
Private Sub SetDefaults(eqNum As Integer)
    ' Set the Variables
    If eqNum = 1 Then
        Me.txtPVar.value = "x"
        Me.txtSVar.value = "y"
    End If

    ' Set Positions
    Me.Controls("txtPVTP" & eqNum).value = "1"
    Me.Controls("txtSVTP" & eqNum).value = "2"
    Me.Controls("txtETP" & eqNum).value = "3"
    Me.Controls("txtCTP" & eqNum).value = "4"

    ' Set Coefficients to 1 and Radicands to empty
    ' Primary Term
    Me.Controls("tglSignFT" & eqNum).Caption = "+"
    Me.Controls("txtANC" & eqNum).value = "1"
    Me.Controls("txtANR" & eqNum).value = "1"
    Me.Controls("txtADC" & eqNum).value = "1"
    Me.Controls("txtADR" & eqNum).value = "1"

    ' Secondary Term
    Me.Controls("tglSignST" & eqNum).Caption = "+"
    Me.Controls("txtBNC" & eqNum).value = "1"
    Me.Controls("txtBNR" & eqNum).value = "1"
    Me.Controls("txtBDC" & eqNum).value = "1"
    Me.Controls("txtBDR" & eqNum).value = "1"

    ' Constant Term
    Me.Controls("tglSignCT" & eqNum).Caption = "+"
    Me.Controls("txtCNC" & eqNum).value = "1"
    Me.Controls("txtCNR" & eqNum).value = "1"
    Me.Controls("txtCDC" & eqNum).value = "1"
    Me.Controls("txtCDR" & eqNum).value = "1"
End Sub
''Private Sub ListView1_DblClick()
'    Dim ws As Worksheet
'    Dim selectedSr As Long
'    Dim foundCell As Range
'    Dim r As Long ' The row number in Excel
'    Me.lblLoadedID.Caption = selectedSr
'    Me.lblLoadedID.Caption = Me.ListView1.SelectedItem.text
'    ' 1. Check if a row is actually selected
'    If Me.ListView1.SelectedItem Is Nothing Then Exit Sub
'
'    ' 2. Get the Sr. No from the first column of the selected row
'    selectedSr = Val(Me.ListView1.SelectedItem.text)
'    Set ws = ThisWorkbook.Sheets("Database")
'
'    ' 3. Find that Sr. No in Column A of the Database
'    Set foundCell = ws.Columns(1).Find(What:=selectedSr, LookIn:=xlValues, LookAt:=xlWhole)
'
'    If Not foundCell Is Nothing Then
'        r = foundCell.Row
'        Me.Tag = r
'        Me.lblLoadedID.Caption = selectedSr
'        ' --- DISABLE EVENTS TO PREVENT CALCULATION OVERLOAD ---
'        IsSwapping = True
'
'        ' 4. Load General Data
'        Me.txtPVar.Value = ws.Cells(r, 2).Value
'        Me.txtSVar.Value = ws.Cells(r, 3).Value
'
'        ' --- LOAD EQUATION 1 ---
'        ' Positions
'        Me.txtPVTP1.Value = ws.Cells(r, 4).Value
'        Me.txtSVTP1.Value = ws.Cells(r, 5).Value
'        Me.txtETP1.Value = ws.Cells(r, 6).Value
'        Me.txtCTP1.Value = ws.Cells(r, 7).Value
'
'        ' Term 1
'        Me.tglSignFT1.Caption = ws.Cells(r, 8).Value
'        Me.txtANC1.Value = ws.Cells(r, 9).Value
'        Me.txtANR1.Value = ws.Cells(r, 10).Value
'        Me.txtADC1.Value = ws.Cells(r, 11).Value
'        Me.txtADR1.Value = ws.Cells(r, 12).Value
'
'        ' Term 2
'        Me.tglSignST1.Caption = ws.Cells(r, 13).Value
'        Me.txtBNC1.Value = ws.Cells(r, 14).Value
'        Me.txtBNR1.Value = ws.Cells(r, 15).Value
'        Me.txtBDC1.Value = ws.Cells(r, 16).Value
'        Me.txtBDR1.Value = ws.Cells(r, 17).Value
'
'        ' Term 3
'        Me.tglSignCT1.Caption = ws.Cells(r, 18).Value
'        Me.txtCNC1.Value = ws.Cells(r, 19).Value
'        Me.txtCNR1.Value = ws.Cells(r, 20).Value
'        Me.txtCDC1.Value = ws.Cells(r, 21).Value
'        Me.txtCDR1.Value = ws.Cells(r, 22).Value
'
'        ' --- LOAD EQUATION 2 ---
'        ' Positions
'        Me.txtPVTP2.Value = ws.Cells(r, 23).Value
'        Me.txtSVTP2.Value = ws.Cells(r, 24).Value
'        Me.txtETP2.Value = ws.Cells(r, 25).Value
'        Me.txtCTP2.Value = ws.Cells(r, 26).Value
'
'        ' Term 1
'        Me.tglSignFT2.Caption = ws.Cells(r, 27).Value
'        Me.txtANC2.Value = ws.Cells(r, 28).Value
'        Me.txtANR2.Value = ws.Cells(r, 29).Value
'        Me.txtADC2.Value = ws.Cells(r, 30).Value
'        Me.txtADR2.Value = ws.Cells(r, 31).Value
'
'        ' Term 2
'        Me.tglSignST2.Caption = ws.Cells(r, 32).Value
'        Me.txtBNC2.Value = ws.Cells(r, 33).Value
'        Me.txtBNR2.Value = ws.Cells(r, 34).Value
'        Me.txtBDC2.Value = ws.Cells(r, 35).Value
'        Me.txtBDR2.Value = ws.Cells(r, 36).Value
'
'        ' Term 3
'        Me.tglSignCT2.Caption = ws.Cells(r, 37).Value
'        Me.txtCNC2.Value = ws.Cells(r, 38).Value
'        Me.txtCNR2.Value = ws.Cells(r, 39).Value
'        Me.txtCDC2.Value = ws.Cells(r, 40).Value
'        Me.txtCDR2.Value = ws.Cells(r, 41).Value
'
'        ' 5. RE-ENABLE EVENTS & REFRESH UI
'        IsSwapping = False
'
'        ' Switch to Page 1
'        Me.MultiPage1.Value = 0
'
'        ' Show the frames
'        Me.fraVariable.Visible = True
'        Me.fraEq1.Visible = True
'        Me.fraEq2.Visible = True
'
'        ' Force UI Refresh
'        SyncAll 1
'        SyncAll 2
'        UpdateQuestionPreview Me
'        RefreshSolutionDisplay
'        MsgBox "Question No. " & selectedSr & " Loaded Successfully!", vbInformation
'    Else
'        MsgBox "Data not found in Database.", vbCritical
'    End If
'End Sub
Private Sub cmdUpdate_Click()

    Dim ws As Worksheet
    Dim foundCell As Range
    Dim r As Long
    Dim eq1 As String
    Dim eq2 As String
    Dim duplicateSr As Long
    
    Set ws = ThisWorkbook.Sheets("Database")
    
    On Error GoTo ErrorHandler
    
    ' ======================================
    ' ALLOW UPDATE ONLY IN EDIT MODE
    ' ======================================
    
    If CurrentMode <> ModeEdit Then Exit Sub
    
    ' ======================================
    ' VALIDATION
    ' ======================================
    
    If Me.lblLoadedID.Caption = "" Then
        MsgBox "Please select a record first.", vbExclamation
        Exit Sub
    End If
    
    ' ======================================
    ' FIND RECORD ROW
    ' ======================================
    
    Set foundCell = ws.Columns(1).Find(What:=Me.lblLoadedID.Caption, LookAt:=xlWhole)
    
    If foundCell Is Nothing Then
        MsgBox "Could not find record in Database.", vbExclamation
        Exit Sub
    End If
    
    r = foundCell.Row
    
    ' ======================================
    ' DUPLICATE CHECK (EXCLUDE CURRENT ROW)
    ' ======================================
    
    eq1 = GetSymbolicEq(Me, 1)
    eq2 = GetSymbolicEq(Me, 2)
    
    duplicateSr = FindDuplicateRecordExcluding(eq1, eq2, r)
    
    If duplicateSr > 0 Then
        MsgBox "Duplicate record found." & vbCrLf & _
               "Same question exists at Sr. No: " & duplicateSr, vbExclamation
        Exit Sub
    End If
    
    ' ======================================
    ' SAVE UPDATED RAW DATA
    ' ======================================
    
    SaveStandardizedData Me, r
    
    ' ======================================
    ' REGENERATE ALL METHOD SOLUTIONS
    ' ======================================
    
    GenerateAllMethodSolutions Me, r
    
    ' ======================================
    ' REFRESH UI
    ' ======================================
    
    UpdateListView
    RefreshSolutionDisplay
    
    UpdateStatus "Record updated successfully.", StatusSuccess
    
    ' ======================================
    ' RETURN TO IDLE MODE
    ' ======================================
    
    SetFormMode ModeIdle
    IsDirty = False
    
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & vbCrLf & _
           Err.Description & vbCrLf & _
           "Source: " & Err.Source, vbCritical

End Sub
Private Sub cmdDelete_Click()

    Dim ws As Worksheet
    Dim foundCell As Range
    Dim response As VbMsgBoxResult
    
    Set ws = ThisWorkbook.Sheets("Database")
    
    ' ======================================
    ' ALLOW DELETE ONLY IN EDIT MODE
    ' ======================================
    
    If CurrentMode <> ModeEdit Then Exit Sub
    
    If Me.lblLoadedID.Caption = "" Then
        MsgBox "Please load a question before deleting.", vbExclamation
        Exit Sub
    End If
    
    ' ======================================
    ' CONFIRMATION
    ' ======================================
    
    response = MsgBox("Are you sure you want to permanently delete Question No. " & _
                      Me.lblLoadedID.Caption & "?", _
                      vbQuestion + vbYesNo, "Confirm Delete")
    
    If response <> vbYes Then Exit Sub
    
    ' ======================================
    ' FIND RECORD
    ' ======================================
    
    Set foundCell = ws.Columns(1).Find(What:=Me.lblLoadedID.Caption, LookAt:=xlWhole)
    
    If foundCell Is Nothing Then
        UpdateStatus "Could not find the record in database.", StatusError
        Exit Sub
    End If
    
    ' ======================================
    ' DELETE ROW
    ' ======================================
    
    foundCell.EntireRow.Delete
    RecalculateQuestionNumbers
    ' ======================================
    ' REFRESH UI
    ' ======================================
    
    UpdateListView
    UpdateStatus "Record deleted successfully.", StatusSuccess
    
    ' ======================================
    ' RESET FORM
    ' ======================================
    
    Me.lblLoadedID.Caption = ""
    Me.Tag = 0
    
    SetDefaults 1
    SetDefaults 2
    
    Me.fraVariable.Visible = False
    Me.fraEq1.Visible = False
    Me.fraEq2.Visible = False
    
    SetFormMode ModeIdle

End Sub

' Add this helper function to your UserForm or a Module
Private Function EncodeBase64(Text As String) As String
    Dim arrData() As Byte: arrData = StrConv(Text, vbFromUnicode)
    Dim objXML As Object: Set objXML = CreateObject("MSXML2.DOMDocument")
    Dim objNode As Object: Set objNode = objXML.createElement("b64")
    
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = Replace(objNode.Text, vbLf, "")
End Function


Public Sub LogAllControlNames()
    Dim ctrl As Control
    Dim fso As Object
    Dim ts As Object
    Dim tempPath As String
    Dim logContent As String
    
    ' 1. Define the path in the Windows Temp folder
    tempPath = Environ("Temp") & "\UserForm_Controls_Log.txt"
    
    ' 2. Build the string of control names
    logContent = "--- Control Log for " & Me.Name & " ---" & vbCrLf
    logContent = logContent & "Generated on: " & Now & vbCrLf & String(40, "-") & vbCrLf
    
    For Each ctrl In Me.Controls
        logContent = logContent & "Name: " & ctrl.Name & " | Type: " & TypeName(ctrl) & vbCrLf
    Next ctrl
    
    ' 3. Write to the file using FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(tempPath, True)
    ts.Write logContent
    ts.Close
    
    ' 4. Open the file automatically for the user
    Shell "notepad.exe " & tempPath, vbNormalFocus
    
    MsgBox "Control names logged to: " & tempPath, vbInformation
End Sub
Public Function GetQuestionLatex(frm As Object) As String

    GetQuestionLatex = _
        GetStep1Latex(frm, 1) & " \; ; \; " & _
        GetStep1Latex(frm, 2)

End Function
Private Sub cmdCopyQuestion_Click()

    Dim qLatex As String
    qLatex = GetQuestionLatex(Me)
    
    If Trim(qLatex) = "" Then
        UpdateStatus "Question Blank.", StatusError
        Exit Sub
    End If
    
    Dim objHTML As Object
    Set objHTML = CreateObject("htmlfile")
    
    objHTML.parentWindow.ClipboardData.SetData "text", qLatex
    
    UpdateStatus "Question copied.", StatusSuccess

End Sub
Public Sub UpdateStatus(msg As String, StatusType As StatusType)

    With Me.lblStatusBar
        
        .Visible = True
        .Font.Bold = True
        .Caption = "   " & msg
        
        Select Case StatusType
            
            Case StatusSuccess
                .BackColor = RGB(220, 255, 220)
                .ForeColor = RGB(0, 120, 0)
                
            Case StatusError
                .BackColor = RGB(255, 220, 220)
                .ForeColor = RGB(180, 0, 0)
                
            Case StatusInfo
                .BackColor = RGB(240, 240, 240)
                .ForeColor = RGB(70, 70, 70)
                
        End Select
        
    End With
    
    ' Start 3 second timer
    StatusClearTime = Timer + 3
    StatusFading = False
    
    Me.Repaint
    
    StartStatusMonitor

End Sub
Private Sub StartStatusMonitor()

    Do While Timer < StatusClearTime
        DoEvents
    Loop
    
    FadeStatusBar

End Sub
Private Sub FadeStatusBar()

    Dim i As Integer
    
    For i = 0 To 10
        
        Me.lblStatusBar.ForeColor = RGB( _
            100 + i * 10, _
            100 + i * 10, _
            100 + i * 10)
            
        DoEvents
        
        Dim t As Double
        t = Timer + 0.05
        Do While Timer < t
            DoEvents
        Loop
        
    Next i
    
    Me.lblStatusBar.Caption = "   Ready"
    Me.lblStatusBar.BackColor = RGB(245, 245, 245)
    Me.lblStatusBar.ForeColor = RGB(70, 70, 70)

End Sub


Private Sub ApplyHoverEffect(btn As MSForms.CommandButton)

    btn.BackColor = RGB(60, 120, 220)

End Sub
Private Sub SetupControls()

    Dim ctrl As MSForms.Control
    Dim objItem As clsEquationControl
    
    Set ControlCollection = New Collection
    
    For Each ctrl In Me.Controls
        
        If TypeName(ctrl) = "TextBox" Or TypeName(ctrl) = "ToggleButton" Then
            
            If ctrl.Name Like "*1" Or ctrl.Name Like "*2" Then
            
                Set objItem = New clsEquationControl
                Set objItem.ParentForm = Me
                objItem.eqNum = val(right(ctrl.Name, 1))
                
                If TypeName(ctrl) = "TextBox" Then
                    Set objItem.txtGroup = ctrl
                ElseIf TypeName(ctrl) = "ToggleButton" Then
                    Set objItem.tglGroup = ctrl
                End If
                
                ControlCollection.Add objItem
                
            End If
            
        End If
        
    Next ctrl

End Sub
Private Sub SetupListView()

    With Me.ListView1
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .ColumnHeaders.Clear
        
        .ColumnHeaders.Add , , "Q. No", 60
        .ColumnHeaders.Add , , "Equation 1", 190
        .ColumnHeaders.Add , , "Equation 2", 190
    End With

    UpdateListView

End Sub
Private Sub SetupStatusBar()

    With Me.lblStatusBar
        .Top = Me.ListView1.Top + Me.ListView1.Height + 2
        .left = 0
        .Width = Me.InsideWidth
        .Height = 24
        .BackColor = RGB(245, 245, 245)
        .ForeColor = RGB(70, 70, 70)
        .Font.Bold = True
        .Caption = "  Ready"
        .Visible = True
    End With

End Sub
Private Sub ApplyButtonIcons()

    Me.cmdAdd.Caption = "Add"
    Me.cmdSave.Caption = "Save"
    Me.cmdUpdate.Caption = "Update"
    Me.cmdDelete.Caption = "Delete"
    Me.cmdReset.Caption = "Reset"
End Sub
Private Sub InitBrowser(br As Object)

    Dim html As String
    
    html = "<!-- saved from url=(0014)about:internet -->" & _
           "<html><head>" & _
           "<meta http-equiv='X-UA-Compatible' content='IE=edge'>" & _
           "<script src='https://cdnjs.cloudflare.com/ajax/libs/mathjax/2.7.7/MathJax.js?config=TeX-AMS_HTML'></script>" & _
           "<script type='text/x-mathjax-config'>" & _
           "MathJax.Hub.Config({ messageStyle: 'none' });" & _
           "</script>" & _
           "</head><body></body></html>"

    br.Navigate "about:blank"

    Do While br.ReadyState <> 4
        DoEvents
    Loop

    br.Document.Open
    br.Document.Write html
    br.Document.Close

End Sub
Private Sub SetActiveMethodButton(activeBtn As MSForms.CommandButton)

    ' Reset all method buttons first
    ResetMethodButtons
    
    ' Highlight active button
    With activeBtn
        .BackColor = RGB(0, 120, 215)      ' Modern blue
        .ForeColor = RGB(255, 255, 255)    ' White text
        .Font.Bold = True
    End With

End Sub
Private Sub ResetMethodButtons()

    With cmdSubstitution
        .BackColor = RGB(240, 240, 240)
        .ForeColor = RGB(0, 0, 0)
        .Font.Bold = False
    End With

    With cmdElimination
        .BackColor = RGB(240, 240, 240)
        .ForeColor = RGB(0, 0, 0)
        .Font.Bold = False
    End With

    With cmdGraphical
        .BackColor = RGB(240, 240, 240)
        .ForeColor = RGB(0, 0, 0)
        .Font.Bold = False
    End With

    With cmdCramer
        .BackColor = RGB(240, 240, 240)
        .ForeColor = RGB(0, 0, 0)
        .Font.Bold = False
    End With

End Sub
Private Sub cmdCramer_Click()

    Dim r As Long
    
    r = val(Me.Tag)
    
    If r <= 0 Then Exit Sub
    
    Dim rawLatex As String
    
    rawLatex = ThisWorkbook.Sheets("Database").Cells(r, 49).value   ' AW
    
    If Trim(rawLatex) = "" Then
        MsgBox "Cramer solution not generated.", vbInformation
        Exit Sub
    End If
    
    ' Render to Solution panel
    RenderMathJax rawLatex, Me, "Solution"
    
    ' Track active method column
    ActiveMethodColumn = 49
    
    ' Highlight button
    SetActiveMethodButton Me.cmdCramer

End Sub
Private Sub cmdSubstitution_Click()

    If val(Me.Tag) = 0 Then Exit Sub
    
    Dim r As Long
    r = val(Me.Tag)
    
    Dim rawLatex As String
    rawLatex = ThisWorkbook.Sheets("Database").Cells(r, 46).value
    
    If rawLatex <> "" Then
        RenderMathJax rawLatex, Me, "Solution"
        ActiveMethodColumn = 46
    End If
    
    SetActiveMethodButton cmdSubstitution

End Sub

Private Sub cmdElimination_Click()

     Dim r As Long
    
    r = val(Me.Tag)
    
    If r <= 0 Then Exit Sub
    
    Dim rawLatex As String
    
    rawLatex = ThisWorkbook.Sheets("Database").Cells(r, 47).value   ' AU
    
    If Trim(rawLatex) = "" Then
        MsgBox "Elimination solution not generated.", vbInformation
        Exit Sub
    End If
    'DoEvents
    'MsgBox rawLatex
    
    ' Render to Solution panel
    RenderMathJax rawLatex, Me, "Solution"
    
    ' Track active method column
    ActiveMethodColumn = 47
    
    ' Highlight button
    SetActiveMethodButton Me.cmdElimination

End Sub
Private Sub cmdGraphical_Click()

    If val(Me.Tag) = 0 Then Exit Sub
    
    Dim r As Long
    r = val(Me.Tag)
    
    Dim rawLatex As String
    rawLatex = ThisWorkbook.Sheets("Database").Cells(r, 48).value
    
    If rawLatex <> "" Then
        RenderMathJax rawLatex, Me, "Solution"
        ActiveMethodColumn = 48
    End If
    
    SetActiveMethodButton cmdGraphical

End Sub

