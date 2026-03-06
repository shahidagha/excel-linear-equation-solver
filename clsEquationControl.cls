VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEquationControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit


Public ParentForm As Object
Public eqNum As Integer
Public WithEvents txtGroup As MSForms.TextBox
Attribute txtGroup.VB_VarHelpID = -1
Public WithEvents tglGroup As MSForms.ToggleButton
Attribute tglGroup.VB_VarHelpID = -1

' --- 1. CAPTURE OLD VALUE FOR SWAPPING ---
Private Sub txtGroup_Enter()

    ' 1?? Store OldVal for position textboxes (swap logic)
    If txtGroup.Name Like "*TP" & eqNum Then
        ParentForm.OldVal = txtGroup.value
    End If

    ' 2?? Auto-select entire text for fast editing
    txtGroup.SelStart = 0
    txtGroup.SelLength = Len(txtGroup.Text)

End Sub

Private Sub txtGroup_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If txtGroup.Name Like "*TP" & eqNum Then
        ParentForm.OldVal = txtGroup.value
    End If
End Sub

' --- 2. TRIGGER SWAP OR SYNC ---
Private Sub txtGroup_Change()
    If Not ParentForm.IsUpdating Then ParentForm.MarkDirty
    ' If the form is bulk-loading data, stop the Class from firing
    If ParentForm.IsUpdating Then Exit Sub
    
    ' If a Position box (TP) changes, handle the swap logic
    If txtGroup.Name Like "*TP" & eqNum Then
        ParentForm.HandleSwap txtGroup
    Else
        ' If a coefficient (ANC, BNC, etc.) changes, update the LaTeX
        ParentForm.SyncAll eqNum
    End If
End Sub

' --- 3. INPUT RESTRICTION ---
Private Sub txtGroup_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Only allow numbers 1-4 in position (TP) boxes
    If txtGroup.Name Like "*TP" & eqNum Then
        ' Allow 1-4 (ASCII 49-52) and Backspace (ASCII 8)
        If Not (KeyAscii >= 49 And KeyAscii <= 52 Or KeyAscii = 8) Then
            KeyAscii = 0
            Beep
        End If
    End If
End Sub

' --- 4. SIGN TOGGLES (+/-) ---
Private Sub tglGroup_Click()
    ParentForm.MarkDirty
    If left(tglGroup.Name, 7) = "tglSign" Then
        ' Toggle the caption
        tglGroup.Caption = IIf(tglGroup.Caption = "+", "-", "+")
        
        ' Get the term key (FT, ST, or CT)
        Dim termKey As String: termKey = Mid(tglGroup.Name, 8, 2)
        Dim realPrefix As String
        
        ' MAP the keys to your EXACT log names
        Select Case termKey
            Case "FT": realPrefix = "PVT" ' Maps FT to txtPVTP
            Case "ST": realPrefix = "SVT" ' Maps ST to txtSVTP
            Case "CT": realPrefix = "CT"  ' Maps CT to txtCTP
        End Select
        
        ' FIX: Constructing the name to match "txtPVTP1", "txtSVTP1", "txtCTP1"
        Dim posControlName As String: posControlName = "txt" & realPrefix & "P" & eqNum
        
        ' Safely get values from ParentForm
        On Error Resume Next ' Guard against missing controls during initialization
        Dim pos As Integer: pos = val(ParentForm.Controls(posControlName).value)
        Dim eqPos As Integer: eqPos = val(ParentForm.Controls("txtETP" & eqNum).value)
        
        ' Apply visibility logic (hides leading + signs)
        ' Ensure this sub exists in your Module1
       HandleSignVisibility tglGroup, pos, eqPos
        On Error GoTo 0
    End If
    
    ' Update the LaTeX display
    If ParentForm Is Nothing Then
        MsgBox "ParentForm is NOTHING"
        Exit Sub
    End If
    ParentForm.SyncAll eqNum
End Sub

Private Sub txtGroup_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                              ByVal Shift As Integer)

    If KeyCode = vbKeyReturn Then
        
        KeyCode = 0
        
        ParentForm.SelectNextControl txtGroup, True, True, True, True
        
    End If

End Sub

