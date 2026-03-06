Attribute VB_Name = "modUIEffects"
Option Explicit

Public Sub ApplyButtonEffects(frm As Object)

    Dim ctrl As Control
    
    For Each ctrl In frm.Controls
        
        If TypeName(ctrl) = "CommandButton" Then
            
            ctrl.SpecialEffect = fmSpecialEffectFlat
            ctrl.BorderStyle = fmBorderStyleSingle
            ctrl.BorderColor = RGB(200, 200, 200)
            
        End If
        
    Next ctrl

End Sub
