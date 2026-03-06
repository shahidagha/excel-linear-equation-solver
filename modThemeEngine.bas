Attribute VB_Name = "modThemeEngine"
Option Explicit
Public IsDarkMode As Boolean
' ==============================
' MODERN LIGHT THEME ENGINE
' ==============================

Public Sub ApplyModernLightTheme(frm As Object)
    StyleContainer frm
End Sub



Private Sub StyleContainer(container As Object)

    Dim ctrl As Object

    For Each ctrl In container.Controls

        Select Case TypeName(ctrl)

            Case "Frame"
                On Error Resume Next
                ctrl.BackColor = RGB(255, 255, 255)
                ctrl.BorderStyle = fmBorderStyleSingle
                ctrl.SpecialEffect = fmSpecialEffectFlat
                On Error GoTo 0
                
                StyleContainer ctrl   ' recursion

            Case "Label"
                On Error Resume Next
                ctrl.ForeColor = RGB(50, 50, 50)
                ctrl.Font.Name = "Segoe UI"
                ctrl.Font.Size = 9
                On Error GoTo 0

            Case "TextBox"
                On Error Resume Next
                ctrl.BackColor = RGB(255, 255, 255)
                ctrl.ForeColor = RGB(0, 0, 0)
                ctrl.BorderStyle = fmBorderStyleSingle
                On Error GoTo 0

            Case "CommandButton"
                On Error Resume Next
                ctrl.BackColor = RGB(45, 95, 180)
                ctrl.ForeColor = RGB(255, 255, 255)
                ctrl.Font.Bold = True
                ctrl.SpecialEffect = fmSpecialEffectFlat
                On Error GoTo 0

            Case "ToggleButton"
                On Error Resume Next
                ctrl.BackColor = RGB(230, 230, 230)
                ctrl.ForeColor = RGB(0, 0, 0)
                On Error GoTo 0

            ' Ignore unsupported ActiveX controls
            Case "WebBrowser", "ListView", "ListView4"
                ' do nothing

        End Select

    Next ctrl

End Sub
