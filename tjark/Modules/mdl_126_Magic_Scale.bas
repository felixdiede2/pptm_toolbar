Attribute VB_Name = "mdl_126_Magic_Scale"
Option Explicit

Sub magic_scale()
    On Error GoTo 9
    
    If ActiveWindow.Selection.ShapeRange.Count < 1 Then
        GoTo 9
    Else
        'MsgBox (ActiveWindow.Selection.TextRange.Font.Size)
        usrScaleFactor.Show
    End If

    If 1 = 0 Then
9:      MsgBox ("Bitte eine Form zum Skalieren auswählen.")
    End If

End Sub

