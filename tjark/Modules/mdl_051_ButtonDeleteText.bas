Attribute VB_Name = "mdl_051_ButtonDeleteText"
Option Explicit

' Prozedur zum Löschen des Textes in allen aktuell angewählten Shapes
' -------------------------------------------------------------------

Sub delete_text()
    On Error Resume Next
    
    Dim S As Shape
    Dim i As Integer
   
    For Each S In ActiveWindow.Selection.ShapeRange
        If S.HasTextFrame Then
            S.TextFrame.TextRange.Text = ""
        End If
        If S.Type = msoGroup Then
            For i = 1 To S.GroupItems.Count
                If S.GroupItems(i).HasTextFrame Then
                    S.GroupItems(i).TextFrame.TextRange.Text = ""
                End If
            Next i
        End If
    Next S
End Sub
