VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usrTabCreate 
   Caption         =   "Bitte eingeben..."
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2505
   OleObjectBlob   =   "usrTabCreate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usrTabCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




' Prozedur zum Erstellen einer Tabelle aus einem Rechteck
' -------------------------------------------------------


Private Sub CommandButton1_Click()

On Error GoTo 7

Dim sTabOverallHeight, sTabOverallWidth, sTabObjectHeight, sTabObjectWidth, sLeftAnchor, sTopAnchor As Single
Dim sSpaceBetweenObjects, sRows, sColumns As Single
Dim sLeftPos, sUpPos As Single

    If (usrTabCreate.sRows.Value > 0) And (usrTabCreate.sRows.Value < 16) Then
        sRows = usrTabCreate.sRows.Value
    Else
        MsgBox ("Nur Werte zwischen 1 und 15 zulässig.")
        GoTo 8
    End If
    If (usrTabCreate.sColumns.Value > 0) And (usrTabCreate.sColumns.Value < 16) Then
        sColumns = usrTabCreate.sColumns.Value
    Else
        MsgBox ("Nur Werte zwischen 1 und 15 zulässig.")
        GoTo 8
    End If


    If ActiveWindow.Selection.ShapeRange.AutoShapeType = msoShapeRectangle Then    ' Prüfung ob exakt ein Rechteck markiert ist
    If ActiveWindow.Selection.ShapeRange.Count = 1 Then
        sSpaceBetweenObjects = 5.66
        sTabOverallHeight = ActiveWindow.Selection.ShapeRange.Height
        sTabOverallWidth = ActiveWindow.Selection.ShapeRange.Width
        sLeftAnchor = ActiveWindow.Selection.ShapeRange.Left
        sTopAnchor = ActiveWindow.Selection.ShapeRange.Top
        ActiveWindow.Selection.ShapeRange.Delete
                
        sTabObjectWidth = (sTabOverallWidth - ((sRows - 1) * sSpaceBetweenObjects)) / sRows
        sTabObjectHeight = (sTabOverallHeight - ((sColumns - 1) * sSpaceBetweenObjects)) / sColumns
        
        For sLeftPos = 0 To (sRows - 1)
            For sUpPos = 0 To (sColumns - 1)
                    Set mydocument = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeRectangle, (sLeftAnchor + (sLeftPos * (sTabObjectWidth + sSpaceBetweenObjects))), (sTopAnchor + (sUpPos * (sTabObjectHeight + sSpaceBetweenObjects))), sTabObjectWidth, sTabObjectHeight)
                    mydocument.Select
                    ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Text = "Text"
                    Call prop_textbox
            Next sUpPos
        Next sLeftPos
        
        
    Else
        MsgBox ("Bitte nur ein Rechteck auswählen.")
    End If
    Else
        MsgBox ("Bitte nur ein Rechteck auswählen.")
    End If
   
   
If 1 = 0 Then
7: MsgBox ("Bitte ein Rechteck mit den Abmessungen " & Chr(10) & Chr(10) & _
            "der zu erstellenden Tabelle auswählen.")
End If
8:
Unload Me

End Sub

Private Sub UserForm_Initialize()

usrTabCreate.sColumns.Value = 2
usrTabCreate.sRows.Value = 2

End Sub
