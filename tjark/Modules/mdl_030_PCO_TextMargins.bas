Attribute VB_Name = "mdl_030_PCO_TextMargins"
'
' PowerPoint 2010 VBA Macro
' Porsche CO
' Copyright PaCE Graphic GbR
' September 2015
'

Sub DeleteMargins()

Dim Selection As Object
Dim NoTextElement As Boolean
    
ErrorExit = False
NoTextElement = False

CheckErrorsSlideSelection
If ErrorExit = True Then
    Exit Sub
End If
CheckErrorViewWrong

    If ActiveWindow.Selection.Type = ppSelectionNone Then
        MsgBox "At least 1 object with text must be selected for this tool. Please, select one or more objects with text and restart tool.", _
                vbInformation, "No selection!"
        ErrorExit = True
        Exit Sub
    End If
    
    Set Selection = ActiveWindow.Selection
    With Selection
        For i = 1 To .ShapeRange.Count
            If Selection.ShapeRange(i).HasTextFrame = msoTrue Then
                If Selection.ShapeRange(i).TextFrame.TextRange.Text <> "" Then
                    Selection.ShapeRange(i).TextFrame.MarginTop = 0
                    Selection.ShapeRange(i).TextFrame.MarginBottom = 0
                    Selection.ShapeRange(i).TextFrame.MarginLeft = 0
                    Selection.ShapeRange(i).TextFrame.MarginRight = 0
                Else
                    NoTextElement = True
                End If
            Else
                NoTextElement = True
            End If
        Next
    End With
    
    If NoTextElement = True Then
        MsgBox "One or more of the selected objects didn't include text. Only the objects with text were be corrected from this tool.", _
                vbInformation, "Wrong selection!"
    End If

End Sub


