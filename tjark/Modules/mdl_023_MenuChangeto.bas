Attribute VB_Name = "mdl_023_MenuChangeto"
Option Explicit

' Prozeduren zum �ndern des Formats von Textfeldern
' -------------------------------------------------
' Ggf. pr�fen ob das angew�hlte Shape ein Rechteck ist
' CI-Eigenschaften (Farbe, Form, Schrift, Position) des Elements anpassen durch anwenden der Eigenschafts-Prozedur auf das Element
' Wenn Element nicht rechteckig oder kein Element angew�hlt, dann Fehler ausgeben

Sub OLD_OC_changeto_AT()
    On Error GoTo 4
    
    If ActiveWindow.Selection.ShapeRange.AutoShapeType <> msoShapeRectangle Then
        MsgBox ("Nur rechteckige Elemente k�nnen zum Actiontitle formatiert werden.")
    Else
        'If ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Text = "" Then
        'ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Text = " "
        Call OLD_prop_AT
    End If
        
    If 0 = 1 Then
4:      MsgBox ("Bitte ein Element zum Formatieren ausw�hlen.")
    End If

End Sub

Sub OLD_OC_changeto_ST()
    On Error GoTo 4
    
    If ActiveWindow.Selection.ShapeRange.AutoShapeType <> msoShapeRectangle Then
        MsgBox ("Nur rechteckige Elemente k�nnen zum Subtitle formatiert werden.")
    Else
        'If ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Text = "" Then
        'ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Text = " "
        Call OLD_prop_ST
    End If
    
    If 0 = 1 Then
4:      MsgBox ("Bitte ein Element zum Formatieren ausw�hlen.")
    End If

End Sub

Sub OC_changeto_header()
    On Error GoTo 4
    
    If (ActiveWindow.Selection.ShapeRange.AutoShapeType <> msoShapePentagon) And (ActiveWindow.Selection.ShapeRange.AutoShapeType <> msoShapeChevron) And (ActiveWindow.Selection.ShapeRange.AutoShapeType <> msoShapeRectangle) Then
        MsgBox ("Nur rechteckige Elemente und Blockpfeile k�nnen zum Zeilenkopf formatiert werden.")
    Else
        'If ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Text = "" Then
        'ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Text = " "
        
        Call prop_header
    End If
    
    If 0 = 1 Then
4:      MsgBox ("Bitte ein Element zum Formatieren ausw�hlen.")
    End If
    
End Sub

Sub OC_changeto_textbox()
    On Error GoTo 4
    
    If ActiveWindow.Selection.ShapeRange.AutoShapeType <> msoShapeRectangle Then
        MsgBox ("Nur rechteckige Elemente k�nnen zur Textbox formatiert werden.")
    Else
        'If ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Text = "" Then
        'ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Text = " "
        Call prop_textbox
    End If
    
    If 0 = 1 Then
4:      MsgBox ("Bitte ein Element zum Formatieren ausw�hlen.")
    End If
    
End Sub

Sub OC_changeto_greybox()
    On Error GoTo 4
    
    If ActiveWindow.Selection.ShapeRange.AutoShapeType <> msoShapeRectangle Then
        MsgBox ("Nur rechteckige Elemente k�nnen zur Graubox formatiert werden.")
    Else
        'If ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Text = "" Then
        'ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Text = " "
        Call prop_greybox
    End If
    
    If 0 = 1 Then
4:      MsgBox ("Bitte ein Element zum Formatieren ausw�hlen.")
    End If

End Sub

Sub OC_changeto_footnote()
    On Error GoTo 4
    
    If ActiveWindow.Selection.ShapeRange.AutoShapeType <> msoShapeRectangle Then
        MsgBox ("Nur rechteckige Elemente k�nnen zur Fusszeile formatiert werden.")
    Else
        'If ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Text = "" Then
        'ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Text = " "
        Call prop_footnote
    End If
    
    If 0 = 1 Then
4:      MsgBox ("Bitte ein Element zum Formatieren ausw�hlen.")
    End If
    
End Sub

Sub OC_changeto_graphicstext()
    On Error GoTo 4
    
    If ActiveWindow.Selection.ShapeRange.AutoShapeType <> msoShapeRectangle Then
        MsgBox ("Nur rechteckige Elemente k�nnen zum Grafiktext formatiert werden.")
    Else
        'If ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Text = "" Then
        'ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Text = " "
        Call prop_graphicstext
    End If
    
    If 0 = 1 Then
4:      MsgBox ("Bitte ein Element zum Formatieren ausw�hlen.")
    End If
    
End Sub

