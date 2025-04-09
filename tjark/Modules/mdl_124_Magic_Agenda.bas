Attribute VB_Name = "mdl_124_Magic_Agenda"
Option Explicit

Public Sub Agenda_Magic_Selection()
    On Error Resume Next
    
    Agenda_Magic ActiveWindow.Selection.ShapeRange
End Sub


Public Sub Agenda_Magic(Elemente As ShapeRange)

    On Error Resume Next
    
    'Dim S As Shape
    'Dim P As Variant
    Dim ebene As Variant
    Dim agenda As Variant
    Dim oben As Single
    Dim abst As Single
    Dim hoch As Single
    Dim objNewShape As Shape
    Dim a, i As Integer
       
    oben = cm2pt_y(4.6)
    abst = 1.5
    hoch = 1
    
    'MsgBox (Elemente.Count)
            
    If Elemente.Count > 0 And Elemente(1).HasTextFrame Then
    
        a = Elemente(1).TextFrame.TextRange.Paragraphs.Count
        
        If a > 12 Then
            MsgBox ("Es werden nur die ersten 12 Agendapunkte verarbeitet.")
            a = 12
        End If
        
        If a > 9 Then
            abst = 1.2
        End If
        
        If a > 11 Then
            hoch = 0.8
            abst = 1
        End If
        
        For i = 1 To a
        
            'P = Elemente(1).TextFrame.TextRange.Paragraphs(i)
            ebene = Elemente(1).TextFrame.TextRange.Paragraphs(i).indentLevel
            agenda = Elemente(1).TextFrame.TextRange.Paragraphs(i).Text
            
            If Asc(Right(agenda, 1)) = 13 Then
                agenda = Left(agenda, Len(agenda) - 1)
            End If
                                            
            Set objNewShape = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeRectangle, 0, 0, 100, 100)
            objNewShape.TextFrame.TextRange.Text = agenda
            objNewShape.Select
            
            If ebene = 1 Then
                Call prop_agenda1
            Else
                Call prop_agenda2
            End If
            
            objNewShape.Top = oben
            objNewShape.Height = hoch * cm2pt
            
            oben = oben + (abst * cm2pt)
            
        Next i
            
    End If
        
    Elemente.Select

End Sub

