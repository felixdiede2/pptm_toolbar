Attribute VB_Name = "mdl_121_Magic_Chevrons"
Option Explicit

Private Const My_Pi As Double = 3.14159265358979
Private Const CAF As Single = 0.18 'Chevron Adjustment Factor


Public Sub Chevron_Magic_Selection()
    On Error Resume Next
    
    Chevron_Magic ActiveWindow.Selection.ShapeRange
End Sub


Public Sub Chevron_Magic(Elemente As ShapeRange)

    On Error Resume Next
    
    Dim AF As Single
    
    If Elemente.Count = 1 Then
        Select Case Elemente(1).AutoShapeType
            Case msoShapeChevron
                AF = Elemente(1).Adjustments.Item(1)
                Elemente(1).AutoShapeType = msoShapePentagon
                Elemente(1).Adjustments.Item(1) = AF
            Case msoShapePentagon
                Chevron_To_Rectangle Elemente
            Case Else
                Rectangle_To_Chevron Elemente
        End Select
    Else
        Select Case Elemente(1).AutoShapeType
            Case msoShapeChevron
                Chevron_To_Rectangle Elemente
            Case msoShapePentagon
                Chevron_To_Rectangle Elemente
            Case Else
                Rectangle_To_Chevron Elemente
        End Select
    End If
    
End Sub



Public Sub Rectangle_To_Chevron_Selection()

    Chevron_To_Rectangle ActiveWindow.Selection.ShapeRange
    Rectangle_To_Chevron ActiveWindow.Selection.ShapeRange

End Sub

Public Sub Rectangle_To_Chevron(Elemente As ShapeRange)
    On Error Resume Next

    Dim i, n, L As Integer
    Dim links As Single
           
    links = 100 * ActiveWindow.Presentation.PageSetup.SlideWidth

    n = Elemente.Count
    
    If n = 1 Then
        L = 2
    Else
        For i = 1 To n
            If Elemente(i).Left < links Then
                L = i
                links = Elemente(i).Left
            End If
        Next i
    End If
    
    For i = 1 To n
        If Elemente(i).AutoShapeType = msoShapeRectangle Then
            If i = L Then
                Elemente(i).AutoShapeType = msoShapePentagon
                Elemente(i).Width = Elemente(i).Width + CAF * Elemente(i).Height
                Elemente(i).Adjustments.Item(1) = CAF
            Else
                Elemente(i).AutoShapeType = msoShapeChevron
                'Elemente(i).Left = Elemente(i).Left - 0.5 * CAF * Elemente(i).Height
                Elemente(i).Width = Elemente(i).Width + CAF * Elemente(i).Height
                Elemente(i).Adjustments.Item(1) = CAF
            End If
        End If
    Next i
    
End Sub

Public Sub Chevron_To_Rectangle_Selection()

    Chevron_To_Rectangle ActiveWindow.Selection.ShapeRange

End Sub


Public Sub Chevron_To_Rectangle(Elemente As ShapeRange)
    On Error Resume Next

    Dim i As Integer
    Dim foo As Single
        
    For i = 1 To Elemente.Count
        Select Case Elemente(i).AutoShapeType
            Case msoShapeChevron
                Elemente(i).Width = Elemente(i).Width - (Elemente(i).Height * Elemente(i).Adjustments.Item(1))
                Elemente(i).AutoShapeType = msoShapeRectangle
            Case msoShapePentagon
                Elemente(i).Width = Elemente(i).Width - (Elemente(i).Height * Elemente(i).Adjustments.Item(1))
                Elemente(i).AutoShapeType = msoShapeRectangle
            Case Else
                foo = 1
        End Select
    Next i
End Sub

