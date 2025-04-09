Attribute VB_Name = "mdl_031_ButtonChevron"
Option Explicit

Public Sub chevron_angle()
    On Error GoTo 1
    
    Dim objChevronShape As Shape
    Dim CAF As Single 'Chevron Adjustment Factor
    Dim AAF1 As Single 'Arrow Adjustment Factor 1
    Dim AAF2 As Single 'Arrow Adjustment Factor 1
    
    CAF = 0.36
    AAF1 = 0.5
    AAF2 = 0.58

    For Each objChevronShape In ActiveWindow.Selection.ShapeRange
        
        'Chevron, Fünfeck
        If (objChevronShape.AutoShapeType = msoShapePentagon) Or (objChevronShape.AutoShapeType = msoShapeChevron) Then
            objChevronShape.Adjustments.Item(1) = CAF
        End If
        
        If (objChevronShape.AutoShapeType = msoShapeParallelogram) Then
            objChevronShape.Adjustments.Item(1) = 2 * CAF
        End If
        
        'Waagerechte Pfeile
        If (objChevronShape.AutoShapeType = msoShapeRightArrow) Then
            objChevronShape.Adjustments.Item(1) = AAF1
            objChevronShape.Adjustments.Item(2) = AAF2
        End If
        If (objChevronShape.AutoShapeType = msoShapeLeftArrow) Or (objChevronShape.AutoShapeType = msoShapeLeftRightArrow) Then
            objChevronShape.Adjustments.Item(1) = AAF1
            objChevronShape.Adjustments.Item(2) = AAF2
        End If
        
        'Senkrechte Pfeile
        If (objChevronShape.AutoShapeType = msoShapeUpArrow) Or (objChevronShape.AutoShapeType = msoShapeUpDownArrow) Then
            objChevronShape.Adjustments.Item(1) = AAF1
            objChevronShape.Adjustments.Item(2) = AAF2
        End If
        If (objChevronShape.AutoShapeType = msoShapeDownArrow) Then
            objChevronShape.Adjustments.Item(1) = AAF1
            objChevronShape.Adjustments.Item(2) = AAF2
        End If
        
    Next
    
    If 0 = 1 Then
1:      MsgBox ("Bitte mindestens ein Blockpfeil- oder Pfeil-Element auswählen.")
    End If
        
End Sub


Public Sub chevron_angle_show()
    On Error GoTo 1
    
    Dim objChevronShape As Shape

    For Each objChevronShape In ActiveWindow.Selection.ShapeRange
        
        'Chevron, Fünfeck
        If (objChevronShape.AutoShapeType = msoShapePentagon) Or (objChevronShape.AutoShapeType = msoShapeChevron) Then
            objChevronShape.TextFrame.TextRange.Text = objChevronShape.Adjustments.Item(1)
        End If
        
        If (objChevronShape.AutoShapeType = msoShapeParallelogram) Then
            objChevronShape.TextFrame.TextRange.Text = objChevronShape.Adjustments.Item(1)
        End If
        
        'Waagerechte Pfeile
        If (objChevronShape.AutoShapeType = msoShapeRightArrow) Then
            objChevronShape.TextFrame.TextRange.Text = objChevronShape.Adjustments.Item(1) & "  " & objChevronShape.Adjustments.Item(2)
        End If
        If (objChevronShape.AutoShapeType = msoShapeLeftArrow) Or (objChevronShape.AutoShapeType = msoShapeLeftRightArrow) Then
            objChevronShape.TextFrame.TextRange.Text = objChevronShape.Adjustments.Item(1) & "  " & objChevronShape.Adjustments.Item(2)
        End If
        
        'Senkrechte Pfeile
        If (objChevronShape.AutoShapeType = msoShapeUpArrow) Or (objChevronShape.AutoShapeType = msoShapeUpDownArrow) Then
            objChevronShape.TextFrame.TextRange.Text = objChevronShape.Adjustments.Item(1) & "  " & objChevronShape.Adjustments.Item(2)
        End If
        If (objChevronShape.AutoShapeType = msoShapeDownArrow) Then
            objChevronShape.TextFrame.TextRange.Text = objChevronShape.Adjustments.Item(1) & "  " & objChevronShape.Adjustments.Item(2)
        End If
        
    Next
    
    If 0 = 1 Then
1:      MsgBox ("Bitte mindestens ein Blockpfeil- oder Pfeil-Element auswählen.")
    End If
        
End Sub
