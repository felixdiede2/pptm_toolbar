Attribute VB_Name = "mdl_071_FitSize"
Option Explicit

Public Sub fit_size()
    On Error GoTo 1
    Dim sngShapeCount As Single

    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        GoTo 1
    Else
    For sngShapeCount = 1 To ActiveWindow.Selection.ShapeRange.Count
        With ActiveWindow.Selection
            .ShapeRange(sngShapeCount).LockAspectRatio = msoFalse
            .ShapeRange(sngShapeCount).Height = .ShapeRange(1).Height
            .ShapeRange(sngShapeCount).Width = .ShapeRange(1).Width
        End With
    Next sngShapeCount
    End If
    
    If 0 = 1 Then
1:      MsgBox ("Bitte mindestens zwei Objekte auswählen. Alle Objekte werden der Größe des erstgewählten Objekts angepasst.")
    End If

End Sub

Public Sub fit_width()
    On Error GoTo 1
    Dim sngShapeCount As Single
    
    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        GoTo 1
    Else
    For sngShapeCount = 1 To ActiveWindow.Selection.ShapeRange.Count
        With ActiveWindow.Selection
            .ShapeRange(sngShapeCount).LockAspectRatio = msoFalse
            .ShapeRange(sngShapeCount).Width = .ShapeRange(1).Width
        End With
    Next sngShapeCount
    End If
    
    If 0 = 1 Then
1:      MsgBox ("Bitte mindestens zwei Objekte auswählen. Alle Objekte werden der Breite des erstgewählten Objekts angepasst.")
    End If

End Sub

Public Sub fit_height()
    On Error GoTo 1
    Dim sngShapeCount As Single
    
    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        GoTo 1
    Else
        For sngShapeCount = 1 To ActiveWindow.Selection.ShapeRange.Count
            With ActiveWindow.Selection
            .ShapeRange(sngShapeCount).LockAspectRatio = msoFalse
            .ShapeRange(sngShapeCount).Height = .ShapeRange(1).Height
            End With
        Next sngShapeCount
    End If
    
    If 0 = 1 Then
1:      MsgBox ("Bitte mindestens zwei Objekte auswählen. Alle Objekte werden der Höhe des erstgewählten Objekts angepasst.")
    End If

End Sub
