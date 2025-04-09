Attribute VB_Name = "mdl_120_FitIn"
Option Explicit

Public Sub Fit_In()
    On Error GoTo 1
    
    If ActiveWindow.Selection.ShapeRange.Count > 2 Then
        GoTo 1
    End If
    
    If ActiveWindow.Selection.ShapeRange.Count < 1 Then
        GoTo 1
    End If
    
    If ActiveWindow.Selection.ShapeRange.Count = 1 Then
        With ActiveWindow.Selection.ShapeRange(1)
            .LockAspectRatio = msoTrue
            .Left = 32.75
            .Top = 105.62
            .Width = 714.5
            If .Height > 385.5 Then
            .Height = 385.5
            End If
            .ZOrder msoBringToFront
        End With
    End If
    
    If ActiveWindow.Selection.ShapeRange.Count = 2 Then
        With ActiveWindow.Selection
           .ShapeRange(1).LockAspectRatio = msoTrue
           .ShapeRange(1).Height = .ShapeRange(2).Height - 20
           
           If .ShapeRange(1).Width > (.ShapeRange(2).Width - 20) Then
                .ShapeRange(1).Width = .ShapeRange(2).Width - 20
           End If
           
           .ShapeRange(1).Left = .ShapeRange(2).Left + 5
           .ShapeRange(1).Top = .ShapeRange(2).Top + 5
           .ShapeRange.Align msoAlignCenters, msoFalse
           .ShapeRange.Align msoAlignMiddles, msoFalse
           .ShapeRange(1).ZOrder msoBringToFront
        End With
    End If
    
    If 0 = 1 Then
1:      MsgBox ("Bitte zwei Objekte auswählen, von denen das Erste in das Zweite eingepasst werden soll.")
    
    End If
End Sub

