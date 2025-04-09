Attribute VB_Name = "mdl_999_cm2pt"
Option Explicit

Public Function cm2pt_x(cm_x As Single) As Single
    On Error Resume Next
    
    cm2pt_x = 0.5 * ActiveWindow.Presentation.PageSetup.SlideWidth + cm_x * cm2pt
End Function

Public Function cm2pt_y(cm_y As Single) As Single
    On Error Resume Next
    
    cm2pt_y = 0.5 * ActiveWindow.Presentation.PageSetup.SlideHeight - cm_y * cm2pt
End Function

Public Function mm2pt_x(mm_x As Single) As Single
    On Error Resume Next
    
    mm2pt_x = 0.5 * ActiveWindow.Presentation.PageSetup.SlideWidth + mm_x * mm2pt
End Function

Public Function mm2pt_y(mm_y As Single) As Single
    On Error Resume Next
    
    mm2pt_y = 0.5 * ActiveWindow.Presentation.PageSetup.SlideHeight - mm_y * mm2pt
End Function

Public Sub PrintShapeSize()
    On Error Resume Next
    
    Dim S As Shape
    Dim T As String
    Dim foo As Single
    
        
    For Each S In ActiveWindow.Selection.ShapeRange
    
        foo = Round((S.Left - 0.5 * ActiveWindow.Presentation.PageSetup.SlideWidth) * pt2cm, 2)
        T = "left:" & (foo)
        
        foo = Round(S.Width * pt2cm, 2)
        T = T & "  width:" & foo
        
        foo = Round((S.Top - 0.5 * ActiveWindow.Presentation.PageSetup.SlideHeight) * pt2cm, 2)
        T = T & "  top:" & foo
        
        foo = Round(S.Height * pt2cm, 2)
        T = T & "  height:" & foo
        
        S.TextFrame.TextRange.Text = T
        S.TextFrame.TextRange.Font.Size = 8
        
    Next S

End Sub

