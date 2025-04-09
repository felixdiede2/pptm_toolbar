Attribute VB_Name = "mdl_082_AlignToPage"
Option Explicit

Sub AusrichtenAbsLinks()

    On Error Resume Next
        
    ActiveWindow.Selection.ShapeRange.Align msoAlignLefts, msoTrue
    
End Sub

Sub AusrichtenAbsHorMitte()

    On Error Resume Next
    
    ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, msoTrue
    
End Sub

Sub AusrichtenAbsRechts()

    On Error Resume Next
        
    ActiveWindow.Selection.ShapeRange.Align msoAlignRights, msoTrue
    
End Sub

Sub AusrichtenAbsOben()

    On Error Resume Next
        
    ActiveWindow.Selection.ShapeRange.Align msoAlignTops, msoTrue
    
End Sub

Sub AusrichtenAbsVerMitte()

    On Error Resume Next
    
    ActiveWindow.Selection.ShapeRange.Align msoAlignMiddles, msoTrue
    
End Sub

Sub AusrichtenAbsUnten()

    On Error Resume Next
        
    ActiveWindow.Selection.ShapeRange.Align msoAlignBottoms, msoTrue
    
End Sub

Sub VerteilenAbsHor()

    On Error Resume Next
        
    ActiveWindow.Selection.ShapeRange.Distribute msoDistributeHorizontally, msoTrue
    
End Sub

Sub VerteilenAbsVer()

    On Error Resume Next
        
    ActiveWindow.Selection.ShapeRange.Distribute msoDistributeVertically, msoTrue
    
End Sub
