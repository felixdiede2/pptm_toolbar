Attribute VB_Name = "mdl_113_Colors_Marke"

'Farben OTTO Marke (OM)

Option Explicit

'Primaerfarbe
Public Const color_OM_rot As Long = 1966290 'RGB(210, 0, 30)
Public Const color_OM_weiss As Long = 16777215  'RGB(255, 255, 255)
Public Const color_OM_schwarz As Long = 0  'RGB(0, 0, 0)

'Sekundaerfarben
Public Const color_OM_grau1 As Long = 12106432 'RGB(192, 186, 184)
Public Const color_OM_grau2 As Long = 7764358 'RGB(134, 121, 118)
Public Const color_OM_blau As Long = 13020235 'RGB(75, 172, 198)
Public Const color_OM_orange As Long = 4626167 'RGB(247, 150, 70)
Public Const color_OM_dunkelrot As Long = 1376402 'RGB(146, 0, 21)




'debug.print RGB(,,)
'Private Const color_OC_ = RGB(,,)

Sub OM_colorscheme_set()
    On Error Resume Next
    
    Dim intMasterCounter As Single
    
    While ActivePresentation.ColorSchemes.Count < 3
        ActivePresentation.ColorSchemes.Add
    Wend
    
    With ActivePresentation.ColorSchemes(3)
        .Colors(ppBackground) = color_OM_weiss
        .Colors(ppForeground) = color_OM_schwarz
        .Colors(ppShadow) = color_OM_schwarz
        .Colors(ppTitle) = color_OM_rot
        .Colors(ppFill) = color_OM_grau1
        .Colors(ppAccent1) = color_OM_grau1
        .Colors(ppAccent2) = color_OM_rot
        .Colors(ppAccent3) = color_OM_grau2
    End With
    
    For intMasterCounter = 1 To ActivePresentation.Designs.Count
        ActivePresentation.Designs(intMasterCounter).SlideMaster.ColorScheme = ActivePresentation.ColorSchemes(3)
    Next intMasterCounter

End Sub

Public Sub OM_extracolors()
    ActivePresentation.ExtraColors.Add color_OM_rot
    ActivePresentation.ExtraColors.Add color_OM_grau1
    ActivePresentation.ExtraColors.Add color_OM_grau2
    ActivePresentation.ExtraColors.Add color_OM_blau
    ActivePresentation.ExtraColors.Add color_OM_orange
    ActivePresentation.ExtraColors.Add color_OM_dunkelrot
    ActivePresentation.ExtraColors.Add color_OM_schwarz
    ActivePresentation.ExtraColors.Add color_OM_weiss
End Sub

Public Sub OM_colors_rot()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OM_rot
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OM_rot
        
        .TextFrame.TextRange.Font.Color.RGB = color_OM_weiss
    End With
    
End Sub

Public Sub OM_colors_dunkelrot()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OM_dunkelrot
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OM_dunkelrot
        
        .TextFrame.TextRange.Font.Color.RGB = color_OM_weiss
    End With
    
End Sub

Public Sub OM_colors_grau1()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OM_grau1
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OM_grau1
        
        .TextFrame.TextRange.Font.Color.RGB = color_OM_schwarz
    End With
    
End Sub

Public Sub OM_colors_grau2()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OM_grau2
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OM_grau2
        
        .TextFrame.TextRange.Font.Color.RGB = color_OM_weiss
    End With
    
End Sub

Public Sub OM_colors_blau()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OM_blau
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OM_blau
        
        .TextFrame.TextRange.Font.Color.RGB = color_OM_weiss
    End With
    
End Sub

Public Sub OM_colors_orange()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OM_orange
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OM_orange
        
        .TextFrame.TextRange.Font.Color.RGB = color_OM_schwarz
    End With
    
End Sub

