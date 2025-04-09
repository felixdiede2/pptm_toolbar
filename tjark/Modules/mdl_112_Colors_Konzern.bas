Attribute VB_Name = "mdl_112_Colors_Konzern"

'Farben Otto Group (OG)

Option Explicit

'Primaerfarben
Public Const color_OG_rot As Long = 2241506 'RGB(226,51,34)
Public Const color_OG_weiss  As Long = 16777215  'RGB(255, 255, 255)
Public Const color_OG_schwarz  As Long = 0  'RGB(0, 0, 0)

'Sekundaerfarben
Public Const color_OG_dunkelrot As Long = 2167222  'RGB(182, 17, 33) 11931937
Public Const color_OG_mittelrot As Long = 2495944  'RGB(200, 21, 38)
Public Const color_OG_blau As Long = 13798656  'RGB(0, 141, 210)
Public Const color_OG_dunkelblau As Long = 8608000  'RGB(0, 89, 131)
Public Const color_OG_hellblau  As Long = 16374941  'RGB(157, 220, 249)

Public Const color_OG_grau1 = 14342874  'RGB(218, 218, 218)
Public Const color_OG_grau2 = 12434877  'RGB(189, 189, 189)
Public Const color_OG_grau3 = 8947848  'RGB(136, 136, 136)
Public Const color_OG_grau4 = 6579300  'RGB(100, 100, 100)


'debug.print RGB(,,)
'Public Const color_OC_ = RGB(,,)

Sub OG_colorscheme_set()
    On Error Resume Next
    
    Dim intMasterCounter As Single
    
    While ActivePresentation.ColorSchemes.Count < 2
        ActivePresentation.ColorSchemes.Add
    Wend
    
    With ActivePresentation.ColorSchemes(2)
        .Colors(ppBackground) = color_OG_weiss
        .Colors(ppForeground) = color_OG_schwarz
        .Colors(ppShadow) = color_OG_schwarz
        .Colors(ppTitle) = color_OG_schwarz
        .Colors(ppFill) = color_OG_grau1
        .Colors(ppAccent1) = color_OG_rot
        .Colors(ppAccent2) = color_OG_blau
        .Colors(ppAccent3) = color_OG_grau1
    End With
    
    For intMasterCounter = 1 To ActivePresentation.Designs.Count
        ActivePresentation.Designs(intMasterCounter).SlideMaster.ColorScheme = ActivePresentation.ColorSchemes(2)
    Next intMasterCounter

End Sub

Public Sub OG_extracolors()
    ActivePresentation.ExtraColors.Add color_OG_dunkelblau
    ActivePresentation.ExtraColors.Add color_OG_hellblau
    ActivePresentation.ExtraColors.Add color_OG_rot
    ActivePresentation.ExtraColors.Add color_OG_dunkelrot
    ActivePresentation.ExtraColors.Add color_OG_mittelrot
    ActivePresentation.ExtraColors.Add color_OG_schwarz
    ActivePresentation.ExtraColors.Add color_OG_weiss
    ActivePresentation.ExtraColors.Add color_OG_grau1
    ActivePresentation.ExtraColors.Add color_OG_grau2
    ActivePresentation.ExtraColors.Add color_OG_grau3
    ActivePresentation.ExtraColors.Add color_OG_grau4
End Sub


Public Sub OG_colors_rot()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OG_rot
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OG_rot
        
        .TextFrame.TextRange.Font.Color.RGB = color_OG_weiss
    End With
    
End Sub

Public Sub OG_colors_dunkelrot()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OG_dunkelrot
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OG_dunkelrot
        
        .TextFrame.TextRange.Font.Color.RGB = color_OG_weiss
    End With
    
End Sub

Public Sub OG_colors_mittelrot()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OG_mittelrot
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OG_mittelrot
        
        .TextFrame.TextRange.Font.Color.RGB = color_OG_weiss
    End With
    
End Sub

Public Sub OG_colors_blau()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OG_blau
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OG_blau
        
        .TextFrame.TextRange.Font.Color.RGB = color_OG_weiss
    End With
    
End Sub

Public Sub OG_colors_dunkelblau()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OG_dunkelblau
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OG_dunkelblau
        
        .TextFrame.TextRange.Font.Color.RGB = color_OG_weiss
    End With
    
End Sub

Public Sub OG_colors_hellblau()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OG_hellblau
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OG_hellblau
        
        .TextFrame.TextRange.Font.Color.RGB = color_OG_schwarz
    End With
    
End Sub

Public Sub OG_colors_grau1()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OG_grau1
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OG_grau1
        
        .TextFrame.TextRange.Font.Color.RGB = color_OG_schwarz
    End With
    
End Sub

Public Sub OG_colors_grau2()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OG_grau2
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OG_grau2
        
        .TextFrame.TextRange.Font.Color.RGB = color_OG_schwarz
    End With
    
End Sub

Public Sub OG_colors_grau3()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OG_grau3
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OG_grau3
        
        .TextFrame.TextRange.Font.Color.RGB = color_OG_schwarz
    End With
    
End Sub

Public Sub OG_colors_grau4()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OG_grau4
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OG_grau4
        
        .TextFrame.TextRange.Font.Color.RGB = color_OG_weiss
    End With
    
End Sub

'Standards

Public Sub OG_colors_textbox()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OG_weiss
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OG_grau1

        .TextFrame.TextRange.Font.Color.RGB = color_OG_schwarz
    End With
    
End Sub

Public Sub OG_colors_legende()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = RGB(255, 255, 0)
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OG_grau1

        .TextFrame.TextRange.Font.Color.RGB = color_OG_schwarz
    End With
    
End Sub

Public Sub OG_colors_transparent()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.ForeColor.RGB = color_OG_weiss
        .Fill.Transparency = 1
        .Fill.Visible = msoFalse
        
        .Line.ForeColor.RGB = color_OG_weiss
        .Line.Visible = msoFalse
        
        .TextFrame.TextRange.Font.Color.RGB = color_OG_grau4
    End With
    
End Sub

Public Sub OG_colors_weiss()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.ForeColor.RGB = color_OG_weiss
        .Fill.Transparency = 0
        .Fill.Visible = msoTrue
        
        .Line.ForeColor.RGB = color_OG_weiss
        .Line.Visible = msoTrue
        
        .TextFrame.TextRange.Font.Color.RGB = color_OG_schwarz
    End With
    
End Sub

Public Sub OG_colors_schwarz()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.ForeColor.RGB = color_OG_schwarz
        .Fill.Transparency = 0
        .Fill.Visible = msoTrue
        
        .Line.ForeColor.RGB = color_OG_schwarz
        .Line.Visible = msoTrue
        
        .TextFrame.TextRange.Font.Color.RGB = color_OG_weiss
    End With
    
End Sub

