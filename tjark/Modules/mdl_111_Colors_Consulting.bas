Attribute VB_Name = "mdl_111_Colors_Consulting"

'Farben Otto Group Consulting - NEUE CI

Option Explicit

'Primaerfarben
Public Const color_OC_rot As Long = 2241506 'RGB(226,51,34)
Public Const color_OC_weiss As Long = 16777215  'RGB(255, 255, 255)
Public Const color_OC_schwarz As Long = 0  'RGB(0, 0, 0)

'Sekundaerfarben
Public Const color_OC_grau1 As Long = 14342874  'RGB(218, 218, 218)
Public Const color_OC_grau2 As Long = 12434877  'RGB(189, 189, 189)
Public Const color_OC_grau3 As Long = 8947848  'RGB(136, 136, 136)
Public Const color_OC_grau4 As Long = 6579300  'RGB(100, 100, 100)

Public Const color_OC_hellblau As Long = 16730674  'RGB(50, 74, 255)
Public Const color_OC_blau As Long = 8199443  'RGB(19, 29, 125)
Public Const color_OC_dunkelblau As Long = 4526091  'RGB(11, 16, 69)

'Weitere Farben
Public Const color_OC_signalrot As Long = 14908674 'RGB(2, 226, 202)
Public Const color_OC_signalgelb As Long = 65535  'RGB(255, 255, 0)
Public Const color_OC_signalgruen As Long = 5287936  'RGB(0, 176, 80)
Public Const color_OC_postit As Long = 65535  'RGB(255, 255, 0)

Sub Sample()
    Dim Col As Long
    
    '~~> RGB to LONG
    Col = RGB(0, 176, 80)
    
    Debug.Print Col
    
End Sub

'debug.print RGB(,,)
'Public Const color_OC_ = RGB(,,)

Sub OC_colorscheme_set()
    On Error Resume Next
    
    Dim intMasterCounter
    
    While ActivePresentation.ColorSchemes.Count < 1
        ActivePresentation.ColorSchemes.Add
    Wend
    
    With ActivePresentation.ColorSchemes(1)
        .Colors(ppBackground) = color_OC_weiss
        .Colors(ppForeground) = color_OC_schwarz
        .Colors(ppShadow) = color_OC_grau2
        .Colors(ppTitle) = color_OC_schwarz
        .Colors(ppFill) = color_OC_grau1
        .Colors(ppAccent1) = color_OC_grau4
        .Colors(ppAccent2) = color_OC_grau3
        .Colors(ppAccent3) = color_OC_hellblau
    End With
    
    For intMasterCounter = 1 To ActivePresentation.Designs.Count
        ActivePresentation.Designs(intMasterCounter).SlideMaster.ColorScheme = ActivePresentation.ColorSchemes(1)
    Next intMasterCounter

End Sub

Public Sub OC_extracolors()
    ActivePresentation.ExtraColors.Add color_OC_weiss
    ActivePresentation.ExtraColors.Add color_OC_schwarz
    ActivePresentation.ExtraColors.Add color_OC_grau1
    ActivePresentation.ExtraColors.Add color_OC_hellblau
    ActivePresentation.ExtraColors.Add color_OC_grau2
    ActivePresentation.ExtraColors.Add color_OC_grau3
    ActivePresentation.ExtraColors.Add color_OC_blau
    ActivePresentation.ExtraColors.Add color_OC_dunkelblau
End Sub

Public Sub extracolors_grau()
    ActivePresentation.ExtraColors.Add RGB(40, 40, 40)
    ActivePresentation.ExtraColors.Add RGB(60, 60, 60)
    ActivePresentation.ExtraColors.Add RGB(80, 80, 80)
    ActivePresentation.ExtraColors.Add RGB(100, 100, 100)
    ActivePresentation.ExtraColors.Add RGB(120, 120, 120)
    ActivePresentation.ExtraColors.Add RGB(140, 140, 140)
    ActivePresentation.ExtraColors.Add RGB(160, 160, 160)
    ActivePresentation.ExtraColors.Add RGB(180, 180, 180)
End Sub


'PRIMÄRFARBEN

Public Sub OC_colors_rot()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OC_rot
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OC_rot

        .TextFrame.TextRange.Font.Color.RGB = color_OC_weiss
    End With
    
End Sub

Public Sub OC_colors_weiss()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OC_weiss
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OC_weiss
        
        .TextFrame.TextRange.Font.Color.RGB = color_OC_schwarz
    End With
    
End Sub

Public Sub OC_colors_schwarz()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OC_schwarz
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OC_schwarz
        
        .TextFrame.TextRange.Font.Color.RGB = color_OC_weiss
    End With
    
End Sub

'AKZENTFARBE

Public Sub OC_colors_grau1()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OC_grau1
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OC_grau1

        .TextFrame.TextRange.Font.Color.RGB = color_OC_schwarz
    End With
    
End Sub

Public Sub OC_colors_grau2()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OC_grau2
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OC_grau2

        .TextFrame.TextRange.Font.Color.RGB = color_OC_schwarz
    End With
    
End Sub

Public Sub OC_colors_grau3()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OC_grau3
    
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OC_grau3
        
        .TextFrame.TextRange.Font.Color.RGB = color_OC_schwarz
    End With
    
End Sub

Public Sub OC_colors_grau4()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OC_grau4
    
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OC_grau4
        
        .TextFrame.TextRange.Font.Color.RGB = color_OC_weiss
    End With
    
End Sub

Public Sub OC_colors_hellblau()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OC_hellblau
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OC_hellblau
        
        .TextFrame.TextRange.Font.Color.RGB = color_OC_weiss
    End With
    
End Sub

Public Sub OC_colors_blau()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OC_blau
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OC_blau
        
        .TextFrame.TextRange.Font.Color.RGB = color_OC_weiss
    End With
    
End Sub

Public Sub OC_colors_dunkelblau()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OC_dunkelblau
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OC_dunkelblau
        
        .TextFrame.TextRange.Font.Color.RGB = color_OC_weiss
    End With
    
End Sub

'Standards

Public Sub OC_colors_textbox()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        '.Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OC_weiss
        
        '.Line.Visible = msoFalse
        .Line.ForeColor.RGB = color_OC_grau2

        .TextFrame.TextRange.Font.Color.RGB = color_OC_schwarz
    End With
    
End Sub

Public Sub OC_colors_textbox_OC()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoFalse
        '.Fill.Transparency = 0
        '.Fill.ForeColor.RGB = color_OC_transparent
        
        .Line.Visible = msoFalse
        '.Line.ForeColor.RGB = color_OC_grau2

        .TextFrame.TextRange.Font.Color.RGB = color_OC_schwarz
    End With
    
End Sub

Public Sub OC_colors_transparent()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.ForeColor.RGB = color_OC_weiss
        .Fill.Transparency = 1
        .Fill.Visible = msoFalse
        
        .Line.ForeColor.RGB = color_OC_weiss
        .Line.Visible = msoFalse
        
        .TextFrame.TextRange.Font.Color.RGB = color_OC_schwarz
    End With
    
End Sub

'SIGNALFARBEN

Public Sub OC_colors_signalrot()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OC_signalrot
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OC_signalrot

        .TextFrame.TextRange.Font.Color.RGB = color_OC_weiss
    End With
    
End Sub

Public Sub OC_colors_signalgelb()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OC_signalgelb
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OC_signalgelb

        .TextFrame.TextRange.Font.Color.RGB = color_OC_schwarz
    End With
    
End Sub

Public Sub OC_colors_signalgruen()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OC_signalgruen
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OC_signalgruen

        .TextFrame.TextRange.Font.Color.RGB = color_OC_weiss
    End With
    
End Sub

Public Sub OC_colors_postit()
    On Error Resume Next
    
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OC_postit
        
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = color_OC_schwarz

        .TextFrame.TextRange.Font.Color.RGB = color_OC_schwarz
    End With
    
End Sub


'SPECIALS



Public Sub OC_Table_Area_Style()
    On Error Resume Next
    
    If ActiveWindow.Selection.ShapeRange.Count < 1 Then
        Exit Sub
    End If
    
    Call OG_colors_textbox
    
    Table_Area_Style_Elements ActiveWindow.Selection.ShapeRange
    
End Sub

Public Sub Table_Area_Style_Elements(Elemente As ShapeRange)
    On Error Resume Next

    Dim Rechteck, Linie As Shape
    Dim RN, LN As String
    
    For Each Rechteck In Elemente
        
        If Rechteck.AutoShapeType = msoShapeRectangle Then
        
            'Rechteck.Fill.ForeColor.RGB = color_OC_weiss
                    
            Rechteck.Line.Visible = msoFalse
            Rechteck.Line.Weight = 0
            
            RN = Rechteck.Name
            
            Set Linie = ActiveWindow.View.Slide.Shapes.AddLine(Rechteck.Left, Rechteck.Top + Rechteck.Height, Rechteck.Left + Rechteck.Width, Rechteck.Top + Rechteck.Height)
                                
            With Linie.Line
                .ForeColor.RGB = color_OC_grau2
                .Weight = 0.75
                .Visible = msoTrue
                .DashStyle = msoLineSolid
            End With
            
            LN = Linie.Name
            
            ActiveWindow.Selection.SlideRange.Shapes.Range(Array(RN, LN)).Group
                    
        End If
            
        Elemente.Select

    Next Rechteck
        
End Sub
