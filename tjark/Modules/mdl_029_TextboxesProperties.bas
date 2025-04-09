Attribute VB_Name = "mdl_029_TextboxesProperties"
Option Explicit
 
Sub OLD_prop_AT()
    On Error Resume Next
    Dim sngt As Single
    
    'Call color7
    
    With ActiveWindow.Selection.ShapeRange
        .Left = cm2pt_x(-12.6)
        .Top = cm2pt_y(7)
        .Width = 25.2 * cm2pt
        .Height = 1.9 * cm2pt
        .Line.Weight = 0.75
        
        With .TextFrame.TextRange.Font
            .Size = 20
            .Name = "Arial"
            .Italic = msoFalse
            .Bold = msoTrue
            .Underline = msoFalse
        End With
        With .TextFrame
            .MarginLeft = 7.08
            .MarginRight = 7.08
            .MarginTop = 3.8
            .MarginBottom = 3.8
            .VerticalAnchor = msoAnchorTop
            .HorizontalAnchor = msoAnchorNone
            .AutoSize = ppAutoSizeNone
            .WordWrap = msoTrue
            .Orientation = msoTextOrientationHorizontal
        End With
        With .TextFrame.Ruler
            For sngt = 1 To 5 Step 1
            .Levels(sngt).LeftMargin = 0
            .Levels(sngt).FirstMargin = 0
            Next
        End With
        With .TextFrame.TextRange.ParagraphFormat
            .Bullet.Visible = msoFalse
            .Alignment = ppAlignLeft
            .SpaceAfter = 0
            .SpaceBefore = 0
            .SpaceWithin = 1
        End With
    End With
End Sub

Sub OLD_prop_ST()
    On Error Resume Next

    'Call color7
    With ActiveWindow.Selection.ShapeRange
        .Left = cm2pt_x(-12.6)
        .Top = cm2pt_y(4.6)
        .Width = 12.1 * cm2pt
        .Height = 1.3 * cm2pt
        .Line.Weight = 0.75
        
        With .TextFrame.TextRange.Font
            .Size = 16
            .Name = "Arial"
            .Italic = msoFalse
            .Bold = msoTrue
            .Underline = msoFalse
        End With
        With .TextFrame
            .MarginLeft = 0.15 * cm2pt
            .MarginRight = 0.15 * cm2pt
            .MarginTop = 0.15 * cm2pt
            .MarginBottom = 0.15 * cm2pt
            .VerticalAnchor = msoAnchorTop
            .HorizontalAnchor = msoAnchorNone
            .AutoSize = ppAutoSizeNone
            .WordWrap = msoFalse
            .Orientation = msoTextOrientationHorizontal
        End With
        With .TextFrame.Ruler
            .Levels(1).LeftMargin = 30
            .Levels(1).FirstMargin = 0
        End With
        With .TextFrame.TextRange.ParagraphFormat
            .Bullet.Visible = msoTrue
            .Bullet.Font.Name = "Webdings"
            .Bullet.Character = "52"
            .Alignment = ppAlignLeft
            .SpaceAfter = 0
            .SpaceBefore = 0
            .SpaceWithin = 1
        End With
    End With
End Sub

Sub prop_header()
    'On Error Resume Next
    Dim sngt As Single

    Call OG_colors_dunkelblau
    
    With ActiveWindow.Selection.ShapeRange
        .Top = cm2pt_y(4.6)
        .Height = 1.3 * cm2pt
        .Line.Weight = 0.75
        .LockAspectRatio = msoFalse
    
        With .TextFrame.TextRange.Font
            .Size = 14
            .Name = "Arial"
            .Italic = msoFalse
            .Bold = msoTrue
            .Underline = msoFalse
        End With
        With .TextFrame
            .MarginLeft = 0.15 * cm2pt
            .MarginRight = 0.15 * cm2pt
            .MarginTop = 0.15 * cm2pt
            .MarginBottom = 0.15 * cm2pt
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorNone
            .AutoSize = ppAutoSizeNone
            .WordWrap = msoTrue
            .Orientation = msoTextOrientationHorizontal
        End With
        With .TextFrame.Ruler
            For sngt = 1 To 5 Step 1
            .Levels(sngt).LeftMargin = 0
            .Levels(sngt).FirstMargin = 0
            Next
        End With
        With .TextFrame.TextRange.ParagraphFormat
            .Bullet.Visible = msoFalse
            .Alignment = ppAlignCenter
            .SpaceAfter = 0
            .SpaceBefore = 0
            .SpaceWithin = 1
        End With
    End With
    


End Sub

Sub prop_textbox_OC()
    On Error Resume Next
    
    Dim S As Shape
    Dim seitenbreite As Single
    Dim seitenhoehe As Single
    
    Dim sCount, i As Integer
     
    Call OC_colors_textbox_OC
    
    'For Each S In ActiveWindow.Selection.ShapeRange
    '    S.Top = cm2pt_y(3.1)
    '    S.Height = 11.3 * cm2pt
    'Next S
    
    seitenbreite = ActiveWindow.Presentation.PageSetup.SlideWidth
    seitenhoehe = ActiveWindow.Presentation.PageSetup.SlideHeight
    
    With ActiveWindow.Selection.ShapeRange
        .Left = 0.5 * seitenbreite - 15.5 * cm2pt
        .Top = 0.5 * seitenhoehe - 5.6 * cm2pt
        .Width = 31 * cm2pt
        .Height = 12.9 * cm2pt
        
        '.Line.Weight = 0.75
        .LockAspectRatio = msoFalse
        
        With .TextFrame.TextRange.Font
            .Size = 14
            .Name = "Arial"
            .Italic = msoFalse
            .Bold = msoFalse
            .Underline = msoFalse
        End With
        With .TextFrame
            .MarginLeft = 0.15 * cm2pt
            .MarginRight = 0.15 * cm2pt
            .MarginTop = 0.15 * cm2pt
            .MarginBottom = 0.15 * cm2pt
            .VerticalAnchor = msoAnchorTop
            .HorizontalAnchor = msoAnchorNone
            .AutoSize = ppAutoSizeNone
            .WordWrap = msoTrue
            .Orientation = msoTextOrientationHorizontal
        End With
        With .TextFrame.Ruler
            .Levels(1).LeftMargin = 21.25
            .Levels(1).FirstMargin = 0
            .Levels(2).LeftMargin = 35.41
            .Levels(2).FirstMargin = 21.25
            .Levels(3).LeftMargin = 49.57
            .Levels(3).FirstMargin = 35.41
        End With
        With .TextFrame.TextRange.ParagraphFormat
            .Alignment = ppAlignLeft
            .SpaceAfter = 0
            .SpaceBefore = 0.5
            .SpaceWithin = 1
            .Bullet.Character = 8226
            .Bullet.Visible = msoTrue
        End With
    End With

    sCount = ActiveWindow.Selection.ShapeRange.Count
    'Die Eigenschaft "Bullet.Character" lässt sich immer nur auf 1 Objekt anwenden (VBA Bug)
    For i = 1 To sCount Step 1
        'ActiveWindow.Selection.ShapeRange(i).TextFrame.TextRange.ParagraphFormat.Bullet.Font.Name = "Webdings"
        ActiveWindow.Selection.ShapeRange(i).TextFrame.TextRange.ParagraphFormat.Bullet.Character = 8226
    Next i
 
End Sub

Sub prop_textbox()
    On Error Resume Next
    
    Dim S As Shape
    
    Dim sCount, i As Integer
     
    Call OC_colors_textbox
    
    'For Each S In ActiveWindow.Selection.ShapeRange
    '    S.Top = cm2pt_y(3.1)
    '    S.Height = 11.3 * cm2pt
    'Next S
    
    With ActiveWindow.Selection.ShapeRange
        '.Left = cm2pt_x(-12.6)
        '.Top = cm2pt_y(3.1)
        '.Width = 12.35 * cm2pt
        '.Height = 11.3 * cm2pt
        
        .Line.Weight = 0.75
        .LockAspectRatio = msoFalse
        
        With .TextFrame.TextRange.Font
            .Size = 14
            .Name = "Arial"
            .Italic = msoFalse
            .Bold = msoFalse
            .Underline = msoFalse
        End With
        With .TextFrame
            .MarginLeft = 0.15 * cm2pt
            .MarginRight = 0.15 * cm2pt
            .MarginTop = 0.15 * cm2pt
            .MarginBottom = 0.15 * cm2pt
            .VerticalAnchor = msoAnchorTop
            .HorizontalAnchor = msoAnchorNone
            .AutoSize = ppAutoSizeNone
            .WordWrap = msoTrue
            .Orientation = msoTextOrientationHorizontal
        End With
        With .TextFrame.Ruler
            .Levels(1).LeftMargin = 21.25
            .Levels(1).FirstMargin = 0
            .Levels(2).LeftMargin = 35.41
            .Levels(2).FirstMargin = 21.25
            .Levels(3).LeftMargin = 49.57
            .Levels(3).FirstMargin = 35.41
        End With
        With .TextFrame.TextRange.ParagraphFormat
            .Alignment = ppAlignLeft
            .SpaceAfter = 0
            .SpaceBefore = 0.5
            .SpaceWithin = 1
            .Bullet.Character = 8226
            .Bullet.Visible = msoTrue
        End With
    End With

    'sCount = ActiveWindow.Selection.ShapeRange.Count
    'Die Eigenschaft "Bullet.Character" lässt sich immer nur auf 1 Objekt anweden (VBA Bug)
    'For i = 1 To sCount Step 1
    '    ActiveWindow.Selection.ShapeRange(i).TextFrame.TextRange.ParagraphFormat.Bullet.Character = 8226
    'Next i
 
End Sub

Sub prop_greybox()
    On Error Resume Next
    Dim sCount, i As Integer
    
    'Call COLOR(2)
    With ActiveWindow.Selection.ShapeRange
    
        .Line.Weight = 0.75
        .LockAspectRatio = msoFalse
        
        With .TextFrame.TextRange.Font
            .Size = 14
            .Name = "Arial"
            .Italic = msoFalse
            .Bold = msoFalse
            .Underline = msoFalse
        End With
        With .TextFrame
            .MarginLeft = 7.08
            .MarginRight = 7.08
            .MarginTop = 3.8
            .MarginBottom = 3.8
            .VerticalAnchor = msoAnchorTop
            .HorizontalAnchor = msoAnchorNone
            .AutoSize = ppAutoSizeNone
            .WordWrap = msoTrue
            .Orientation = msoTextOrientationHorizontal
        End With
        With .TextFrame.Ruler
            .Levels(1).LeftMargin = 21.25
            .Levels(1).FirstMargin = 0
            .Levels(2).LeftMargin = 35.41
            .Levels(2).FirstMargin = 21.25
            .Levels(3).LeftMargin = 49.57
            .Levels(3).FirstMargin = 35.41
        End With
        With .TextFrame.TextRange.ParagraphFormat
            .Alignment = ppAlignLeft
            .SpaceAfter = 0
            .SpaceBefore = 0.5
            .SpaceWithin = 1
            .Bullet.Character = 8226
            .Bullet.Visible = msoTrue
        End With
    End With
    
    sCount = ActiveWindow.Selection.ShapeRange.Count
    'Die Eigenschaft "Bullet.Character" lässt sich immer nur auf 1 Objekt anweden
    For i = 1 To sCount Step 1
        ActiveWindow.Selection.ShapeRange(i).TextFrame.TextRange.ParagraphFormat.Bullet.Character = 8226
    Next i
 
End Sub

Sub prop_footnote()
    On Error Resume Next
    Dim sngt As Integer
    
    Call OG_colors_transparent
    
    With ActiveWindow.Selection.ShapeRange
        .Left = cm2pt_x(-14.35)
        .Top = cm2pt_y(-7.3)
        .Width = 25 * cm2pt
        .Height = 0.75 * cm2pt
        .LockAspectRatio = msoFalse
        
        '.Line.Weight = 0.75
        
        With .TextFrame.TextRange.Font
            .Size = 10
            .Name = "Arial"
            .Italic = msoFalse
            .Bold = msoFalse
            .Underline = msoFalse
        End With
        With .TextFrame
            .MarginLeft = 0.1 * cm2pt
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
            .VerticalAnchor = msoAnchorBottom
            .HorizontalAnchor = msoAnchorNone
            .AutoSize = ppAutoSizeNone
            .WordWrap = msoFalse
            .Orientation = msoTextOrientationHorizontal
        End With
        With .TextFrame.Ruler
            For sngt = 1 To 5 Step 1
            .Levels(sngt).LeftMargin = 0
            .Levels(sngt).FirstMargin = 0
            Next
        End With
        With .TextFrame.TextRange.ParagraphFormat
            .Bullet.Visible = msoFalse
            .Alignment = ppAlignLeft
            .SpaceAfter = 0
            .SpaceBefore = 0
            .SpaceWithin = 0.9
        End With
    End With
End Sub

Sub prop_graphicstext()
    On Error Resume Next
    Dim sngt As Integer
    
    Call OG_colors_transparent
    
    With ActiveWindow.Selection.ShapeRange
    
        .Line.Weight = 0.75
        .LockAspectRatio = msoFalse
        
        With .TextFrame.TextRange.Font
            .Size = 12
            .Name = "Arial"
            .Italic = msoFalse
            .Bold = msoTrue
            .Underline = msoFalse
        End With
        With .TextFrame
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
            .VerticalAnchor = msoAnchorTop
            .HorizontalAnchor = msoAnchorNone
            .AutoSize = ppAutoSizeShapeToFitText
            .WordWrap = msoFalse
            .Orientation = msoTextOrientationHorizontal
        End With
        With .TextFrame.Ruler
            For sngt = 1 To 5 Step 1
            .Levels(sngt).LeftMargin = 0
            .Levels(sngt).FirstMargin = 0
            Next
        End With
        With .TextFrame.TextRange.ParagraphFormat
            .Bullet.Visible = msoFalse
            .Alignment = ppAlignCenter
            .SpaceAfter = 0
            .SpaceBefore = 0
            .SpaceWithin = 1
        End With
    End With
End Sub

Sub prop_navigator()
    On Error Resume Next
    Dim sngt As Integer

    Call OC_colors_rot
    
    With ActiveWindow.Selection.ShapeRange
        .Line.Visible = msoTrue
        .Line.Weight = 0.75
        .Line.ForeColor.RGB = color_OG_weiss
    
        .Left = cm2pt_x(12.6) - 20
        .Top = cm2pt_y(6.95)
        .Width = 20
        .Height = 20
        .LockAspectRatio = msoTrue
        With .TextFrame.TextRange.Font
            .Size = 14
            .Name = "Arial"
            .Italic = msoFalse
            .Bold = msoTrue
            .Underline = msoFalse
        End With
        With .TextFrame
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorCenter
            .AutoSize = ppAutoSizeNone
            .WordWrap = msoFalse
            .Orientation = msoTextOrientationHorizontal
        End With
        With .TextFrame.Ruler
            For sngt = 1 To 5 Step 1
            .Levels(sngt).LeftMargin = 0
            .Levels(sngt).FirstMargin = 0
            Next
        End With
        With .TextFrame.TextRange.ParagraphFormat
            .Bullet.Visible = msoFalse
            .Alignment = ppAlignLeft
            .SpaceAfter = 0
            .SpaceBefore = 0
            .SpaceWithin = 1
        End With
    End With
End Sub

Sub prop_legend()
    On Error Resume Next
    Dim sngt As Integer

    Call OG_colors_legende
        
    With ActiveWindow.Selection.ShapeRange
        .Left = cm2pt_x(12.6 - 5.63)
        .Top = cm2pt_y(4.6)
        .Width = 5.63 * cm2pt
        .Height = 1 * cm2pt
        
        .Line.Weight = 0.75

        With .Shadow
            .Type = msoShadow21
            .Blur = 4
            .ForeColor.RGB = RGB(0, 0, 0)
            .OffsetX = 2
            .OffsetY = 2
            .Style = msoShadowStyleOuterShadow
            .Visible = msoTrue
            .Transparency = 0.6
        End With
        With .TextFrame.TextRange.Font
            .Size = 14
            .Name = "Arial"
            .Italic = msoFalse
            .Bold = msoFalse
            .Underline = msoFalse
        End With
        With .TextFrame
            .MarginLeft = 7.08
            .MarginRight = 7.08
            .MarginTop = 3.8
            .MarginBottom = 3.8
            .VerticalAnchor = msoAnchorTop
            .HorizontalAnchor = msoAnchorNone
            .AutoSize = ppAutoSizeNone
            .WordWrap = msoTrue
            .Orientation = msoTextOrientationHorizontal
        End With
        With .TextFrame.Ruler
            For sngt = 1 To 5 Step 1
            .Levels(sngt).LeftMargin = 0
            .Levels(sngt).FirstMargin = 0
            Next
        End With
        With .TextFrame.TextRange.ParagraphFormat
            .Bullet.Visible = msoFalse
            .Alignment = ppAlignLeft
            .SpaceAfter = 0
            .SpaceBefore = 0
            .SpaceWithin = 1
        End With
    End With
End Sub

Sub prop_addon_vertraulich()
    On Error Resume Next
    Dim sngt As Integer
    
    Call OG_colors_transparent
    
    With ActiveWindow.Selection.ShapeRange
        .Left = ActiveWindow.Presentation.PageSetup.SlideWidth / 2 - 5 * cm2pt
        '.Top = cm2pt_y(-8.33)
        .Top = ActiveWindow.Presentation.PageSetup.SlideHeight - 0.8 * cm2pt
        .Width = 10 * cm2pt
        .Height = 0.8 * cm2pt
        .LockAspectRatio = msoFalse
        '.Line.Weight = 0.75
        
        With .TextFrame.TextRange.Font
            .Color.RGB = color_OG_rot
            .Size = 12
            .Name = "Arial"
            .Italic = msoFalse
            .Bold = msoTrue
            .Underline = msoFalse
        End With
        With .TextFrame
            .MarginLeft = 7.08
            .MarginRight = 7.08
            .MarginTop = 3.8
            .MarginBottom = 3.8
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorNone
            .AutoSize = ppAutoSizeNone
            .WordWrap = msoFalse
            .Orientation = msoTextOrientationHorizontal
        End With
        With .TextFrame.Ruler
            For sngt = 1 To 5 Step 1
            .Levels(sngt).LeftMargin = 0
            .Levels(sngt).FirstMargin = 0
            Next
        End With
        With .TextFrame.TextRange.ParagraphFormat
            .Bullet.Visible = msoFalse
            .Alignment = ppAlignCenter
            .SpaceAfter = 0
            .SpaceBefore = 0
            .SpaceWithin = 1
        End With
    End With
End Sub

Sub prop_addon_arbeitsstand()
    On Error Resume Next
    Dim sngt As Integer
    
    Call OC_colors_weiss
    
    With ActiveWindow.Selection.ShapeRange
        .Left = 0
        .Top = ActiveWindow.Presentation.PageSetup.SlideHeight / 2 - 3 * cm2pt
        .Width = 0.6 * cm2pt
        .Height = 6 * cm2pt
        '.Height = ActiveWindow.Presentation.PageSetup.SlideHeight
        '.Line.Weight = 0.75
        .LockAspectRatio = msoFalse
        
        With .TextFrame.TextRange.Font
            .Color.RGB = color_OG_rot
            .Size = 12
            .Name = "Arial"
            .Italic = msoFalse
            .Bold = msoTrue
            .Underline = msoFalse
        End With
        With .TextFrame
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorCenter
            .AutoSize = ppAutoSizeNone
            .WordWrap = msoFalse
            .Orientation = msoTextOrientationUpward
        End With
        With .TextFrame.Ruler
            For sngt = 1 To 5 Step 1
            .Levels(sngt).LeftMargin = 0
            .Levels(sngt).FirstMargin = 0
            Next
        End With
        With .TextFrame.TextRange.ParagraphFormat
            .Bullet.Visible = msoFalse
            .Alignment = ppAlignCenter
            .SpaceAfter = 0
            .SpaceBefore = 0
            .SpaceWithin = 1
        End With
    End With
End Sub

Sub prop_addon_backup()
    On Error Resume Next
    Dim sngt As Integer
    Dim seitenbreite, seitenhoehe As Single
        
    seitenbreite = ActiveWindow.Presentation.PageSetup.SlideWidth
    seitenhoehe = ActiveWindow.Presentation.PageSetup.SlideHeight

    With ActiveWindow.Selection.ShapeRange
        .Width = 70
        .Height = 0.79 * cm2pt
        '.Left = 30.2 * cm2pt
        .Left = (0.5 * seitenbreite) + (15.5 * cm2pt) - 70
        .Top = 2.967 * cm2pt
        
        .Line.Visible = msoFalse
        
        .Fill.Visible = msoFalse
        .Fill.Transparency = 0
        .Fill.ForeColor.RGB = color_OG_weiss
        
        With .TextFrame.TextRange.Font
            .Color.RGB = color_OC_rot
            .Size = 14
            .Name = "Arial"
            .Italic = msoFalse
            .Bold = msoTrue
            .Underline = msoFalse
        End With
        With .TextFrame
            .MarginLeft = 2.83
            .MarginRight = 2.83
            .MarginTop = 0.05 * cm2pt
            .MarginBottom = 0.05 * cm2pt
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorNone
            .AutoSize = ppAutoSizeShapeToFitText
            .WordWrap = msoFalse
            .Orientation = msoTextOrientationHorizontal
        End With
        'With .TextFrame.Ruler
        '    For sngt = 1 To 5 Step 1
        '    .Levels(sngt).LeftMargin = 0
        '    .Levels(sngt).FirstMargin = 0
        '    Next
        'End With
        With .TextFrame.TextRange.ParagraphFormat
            .Bullet.Visible = msoFalse
            .Alignment = ppAlignRight
            .SpaceAfter = 0
            .SpaceBefore = 0
            .SpaceWithin = 1
        End With
    End With
End Sub

Sub prop_triangle()
    On Error Resume Next
    Dim sngt As Integer
    Dim seitenbreite As Single
    Dim seitenhoehe As Single
    
    Call OC_colors_schwarz
    
    seitenbreite = ActiveWindow.Presentation.PageSetup.SlideWidth
    seitenhoehe = ActiveWindow.Presentation.PageSetup.SlideHeight
    
    With ActiveWindow.Selection.ShapeRange
        .Left = 0.5 * seitenbreite
        .Top = 0.5 * seitenhoehe
        .Width = 1 * cm2pt
        .Height = 0.5 * cm2pt
        .Rotation = 90
        '.Line.Weight = 0.75
        
        With .TextFrame.TextRange.Font
            .Size = 12
            .Name = "Arial"
            .Italic = msoFalse
            .Bold = msoFalse
            .Underline = msoFalse
        End With
        With .TextFrame
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorNone
            .AutoSize = ppAutoSizeNone
            .WordWrap = msoFalse
            .Orientation = msoTextOrientationHorizontal
        End With
        With .TextFrame.Ruler
            For sngt = 1 To 5 Step 1
            .Levels(sngt).LeftMargin = 0
            .Levels(sngt).FirstMargin = 0
            Next
        End With
        With .TextFrame.TextRange.ParagraphFormat
            .Bullet.Visible = msoFalse
            .Alignment = ppAlignCenter
            .SpaceAfter = 0
            .SpaceBefore = 0
            .SpaceWithin = 1
        End With
    End With
End Sub

Sub prop_agenda1()
    On Error Resume Next
    Dim sngt As Integer

    Call OG_colors_dunkelblau
    With ActiveWindow.Selection.ShapeRange
        .Left = cm2pt_x(-12.6)
        .Top = cm2pt_y(4.6)
        .Width = 25.2 * cm2pt
        .Height = 1 * cm2pt
        .Line.Weight = 0.75
        With .TextFrame.TextRange.Font
            .Size = 18
            .Name = "Arial"
            .Italic = msoFalse
            .Bold = msoTrue
            .Underline = msoFalse
        End With
        With .TextFrame
            .MarginLeft = 7.08
            .MarginRight = 7.08
            .MarginTop = 3.8
            .MarginBottom = 3.8
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorNone
            .AutoSize = ppAutoSizeNone
            .WordWrap = msoFalse
            .Orientation = msoTextOrientationHorizontal
        End With
        With .TextFrame.Ruler
            .Levels(1).LeftMargin = 21.81
            .Levels(1).FirstMargin = 0
        End With
        With .TextFrame.TextRange.ParagraphFormat
            .Bullet.Visible = msoFalse
            '.Bullet.Font.Name = "Webdings"
            '.Bullet.Character = "52"
            .Alignment = ppAlignLeft
            .SpaceAfter = 0
            .SpaceBefore = 0
            .SpaceWithin = 1
        End With
    End With
End Sub

Sub prop_agenda2()
    On Error Resume Next
    
    Call OG_colors_grau1
    With ActiveWindow.Selection.ShapeRange
        .Left = cm2pt_x(-12.6)
        .Top = cm2pt_y(3.1)
        .Width = 25.2 * cm2pt
        .Height = 1 * cm2pt
        .Line.Weight = 0.75
        With .TextFrame.TextRange.Font
            .Size = 16
            .Name = "Arial"
            .Italic = msoFalse
            .Bold = msoTrue
            .Underline = msoFalse
        End With
        With .TextFrame
            .MarginLeft = 7.08
            .MarginRight = 7.08
            .MarginTop = 3.8
            .MarginBottom = 3.8
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorNone
            .AutoSize = ppAutoSizeNone
            .WordWrap = msoFalse
            .Orientation = msoTextOrientationHorizontal
        End With
        With .TextFrame.Ruler
            .Levels(1).LeftMargin = 0.75 * cm2pt
            .Levels(1).FirstMargin = 0
        End With
        With .TextFrame.TextRange.ParagraphFormat
            .Bullet.Visible = msoFalse
            .Bullet.Font.Name = "Webdings"
            .Bullet.Character = 52
            .Alignment = ppAlignLeft
            .SpaceAfter = 0
            .SpaceBefore = 0
            .SpaceWithin = 1
        End With
    End With
End Sub

Sub prop_reminder()
    On Error Resume Next
    Dim sngt As Integer
    Dim seitenbreite As Single
    
    seitenbreite = ActiveWindow.Presentation.PageSetup.SlideWidth
    
    With ActiveWindow.Selection.ShapeRange
        '.Left = cm2pt_x(6)
        '.Top = cm2pt_y(9.43)
        .Left = seitenbreite - 10 * cm2pt
        .Top = 0
        .Width = 10 * cm2pt
        .Height = 2 * cm2pt
        .Line.Weight = 0.75
        .Fill.ForeColor.RGB = RGB(200, 0, 0)

        'With .Shadow
        '    .Type = msoShadow21
        '    .Blur = 4
        '    .ForeColor.RGB = RGB(0, 0, 0)
        '    .OffsetX = 2
        '    .OffsetY = 2
        '    .Style = msoShadowStyleOuterShadow
        '    .Visible = msoTrue
        '    .Transparency = 0.6
        'End With

        With .TextFrame.TextRange.Font
            .Color.RGB = color_OG_weiss
            .Size = 14
            .Name = "Arial"
            .Italic = msoFalse
            .Bold = msoTrue
            .Underline = msoFalse
        End With
        With .TextFrame
            .MarginLeft = 0.1 * cm2pt
            .MarginRight = 0.1 * cm2pt
            .MarginTop = 0.1 * cm2pt
            .MarginBottom = 0.1 * cm2pt
            .VerticalAnchor = msoAnchorTop
            '.VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorNone
            .AutoSize = ppAutoSizeShapeToFitText
            '.AutoSize = ppAutoSizeNone
            .WordWrap = msoTrue
            .Orientation = msoTextOrientationHorizontal
        End With
        .Top = 0
        With .TextFrame.Ruler
            For sngt = 1 To 5 Step 1
            .Levels(sngt).LeftMargin = 0
            .Levels(sngt).FirstMargin = 0
            Next
        End With
        With .TextFrame.TextRange.ParagraphFormat
            .Bullet.Visible = msoFalse
            .Alignment = ppAlignCenter
            .SpaceAfter = 0
            .SpaceBefore = 0
            .SpaceWithin = 0.9
        End With
    End With
End Sub

Sub prop_postit()
    On Error Resume Next
    Dim sngt As Integer
    Dim seitenbreite As Single
    
    seitenbreite = ActiveWindow.Presentation.PageSetup.SlideWidth
    
    With ActiveWindow.Selection.ShapeRange
        .Left = seitenbreite - 10 * cm2pt
        .Top = 0
        .Width = 10 * cm2pt
        .Height = 2 * cm2pt
        .Line.Weight = 0.75
        .Fill.ForeColor.RGB = RGB(255, 255, 102)
        
        'With .Shadow
        '    .Type = msoShadow21
        '    .Blur = 4
        '    .ForeColor.RGB = RGB(0, 0, 0)
        '    .OffsetX = 2
        '    .OffsetY = 2
        '    .Style = msoShadowStyleOuterShadow
        '    .Visible = msoTrue
        '    .Transparency = 0.6
        'End With

        With .TextFrame.TextRange.Font
            .Color.RGB = color_OG_schwarz
            .Size = 14
            .Name = "Arial"
            .Italic = msoFalse
            .Bold = msoTrue
            .Underline = msoFalse
        End With
        
        With .TextFrame
            .MarginLeft = 0.1 * cm2pt
            .MarginRight = 0.1 * cm2pt
            .MarginTop = 0.1 * cm2pt
            .MarginBottom = 0.1 * cm2pt
            .VerticalAnchor = msoAnchorTop
            '.VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorNone
            .AutoSize = ppAutoSizeShapeToFitText
            '.AutoSize = ppAutoSizeNone
            .WordWrap = msoTrue
            .Orientation = msoTextOrientationHorizontal
        End With
        .Top = 0
        With .TextFrame.Ruler
            For sngt = 1 To 5 Step 1
            .Levels(sngt).LeftMargin = 0
            .Levels(sngt).FirstMargin = 0
            Next
        End With
        
        With .TextFrame.TextRange.ParagraphFormat
            .Bullet.Visible = msoFalse
            .Alignment = ppAlignCenter
            .SpaceAfter = 0
            .SpaceBefore = 0
            .SpaceWithin = 0.9
        End With
    End With
End Sub

