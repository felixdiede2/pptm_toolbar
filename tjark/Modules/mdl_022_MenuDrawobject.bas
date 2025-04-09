Attribute VB_Name = "mdl_022_MenuDrawobject"
' Prozeduren zum Erzeugen von Objekten im aktuellen Slide
' -------------------------------------------------------

Option Explicit

Private objNewShape As Shape

Sub OC_drawobject_navigator()
    On Error GoTo 1
    
    Set objNewShape = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeRectangle, 0, 0, 100, 100)
    objNewShape.TextFrame.TextRange.Text = "1"
    objNewShape.Select
    Call prop_navigator
        
    If 0 = 1 Then
1:    MsgBox ("Bitte eine Folie anwählen, auf der das Element erstellt werden soll.")
    End If
    
End Sub

Sub OC_drawobject_navigator_triangle_left()
    On Error GoTo 1
    
    Dim objAicTriangle, objAicOval, objAicGroup As Shape


    Set objAicTriangle = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeRightTriangle, 0, 0, 2.3 * cm2pt, 2.3 * cm2pt)
    objAicTriangle.LockAspectRatio = msoTrue
    objAicTriangle.Select
    objAicTriangle.Rotation = 90
    Call OC_colors_rot
    
    Set objAicOval = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeOval, 0.1 * cm2pt, 0.1 * cm2pt, 1.1 * cm2pt, 1.1 * cm2pt)
    objAicOval.LockAspectRatio = msoTrue
    objAicOval.Select
    Call OC_colors_weiss
    
    With objAicOval.TextFrame.TextRange.Font
        .Size = 20
        .Name = "Arial"
        .Italic = msoFalse
        .Bold = msoTrue
        .Underline = msoFalse
    End With
    
    With objAicOval.TextFrame
        .MarginBottom = 0
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .Orientation = msoTextOrientationHorizontal
        .VerticalAnchor = msoAnchorMiddle
        .HorizontalAnchor = msoAnchorCenter
    End With
    
    With objAicOval.TextFrame.TextRange.ParagraphFormat
            .Bullet.Visible = msoFalse
            .Alignment = ppAlignLeft
            .SpaceAfter = 0
            .SpaceBefore = 0
            .SpaceWithin = 1
    End With
    
    objAicOval.Select
    objAicTriangle.Select (msoFalse)
    
    Set objAicGroup = ActiveWindow.Selection.ShapeRange.Group
    objAicGroup.LockAspectRatio = msoTrue
    
    If 0 = 1 Then
1:    MsgBox ("Bitte eine Folie anwählen, auf der das Element erstellt werden soll.")
    End If
    
End Sub

Sub OC_drawobject_navigator_triangle_right()
    On Error GoTo 1
    
    Dim objAicTriangle, objAicOval, objAicGroup As Shape


    Set objAicTriangle = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeRightTriangle, 0, 0, 2.3 * cm2pt, 2.3 * cm2pt)
    objAicTriangle.LockAspectRatio = msoTrue
    objAicTriangle.Select
    objAicTriangle.Rotation = 90
    Call OC_colors_rot
    
    Set objAicOval = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeOval, 0.1 * cm2pt, 0.1 * cm2pt, 1.1 * cm2pt, 1.1 * cm2pt)
    objAicOval.LockAspectRatio = msoTrue
    objAicOval.Select
    Call OC_colors_weiss
    
    With objAicOval.TextFrame.TextRange.Font
        .Size = 20
        .Name = "Arial"
        .Italic = msoFalse
        .Bold = msoTrue
        .Underline = msoFalse
    End With
    
    With objAicOval.TextFrame
        .MarginBottom = 0
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .Orientation = msoTextOrientationHorizontal
        .VerticalAnchor = msoAnchorMiddle
        .HorizontalAnchor = msoAnchorCenter
    End With
    
    With objAicOval.TextFrame.TextRange.ParagraphFormat
            .Bullet.Visible = msoFalse
            .Alignment = ppAlignLeft
            .SpaceAfter = 0
            .SpaceBefore = 0
            .SpaceWithin = 1
    End With
    
    objAicOval.Select
    objAicTriangle.Select (msoFalse)
    
    Set objAicGroup = ActiveWindow.Selection.ShapeRange.Group
    objAicGroup.LockAspectRatio = msoTrue
    objAicGroup.Flip msoFlipHorizontal
    objAicGroup.Select
    Call AusrichtenAbsRechts
    ActiveWindow.Selection.Unselect
        
    If 0 = 1 Then
1:    MsgBox ("Bitte eine Folie anwählen, auf der das Element erstellt werden soll.")
    End If
    
End Sub

Sub OC_drawobject_linepoint()
    On Error GoTo 1
    
    Set objNewShape = ActiveWindow.View.Slide.Shapes.AddLine(0, ActiveWindow.Presentation.PageSetup.SlideHeight / 3, 100, ActiveWindow.Presentation.PageSetup.SlideHeight / 3)
    
    With objNewShape.Line
        .Visible = msoTrue
        .DashStyle = msoLineSolid
        .Weight = 2.25
        .ForeColor.RGB = color_OC_rot
        .BeginArrowheadStyle = msoArrowheadNone
        '.EndArrowheadLength = msoArrowheadLong
        .EndArrowheadStyle = msoArrowheadOval
        '.EndArrowheadWidth = msoArrowheadWide
    End With
    'objNewLine.Select
    
    If 0 = 1 Then
1:    MsgBox ("Bitte eine Folie anwählen, auf der das Element erstellt werden soll.")
    End If
    
End Sub

Sub OC_drawobject_legend()
    On Error GoTo 1
    
    Set objNewShape = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeRectangularCallout, 0, 0, 100, 100)
    objNewShape.TextFrame.TextRange.Text = "Legende"
    objNewShape.Select
    Call prop_legend
    
    If 0 = 1 Then
1:    MsgBox ("Bitte eine Folie anwählen, auf der das Element erstellt werden soll.")
    End If
    
End Sub

Sub OC_drawobject_addon_vertraulich()
    On Error GoTo 1
    Dim rngDateTime As TextRange
    
    Set objNewShape = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeRectangle, 0, 0, 100, 100)
    objNewShape.TextFrame.TextRange.Text = "- VERTRAULICH -"
    objNewShape.Select
    Call prop_addon_vertraulich
    
    If 0 = 1 Then
1:    MsgBox ("Bitte eine Folie anwählen, auf der das Element erstellt werden soll.")
    End If

End Sub

Sub OC_drawobject_addon_arbeitsstand()
    On Error GoTo 1
    Dim rngDateTime As TextRange
    Dim objNewLine(9), objAicGroup As Shape
    
    Set objNewShape = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeRectangle, 0, 0, 100, 100)
    'objNewShape.TextFrame.TextRange.Paragraphs(1).Sentences(1).InsertBefore.InsertDateTime ppDateTimeMdyy
    objNewShape.TextFrame.TextRange.InsertDateTime ppDateTimeMdyy
    objNewShape.TextFrame.TextRange.Text = "Arbeitsstand " & objNewShape.TextFrame.TextRange.Text
    objNewShape.Select
    Call prop_addon_arbeitsstand
    
    Set objNewLine(1) = ActiveWindow.View.Slide.Shapes.AddLine(0.3 * cm2pt, 0, 0.3 * cm2pt, ActiveWindow.Presentation.PageSetup.SlideHeight / 2 - 3 * cm2pt)
    
    With objNewLine(1).Line
        .Visible = msoTrue
        .DashStyle = msoLineSolid
        .Weight = 1.5
        .ForeColor.RGB = color_OC_rot
        .BeginArrowheadStyle = msoArrowheadNone
        '.EndArrowheadLength = msoArrowheadLong
        .EndArrowheadStyle = msoArrowheadOval
        '.EndArrowheadWidth = msoArrowheadWide
    End With
    
    Set objNewLine(2) = ActiveWindow.View.Slide.Shapes.AddLine(0.3 * cm2pt, ActiveWindow.Presentation.PageSetup.SlideHeight, 0.3 * cm2pt, ActiveWindow.Presentation.PageSetup.SlideHeight / 2 + 3 * cm2pt)
    
    With objNewLine(2).Line
        .Visible = msoTrue
        .DashStyle = msoLineSolid
        .Weight = 1.5
        .ForeColor.RGB = color_OC_rot
        .BeginArrowheadStyle = msoArrowheadNone
        '.EndArrowheadLength = msoArrowheadLong
        .EndArrowheadStyle = msoArrowheadOval
        '.EndArrowheadWidth = msoArrowheadWide
    End With
    
    objNewShape.Select
    objNewLine(1).Select (msoFalse)
    objNewLine(2).Select (msoFalse)
    
    Set objAicGroup = ActiveWindow.Selection.ShapeRange.Group
    ActiveWindow.Selection.Unselect
    
    If 0 = 1 Then
1:    MsgBox ("Bitte eine Folie anwählen, auf der das Element erstellt werden soll.")
    End If

End Sub


Sub OC_drawobject_addon_backup()
    On Error GoTo 1
    
    Dim objNewLine(9), objAicGroup As Shape

    Set objNewShape = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeLeftRightArrow, 0, 0, 100, 100)
    objNewShape.TextFrame.TextRange.Text = "Backup"
    objNewShape.Adjustments.Item(1) = 1
    objNewShape.Adjustments.Item(2) = 0
    objNewShape.Select
    Call prop_addon_backup
    
    Set objNewLine(1) = ActiveWindow.Selection.SlideRange.Shapes.AddConnector(msoConnectorStraight, 259.62, 170.75, 56.62, 0#)
    objNewLine(1).Line.Visible = msoTrue
    objNewLine(1).Line.ForeColor.RGB = color_OC_rot
    objNewLine(1).ConnectorFormat.BeginConnect objNewShape, 3
    objNewLine(1).ConnectorFormat.EndConnect objNewShape, 1
    objNewLine(1).Line.BeginArrowheadStyle = msoArrowheadNone
    objNewLine(1).Line.EndArrowheadStyle = msoArrowheadNone
        
    Set objNewLine(2) = ActiveWindow.Selection.SlideRange.Shapes.AddConnector(msoConnectorStraight, 259.62, 190#, 70#, 0#)
    objNewLine(2).Line.Visible = msoTrue
    objNewLine(2).Line.ForeColor.RGB = color_OC_rot
    objNewLine(2).ConnectorFormat.BeginConnect objNewShape, 5
    objNewLine(2).ConnectorFormat.EndConnect objNewShape, 7
    objNewLine(2).Line.BeginArrowheadStyle = msoArrowheadNone
    objNewLine(2).Line.EndArrowheadStyle = msoArrowheadNone

    objNewShape.Select
    objNewLine(1).Select (msoFalse)
    objNewLine(2).Select (msoFalse)
    
    Set objAicGroup = ActiveWindow.Selection.ShapeRange.Group
    ActiveWindow.Selection.Unselect

    If 0 = 1 Then
1:    MsgBox ("Bitte eine Folie anwählen, auf der das Element erstellt werden soll.")
    End If
    
End Sub

Sub OC_drawobject_triangle()
    On Error GoTo 1
    
    Set objNewShape = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeIsoscelesTriangle, 0, 0, 100, 100)
    objNewShape.TextFrame.TextRange.Text = ""
    objNewShape.Select
    Call prop_triangle
    objNewShape.LockAspectRatio = msoTrue
    
    If 0 = 1 Then
1:    MsgBox ("Bitte eine Folie anwählen, auf der das Element erstellt werden soll.")
    End If
    
End Sub

Sub OC_drawobject_arrow_in_circle()
    On Error GoTo 1
    
    Dim objAicOval, objAicArrow, objAicGroup As Shape

    Set objAicOval = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeOval, 378, 281.9, 24.08, 24.08)
    objAicOval.LockAspectRatio = msoTrue
    objAicOval.Select
    Call OG_colors_blau
    Set objAicArrow = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeRightArrow, 382.2, 289, 15.9, 10.2)
    objAicArrow.LockAspectRatio = msoTrue
    objAicArrow.Select
    Call OC_colors_weiss
    objAicOval.Select
    objAicArrow.Select (msoFalse)
    
    Set objAicGroup = ActiveWindow.Selection.ShapeRange.Group
    objAicGroup.LockAspectRatio = msoTrue
    
    If 0 = 1 Then
1:    MsgBox ("Bitte eine Folie anwählen, auf der das Element erstellt werden soll.")
    End If

End Sub

Sub OC_drawobject_info()
    On Error GoTo 1
    
    Dim objUncertainOval As Shape

    Set objUncertainOval = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeOval, 378, 281.9, 24.08, 24.08)
    objUncertainOval.Select
    objUncertainOval.TextFrame.TextRange.Text = "i"
    Call OG_colors_dunkelblau
    
    With objUncertainOval
        .LockAspectRatio = msoTrue
        With .TextFrame.TextRange.Font
            .Size = 24
            .Name = "Times New Roman"
            .Italic = msoFalse
            .Bold = msoTrue
            .Underline = msoFalse
        End With
        With .TextFrame.TextRange.ParagraphFormat
            .Bullet.Visible = msoFalse
            .Alignment = ppAlignCenter
            .SpaceAfter = 0
            .SpaceBefore = 0.5
            .SpaceWithin = 1
        End With
        With .TextFrame.Ruler
            .Levels(1).LeftMargin = 0
            .Levels(1).FirstMargin = 0
        End With
        With .TextFrame
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0.1 * cm2pt
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorNone
            .AutoSize = ppAutoSizeNone
            .WordWrap = msoTrue
            .Orientation = msoTextOrientationHorizontal
        End With
    End With
    
    If 0 = 1 Then
1:    MsgBox ("Bitte eine Folie anwählen, auf der das Element erstellt werden soll.")
    End If

End Sub

Sub OC_drawobject_conflict()
    On Error GoTo 1
    
    Dim objConflictOval As Shape

    Set objConflictOval = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeOval, 378, 281.9, 24.08, 24.08)
    objConflictOval.Select
    objConflictOval.TextFrame.TextRange.Text = "7"
    Call OG_colors_rot
    
    With objConflictOval
        .LockAspectRatio = msoTrue
        With .TextFrame.TextRange.Font
            .Size = 20
            .Name = "Wingdings 3"
            .Italic = msoFalse
            .Bold = msoTrue
            .Underline = msoFalse
        End With
        With .TextFrame.TextRange.ParagraphFormat
            .Bullet.Visible = msoFalse
            .Alignment = ppAlignCenter
            .SpaceAfter = 0
            .SpaceBefore = 0.5
            .SpaceWithin = 1
        End With
        With .TextFrame.Ruler
            .Levels(1).LeftMargin = 0
            .Levels(1).FirstMargin = 0
        End With
        With .TextFrame
            .MarginLeft = 0
            .MarginRight = 3.5
            .MarginTop = 0
            .MarginBottom = 0
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorNone
            .AutoSize = ppAutoSizeNone
            .WordWrap = msoTrue
            .Orientation = msoTextOrientationHorizontal
        End With
    End With
    
    If 0 = 1 Then
1:    MsgBox ("Bitte eine Folie anwählen, auf der das Element erstellt werden soll.")
    End If
    
End Sub

Sub OC_drawobject_uncertain()
    On Error GoTo 1
    
    Dim objUncertainOval As Shape

    Set objUncertainOval = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeOval, 378, 281.9, 24.08, 24.08)
    objUncertainOval.Select
    objUncertainOval.TextFrame.TextRange.Text = "?"
    Call OG_colors_dunkelblau
    
    With objUncertainOval
        .LockAspectRatio = msoTrue
        With .TextFrame.TextRange.Font
            .Size = 20
            .Name = "Arial"
            .Italic = msoFalse
            .Bold = msoTrue
            .Underline = msoFalse
        End With
        With .TextFrame.TextRange.ParagraphFormat
            .Bullet.Visible = msoFalse
            .Alignment = ppAlignCenter
            .SpaceAfter = 0
            .SpaceBefore = 0.5
            .SpaceWithin = 1
        End With
        With .TextFrame.Ruler
            .Levels(1).LeftMargin = 0
            .Levels(1).FirstMargin = 0
        End With
        With .TextFrame
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorNone
            .AutoSize = ppAutoSizeNone
            .WordWrap = msoTrue
            .Orientation = msoTextOrientationHorizontal
        End With
    End With
    
    If 0 = 1 Then
1:    MsgBox ("Bitte eine Folie anwählen, auf der das Element erstellt werden soll.")
    End If

End Sub

Sub OC_drawobject_exclamation()
    On Error GoTo 1
    
    Dim objUncertainOval As Shape

    Set objUncertainOval = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeOval, 378, 281.9, 24.08, 24.08)
    objUncertainOval.Select
    objUncertainOval.TextFrame.TextRange.Text = "!"
    Call OG_colors_dunkelblau
    
    With objUncertainOval
        .LockAspectRatio = msoTrue
        With .TextFrame.TextRange.Font
            .Size = 20
            .Name = "Arial"
            .Italic = msoFalse
            .Bold = msoTrue
            .Underline = msoFalse
        End With
        With .TextFrame.TextRange.ParagraphFormat
            .Bullet.Visible = msoFalse
            .Alignment = ppAlignCenter
            .SpaceAfter = 0
            .SpaceBefore = 0.5
            .SpaceWithin = 1
        End With
        With .TextFrame.Ruler
            .Levels(1).LeftMargin = 0
            .Levels(1).FirstMargin = 0
        End With
        With .TextFrame
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorNone
            .AutoSize = ppAutoSizeNone
            .WordWrap = msoTrue
            .Orientation = msoTextOrientationHorizontal
        End With
    End With
    
    If 0 = 1 Then
1:    MsgBox ("Bitte eine Folie anwählen, auf der das Element erstellt werden soll.")
    End If

End Sub

Sub OC_drawobject_checkmark()
    On Error GoTo 1
    Dim objOval As Shape

    Set objOval = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeOval, 378, 281.9, 24.08, 24.08)
    objOval.Select
    objOval.TextFrame.TextRange.Text = "ü"
    Call OG_colors_grau2
    
    With objOval
        .LockAspectRatio = msoTrue
        With .TextFrame.TextRange.Font
            .Size = 18
            .Name = "Wingdings"
            .Italic = msoFalse
            .Bold = msoTrue
            .Underline = msoFalse
        End With
        With .TextFrame.TextRange.ParagraphFormat
            .Bullet.Visible = msoFalse
            .Alignment = ppAlignCenter
            .SpaceAfter = 0
            .SpaceBefore = 0.5
            .SpaceWithin = 1
        End With
        With .TextFrame.Ruler
            .Levels(1).LeftMargin = 0
            .Levels(1).FirstMargin = 0
        End With
        With .TextFrame
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorNone
            .AutoSize = ppAutoSizeNone
            .WordWrap = msoTrue
            .Orientation = msoTextOrientationHorizontal
        End With
    End With

    If 0 = 1 Then
1:    MsgBox ("Bitte eine Folie anwählen, auf der das Element erstellt werden soll.")
    End If

End Sub

Sub OC_drawobject_advantage()
    On Error GoTo 1
    Dim objOval As Shape

    Set objOval = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeOval, 378, 281.9, 24.08, 24.08)
    objOval.Select
    objOval.TextFrame.TextRange.Text = "+"
    objOval.Fill.ForeColor.RGB = RGB(0, 176, 80)
    objOval.Line.ForeColor.RGB = RGB(0, 176, 80)
    
    With objOval
        .LockAspectRatio = msoTrue
        With .TextFrame.TextRange.Font
            .Color.RGB = color_OG_weiss
            .Size = 20
            .Name = "Arial"
            .Italic = msoFalse
            .Bold = msoTrue
            .Underline = msoFalse
        End With
        With .TextFrame.TextRange.ParagraphFormat
            .Bullet.Visible = msoFalse
            .Alignment = ppAlignCenter
            .SpaceAfter = 0
            .SpaceBefore = 0.5
            .SpaceWithin = 1
        End With
        With .TextFrame.Ruler
            .Levels(1).LeftMargin = 0
            .Levels(1).FirstMargin = 0
        End With
        With .TextFrame
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorNone
            .AutoSize = ppAutoSizeNone
            .WordWrap = msoTrue
            .Orientation = msoTextOrientationHorizontal
        End With
    End With
    
    If 0 = 1 Then
1:    MsgBox ("Bitte eine Folie anwählen, auf der das Element erstellt werden soll.")
    End If

End Sub

Sub OC_drawobject_disadvantage()
    On Error GoTo 1
    Dim objOval As Shape
    
    Set objOval = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeOval, 378, 281.9, 24.08, 24.08)
    objOval.Select
    objOval.TextFrame.TextRange.Text = "–"
    objOval.Fill.ForeColor.RGB = RGB(226, 51, 34)
    objOval.Line.ForeColor.RGB = RGB(226, 51, 34)
    
    With objOval
        .LockAspectRatio = msoTrue
        With .TextFrame.TextRange.Font
            .Color.RGB = color_OG_weiss
            .Size = 20
            .Name = "Arial"
            .Italic = msoFalse
            .Bold = msoTrue
            .Underline = msoFalse
        End With
        With .TextFrame.TextRange.ParagraphFormat
            .Bullet.Visible = msoFalse
            .Alignment = ppAlignCenter
            .SpaceAfter = 0
            .SpaceBefore = 0.5
            .SpaceWithin = 1
        End With
        With .TextFrame.Ruler
            .Levels(1).LeftMargin = 0
            .Levels(1).FirstMargin = 0
        End With
        With .TextFrame
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 3.5
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorNone
            .AutoSize = ppAutoSizeNone
            .WordWrap = msoTrue
            .Orientation = msoTextOrientationHorizontal
        End With
    End With
    
    If 0 = 1 Then
1:    MsgBox ("Bitte eine Folie anwählen, auf der das Element erstellt werden soll.")
    End If
End Sub

Sub OC_drawobject_agenda1()
    On Error GoTo 1
    
    Set objNewShape = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeRectangle, 0, 0, 100, 100)
    objNewShape.TextFrame.TextRange.Text = "Oberpunkt"
    objNewShape.Select
    Call prop_agenda1
    
    If 0 = 1 Then
1:    MsgBox ("Bitte eine Folie anwählen, auf der das Element erstellt werden soll.")
    End If

End Sub

Sub OC_drawobject_agenda2()
    On Error GoTo 1
    
    Set objNewShape = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeRectangle, 0, 0, 100, 100)
    objNewShape.TextFrame.TextRange.Text = "Unterpunkt"
    objNewShape.Select
    Call prop_agenda2
    
    If 0 = 1 Then
1:    MsgBox ("Bitte eine Folie anwählen, auf der das Element erstellt werden soll.")
    End If

End Sub

Sub OC_drawobject_reminder()
    On Error GoTo 1
    
    Set objNewShape = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeRectangle, 0, 0, 100, 100)
    objNewShape.TextFrame.TextRange.Text = "Chart" & Chr(13) & "überarbeiten"
    objNewShape.Select
    Call prop_reminder
        
    If 0 = 1 Then
1:    MsgBox ("Bitte eine Folie anwählen, auf der das Element erstellt werden soll.")
    End If

End Sub

Sub OC_drawobject_postit()
    On Error GoTo 1
    
    Set objNewShape = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeRectangle, 0, 0, 100, 100)
    objNewShape.TextFrame.TextRange.Text = "To Do"
    objNewShape.Select
    Call prop_postit
        
    If 0 = 1 Then
1:    MsgBox ("Bitte eine Folie anwählen, auf der das Element erstellt werden soll.")
    End If

End Sub
