Attribute VB_Name = "mdl_021_MenuDrawscheme"
' Prozeduren zum Erzeugen von Textboxen im aktuellen Slide
' --------------------------------------------------------

Option Explicit

Private objNewShape As Shape

Sub OLD_OC_drawscheme_AT()
    On Error GoTo 1
    
    Set objNewShape = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeRectangle, 0, 0, 100, 100)
    objNewShape.TextFrame.TextRange.Text = "Title"
    objNewShape.Select
    
    Call OLD_prop_AT
    
    If 0 = 1 Then
1:    MsgBox ("Bitte eine Folie anw‰hlen, auf der das Element erstellt werden soll.")
    End If
    
End Sub

Sub OLD_OC_drawscheme_ST()
    On Error GoTo 1
    
    Set objNewShape = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeRectangle, 0, 0, 100, 100)
    objNewShape.TextFrame.TextRange.Text = "Sub-Title"
    objNewShape.Select
    Call OLD_prop_ST
    
    If 0 = 1 Then
1:    MsgBox ("Bitte eine Folie anw‰hlen, auf der das Element erstellt werden soll.")
    End If
    
End Sub

Sub OC_drawscheme_header()
    On Error GoTo 1
    
    Set objNewShape = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeRectangle, cm2pt_x(-12.6), cm2pt_y(4.6), 12.35 * cm2pt, 1.3 * cm2pt)
    objNewShape.TextFrame.TextRange.Text = "‹berschrift"
    objNewShape.Select
    
    Call prop_header
    
    If 0 = 1 Then
1:    MsgBox ("Bitte eine Folie anw‰hlen, auf der das Element erstellt werden soll.")
    End If
    
End Sub

Sub OC_drawscheme_textbox()

    On Error GoTo 1
    Set objNewShape = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeRectangle, cm2pt_x(-12.6), cm2pt_y(3.1), 12.35 * cm2pt, 11.3 * cm2pt)
    objNewShape.TextFrame.TextRange.Text = "Text"
    objNewShape.Select
    Call prop_textbox_OC
    
    If 0 = 1 Then
1:    MsgBox ("Bitte eine Folie anw‰hlen, auf der das Element erstellt werden soll.")
    End If

End Sub

Sub OC_drawscheme_greybox()
    On Error GoTo 1
    
    Set objNewShape = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeRectangle, 33, 150.7, 354.1, 286.4)
    objNewShape.TextFrame.TextRange.Text = "Text"
    objNewShape.Select
    Call prop_greybox
    
    If 0 = 1 Then
1:    MsgBox ("Bitte eine Folie anw‰hlen, auf der das Element erstellt werden soll.")
    End If

End Sub

Sub OC_drawscheme_footnote()
    On Error GoTo 1
    
    Set objNewShape = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeRectangle, 0, 0, 100, 100)
    objNewShape.TextFrame.TextRange.Text = "1) Fuﬂnote"
    objNewShape.Select
    Call prop_footnote
    
    If 0 = 1 Then
1:    MsgBox ("Bitte eine Folie anw‰hlen, auf der das Element erstellt werden soll.")
    End If

End Sub

Sub OC_drawscheme_graphicstext()
    On Error GoTo 1
    
    Set objNewShape = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeRectangle, 246.75, 279.62, 141.64, 22.66)
    objNewShape.TextFrame.TextRange.Text = "Beispieltext"
    objNewShape.Select
    Call prop_graphicstext
    
    If 0 = 1 Then
1:    MsgBox ("Bitte eine Folie anw‰hlen, auf der das Element erstellt werden soll.")
    End If

End Sub

Sub OC_drawscheme_sowhatbox()
    On Error GoTo 1
    
    Dim objSowhatRectangle As Shape
    Dim objSowhat_Oval As Shape
    Dim objSowhat_Arrow As Shape
    
    Set objSowhatRectangle = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeRectangle, 33, 442.8, 714.2, 47.6)
    objSowhatRectangle.Select
    ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Text = "Text"
    'Call COLOR(2)
    With ActiveWindow.Selection.ShapeRange
        With .TextFrame.TextRange.Font
            .Size = 14
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
            .WordWrap = msoTrue
            .Orientation = msoTextOrientationHorizontal
        End With
        With .TextFrame.Ruler
            .Levels(1).LeftMargin = 56.6
            .Levels(1).FirstMargin = 56.6
        End With
        With .TextFrame.TextRange.ParagraphFormat
            .Bullet.Visible = msoFalse
            .Alignment = ppAlignLeft
            .SpaceAfter = 0
            .SpaceBefore = 0
            .SpaceWithin = 1
        End With
    End With
        
    Set objSowhat_Oval = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeOval, 50.7, 454.4, 24.08, 24.08)
    objSowhat_Oval.Select
    'Call COLOR(4)
    
    Set objSowhat_Arrow = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeRightArrow, 55, 461.5, 15.9, 10.2)
    objSowhat_Arrow.Select
    'Call COLOR(7)
    
    objSowhatRectangle.Select
    objSowhat_Oval.Select (msoFalse)
    objSowhat_Arrow.Select (msoFalse)
    ActiveWindow.Selection.ShapeRange.Group

    If 0 = 1 Then
1:    MsgBox ("Bitte eine Folie anw‰hlen, auf der das Element erstellt werden soll.")
    End If
End Sub

