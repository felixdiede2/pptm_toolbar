Attribute VB_Name = "mdl_011_DesignBorders_Set"
Option Explicit

' Save border set marker
Dim bolBorderSet As Boolean

'Save border shapes and counter
Dim objBorderShape(99, 5) As Shape
Dim sngCounter1, sngMasterCounter As Single
Dim classOnSave As New mdlBeforeSave

Sub designborders_set()
    On Error Resume Next
        
    Dim seitenbreite As Single
    Dim seitenhoehe As Single
    
    'Initial delete
    For sngMasterCounter = 1 To ActivePresentation.Designs.Count
        For sngCounter1 = 1 To 5
            objBorderShape(sngMasterCounter, sngCounter1).Delete
        Next sngCounter1
    Next sngMasterCounter
    
    'Activate class module
    Set classOnSave.DeleteDesignBorders = Application
      
    'Create and format borders
    For sngMasterCounter = 1 To ActivePresentation.Designs.Count
    
        seitenbreite = ActiveWindow.Presentation.PageSetup.SlideWidth
        seitenhoehe = ActiveWindow.Presentation.PageSetup.SlideHeight
        
        Set objBorderShape(sngMasterCounter, 1) = ActivePresentation.Designs(sngMasterCounter).SlideMaster.Shapes.AddShape(msoShapeRectangle, 0, 0, seitenbreite, (0.5 * seitenhoehe) - (5.6 * cm2pt))
        Set objBorderShape(sngMasterCounter, 2) = ActivePresentation.Designs(sngMasterCounter).SlideMaster.Shapes.AddShape(msoShapeRectangle, 0, 0, 0.5 * seitenbreite - 15.5 * cm2pt, seitenhoehe)
        Set objBorderShape(sngMasterCounter, 3) = ActivePresentation.Designs(sngMasterCounter).SlideMaster.Shapes.AddShape(msoShapeRectangle, 0.5 * seitenbreite + 15.5 * cm2pt, 0, 0.5 * seitenbreite - 15.5 * cm2pt, seitenhoehe)
        Set objBorderShape(sngMasterCounter, 4) = ActivePresentation.Designs(sngMasterCounter).SlideMaster.Shapes.AddShape(msoShapeRectangle, 0, 0.5 * seitenhoehe + 7.3 * cm2pt, seitenbreite, 0.5 * seitenhoehe - 7.3 * cm2pt)
        Set objBorderShape(sngMasterCounter, 5) = ActivePresentation.Designs(sngMasterCounter).SlideMaster.Shapes.AddShape(msoShapeRectangle, cm2pt_x(-7), cm2pt_y(-8.6), 14 * cm2pt, 0.9 * cm2pt)
    
        For sngCounter1 = 1 To 4
            With objBorderShape(sngMasterCounter, sngCounter1)
                .Fill.Visible = msoTrue
                .Fill.Transparency = 0
                .Fill.Patterned msoPatternDarkDownwardDiagonal
                .Fill.ForeColor.RGB = color_OG_blau
                .Fill.BackColor.RGB = color_OG_weiss
                
                .Line.Visible = msoFalse
            End With
        Next sngCounter1
    
        'Format textbox
        With objBorderShape(sngMasterCounter, 5)
                .Fill.Visible = msoTrue
                .Fill.Transparency = 0
                .Fill.ForeColor.RGB = color_OG_weiss
                .Line.Visible = msoTrue
                .Line.ForeColor.RGB = color_OG_weiss
                .TextFrame.TextRange.Text = "Der Gestaltungsrahmen wird beim Speichern gelöscht"
            With .TextFrame.TextRange.Font
                .Color.RGB = color_OG_rot
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
                .Levels(1).LeftMargin = 0
                .Levels(1).FirstMargin = 0
            End With
            With .TextFrame.TextRange.ParagraphFormat
                .Bullet.Visible = msoFalse
                .Alignment = ppAlignCenter
                .SpaceAfter = 0
                .SpaceBefore = 0
                .SpaceWithin = 1
            End With
        End With
     
    Next sngMasterCounter
    bolBorderSet = msoTrue
End Sub

Sub designborders_delete()
    On Error Resume Next
    
    If bolBorderSet = msoTrue Then
        For sngMasterCounter = 1 To ActivePresentation.Designs.Count
            For sngCounter1 = 1 To 5
                objBorderShape(sngMasterCounter, sngCounter1).Delete
            Next sngCounter1
        Next sngMasterCounter
    bolBorderSet = msoFalse
    End If

End Sub
