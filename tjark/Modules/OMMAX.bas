Attribute VB_Name = "OMMAX"
Option Explicit

'Diese beiden Zeilen gehören zu PositionPickup und PositionApply
Public sngleft As Single
Public sngtop As Single

'Der folgende Block gehört zu den BulletLevel-Tools
Dim startPos As Long, endPos As Long, lastPos As Long, indentPos As Long
Dim IndentFirstMargin As Double, IndentLeftMargin As Double, bulletVisible As Boolean, textBold As Boolean
Dim bulletCharacter As Long, bulletColorRed As Long, bulletColorGreen As Long, bulletColorBlue As Long
Dim fontColorRed As Long, fontColorGreen As Long, fontColorBlue As Long, bulletFontName As String, bulletSize As Double
Dim colTable As Long, rowTable As Long
Const GroupShapesErrorMessage = "Please ungroup shapes before using this macro"
Const MultiShapesErrorMessage = "Please select only one shape"
Const AllowedShapesErrorMessage = "Please select a shape to use this macro, not a placeholder"
Const CursorPosErrorMessage = "Please do not place the cursor to final position of the text, when using this macro"

'Diese beiden sind unerlässlich für Mac-Kompatibilität

Function isMac() As Boolean
isMac = False
If InStr(Application.OperatingSystem, "Macintosh") Then
isMac = True
End If
End Function

Function GetUserNameMac() As String
    Dim sMyScript As String

    sMyScript = "set userName to short user name of (system info)" & vbNewLine & "return userName"

    GetUserNameMac = MacScript(sMyScript)
End Function

'AAA - Anzupassende

'Dies sind die Parameter, auf die von vielen Shapes zurückgegriffen wird

Private Sub ParameterBoxBodyShadow()

Dim shp As Shape
Dim i As Integer
    
Set shp = ActiveWindow.Selection.ShapeRange(1)

With shp
    .Fill.Visible = msoTrue
    .Fill.Transparency = 0
    .Fill.ForeColor.RGB = RGB(255, 255, 255)
    .Line.Visible = msoTrue
    .Line.ForeColor.RGB = RGB(255, 255, 255)
    .Line.Weight = 0.75
    .Shadow.Style = msoShadowStyleOuterShadow
    .Shadow.ForeColor.RGB = RGB(0, 0, 0)
    .Shadow.Transparency = 0.6
    .Shadow.Size = 100
    .Shadow.Blur = 4
    .Shadow.OffsetX = 2.1
    .Shadow.OffsetY = 2.1

With .TextFrame2
    .TextRange.Text = "Text 1" & vbCr & "Text 2" & vbCr & "Text 3"
    .VerticalAnchor = msoAnchorTop
    .MarginBottom = 7.0866097
    .MarginLeft = 7.0866097
    .MarginRight = 7.0866097
    .MarginTop = 7.0866097
    .WordWrap = msoTrue
With .TextRange
    .Font.Size = 14
    .Font.Name = "Arial"
    .Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    .Font.Bold = msoFalse
    .Font.Italic = msoFalse
    .Font.UnderlineStyle = msoNoUnderline
    .ParagraphFormat.Alignment = ppAlignLeft
    .ParagraphFormat.Bullet.UseTextColor = msoFalse
    .ParagraphFormat.Bullet.UseTextFont = msoFalse
    .ParagraphFormat.Bullet.Font.Fill.ForeColor.RGB = RGB(9, 91, 164)
    .Characters(1, 6).Font.Bold = msoTrue
                For i = 2 To 3
With .Paragraphs(i).ParagraphFormat.Bullet
    .Visible = msoTrue
End With
                Next
With .Paragraphs(2)
    .ParagraphFormat.indentLevel = 2
    .ParagraphFormat.Bullet.Character = 8226
    .ParagraphFormat.Bullet.RelativeSize = 1
End With
With .Paragraphs(3)
    .ParagraphFormat.indentLevel = 3
    .ParagraphFormat.Bullet.Character = 8211
    .ParagraphFormat.Bullet.RelativeSize = 1
End With
End With
End With
With .TextFrame
With .Ruler
     .Levels(2).FirstMargin = 14.173219
     .Levels(2).LeftMargin = 28.346439
     .Levels(3).FirstMargin = 28.346439
     .Levels(3).LeftMargin = 42.519658
End With
End With
End With

End Sub

Private Sub ParameterBoxHeaderFlat()

Dim shp As Shape

Set shp = ActiveWindow.Selection.ShapeRange(1)

    shp.Fill.Visible = msoTrue
    shp.Fill.Transparency = 0
    shp.Fill.ForeColor.RGB = RGB(9, 91, 164)
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(9, 91, 164)
    shp.Line.Weight = 0.75
    
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Characters.Text = "Header"
    shp.TextFrame2.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.Font.Size = 16
    shp.TextFrame2.TextRange.Font.Name = "Arial"
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.Font.Italic = msoFalse
    shp.TextFrame2.TextRange.Font.UnderlineStyle = msoNoUnderline
    shp.TextFrame2.Orientation = msoTextOrientationHorizontal
    shp.TextFrame2.MarginBottom = 7.0866097
    shp.TextFrame2.MarginLeft = 7.0866097
    shp.TextFrame2.MarginRight = 7.0866097
    shp.TextFrame2.MarginTop = 7.0866097
    shp.TextFrame2.WordWrap = msoTrue
    
End Sub

Private Sub ParameterBoxHeaderShadow()

Dim shp As Shape

Set shp = ActiveWindow.Selection.ShapeRange(1)

    shp.Fill.Visible = msoTrue
    shp.Fill.Transparency = 0
    shp.Fill.ForeColor.RGB = RGB(221, 221, 221)
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(221, 221, 221)
    shp.Line.Weight = 0.75
    shp.Shadow.Style = msoShadowStyleOuterShadow
    shp.Shadow.ForeColor.RGB = RGB(0, 0, 0)
    shp.Shadow.Transparency = 0.6
    shp.Shadow.Size = 100
    shp.Shadow.Blur = 4
    shp.Shadow.OffsetX = 2.1
    shp.Shadow.OffsetY = 2.1

    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(9, 91, 164)
    shp.TextFrame2.TextRange.Characters.Text = "Header"
    shp.TextFrame2.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.Font.Size = 16
    shp.TextFrame2.TextRange.Font.Name = "Arial"
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.Font.Italic = msoFalse
    shp.TextFrame2.TextRange.Font.UnderlineStyle = msoNoUnderline
    shp.TextFrame2.MarginBottom = 7.0866097
    shp.TextFrame2.MarginLeft = 7.0866097
    shp.TextFrame2.MarginRight = 7.0866097
    shp.TextFrame2.MarginTop = 7.0866097
    shp.TextFrame2.WordWrap = msoTrue
    
End Sub

Private Sub ParameterColumnBodyStandard()

Dim shp As Shape
Dim i As Integer
    
Set shp = ActiveWindow.Selection.ShapeRange(1)

With shp
    .Fill.Visible = msoFalse
    .Line.Visible = msoFalse
With .TextFrame2
    .TextRange.Text = "Text 1" & vbCr & "Text 2" & vbCr & "Text 3"
    .VerticalAnchor = msoAnchorTop
    .MarginBottom = 7.0866097
    .MarginLeft = 7.0866097
    .MarginRight = 7.0866097
    .MarginTop = 7.0866097
    .WordWrap = msoTrue
With .TextRange
    .Font.Size = 14
    .Font.Name = "Arial"
    .Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    .Font.Bold = msoFalse
    .Font.Italic = msoFalse
    .Font.UnderlineStyle = msoNoUnderline
    .ParagraphFormat.Alignment = ppAlignLeft
    .ParagraphFormat.Bullet.UseTextColor = msoFalse
    .ParagraphFormat.Bullet.UseTextFont = msoFalse
    .ParagraphFormat.Bullet.Font.Fill.ForeColor.RGB = RGB(9, 91, 164)
    .Characters(1, 6).Font.Bold = msoTrue
                For i = 2 To 3
With .Paragraphs(i).ParagraphFormat.Bullet
    .Visible = msoTrue
End With
                Next
With .Paragraphs(2)
    .ParagraphFormat.indentLevel = 2
    .ParagraphFormat.Bullet.Character = 8226
    .ParagraphFormat.Bullet.RelativeSize = 1
End With
With .Paragraphs(3)
    .ParagraphFormat.indentLevel = 3
    .ParagraphFormat.Bullet.Character = 8211
    .ParagraphFormat.Bullet.RelativeSize = 1
End With
End With
End With
With .TextFrame
With .Ruler
     .Levels(2).FirstMargin = 14.173219
     .Levels(2).LeftMargin = 28.346439
     .Levels(3).FirstMargin = 28.346439
     .Levels(3).LeftMargin = 42.519658
End With
End With
End With

End Sub

Private Sub ParameterColumnHeader()

Dim shp As Shape

Set shp = ActiveWindow.Selection.ShapeRange(1)

    shp.Fill.Visible = msoTrue
    shp.Fill.Transparency = 0
    shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.Line.Visible = msoFalse
    shp.Shadow.Style = msoShadowStyleOuterShadow
    shp.Shadow.Type = msoShadow21
    shp.Shadow.ForeColor.RGB = RGB(9, 91, 164)
    shp.Shadow.Transparency = 0
    shp.Shadow.Size = 99
    shp.Shadow.Blur = 0
    shp.Shadow.OffsetX = 0
    shp.Shadow.OffsetY = 2

    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    shp.TextFrame2.TextRange.Characters.Text = "Header"
    shp.TextFrame2.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.Font.Size = 16
    shp.TextFrame2.TextRange.Font.Name = "Arial"
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.Font.Italic = msoFalse
    shp.TextFrame2.TextRange.Font.UnderlineStyle = msoNoUnderline
    shp.TextFrame2.MarginBottom = 7.0866097
    shp.TextFrame2.MarginLeft = 7.0866097
    shp.TextFrame2.MarginRight = 7.0866097
    shp.TextFrame2.MarginTop = 7.0866097
    shp.TextFrame2.WordWrap = msoTrue
    
End Sub

Private Sub ParameterNumberBall()

Dim shp As Shape
    
Set shp = ActiveWindow.Selection.ShapeRange(1)

    shp.Fill.Visible = msoTrue
    shp.Fill.Transparency = 0
    shp.Fill.ForeColor.RGB = RGB(89, 171, 244)
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(89, 171, 244)

    shp.TextFrame2.TextRange.Font.Size = 14
    shp.TextFrame2.MarginBottom = 0
    shp.TextFrame2.MarginLeft = 0
    shp.TextFrame2.MarginRight = 0
    shp.TextFrame2.MarginTop = 0

End Sub

Private Sub ParameterNumberBigSolo()

Dim shp As Shape
    
Set shp = ActiveWindow.Selection.ShapeRange(1)

shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(150, 150, 150)
shp.TextFrame2.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
shp.TextFrame2.TextRange.Font.Size = 28
shp.TextFrame2.MarginBottom = 0
shp.TextFrame2.MarginLeft = 0
shp.TextFrame2.MarginRight = 0
shp.TextFrame2.MarginTop = 0
shp.TextFrame2.WordWrap = msoFalse

End Sub

Private Sub ParameterNumberSquare()

Dim shp As Shape
    
Set shp = ActiveWindow.Selection.ShapeRange(1)

shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(225, 144, 12)
shp.TextFrame2.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
shp.TextFrame2.TextRange.Font.Bold = msoTrue
shp.TextFrame2.TextRange.Font.Italic = msoFalse
shp.TextFrame2.TextRange.Font.UnderlineStyle = msoNoUnderline
shp.TextFrame2.MarginBottom = 0
shp.TextFrame2.MarginLeft = 0
shp.TextFrame2.MarginRight = 0
shp.TextFrame2.MarginTop = 0

End Sub

Private Sub ParameterShapeStandard()

Dim shp As Shape
    
Set shp = ActiveWindow.Selection.ShapeRange(1)

    shp.Fill.Visible = msoTrue
    shp.Fill.Transparency = 0
    shp.Fill.ForeColor.RGB = RGB(221, 221, 221)
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(221, 221, 221)
    shp.Line.Weight = 0.75
    
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    shp.TextFrame2.TextRange.Characters.Text = "Text"
    shp.TextFrame2.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.Font.Size = 14
    shp.TextFrame2.TextRange.Font.Name = "Arial"
    shp.TextFrame2.TextRange.Font.Bold = msoFalse
    shp.TextFrame2.TextRange.Font.Italic = msoFalse
    shp.TextFrame2.TextRange.Font.UnderlineStyle = msoNoUnderline
    shp.TextFrame2.Orientation = msoTextOrientationHorizontal
    shp.TextFrame2.MarginBottom = 7.0866097
    shp.TextFrame2.MarginLeft = 7.0866097
    shp.TextFrame2.MarginRight = 7.0866097
    shp.TextFrame2.MarginTop = 7.0866097
    shp.TextFrame2.WordWrap = msoTrue

End Sub

'Ab hier gehört alles zu den BulletLevel-Tools

Function CheckAllowIndent() As Boolean
    CheckAllowIndent = True
    On Error Resume Next
    If ActiveWindow.Selection.ShapeRange.Count > 1 Then CheckAllowIndent = False
End Function

Function GroupAllowShapesType() As Boolean
    GroupAllowShapesType = True
    On Error Resume Next
    If (ActiveWindow.Selection.ShapeRange.Type = msoGroup) Then GroupAllowShapesType = False
End Function

Function CheckAllowShapesType()
    CheckAllowShapesType = True
    On Error Resume Next
    If Not (ActiveWindow.Selection.ShapeRange.Type = msoAutoShape _
        Or ActiveWindow.Selection.ShapeRange.Type = msoTable _
        Or ActiveWindow.Selection.ShapeRange.Type = msoTextBox) Then CheckAllowShapesType = False
End Function

Sub getStartEndPosition(rngSelection As Selection)

    '*** This sub identifies the start paragraph number and length of paragraphs selected.
Dim pCount, i
    startPos = 0
    endPos = 0
    pCount = 0
    
    With rngSelection
        '*** To check if the selection type is cursor or region.
        If .TextRange.Paragraphs.Count = 0 Then
            For i = 1 To .ShapeRange.Item(1).TextFrame.TextRange.Paragraphs.Count
                If .ShapeRange.Item(1).TextFrame.TextRange.Paragraphs(i).Start + _
                    .ShapeRange.Item(1).TextFrame.TextRange.Paragraphs(i).Length > .TextRange.Start Then
                    startPos = i 'Get the start position of the paragraph
                    lastPos = i
                    pCount = pCount + 1
                    endPos = pCount 'Get the length of paragraph
                    Exit For
                End If
            Next
        Else
            For i = 1 To .ShapeRange.Item(1).TextFrame.TextRange.Paragraphs.Count
                If .ShapeRange.Item(1).TextFrame.TextRange.Paragraphs(i).Start + _
                    .ShapeRange.Item(1).TextFrame.TextRange.Paragraphs(i).Length > .TextRange.Start Then
                    If pCount >= .TextRange.Paragraphs.Count Then Exit For 'if the selected range is greater than the actual count, then exit for
                    If startPos = 0 Then startPos = i 'Get the start position of the paragraph
                    lastPos = i
                    pCount = pCount + 1
                    endPos = pCount 'Get the length of paragraphs
                End If
            Next
        End If
    End With

End Sub

Private Sub CurrentTextParaPositionWithinTable(rngSelection As Selection)
Dim oTbl As Table
Dim oTR As TextRange2
Dim ColIndex As Long
Dim RowIndex As Long
Dim i As Long, paraCount As Long
Dim pCount As Long
startPos = 0
endPos = 0
pCount = 0
If rngSelection.ShapeRange(1).HasTable Then
    Set oTbl = rngSelection.ShapeRange.Table
    For RowIndex = 1 To oTbl.Rows.Count
        For ColIndex = 1 To oTbl.Columns.Count
            If oTbl.Cell(RowIndex, ColIndex).Selected Then
                Set oTR = oTbl.Cell(RowIndex, ColIndex).Shape.TextFrame2.TextRange
                With ActiveWindow.Selection.TextRange
                    paraCount = .Paragraphs.Count
                    If .Paragraphs.Count = 0 Then paraCount = 1
                    For i = 1 To oTR.Paragraphs.Count
                        If oTR.Paragraphs(i).Start + oTR.Paragraphs(i).Length > .Start Then
                            If pCount >= paraCount Then Exit For 'if the selected range is greater than the actual count, then exit for
                            If startPos = 0 Then startPos = i 'Get the start position of the paragraph
                            lastPos = i
                            pCount = pCount + 1
                            endPos = pCount 'Get the length of paragraphs
                            colTable = ColIndex
                            rowTable = RowIndex
                        End If
                    Next i
                End With
            End If
        Next
    Next
End If

End Sub

Sub IndentBlock(Optional selectType As String = "Null")
'*** This sub sets the indents if the whole shape is selected.

Dim i As Integer, indentLevel As Integer

With ActiveWindow.Selection
    'Iterate through each paragraph in the block and set the indents.
    For i = 1 To .TextRange.Paragraphs.Count
        With .ShapeRange.Item(1).TextFrame.TextRange.Paragraphs(Start:=i, Length:=1)
            'Set the Indent Level for each paragraph
            .indentLevel = indentPos
            'Set Font RGB for the text.
            'Uncomment the next 6 lines, when text should always become black
            'With .Font
                '.Name = "Arial"
                'If textBold = False Then .Bold = msoFalse
                'If textBold = True Then .Bold = msoCTrue
                '.Color.RGB = RGB(fontColorRed, fontColorGreen, fontColorBlue)
            'End With
        End With
        'Set the bullets, left margin and first margin for the block with TextFrame2
        With .ShapeRange.Item(1).TextFrame2.TextRange.Paragraphs(Start:=i, Length:=1)
            .ParagraphFormat.Alignment = ppAlignLeft
            
            With ActiveWindow.Selection.ShapeRange.Item(1).TextFrame.TextRange.ParagraphFormat.Bullet
                .Type = ppBulletUnnumbered
                If bulletVisible = True Then .Visible = msoCTrue
                If bulletVisible = False Then .Visible = msoFalse
                .UseTextColor = msoFalse
                .UseTextFont = msoFalse
                With .Font
                    .Name = bulletFontName 'Changed to Global Variable "Arial"
                    If textBold = False Then .Bold = msoFalse
                    If textBold = True Then .Bold = msoCTrue
                    .Color.RGB = RGB(bulletColorRed, bulletColorGreen, bulletColorBlue)
                End With
                .RelativeSize = bulletSize
                If bulletCharacter > 0 Then .Character = bulletCharacter
            End With
            'Set Indent Margin for each paragraph
            Let indentLevel = ActiveWindow.Selection.ShapeRange.Item(1).TextFrame.TextRange.Paragraphs(Start:=i, Length:=1).indentLevel
            Let .Parent.Ruler.Levels(indentLevel).FirstMargin = IndentFirstMargin ' 28.346439
            Let .Parent.Ruler.Levels(indentLevel).LeftMargin = IndentLeftMargin  '42.519658
        End With
    Next
End With
End Sub

Sub IndentBlockTable(Optional selectType As String = "Null")
'*** This sub sets the indents if the whole table is selected

Dim i As Integer, indentLevel As Integer
Dim oTbl As Table
Dim oTR As TextRange2
Dim ColIndex As Long
Dim RowIndex As Long
Dim paraCount As Long
Dim pCount As Long
startPos = 0
endPos = 0
pCount = 0

If ActiveWindow.Selection.ShapeRange(1).HasTable Then
    Set oTbl = ActiveWindow.Selection.ShapeRange.Table
    For RowIndex = 1 To oTbl.Rows.Count
        For ColIndex = 1 To oTbl.Columns.Count
            If oTbl.Cell(RowIndex, ColIndex).Selected Then
                Set oTR = oTbl.Cell(RowIndex, ColIndex).Shape.TextFrame2.TextRange
                With oTR
                   colTable = ColIndex
                    rowTable = RowIndex
                    For i = 1 To oTR.Paragraphs.Count
                        With oTbl.Cell(rowTable, colTable).Shape.TextFrame.TextRange.Paragraphs(Start:=i, Length:=1)
                            'Set the Indent Level for each paragraph
                            .indentLevel = indentPos
                            'Set Font RGB for the text.
                            'Uncomment the next 6 lines, when text should always become black
                            'With .Font
                                '.Name = "Arial"
                                'If textBold = False Then .Bold = msoFalse
                                'If textBold = True Then .Bold = msoCTrue
                                '.Color.RGB = RGB(fontColorRed, fontColorGreen, fontColorBlue)
                            'End With
                        End With
                        'Set the bullets, left margin and first margin for the block with TextFrame2
                        With oTbl.Cell(rowTable, colTable).Shape.TextFrame.TextRange.Paragraphs(Start:=i, Length:=1)
                            .ParagraphFormat.Alignment = ppAlignLeft
                            
                            With oTbl.Cell(rowTable, colTable).Shape.TextFrame.TextRange.ParagraphFormat.Bullet
                                .Type = ppBulletUnnumbered
                                If bulletVisible = True Then .Visible = msoCTrue
                                If bulletVisible = False Then .Visible = msoFalse
                                .UseTextColor = msoFalse
                                .UseTextFont = msoFalse
                                With .Font
                                    .Name = bulletFontName 'Changed to Global Variable "Arial"
                                    If textBold = False Then .Bold = msoFalse
                                    If textBold = True Then .Bold = msoCTrue
                                    .Color.RGB = RGB(bulletColorRed, bulletColorGreen, bulletColorBlue)
                                End With
                                .RelativeSize = bulletSize
                                If bulletCharacter > 0 Then .Character = bulletCharacter
                            End With
                            'Set Indent Margin for each paragraph
                            Let indentLevel = oTbl.Cell(rowTable, colTable).Shape.TextFrame.TextRange.Paragraphs(Start:=i, Length:=1).indentLevel
                            Let .Parent.Ruler.Levels(indentLevel).FirstMargin = IndentFirstMargin ' 28.346439
                            Let .Parent.Ruler.Levels(indentLevel).LeftMargin = IndentLeftMargin  '42.519658
                        End With
                    Next
                End With
            End If
        Next
    Next
End If
End Sub

Sub IndentSelection(Optional selectType As String = "Null")
'*** This sub sets the indents if parts of the text are selected.

Dim lastCount As Long, i As Integer, indentLevel As Integer
lastCount = 1
    
    With ActiveWindow.Selection
        If .Type = ppSelectionText Then
            If startPos <= 0 Then MsgBox "Please do not place the cursor to final position of the text box", vbOKOnly, Application.ActiveWindow.Caption: Exit Sub
            'Iterate through each paragraph in the selected text and set the indents.
            For i = startPos To lastPos
                'Set the Indent Level for each paragraph
                With ActiveWindow.Selection.ShapeRange.Item(1).TextFrame.TextRange.Paragraphs(Start:=i, Length:=1)
                    .indentLevel = indentPos
            'Set Font RGB for the text.
            'Uncomment the next 6 lines, when text should always become black
                    'With .Font
                        '.Name = "Arial"
                        'If textBold = False Then .Bold = msoFalse
                        'If textBold = True Then .Bold = msoCTrue
                        '.Color.RGB = RGB(fontColorRed, fontColorGreen, fontColorBlue)
                    'End With
                End With
                'Set the bullets, left margin and first margin for the block with TextFrame2
                With .ShapeRange.Item(1).TextFrame2.TextRange.Paragraphs(Start:=i, Length:=1)
                    .ParagraphFormat.Alignment = ppAlignLeft
                    With ActiveWindow.Selection.ShapeRange.Item(1).TextFrame.TextRange.Paragraphs(Start:=i, Length:=1).ParagraphFormat.Bullet
                        .Type = ppBulletUnnumbered
                        If bulletVisible = True Then .Visible = msoCTrue
                        If bulletVisible = False Then .Visible = msoFalse
                        .UseTextColor = msoFalse
                        .UseTextFont = msoFalse
                        With .Font
                            .Name = bulletFontName ' Changed to Global Variable"Wingdings"
                            If textBold = False Then .Bold = msoFalse
                            If textBold = True Then .Bold = msoCTrue
                            .Color.RGB = RGB(bulletColorRed, bulletColorGreen, bulletColorBlue)
                        End With
                        .RelativeSize = bulletSize
                        If bulletCharacter > 0 Then .Character = bulletCharacter
                    End With
                    Let indentLevel = ActiveWindow.Selection.ShapeRange.Item(1).TextFrame.TextRange.Paragraphs(Start:=i, Length:=1).indentLevel
                    If indentLevel > 0 Then
                        Let .Parent.Ruler.Levels(indentLevel).FirstMargin = IndentFirstMargin
                        Let .Parent.Ruler.Levels(indentLevel).LeftMargin = IndentLeftMargin
                    Else
                        Let .Parent.Ruler.Levels(i).FirstMargin = IndentFirstMargin
                        Let .Parent.Ruler.Levels(i).LeftMargin = IndentLeftMargin
                    End If
                End With
            Next
        End If
    End With
End Sub

Sub IndentSelectionTable(Optional selectType As String = "Null")
'*** This sub sets the indents if parts of the text inside the table are selected.

Dim lastCount As Long, i As Integer, indentLevel As Integer
Dim oTbl As Table
Dim oTR As TextRange2
lastCount = 1

    With ActiveWindow.Selection
        Set oTbl = .ShapeRange.Table
        Set oTR = oTbl.Cell(rowTable, colTable).Shape.TextFrame2.TextRange
        If .Type = ppSelectionText Then
            If startPos <= 0 Then MsgBox "Please do not place the cursor to final position of the text box", vbOKOnly, Application.ActiveWindow.Caption: Exit Sub
            'Iterate through each paragraph in the selected text and set the indents.
            For i = startPos To lastPos
                'Set the Indent Level for each paragraph
                With oTbl.Cell(rowTable, colTable).Shape.TextFrame.TextRange.Paragraphs(Start:=i, Length:=1)
                    .indentLevel = indentPos
            'Set Font RGB for the text.
            'Uncomment the next 6 lines, when text should always become black
                    'With .Font
                        '.Name = "Arial"
                        'If textBold = False Then .Bold = msoFalse
                        'If textBold = True Then .Bold = msoCTrue
                        '.Color.RGB = RGB(fontColorRed, fontColorGreen, fontColorBlue)
                    'End With
                End With
                'Set the bullets, left margin and first margin for the block with TextFrame2
                With oTR.Paragraphs(Start:=i, Length:=1)
                    .ParagraphFormat.Alignment = ppAlignLeft
                    With oTbl.Cell(rowTable, colTable).Shape.TextFrame.TextRange.Paragraphs(Start:=i, Length:=1).ParagraphFormat.Bullet
                        .Type = ppBulletUnnumbered
                        If bulletVisible = True Then .Visible = msoCTrue
                        If bulletVisible = False Then .Visible = msoFalse
                        .UseTextColor = msoFalse
                        .UseTextFont = msoFalse
                        With .Font
                            .Name = bulletFontName ' Changed to Global Variable"Wingdings"
                            If textBold = False Then .Bold = msoFalse
                            If textBold = True Then .Bold = msoCTrue
                            .Color.RGB = RGB(bulletColorRed, bulletColorGreen, bulletColorBlue)
                        End With
                        .RelativeSize = bulletSize
                        If bulletCharacter > 0 Then .Character = bulletCharacter
                    End With
                    Let indentLevel = oTbl.Cell(rowTable, colTable).Shape.TextFrame.TextRange.Paragraphs(Start:=i, Length:=1).indentLevel
                    If indentLevel > 0 Then
                        Let .Parent.Ruler.Levels(indentLevel).FirstMargin = IndentFirstMargin
                        Let .Parent.Ruler.Levels(indentLevel).LeftMargin = IndentLeftMargin
                    Else
                        Let .Parent.Ruler.Levels(i).FirstMargin = IndentFirstMargin
                        Let .Parent.Ruler.Levels(i).LeftMargin = IndentLeftMargin
                    End If
                End With
            Next
        End If
    End With
End Sub

'Hier endet das Zubehör zu den BulletLevel-Tools

Public Sub Backup(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
Dim Pres As Presentation
Dim rng As SlideRange
Dim sld As Slide
Dim phd As Shape
Dim shp As Shape
Dim iCount As Integer
Dim lngPos As Long
Dim LeftRight As Single
Dim TopBottom As Single

lngPos = ActiveWindow.Selection.SlideRange(1).SlideIndex

Set Pres = ActivePresentation
    For Each sld In ActivePresentation.Slides
        If sld.Tags("BACKUPDIVIDER") = "YES" Then
            iCount = iCount + 1
        End If
    Next sld

Set rng = ActiveWindow.Selection.SlideRange
    
Select Case iCount
Case Is > 0
    rng.Cut
    DoEvents
    Pres.Slides.Paste -1
    DoEvents
Case 0
    Set sld = Pres.Slides.Add(Pres.Slides.Count + 1, ppLayoutBlank)
    sld.Tags.Add "BACKUPDIVIDER", "YES"
    
    LeftRight = flexWidth(ActivePresentation)
    TopBottom = flexHeight(ActivePresentation)
    
    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)
    Set shp = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=10, Width:=10, Height:=41.952729)
    shp.Fill.Visible = msoTrue
    shp.Fill.Transparency = 0
    shp.Fill.ForeColor.RGB = RGB(9, 91, 164)
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(9, 91, 164)
    shp.Line.Weight = 0.75

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    shp.TextFrame2.TextRange.Font.Size = 16
Else
    shp.TextFrame2.TextRange.Font.Size = 18
End If

    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Text = "Backup"
    shp.TextFrame2.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignLeft
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.Font.Name = "Arial"
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.Font.Italic = msoFalse
    shp.TextFrame2.TextRange.Font.UnderlineStyle = msoNoUnderline
    shp.TextFrame2.Orientation = msoTextOrientationHorizontal
    shp.TextFrame2.MarginBottom = 7.0866097
    shp.TextFrame2.MarginLeft = (LeftRight - phd.Width) / 2
    shp.TextFrame2.MarginRight = 7.0866097
    shp.TextFrame2.MarginTop = 7.0866097
    shp.TextFrame2.WordWrap = msoTrue
    
    shp.Width = LeftRight
    shp.Top = TopBottom / 2 - shp.Height / 2
    
    rng.Cut
    DoEvents
    Pres.Slides.Paste -1
    DoEvents
End Select

ActiveWindow.View.GotoSlide (lngPos)

End Sub

Public Sub CommAdd(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

    Dim shp As Shape
    Dim sld As Slide
    'Comment field

On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
    MsgBox "This function cannot be used for several slides at the same time"
    Exit Sub
Else

    Set sld = Application.ActiveWindow.View.Slide
    Set shp = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=104.88182, Width:=198.42507, Height:=28.913368)
    shp.Fill.Visible = msoTrue
    shp.Fill.Transparency = 0
    shp.Fill.ForeColor.RGB = RGB(211, 61, 95)
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(255, 255, 255)
    shp.Line.Weight = 0.75
    shp.Tags.Add "COMMENT", "YES"
    shp.Shadow.Style = msoShadowStyleOuterShadow
    shp.Shadow.ForeColor.RGB = RGB(0, 0, 0)
    shp.Shadow.Transparency = 0.6
    shp.Shadow.Size = 100
    shp.Shadow.Blur = 4
    shp.Shadow.OffsetX = 2.1
    shp.Shadow.OffsetY = 2.1
    shp.Select
    
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Characters.Text = "Comment: "
    shp.TextFrame2.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignLeft
    shp.TextFrame2.VerticalAnchor = msoAnchorTop
    shp.TextFrame2.TextRange.Font.Size = 12
    shp.TextFrame2.TextRange.Font.Name = "Arial"
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.Font.Italic = msoFalse
    shp.TextFrame2.TextRange.Font.UnderlineStyle = msoNoUnderline
    shp.TextFrame2.Orientation = msoTextOrientationHorizontal
    shp.TextFrame2.MarginBottom = 7.0866097
    shp.TextFrame2.MarginLeft = 7.0866097
    shp.TextFrame2.MarginRight = 7.0866097
    shp.TextFrame2.MarginTop = 7.0866097
    shp.TextFrame2.WordWrap = msoTrue
    shp.TextFrame2.AutoSize = ppAutoSizeShapeToFitText
    shp.TextFrame2.TextRange.Select

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub ConfidON(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    'Confidential stamp on Master edited for Ommax
    
    Set shp = Application.ActivePresentation.SlideMaster.Shapes.AddShape(Type:=msoShapeRectangle, Left:=307.2754, Top:=524.40912, Width:=105.44875, Height:=15.590541)

    shp.TextFrame2.TextRange.Text = "CONFIDENTIAL"
    shp.TextFrame2.TextRange.Font.Name = "Gotham Medium"
    shp.TextFrame2.TextRange.Font.Size = 12
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(225, 65, 180)
    shp.TextFrame2.TextRange.Font.Bold = msoFalse
    shp.TextFrame2.TextRange.Font.Italic = msoFalse
    shp.TextFrame2.TextRange.Font.UnderlineStyle = msoNoUnderline
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = ppAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.MarginBottom = 5.6692878
    shp.TextFrame2.MarginLeft = 5.6692878
    shp.TextFrame2.MarginRight = 5.6692878
    shp.TextFrame2.MarginTop = 5.6692878
    shp.TextFrame2.AutoSize = msoAutoSizeNone
    shp.TextFrame2.WordWrap = msoFalse
    
    shp.Fill.Visible = msoFalse
    shp.Fill.Transparency = 0
    shp.Line.Visible = msoFalse
    shp.Tags.Add "CONFIDMASTER", "YES"

End Sub

Public Sub DraftON(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    'Draft stamp on Master edited for Ommax

    Set shp = Application.ActivePresentation.SlideMaster.Shapes.AddShape(Type:=msoShapeRectangle, Left:=540.85005, Top:=0, Width:=150.80305, Height:=15.590541)

    shp.TextFrame2.TextRange.Text = "DRAFT – for discussion only"
    shp.TextFrame2.TextRange.Font.Name = "Gotham Medium"
    shp.TextFrame2.TextRange.Font.Size = 10
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(225, 65, 180)
    shp.TextFrame2.TextRange.Font.Bold = msoFalse
    shp.TextFrame2.TextRange.Font.Italic = msoFalse
    shp.TextFrame2.TextRange.Font.UnderlineStyle = msoNoUnderline
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = ppAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.MarginBottom = 5.6692878
    shp.TextFrame2.MarginLeft = 5.6692878
    shp.TextFrame2.MarginRight = 5.6692878
    shp.TextFrame2.MarginTop = 5.6692878
    shp.TextFrame2.AutoSize = msoAutoSizeNone
    shp.TextFrame2.WordWrap = msoFalse

    shp.Fill.Visible = msoFalse
    shp.Fill.Transparency = 0
    shp.Line.Visible = msoFalse
    shp.Tags.Add "DRAFTMASTER", "YES"

End Sub

Public Sub Examples(control As IRibbonControl)

    ActivePresentation.FollowHyperlink Address:="http://www.presix.de/downloads/Presix-Template-Examples.pdf"

End Sub

Public Sub Favorites(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim target As Presentation

    Set target = Presentations.Open("C:\Users\Chef\Desktop\PRESIX_Favorites.pptx")

End Sub

Public Sub FooterAdd(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim phd As Shape
    Dim shp As Shape
    Dim L As Long
    Dim TopBottom As Single
 'Footer on Master
 
    For L = ActivePresentation.SlideMaster.Shapes.Count To 1 Step -1
        If ActivePresentation.SlideMaster.Shapes(L).Tags("FOOTERMASTER") = "YES" Then
            ActivePresentation.SlideMaster.Shapes(L).Delete
        Else
        
        End If
    Next L

TopBottom = flexHeight(ActivePresentation)

    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)
    Set shp = Application.ActivePresentation.SlideMaster.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=10, Height:=10.488182)


    shp.TextFrame2.TextRange.Text = ActivePresentation.Name
    shp.TextFrame2.TextRange.Font.Name = "Arial"
    shp.TextFrame2.TextRange.Font.Size = 7
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(150, 150, 150)
    shp.TextFrame2.TextRange.Font.Bold = msoFalse
    shp.TextFrame2.TextRange.Font.Italic = msoFalse
    shp.TextFrame2.TextRange.Font.UnderlineStyle = msoNoUnderline
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.MarginBottom = 0
    shp.TextFrame2.MarginLeft = 0
    shp.TextFrame2.MarginRight = 0
    shp.TextFrame2.MarginTop = 0
    shp.TextFrame2.WordWrap = msoFalse
    
    shp.Fill.Visible = msoFalse
    shp.Fill.Transparency = 0
    shp.Line.Visible = msoFalse
    shp.Tags.Add "FOOTERMASTER", "YES"
    shp.Left = phd.Left
    shp.Width = phd.Width

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    shp.Top = TopBottom - (shp.Height * 3 / 2)
Else
    shp.Top = TopBottom - (shp.Height * 2)
End If

End Sub

Public Sub FormatSub(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    
    If ActiveWindow.Selection.Type = ppSelectionNone Then
        'No slide or not shape is selected by user
        MsgBox "Please select (only) a Title Placeholder"
        
    ElseIf ActiveWindow.Selection.Type = ppSelectionSlides Then
        'User selects nothing: Macros treats subtitle in title placeholder as wanted
        If ActiveWindow.Selection.SlideRange.Shapes.Placeholders.Count > 0 Then
            Call applyFormatting(ActiveWindow.Selection.SlideRange.Shapes.Placeholders(1))
        Else
        MsgBox "Please select (only) a Title Placeholder"
        End If
        
    ElseIf ActiveWindow.Selection.ShapeRange.Count = 1 Then
        If ActiveWindow.Selection.ShapeRange(1).Type = msoPlaceholder Then
            If ActiveWindow.Selection.ShapeRange(1).PlaceholderFormat.Type = ppPlaceholderTitle Or ActiveWindow.Selection.ShapeRange(1).PlaceholderFormat.Type = ppPlaceholderCenterTitle Then
                'User selects title placeholder: Macro treats subtitle in title placeholder as wanted
                Call applyFormatting(ActiveWindow.Selection.ShapeRange(1))
            Else
            MsgBox "Please select (only) a Title Placeholder"
            End If
        Else
        MsgBox "Please select (only) a Title Placeholder"
        End If
        
    ElseIf ActiveWindow.Selection.ShapeRange.Count > 1 Then
        'User selects title placeholder and other shape: Message box appears, nothing else happens
        'User selects one or more shapes, but none is a title placeholder:  Message box appears, nothing else happens
        
        MsgBox "Please select (only) a Title Placeholder"
    End If
    
End Sub

'apply Formatting gehört zu FormatSub

Sub applyFormatting(targetShape As Object)
    
    Dim shText$, startIdx%
    
    shText = targetShape.TextFrame.TextRange.Text
    
    startIdx = InStr(1, shText, Chr(11))
    
    If startIdx = 0 Then Exit Sub
    
    With targetShape.TextFrame.TextRange.Characters(startIdx + 1, Len(shText) - startIdx)
        .Font.Bold = msoFalse
        .Font.Italic = msoFalse
        .Font.Underline = msoFalse
        .Font.Size = 16
        .Font.Name = "Arial"
        .Font.Color.RGB = RGB(9, 91, 164)
    End With
    
End Sub

Public Sub Harvey0of8(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

    Dim sld As Slide
    Dim phd As Shape
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim Ball As Shape
    
On Error GoTo ErrMsg

    Set sld = Application.ActiveWindow.View.Slide
    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    ElseIf ActiveWindow.Selection.Type <> ppSelectionShapes And ActiveWindow.Selection.Type <> ppSelectionText Then
        GoTo DoTheShit
    ElseIf ActiveWindow.Selection.ShapeRange.Count > 1 Then
        MsgBox "Please select only one Harvey ball for replacement"
        Exit Sub
    ElseIf ActiveWindow.Selection.ShapeRange.Count = 1 And ActiveWindow.Selection.ShapeRange.Type <> msoGroup And ActiveWindow.Selection.ShapeRange(1).Tags("HARVEY") <> "YES" Then
        MsgBox "Please select a Harvey ball for replacement"
        Exit Sub
    Else
        GoTo DoBoth
    End If
    
DoBoth:
    Set shp1 = ActiveWindow.Selection.ShapeRange(1)

    Set Ball = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=14.173219, Top:=113.38576, Width:=28.346439, Height:=28.346439)
    With Ball
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
        End With
        .Tags.Add "HARVEY", "YES"
        .Select
    End With
    
            ActiveWindow.Selection.ShapeRange(1).Height = shp1.Height
            ActiveWindow.Selection.ShapeRange(1).Width = shp1.Width
            ActiveWindow.Selection.ShapeRange(1).Line.Weight = shp1.Line.Weight
            ActiveWindow.Selection.ShapeRange(1).Top = shp1.Top
            ActiveWindow.Selection.ShapeRange(1).Left = shp1.Left

            shp1.Delete
            
            Ball.LockAspectRatio = msoTrue
            Ball.Select (msoTrue)
    
            Exit Sub
            
DoTheShit:
    Set Ball = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=14.173219, Top:=113.38576, Width:=28.346439, Height:=28.346439)
    With Ball
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
        End With
        .Tags.Add "HARVEY", "YES"
        .LockAspectRatio = msoTrue
        .Top = phd.Top
    End With
        
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub Harvey1of8(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

    Dim sld As Slide
    Dim phd As Shape
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim Ball As Shape
    Dim Harv As Shape
    
On Error GoTo ErrMsg

    Set sld = Application.ActiveWindow.View.Slide
    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    ElseIf ActiveWindow.Selection.Type <> ppSelectionShapes And ActiveWindow.Selection.Type <> ppSelectionText Then
        GoTo DoTheShit
    ElseIf ActiveWindow.Selection.ShapeRange.Count > 1 Then
        MsgBox "Please select only one Harvey ball for replacement"
        Exit Sub
    ElseIf ActiveWindow.Selection.ShapeRange.Count = 1 And ActiveWindow.Selection.ShapeRange.Type <> msoGroup And ActiveWindow.Selection.ShapeRange(1).Tags("HARVEY") <> "YES" Then
        MsgBox "Please select a Harvey ball for replacement"
        Exit Sub
    Else
        GoTo DoBoth
    End If

DoBoth:
    Set shp1 = ActiveWindow.Selection.ShapeRange(1)

    Set Ball = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=14.173219, Top:=155.90541, Width:=28.346439, Height:=28.346439)
    With Ball
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
        End With
    End With
    
    Set Harv = sld.Shapes.AddShape(Type:=msoShapeArc, Left:=28.346439, Top:=155.90541, Width:=14.173219, Height:=14.173219)
    With Harv
        .Adjustments(1) = 270
        .Adjustments(2) = 315
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
        End With
    End With
    
            Ball.Select (True)
            Harv.Select (False)
            ActiveWindow.Selection.ShapeRange.Group.Select


            ActiveWindow.Selection.ShapeRange(1).Height = shp1.Height
            ActiveWindow.Selection.ShapeRange(1).Width = shp1.Width
            ActiveWindow.Selection.ShapeRange(1).Line.Weight = shp1.Line.Weight
            ActiveWindow.Selection.ShapeRange(1).Top = shp1.Top
            ActiveWindow.Selection.ShapeRange(1).Left = shp1.Left
            ActiveWindow.Selection.ShapeRange.LockAspectRatio = msoTrue
            
            shp1.Delete
            
            Exit Sub
            
DoTheShit:
    Set Ball = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=14.173219, Top:=155.90541, Width:=28.346439, Height:=28.346439)
    With Ball
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
        End With
    End With
    
    Set Harv = sld.Shapes.AddShape(Type:=msoShapeArc, Left:=28.346439, Top:=155.90541, Width:=14.173219, Height:=14.173219)
    With Harv
        .Adjustments(1) = 270
        .Adjustments(2) = 315
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
        End With
    End With
    
    Ball.Select (True)
    Harv.Select (False)
    ActiveWindow.Selection.ShapeRange.Group.Select
    ActiveWindow.Selection.ShapeRange.LockAspectRatio = msoTrue
    ActiveWindow.Selection.ShapeRange.Top = phd.Top + 30
    ActiveWindow.Selection.Unselect
    
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub Harvey2of8(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

    Dim sld As Slide
    Dim phd As Shape
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim Ball As Shape
    Dim Harv As Shape
    
On Error GoTo ErrMsg

    Set sld = Application.ActiveWindow.View.Slide
    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    ElseIf ActiveWindow.Selection.Type <> ppSelectionShapes And ActiveWindow.Selection.Type <> ppSelectionText Then
        GoTo DoTheShit
    ElseIf ActiveWindow.Selection.ShapeRange.Count > 1 Then
        MsgBox "Please select only one Harvey ball for replacement"
        Exit Sub
    ElseIf ActiveWindow.Selection.ShapeRange.Count = 1 And ActiveWindow.Selection.ShapeRange.Type <> msoGroup And ActiveWindow.Selection.ShapeRange(1).Tags("HARVEY") <> "YES" Then
        MsgBox "Please select a Harvey ball for replacement"
        Exit Sub
    Else
        GoTo DoBoth
    End If

DoBoth:
    Set shp1 = ActiveWindow.Selection.ShapeRange(1)

    Set Ball = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=14.173219, Top:=198.42507, Width:=28.346439, Height:=28.346439)
    With Ball
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
        End With
    End With
    
    Set Harv = sld.Shapes.AddShape(Type:=msoShapeArc, Left:=28.346439, Top:=198.42507, Width:=14.173219, Height:=14.173219)
    With Harv
        .Adjustments(1) = 270
        .Adjustments(2) = 0
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
        End With
    End With
    
            Ball.Select (True)
            Harv.Select (False)
            ActiveWindow.Selection.ShapeRange.Group.Select


            ActiveWindow.Selection.ShapeRange(1).Height = shp1.Height
            ActiveWindow.Selection.ShapeRange(1).Width = shp1.Width
            ActiveWindow.Selection.ShapeRange(1).Line.Weight = shp1.Line.Weight
            ActiveWindow.Selection.ShapeRange(1).Top = shp1.Top
            ActiveWindow.Selection.ShapeRange(1).Left = shp1.Left
            ActiveWindow.Selection.ShapeRange.LockAspectRatio = msoTrue
            
            shp1.Delete
            
            Exit Sub
            
DoTheShit:
    Set Ball = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=14.173219, Top:=198.42507, Width:=28.346439, Height:=28.346439)
    With Ball
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
        End With
    End With
    
    Set Harv = sld.Shapes.AddShape(Type:=msoShapeArc, Left:=28.346439, Top:=198.42507, Width:=14.173219, Height:=14.173219)
    With Harv
        .Adjustments(1) = 270
        .Adjustments(2) = 0
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
        End With
    End With
    
    Ball.Select (True)
    Harv.Select (False)
    ActiveWindow.Selection.ShapeRange.Group.Select
    ActiveWindow.Selection.ShapeRange.LockAspectRatio = msoTrue
    ActiveWindow.Selection.ShapeRange.Top = phd.Top + 30 + 30
    ActiveWindow.Selection.Unselect
    
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub Harvey3of8(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

    Dim sld As Slide
    Dim phd As Shape
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim Ball As Shape
    Dim Harv As Shape
    
On Error GoTo ErrMsg

    Set sld = Application.ActiveWindow.View.Slide
    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    ElseIf ActiveWindow.Selection.Type <> ppSelectionShapes And ActiveWindow.Selection.Type <> ppSelectionText Then
        GoTo DoTheShit
    ElseIf ActiveWindow.Selection.ShapeRange.Count > 1 Then
        MsgBox "Please select only one Harvey ball for replacement"
        Exit Sub
    ElseIf ActiveWindow.Selection.ShapeRange.Count = 1 And ActiveWindow.Selection.ShapeRange.Type <> msoGroup And ActiveWindow.Selection.ShapeRange(1).Tags("HARVEY") <> "YES" Then
        MsgBox "Please select a Harvey ball for replacement"
        Exit Sub
    Else
        GoTo DoBoth
    End If

DoBoth:
    Set shp1 = ActiveWindow.Selection.ShapeRange(1)

    Set Ball = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=14.173219, Top:=240.94473, Width:=28.346439, Height:=28.346439)
    With Ball
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
        End With
    End With
    
    Set Harv = sld.Shapes.AddShape(Type:=msoShapeArc, Left:=28.346439, Top:=240.94473, Width:=14.173219, Height:=14.173219)
    With Harv
        .Adjustments(1) = 270
        .Adjustments(2) = 45
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
        End With
    End With
    
            Ball.Select (True)
            Harv.Select (False)
            ActiveWindow.Selection.ShapeRange.Group.Select


            ActiveWindow.Selection.ShapeRange(1).Height = shp1.Height
            ActiveWindow.Selection.ShapeRange(1).Width = shp1.Width
            ActiveWindow.Selection.ShapeRange(1).Line.Weight = shp1.Line.Weight
            ActiveWindow.Selection.ShapeRange(1).Top = shp1.Top
            ActiveWindow.Selection.ShapeRange(1).Left = shp1.Left
            ActiveWindow.Selection.ShapeRange.LockAspectRatio = msoTrue
            
            shp1.Delete
            
            Exit Sub
            
DoTheShit:
    Set Ball = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=14.173219, Top:=240.94473, Width:=28.346439, Height:=28.346439)
    With Ball
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
        End With
    End With
    
    Set Harv = sld.Shapes.AddShape(Type:=msoShapeArc, Left:=28.346439, Top:=240.94473, Width:=14.173219, Height:=14.173219)
    With Harv
        .Adjustments(1) = 270
        .Adjustments(2) = 45
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
        End With
    End With
    
    Ball.Select (True)
    Harv.Select (False)
    ActiveWindow.Selection.ShapeRange.Group.Select
    ActiveWindow.Selection.ShapeRange.LockAspectRatio = msoTrue
    ActiveWindow.Selection.ShapeRange.Top = phd.Top + 30 + 30 + 30
    ActiveWindow.Selection.Unselect
    
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub Harvey4of8(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

    Dim sld As Slide
    Dim phd As Shape
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim Ball As Shape
    Dim Harv As Shape
    
On Error GoTo ErrMsg

    Set sld = Application.ActiveWindow.View.Slide
    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    ElseIf ActiveWindow.Selection.Type <> ppSelectionShapes And ActiveWindow.Selection.Type <> ppSelectionText Then
        GoTo DoTheShit
    ElseIf ActiveWindow.Selection.ShapeRange.Count > 1 Then
        MsgBox "Please select only one Harvey ball for replacement"
        Exit Sub
    ElseIf ActiveWindow.Selection.ShapeRange.Count = 1 And ActiveWindow.Selection.ShapeRange.Type <> msoGroup And ActiveWindow.Selection.ShapeRange(1).Tags("HARVEY") <> "YES" Then
        MsgBox "Please select a Harvey ball for replacement"
        Exit Sub
    Else
        GoTo DoBoth
    End If

DoBoth:
    Set shp1 = ActiveWindow.Selection.ShapeRange(1)

    Set Ball = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=14.173219, Top:=283.46439, Width:=28.346439, Height:=28.346439)
    With Ball
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
        End With
    End With
    
    Set Harv = sld.Shapes.AddShape(Type:=msoShapeArc, Left:=28.346439, Top:=283.46439, Width:=14.173219, Height:=14.173219)
    With Harv
        .Adjustments(1) = 270
        .Adjustments(2) = 90
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
        End With
    End With
    
            Ball.Select (True)
            Harv.Select (False)
            ActiveWindow.Selection.ShapeRange.Group.Select


            ActiveWindow.Selection.ShapeRange(1).Height = shp1.Height
            ActiveWindow.Selection.ShapeRange(1).Width = shp1.Width
            ActiveWindow.Selection.ShapeRange(1).Line.Weight = shp1.Line.Weight
            ActiveWindow.Selection.ShapeRange(1).Top = shp1.Top
            ActiveWindow.Selection.ShapeRange(1).Left = shp1.Left
            ActiveWindow.Selection.ShapeRange.LockAspectRatio = msoTrue
            
            shp1.Delete
            
            Exit Sub
            
DoTheShit:
    Set Ball = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=14.173219, Top:=283.46439, Width:=28.346439, Height:=28.346439)
    With Ball
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
        End With
    End With
    
    Set Harv = sld.Shapes.AddShape(Type:=msoShapeArc, Left:=28.346439, Top:=283.46439, Width:=14.173219, Height:=14.173219)
    With Harv
        .Adjustments(1) = 270
        .Adjustments(2) = 90
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
        End With
    End With
    
    Ball.Select (True)
    Harv.Select (False)
    ActiveWindow.Selection.ShapeRange.Group.Select
    ActiveWindow.Selection.ShapeRange.LockAspectRatio = msoTrue
    ActiveWindow.Selection.ShapeRange.Top = phd.Top + 30 + 30 + 30 + 30
    ActiveWindow.Selection.Unselect
    
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub Harvey5of8(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

    Dim sld As Slide
    Dim phd As Shape
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim Ball As Shape
    Dim Harv As Shape
    
On Error GoTo ErrMsg

    Set sld = Application.ActiveWindow.View.Slide
    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    ElseIf ActiveWindow.Selection.Type <> ppSelectionShapes And ActiveWindow.Selection.Type <> ppSelectionText Then
        GoTo DoTheShit
    ElseIf ActiveWindow.Selection.ShapeRange.Count > 1 Then
        MsgBox "Please select only one Harvey ball for replacement"
        Exit Sub
    ElseIf ActiveWindow.Selection.ShapeRange.Count = 1 And ActiveWindow.Selection.ShapeRange.Type <> msoGroup And ActiveWindow.Selection.ShapeRange(1).Tags("HARVEY") <> "YES" Then
        MsgBox "Please select a Harvey ball for replacement"
        Exit Sub
    Else
        GoTo DoBoth
    End If

DoBoth:
    Set shp1 = ActiveWindow.Selection.ShapeRange(1)
    
    Set Ball = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=14.173219, Top:=325.98405, Width:=28.346439, Height:=28.346439)
    With Ball
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
        End With
    End With
    
    Set Harv = sld.Shapes.AddShape(Type:=msoShapeArc, Left:=28.346439, Top:=325.98405, Width:=14.173219, Height:=14.173219)
    With Harv
        .Adjustments(1) = 270
        .Adjustments(2) = 135
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
        End With
    End With
    
            Ball.Select (True)
            Harv.Select (False)
            ActiveWindow.Selection.ShapeRange.Group.Select


            ActiveWindow.Selection.ShapeRange(1).Height = shp1.Height
            ActiveWindow.Selection.ShapeRange(1).Width = shp1.Width
            ActiveWindow.Selection.ShapeRange(1).Line.Weight = shp1.Line.Weight
            ActiveWindow.Selection.ShapeRange(1).Top = shp1.Top
            ActiveWindow.Selection.ShapeRange(1).Left = shp1.Left
            ActiveWindow.Selection.ShapeRange.LockAspectRatio = msoTrue
            
            shp1.Delete
            
            Exit Sub
            
DoTheShit:
    Set Ball = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=14.173219, Top:=325.98405, Width:=28.346439, Height:=28.346439)
    With Ball
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
        End With
    End With
    
    Set Harv = sld.Shapes.AddShape(Type:=msoShapeArc, Left:=28.346439, Top:=325.98405, Width:=14.173219, Height:=14.173219)
    With Harv
        .Adjustments(1) = 270
        .Adjustments(2) = 135
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
        End With
    End With
    
    Ball.Select (True)
    Harv.Select (False)
    ActiveWindow.Selection.ShapeRange.Group.Select
    ActiveWindow.Selection.ShapeRange.LockAspectRatio = msoTrue
    ActiveWindow.Selection.ShapeRange.Top = phd.Top + 30 + 30 + 30 + 30 + 30
    ActiveWindow.Selection.Unselect
    
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub Harvey6of8(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

    Dim sld As Slide
    Dim phd As Shape
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim Ball As Shape
    Dim Harv As Shape
    
On Error GoTo ErrMsg

    Set sld = Application.ActiveWindow.View.Slide
    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    ElseIf ActiveWindow.Selection.Type <> ppSelectionShapes And ActiveWindow.Selection.Type <> ppSelectionText Then
        GoTo DoTheShit
    ElseIf ActiveWindow.Selection.ShapeRange.Count > 1 Then
        MsgBox "Please select only one Harvey ball for replacement"
        Exit Sub
    ElseIf ActiveWindow.Selection.ShapeRange.Count = 1 And ActiveWindow.Selection.ShapeRange.Type <> msoGroup And ActiveWindow.Selection.ShapeRange(1).Tags("HARVEY") <> "YES" Then
        MsgBox "Please select a Harvey ball for replacement"
        Exit Sub
    Else
        GoTo DoBoth
    End If

DoBoth:
    Set shp1 = ActiveWindow.Selection.ShapeRange(1)

    Set Ball = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=14.173219, Top:=368.5037, Width:=28.346439, Height:=28.346439)
    With Ball
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
        End With
    End With
    
    Set Harv = sld.Shapes.AddShape(Type:=msoShapeArc, Left:=28.346439, Top:=368.5037, Width:=14.173219, Height:=14.173219)
    With Harv
        .Adjustments(1) = 270
        .Adjustments(2) = 180
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
        End With
    End With
    
            Ball.Select (True)
            Harv.Select (False)
            ActiveWindow.Selection.ShapeRange.Group.Select


            ActiveWindow.Selection.ShapeRange(1).Height = shp1.Height
            ActiveWindow.Selection.ShapeRange(1).Width = shp1.Width
            ActiveWindow.Selection.ShapeRange(1).Line.Weight = shp1.Line.Weight
            ActiveWindow.Selection.ShapeRange(1).Top = shp1.Top
            ActiveWindow.Selection.ShapeRange(1).Left = shp1.Left
            ActiveWindow.Selection.ShapeRange.LockAspectRatio = msoTrue
            
            shp1.Delete
            
            Exit Sub
            
DoTheShit:
    Set Ball = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=14.173219, Top:=368.5037, Width:=28.346439, Height:=28.346439)
    With Ball
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
        End With
    End With
    
    Set Harv = sld.Shapes.AddShape(Type:=msoShapeArc, Left:=28.346439, Top:=368.5037, Width:=14.173219, Height:=14.173219)
    With Harv
        .Adjustments(1) = 270
        .Adjustments(2) = 180
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
        End With
    End With
    
    Ball.Select (True)
    Harv.Select (False)
    ActiveWindow.Selection.ShapeRange.Group.Select
    ActiveWindow.Selection.ShapeRange.LockAspectRatio = msoTrue
    ActiveWindow.Selection.ShapeRange.Top = phd.Top + 30 + 30 + 30 + 30 + 30 + 30
    ActiveWindow.Selection.Unselect
    
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub Harvey7of8(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

    Dim sld As Slide
    Dim phd As Shape
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim Ball As Shape
    Dim Harv As Shape
    
On Error GoTo ErrMsg
    
    Set sld = Application.ActiveWindow.View.Slide
    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    ElseIf ActiveWindow.Selection.Type <> ppSelectionShapes And ActiveWindow.Selection.Type <> ppSelectionText Then
        GoTo DoTheShit
    ElseIf ActiveWindow.Selection.ShapeRange.Count > 1 Then
        MsgBox "Please select only one Harvey ball for replacement"
        Exit Sub
    ElseIf ActiveWindow.Selection.ShapeRange.Count = 1 And ActiveWindow.Selection.ShapeRange.Type <> msoGroup And ActiveWindow.Selection.ShapeRange(1).Tags("HARVEY") <> "YES" Then
        MsgBox "Please select a Harvey ball for replacement"
        Exit Sub
    Else
        GoTo DoBoth
    End If

DoBoth:
    Set shp1 = ActiveWindow.Selection.ShapeRange(1)

    Set Ball = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=14.173219, Top:=411.02336, Width:=28.346439, Height:=28.346439)
    With Ball
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
        End With
    End With
    
    Set Harv = sld.Shapes.AddShape(Type:=msoShapeArc, Left:=28.346439, Top:=411.02336, Width:=14.173219, Height:=14.173219)
    With Harv
        .Adjustments(1) = 270
        .Adjustments(2) = 225
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
        End With
    End With
    
            Ball.Select (True)
            Harv.Select (False)
            ActiveWindow.Selection.ShapeRange.Group.Select


            ActiveWindow.Selection.ShapeRange(1).Height = shp1.Height
            ActiveWindow.Selection.ShapeRange(1).Width = shp1.Width
            ActiveWindow.Selection.ShapeRange(1).Line.Weight = shp1.Line.Weight
            ActiveWindow.Selection.ShapeRange(1).Top = shp1.Top
            ActiveWindow.Selection.ShapeRange(1).Left = shp1.Left
            ActiveWindow.Selection.ShapeRange.LockAspectRatio = msoTrue
            
            shp1.Delete
            
            Exit Sub
            
DoTheShit:
    Set Ball = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=14.173219, Top:=411.02336, Width:=28.346439, Height:=28.346439)
    With Ball
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
        End With
    End With
    
    Set Harv = sld.Shapes.AddShape(Type:=msoShapeArc, Left:=28.346439, Top:=411.02336, Width:=14.173219, Height:=14.173219)
    With Harv
        .Adjustments(1) = 270
        .Adjustments(2) = 225
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
        End With
    End With
    
    Ball.Select (True)
    Harv.Select (False)
    ActiveWindow.Selection.ShapeRange.Group.Select
    ActiveWindow.Selection.ShapeRange.LockAspectRatio = msoTrue
    ActiveWindow.Selection.ShapeRange.Top = phd.Top + 30 + 30 + 30 + 30 + 30 + 30 + 30
    ActiveWindow.Selection.Unselect
    
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub Harvey8of8(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

    Dim sld As Slide
    Dim phd As Shape
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim Ball As Shape
    
On Error GoTo ErrMsg

    Set sld = Application.ActiveWindow.View.Slide
    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    ElseIf ActiveWindow.Selection.Type <> ppSelectionShapes And ActiveWindow.Selection.Type <> ppSelectionText Then
        GoTo DoTheShit
    ElseIf ActiveWindow.Selection.ShapeRange.Count > 1 Then
        MsgBox "Please select only one Harvey ball for replacement"
        Exit Sub
    ElseIf ActiveWindow.Selection.ShapeRange.Count = 1 And ActiveWindow.Selection.ShapeRange.Type <> msoGroup And ActiveWindow.Selection.ShapeRange(1).Tags("HARVEY") <> "YES" Then
        MsgBox "Please select a Harvey ball for replacement"
        Exit Sub
    Else
        GoTo DoBoth
    End If
    
DoBoth:
    Set shp1 = ActiveWindow.Selection.ShapeRange(1)
        
    Set Ball = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=14.173219, Top:=453.54302, Width:=28.346439, Height:=28.346439)
    With Ball
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
        End With
        .Tags.Add "HARVEY", "YES"
        .Select
    End With
    
            ActiveWindow.Selection.ShapeRange(1).Height = shp1.Height
            ActiveWindow.Selection.ShapeRange(1).Width = shp1.Width
            ActiveWindow.Selection.ShapeRange(1).Line.Weight = shp1.Line.Weight
            ActiveWindow.Selection.ShapeRange(1).Top = shp1.Top
            ActiveWindow.Selection.ShapeRange(1).Left = shp1.Left

            shp1.Delete
            
            Ball.LockAspectRatio = msoTrue
            Ball.Select (msoTrue)
    
            Exit Sub
            
DoTheShit:
    Set Ball = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=14.173219, Top:=453.54302, Width:=28.346439, Height:=28.346439)
    With Ball
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
            .Weight = 0.75
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 64, 146)
        End With
        .Tags.Add "HARVEY", "YES"
        .LockAspectRatio = msoTrue
        .Top = phd.Top + 30 + 30 + 30 + 30 + 30 + 30 + 30 + 30
    End With
        
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub Library(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim target As Presentation
    Dim strpath As String

If ActivePresentation.Path = "" Then
        MsgBox "Please save presentation"
    Exit Sub
End If
    
On Error Resume Next
    
If ActiveWindow.Selection.SlideRange.Count = 0 Then
    ActiveWindow.ViewType = ppViewNotesPage
    ActiveWindow.ViewType = ppViewNormal
End If

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
    MsgBox "This function cannot be used for several slides at the same time"
    Exit Sub
End If
    
On Error GoTo ErMsg
    
    strpath = Environ$("appdata") & "\Microsoft\AddIns\"

'Der Mac-Pfad erfordert, dass auf Macs auf dem Desktop ein Ordner namens "PRESIX-Addin" liegt,
'und die Zieldatei sich darin befindet.
'Entsprechend sind hier die Dateinamen anzupassen.

If (isMac()) Then
  Set target = Presentations.Open("/Users/" & (GetUserNameMac) & "/Desktop/PRESIX-Addin/VWFS-Library-EN.pptx")
Else
    Set target = Presentations.Open(strpath & "OMMAX_Addin_Library.pptx")
End If

Exit Sub
    
ErMsg:
    MsgBox "Error"
    
End Sub

Public Sub NotesAdd(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

    Dim shp As Shape
    Dim sld As Slide
    Dim LeftRight As Single
    'Comment field - Note bei Ommax
    
On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else

LeftRight = flexWidth(ActivePresentation)

    Set sld = Application.ActiveWindow.View.Slide
    Set shp = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=217.41719, Height:=91.842462)
    shp.Left = LeftRight - shp.Width
    shp.Fill.Visible = msoTrue
    shp.Fill.Transparency = 0
    shp.Fill.ForeColor.RGB = RGB(228, 228, 228)
    shp.Line.Visible = msoFalse
    shp.Tags.Add "COMMENT", "YES"
    shp.Select
    
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(3, 3, 3)
    shp.TextFrame2.TextRange.Characters.Text = "Note:" & vbCr & "Text here"
    shp.TextFrame2.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignLeft
    shp.TextFrame2.VerticalAnchor = msoAnchorTop
    shp.TextFrame2.TextRange.Font.Size = 10
    shp.TextFrame2.TextRange.Font.Name = "Proxima Nova"
    shp.TextFrame2.TextRange.Font.Bold = msoFalse
    shp.TextFrame2.TextRange.Font.Italic = msoFalse
    shp.TextFrame2.TextRange.Font.UnderlineStyle = msoNoUnderline
    shp.TextFrame2.Orientation = msoTextOrientationHorizontal
    shp.TextFrame2.MarginBottom = 5.6692878
    shp.TextFrame2.MarginLeft = 8.5039316
    shp.TextFrame2.MarginRight = 8.5039316
    shp.TextFrame2.MarginTop = 5.6692878
    shp.TextFrame2.WordWrap = msoTrue
    shp.TextFrame2.TextRange.Characters(1, 5).Font.Bold = msoTrue
    'shp.TextFrame.AutoSize = ppAutoSizeShapeToFitText
    'shp.TextFrame.TextRange.Select

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub Link01(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

    ActivePresentation.FollowHyperlink Address:="http://www.presix.de"

End Sub

Public Sub Maps(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim target As Presentation

    Set target = Presentations.Open("C:\Users\Chef\Desktop\PRESIX_Maps.pptx")

End Sub

'Die Funktionen getLeft und getTop gehören zu PNAdd

Function getLeft(oPres As Presentation) As Single
    getLeft = oPres.PageSetup.SlideWidth - 40
End Function

Function getTop(oPres As Presentation) As Single
    getTop = oPres.PageSetup.SlideHeight - 30
End Function
 
Public Sub PNAdd(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

    Dim sld As Slide
    Dim shp As Shape
    Dim x As Long
    Dim thisLeft As Single
    Dim thisTop As Single
    thisLeft = getLeft(ActivePresentation)
    thisTop = getTop(ActivePresentation)
    For Each sld In ActivePresentation.Slides
        x = sld.SlideIndex
'Bei Sld.SlideIndex muss ein " - 1" dahinter, wenn Die Präsi mit 0 beginnt
        Set shp = sld.Shapes.AddShape(msoShapeRectangle, thisLeft, thisTop, 30, 20)
        shp.Fill.ForeColor.RGB = RGB(255, 255, 0)
        shp.Line.ForeColor.RGB = RGB(0, 0, 0)
        shp.Line.Weight = 0.75
        shp.TextFrame2.TextRange.Text = "" & x & ""
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        shp.TextFrame2.TextRange.Font.Name = "Arial"
        shp.TextFrame2.TextRange.Font.Size = 12
        shp.TextFrame2.TextRange.Font.Bold = msoTrue
        shp.TextFrame2.TextRange.Paragraphs.ParagraphFormat.Alignment = msoAlignCenter
        shp.Name = "Slidexx"
    Next
End Sub

Public Sub PresixInfo(control As IRibbonControl)

    MsgBox "PRESIX-Addin für OMMAX - Version 1.08.00" & vbCrLf & "(C) Friedemann Laubach 2015 - 2020" & vbCrLf & "www.presix.de"

End Sub

Public Sub Sorry(control As IRibbonControl)

    MsgBox "As there is no slide library linked with this demo version, this function is not available. We're sorry."

End Sub

'BBB - Fixierte

Private Sub FakeObject() 'wichtig für aus externen Dokumenten einzufügende Folien bzw. Objekte
    Dim sld As Slide
    Dim shp As Shape

    On Error Resume Next
    Set sld = Application.ActiveWindow.View.Slide
    Set shp = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=1, Height:=1)
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
    shp.Name = "FakeObject"
    
    For Each sld In ActivePresentation.Slides
        If shp.Name = "FakeObject" Then shp.Delete
    Next
End Sub

Public Sub AlgSizeC(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
Dim shp As Shape
Dim shp1 As Shape

    On Error GoTo err

    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please select at least two shapes"
        Exit Sub
    End If

Set shp1 = ActiveWindow.Selection.ShapeRange(1)
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Width = shp1.Width
        shp.Left = shp1.Left + ((shp1.Width - shp.Width) / 2)
    Next shp
    Exit Sub
    
err:
    MsgBox "Please select at least two shapes"
         
End Sub

Public Sub AlgSizeM(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
Dim shp As Shape
Dim shp1 As Shape

    On Error GoTo err

    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please select at least two shapes"
        Exit Sub
    End If

Set shp1 = ActiveWindow.Selection.ShapeRange(1)
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Height = shp1.Height
        shp.Top = shp1.Top + ((shp1.Height - shp.Height) / 2)
    Next shp
    Exit Sub
    
err:
    MsgBox "Please select at least two shapes"
         
End Sub

Public Sub AlignB(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
Dim shp As Shape
Dim shp1 As Shape

    On Error GoTo err

    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please select at least two shapes"
        Exit Sub
    End If

Set shp1 = ActiveWindow.Selection.ShapeRange(1)
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Top = shp1.Top + (shp1.Height - shp.Height)
    Next shp
    Exit Sub
    
err:
    MsgBox "Please select at least two shapes"
         
End Sub

Public Sub AlignC(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
Dim shp As Shape
Dim shp1 As Shape

    On Error GoTo err

    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please select at least two shapes"
        Exit Sub
    End If

Set shp1 = ActiveWindow.Selection.ShapeRange(1)
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Left = shp1.Left + ((shp1.Width - shp.Width) / 2)
    Next shp
    Exit Sub
    
err:
    MsgBox "Please select at least two shapes"
         
End Sub

Public Sub AlignL(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
Dim shp As Shape
Dim shp1 As Shape

    On Error GoTo err

    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please select at least two shapes"
        Exit Sub
    End If

Set shp1 = ActiveWindow.Selection.ShapeRange(1)
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Left = shp1.Left
    Next shp
    Exit Sub
    
err:
    MsgBox "Please select at least two shapes"
         
End Sub

Public Sub AlignM(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
Dim shp As Shape
Dim shp1 As Shape

    On Error GoTo err

    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please select at least two shapes"
        Exit Sub
    End If

Set shp1 = ActiveWindow.Selection.ShapeRange(1)
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Top = shp1.Top + ((shp1.Height - shp.Height) / 2)
    Next shp
    Exit Sub
    
err:
    MsgBox "Please select at least two shapes"
         
End Sub

Public Sub AlignR(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
Dim shp As Shape
Dim shp1 As Shape

    On Error GoTo err

    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please select at least two shapes"
        Exit Sub
    End If

Set shp1 = ActiveWindow.Selection.ShapeRange(1)
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Left = shp1.Left + (shp1.Width - shp.Width)
    Next shp
    Exit Sub
    
err:
    MsgBox "Please select at least two shapes"
         
End Sub

Public Sub AlignT(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
Dim shp As Shape
Dim shp1 As Shape

    On Error GoTo err

    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please select at least two shapes"
        Exit Sub
    End If

Set shp1 = ActiveWindow.Selection.ShapeRange(1)
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Top = shp1.Top
    Next shp
    Exit Sub
    
err:
    MsgBox "Please select at least two shapes"
         
End Sub

Public Sub AnimAllDel(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim oeff As Effect
    Dim i As Integer
    Dim osld As Slide
    
If MsgBox("Do you want to delete ALL animations from entire presentation?", vbYesNo) <> vbYes Then Exit Sub

    For Each osld In ActivePresentation.Slides
        If Val(Application.Version) < 10 Then
            For i = 1 To osld.Shapes.Count
                osld.Shapes(i).AnimationSettings.Animate = msoFalse
            Next i
        Else
            For i = osld.TimeLine.MainSequence.Count To 1 Step -1
                osld.TimeLine.MainSequence(i).Delete
            Next i
            
'Remove triggers
            For i = osld.TimeLine.InteractiveSequences.Count To 1 Step -1
                For Each oeff In osld.TimeLine.InteractiveSequences(i)
                    oeff.Delete
                Next oeff
            Next i
        End If
        
    Next osld
 End Sub
 
 Public Sub ArrowTip(control As IRibbonControl)
     'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    Dim shp1 As Shape
    Dim lCount As Long
    
On Error GoTo err

    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please select at least two shapes"
        Exit Sub
    End If
    
    Set shp1 = ActiveWindow.Selection.ShapeRange(1)

    For Each shp In ActiveWindow.Selection.ShapeRange
    For lCount = 1 To shp1.Adjustments.Count
        shp.Adjustments(lCount) = shp1.Adjustments(lCount)
    Next
    Next shp
    Exit Sub
    
err:
    MsgBox "Please select at least two shapes"
    
End Sub

Public Sub ArrowTipHeight(control As IRibbonControl)
    'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    Dim shp1 As Shape
    Dim lCount As Long
    
On Error GoTo err

    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please select at least two shapes"
        Exit Sub
    End If
    
    Set shp1 = ActiveWindow.Selection.ShapeRange(1)

    For Each shp In ActiveWindow.Selection.ShapeRange
        'Shp.Width = ActiveWindow.Selection.ShapeRange(1).Width
'Wenn Height anstatt Width, die obere Zeile deaktivieren und die folgende aktivieren
        shp.Height = ActiveWindow.Selection.ShapeRange(1).Height
    For lCount = 1 To shp1.Adjustments.Count
        shp.Adjustments(lCount) = shp1.Adjustments(lCount)
    Next
    Next shp
    Exit Sub
    
err:
    MsgBox "Please select at least two shapes"
    
End Sub

Public Sub ArrowTipWidth(control As IRibbonControl)
    'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    Dim shp1 As Shape
    Dim lCount As Long
    
On Error GoTo err

    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please select at least two shapes"
        Exit Sub
    End If
    
    Set shp1 = ActiveWindow.Selection.ShapeRange(1)

    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Width = ActiveWindow.Selection.ShapeRange(1).Width
'Wenn Height anstatt Width, die obere Zeile deaktivieren und die folgende aktivieren
        'shp.Height = ActiveWindow.Selection.ShapeRange(1).Height
    For lCount = 1 To shp1.Adjustments.Count
        shp.Adjustments(lCount) = shp1.Adjustments(lCount)
    Next
    Next shp
    Exit Sub
    
err:
    MsgBox "Please select at least two shapes"
    
End Sub

Public Sub AutoSize(control As IRibbonControl)
    'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    On Error GoTo err
    For Each shp In ActiveWindow.Selection.ShapeRange
    If shp.TextFrame.AutoSize = ppAutoSizeNone Then
        shp.TextFrame.AutoSize = ppAutoSizeShapeToFitText
        Else
    If shp.TextFrame.AutoSize = ppAutoSizeShapeToFitText Then
        shp.TextFrame.AutoSize = ppAutoSizeNone
        Else
    If shp.TextFrame.AutoSize = ppAutoSizeMixed Then
    'Nothing
    End If
    End If
    End If
    Next shp
    Exit Sub
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub AutoSub(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim Message As String
    Dim Title As String
    Dim Default As String
    Dim txtRange As TextRange2
    Dim txtrangeSS As TextRange2
    Dim myValue As String
     
    'Input box
    Message = "Insert your subscripted text"
    Title = "Subscript Input"
    Default = "1"
    myValue = InputBox(Message, Title, Default) 'Auf Wunsch hier ein + " " dahinter, das löst das H2O-Problem, aber auf Kosten eines Leerschritts
     
    'Handles if User cancels
    If StrPtr(myValue) = False Then GoTo UserCancels
     
    On Error GoTo err
     
    Set txtRange = ActiveWindow.Selection.TextRange2
    Set txtrangeSS = txtRange.InsertAfter(myValue)
    txtrangeSS.Font.Subscript = True
    'turn OFF SS
    CommandBars.ExecuteMso ("Subscript")

    Exit Sub
    
err:
    Exit Sub
    
UserCancels:
    Exit Sub

End Sub

Public Sub AutoSuper(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim Message As String
    Dim Title As String
    Dim Default As String
    Dim txtRange As TextRange2
    Dim txtrangeSS As TextRange2
    Dim myValue As String

    'Input box
    Message = "Insert your superscripted text"
    Title = "Superscript Input"
    Default = "1"
    myValue = InputBox(Message, Title, Default) 'Auf Wunsch hier ein + " " dahinter, das löst das H2O-Problem, aber auf Kosten eines Leerschritts
     
    'Handles if User cancels
    If StrPtr(myValue) = False Then GoTo UserCancels
     
    On Error GoTo err
     
    Set txtRange = ActiveWindow.Selection.TextRange2
    Set txtrangeSS = txtRange.InsertAfter(myValue)
    txtrangeSS.Font.Superscript = True
    'turn OFF SS
    CommandBars.ExecuteMso ("Superscript")
     
    Exit Sub
    
err:
    Exit Sub
    
UserCancels:
    Exit Sub

End Sub

Public Sub BackNorm(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
With ActiveWindow.Selection.SlideRange
     .FollowMasterBackground = msoTrue
End With

End Sub

Public Sub ByteCountDel(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

Dim sld As Slide
Dim L As Long

On Error Resume Next

For Each sld In ActivePresentation.Slides
    For L = sld.Shapes.Count To 1 Step -1
        If sld.Shapes(L).Tags("BYTECOUNT") = "YES" Then sld.Shapes(L).Delete
    Next L
Next sld

End Sub

Public Sub CommDel(control As IRibbonControl)
    'YYY If Not Init Then Exit Sub 'ZZZ
    Dim sld As Slide
    Dim L As Long
    If MsgBox("Do you want to delete ALL comments from the entire presentation?", vbYesNo) <> vbYes Then Exit Sub
    On Error Resume Next
    For Each sld In ActivePresentation.Slides
        For L = sld.Shapes.Count To 1 Step -1
            If sld.Shapes(L).Tags("COMMENT") = "YES" Then sld.Shapes(L).Delete
        Next L
    Next sld
End Sub

Public Sub ConfidOFF(control As IRibbonControl)
    'YYY If Not Init Then Exit Sub 'ZZZ
    Dim L As Long

    For L = ActivePresentation.SlideMaster.Shapes.Count To 1 Step -1
        If ActivePresentation.SlideMaster.Shapes(L).Tags("CONFIDMASTER") = "YES" Then
            ActivePresentation.SlideMaster.Shapes(L).Delete
        End If
    Next L
        Exit Sub

End Sub

Public Sub CopySelToNew(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

If (isMac()) Then
    MsgBox "Mac does not support this function. We are sorry."
Else
    Call CopySelToNewWin
End If

End Sub

Private Sub CopySelToNewWin()

Dim otemp As Presentation
Dim oPres As Presentation
Dim strName As String
Dim L As Long
Dim raySlides() As Long
Dim SlideNumbers() As Long

    On Error GoTo ErMsg

Set oPres = ActivePresentation

Call zaptags(oPres)

strName = CreateObject("Scripting.FileSystemObject").GetBaseName(oPres.Name) & " S_"
ReDim raySlides(1 To ActiveWindow.Selection.SlideRange.Count)
ReDim SlideNumbers(1 To ActiveWindow.Selection.SlideRange.Count)
For L = 1 To ActiveWindow.Selection.SlideRange.Count
SlideNumbers(L) = ActiveWindow.Selection.SlideRange(L).SlideIndex
ActiveWindow.Selection.SlideRange(L).Tags.Add "SELECTED", "YES"
Next L

Call QuickSort(SlideNumbers, LBound(SlideNumbers), UBound(SlideNumbers))

For L = LBound(SlideNumbers) To UBound(SlideNumbers)
If L <> UBound(SlideNumbers) Then
strName = strName & SlideNumbers(L) & "+"
Else
strName = strName & SlideNumbers(L)
End If
Next L

' make a copy
oPres.SaveCopyAs Environ("TEMP") & "\" & strName & ".pptx"
'open the copy
Set otemp = Presentations.Open(Environ("TEMP") & "\" & strName & ".pptx")
'delete unwanted slides
For L = otemp.Slides.Count To 1 Step -1
Debug.Print otemp.Slides(L).Tags("SELECTED")
If otemp.Slides(L).Tags("SELECTED") <> "YES" Then otemp.Slides(L).Delete
Next L
otemp.Save

Exit Sub
    
ErMsg:
    MsgBox "Please do not place the cursor between two slides"
 
End Sub

Sub zaptags(oPres) 'gehört zu CopySelNew und MailSel und MailSelPDF
Dim osld As Slide
On Error Resume Next
For Each osld In oPres.Slides
osld.Tags.Delete ("SELECTED")
Next osld
End Sub

Private Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long) 'gehört zu CopySelNew und MailSel und MailSelPDF

  Dim pivot As Variant
  Dim tmpSwap As Variant
  Dim tmpLow As Long
  Dim tmpHi As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)

     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If

  Wend

  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi

End Sub

'changeme gehört zum folgenden DoubleBlank

Sub changeme(sFindMe As String, sSwapme As String)

Dim osld As Slide
Dim oshp As Shape
Dim otemp As TextRange
Dim otext As TextRange
Dim Inewstart As Integer
Dim i As Long
Dim j As Long
Dim x As Long

For Each osld In ActiveWindow.Selection.SlideRange
    For Each oshp In osld.Shapes
        If oshp.HasTextFrame Then
            If oshp.TextFrame.HasText Then
                Set otext = oshp.TextFrame.TextRange
                Set otemp = otext.Replace(sFindMe, sSwapme, , msoFalse, msoFalse)
                Do While Not otemp Is Nothing
                Inewstart = otemp.Start + otemp.Length
                Set otemp = otext.Replace(sFindMe, sSwapme, Inewstart, msoFalse, msoFalse)
                Loop
            End If
        End If
        
        If oshp.HasTable Then
            For i = 1 To oshp.Table.Rows.Count
            For j = 1 To oshp.Table.Columns.Count
                Set otext = oshp.Table.Rows.Item(i).Cells(j).Shape.TextFrame.TextRange
                Set otemp = otext.Replace(sFindMe, sSwapme, , msoFalse, msoFalse)
                Do While Not otemp Is Nothing
                Inewstart = otemp.Start + otemp.Length
                Set otemp = otext.Replace(sFindMe, sSwapme, Inewstart, msoFalse, msoFalse)
                Loop
            Next j
            Next i
        End If
    Next oshp
Next osld
 
For Each osld In ActiveWindow.Selection.SlideRange
    For Each oshp In osld.Shapes
        With oshp
            Select Case .Type
                Case Is = msoGroup
                For x = 1 To .GroupItems.Count
                    If .GroupItems(x).HasTextFrame Then
                        If .GroupItems(x).TextFrame.HasText Then
                            Set otext = oshp.GroupItems(x).TextFrame.TextRange
                            Set otemp = otext.Replace(sFindMe, sSwapme, , msoFalse, msoFalse)
                            Do While Not otemp Is Nothing
                            Inewstart = otemp.Start + otemp.Length
                            Set otemp = otext.Replace(sFindMe, sSwapme, Inewstart, msoFalse, msoFalse)
                            Loop
                        End If
                    End If
                Next x
            End Select
        End With
    Next oshp
Next
 
End Sub

Public Sub DoubleBlank(control As IRibbonControl)
    'YYY If Not Init Then Exit Sub 'ZZZ
 Dim sFindMe As String
 Dim sSwapme As String
 
    On Error GoTo ErMsg
 
 sFindMe = "   "
 'change this to suit
 sSwapme = " "
 Call changeme(sFindMe, sSwapme)
 sFindMe = "  "
 'change this to suit
 sSwapme = " "
 Call changeme(sFindMe, sSwapme)
 
Exit Sub
    
ErMsg:
    MsgBox "Please do not place the cursor between two slides"
 
End Sub

Public Sub DraftOFF(control As IRibbonControl)
    'YYY If Not Init Then Exit Sub 'ZZZ
    Dim L As Long

    For L = ActivePresentation.SlideMaster.Shapes.Count To 1 Step -1
        If ActivePresentation.SlideMaster.Shapes(L).Tags("DRAFTMASTER") = "YES" Then
            ActivePresentation.SlideMaster.Shapes(L).Delete
        End If
    Next L
        Exit Sub

End Sub

Public Sub DubShapeStyle(control As IRibbonControl)
    'YYY If Not Init Then Exit Sub 'ZZZ
Dim shp As Shape
Dim shp1 As Shape

    On Error GoTo err

    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
    MsgBox "Please select at least two shapes (no tables)"
        Exit Sub
    End If

Set shp1 = ActiveWindow.Selection.ShapeRange(1)

shp1.PickUp

    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Apply
    Next shp
    Exit Sub

err:
    MsgBox "Please select at least two shapes (no tables)"
    
End Sub

Public Sub DubTextOnly(control As IRibbonControl)
    'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    Dim shp1 As Shape
    Dim i As Integer

    On Error GoTo err

    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please select at least two shapes (no tables)"
        Exit Sub
    End If
    
Set shp1 = ActiveWindow.Selection.ShapeRange(1)

shp1.TextFrame2.TextRange.Copy
DoEvents
shp1.Tags.Add "Deselect", "yes"

    For Each shp In ActiveWindow.Selection.ShapeRange
        If shp.Tags("Deselect") = "yes" Then
        Else
        With shp
        With .TextFrame
            For i = 1 To 9
            With .Ruler
                .Levels(i).FirstMargin = 0
                .Levels(i).LeftMargin = 0
            End With
            Next
            End With
            With .TextFrame2
            With .TextRange
                .ParagraphFormat.Bullet.Type = ppBulletNone
                .PasteSpecial msoClipboardFormatPlainText
            End With
        End With
        End With
        DoEvents
        End If
    Next shp
    For Each shp In ActiveWindow.Selection.ShapeRange
        If shp.Tags("Deselect") = "yes" Then
        shp.Tags.Delete "Deselect"
        End If
    Next shp
    Exit Sub
    
err:
    MsgBox "Please select at least two shapes (no tables)"
    
End Sub

Public Sub DubTextWith(control As IRibbonControl)
    'YYY If Not Init Then Exit Sub 'ZZZ
Dim shp As Shape
Dim shp1 As Shape

    On Error GoTo err

    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please select at least two shapes (no tables)"
        Exit Sub
    End If

Set shp1 = ActiveWindow.Selection.ShapeRange(1)

shp1.PickUp
shp1.TextFrame.TextRange.Copy
DoEvents

    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Apply
        shp.TextFrame.TextRange.Paste
        DoEvents
        
    Next shp
    Exit Sub

err:
    MsgBox "Please select at least two shapes (no tables)"
    
End Sub

Public Sub EdgeB(control As IRibbonControl)
    'YYY If Not Init Then Exit Sub 'ZZZ
Dim shp As Shape
Dim shp1 As Shape

    On Error GoTo err

    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please select at least two shapes"
        Exit Sub
    End If

Set shp1 = ActiveWindow.Selection.ShapeRange(1)
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Height = shp1.Height - (shp.Top - shp1.Top)
    Next shp
    Exit Sub
    
err:
    MsgBox "Please select at least two shapes"
         
End Sub

Public Sub EdgeL(control As IRibbonControl)
    'YYY If Not Init Then Exit Sub 'ZZZ
Dim shp As Shape
Dim shp1 As Shape

    On Error GoTo err

    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please select at least two shapes"
        Exit Sub
    End If

Set shp1 = ActiveWindow.Selection.ShapeRange(1)
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Width = shp1.Width - (shp1.Width - (shp.Width + (shp.Left - shp1.Left)))
        shp.Left = shp1.Left
    Next shp
    Exit Sub
    
err:
    MsgBox "Please select at least two shapes"
         
End Sub

Public Sub EdgeR(control As IRibbonControl)
    'YYY If Not Init Then Exit Sub 'ZZZ
Dim shp As Shape
Dim shp1 As Shape

    On Error GoTo err

    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please select at least two shapes"
        Exit Sub
    End If

Set shp1 = ActiveWindow.Selection.ShapeRange(1)
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Width = shp1.Width - (shp.Left - shp1.Left)
    Next shp
    Exit Sub
    
err:
    MsgBox "Please select at least two shapes"
         
End Sub

Public Sub EdgeT(control As IRibbonControl)
    'YYY If Not Init Then Exit Sub 'ZZZ
Dim shp As Shape
Dim shp1 As Shape

    On Error GoTo err

    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please select at least two shapes"
        Exit Sub
    End If

Set shp1 = ActiveWindow.Selection.ShapeRange(1)
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Height = shp1.Height - (shp1.Height - (shp.Height + (shp.Top - shp1.Top)))
        shp.Top = shp1.Top
    Next shp
    Exit Sub
    
err:
    MsgBox "Please select at least two shapes"
         
End Sub

Public Sub FooterDel(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

    Dim L As Long

    For L = ActivePresentation.SlideMaster.Shapes.Count To 1 Step -1
        If ActivePresentation.SlideMaster.Shapes(L).Tags("FOOTERMASTER") = "YES" Then
            ActivePresentation.SlideMaster.Shapes(L).Delete
        End If
    Next L
        Exit Sub

End Sub

Public Sub HyperDel(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
   
    Dim oSl As Slide
    Dim x As Long

    For Each oSl In ActiveWindow.Selection.SlideRange
        For x = oSl.Hyperlinks.Count To 1 Step -1
            oSl.Hyperlinks(x).Delete
        Next
    Next

End Sub

Public Sub MailPres(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

If (isMac()) Then
    MsgBox "As Outlook for Mac does not support VBA programming, this function is unfortunately unanvailable. We are sorry."
Else
    Call MailPresWin
End If

End Sub

Private Sub MailPresWin()

Dim SourcePres As Presentation
Dim TargetPres As Presentation
Dim OutlookApp As Object
Dim OutlookMessage As Object
Dim TempFileName As Variant
Dim ExternalLinks As Variant
Dim TempFilePath As String
Dim FileExtStr As String
Dim DefaultName As String
Dim UserAnswer As Long

Set SourcePres = ActivePresentation

'Determine Temporary File Path
  TempFilePath = Environ$("temp") & "\"

'Determine Default File Name for InputBox
  If SourcePres.Saved Then
    DefaultName = Left(SourcePres.Name, InStrRev(SourcePres.Name, ".") - 1)
  Else
    DefaultName = SourcePres.Name
  End If

'Ask user for a file name
On Error GoTo err
  TempFileName = InputBox("Please insert a name for your attachment and avoid using special characters", "Input box", Default:=DefaultName)
    
    If TempFileName = False Then Exit Sub 'Handle if user cancels
  
'Determine File Extension
  If SourcePres.Saved = True Then
    FileExtStr = "." & LCase(Right(SourcePres.Name, Len(SourcePres.Name) - InStrRev(SourcePres.Name, ".", , 1)))
  Else
    FileExtStr = ".pptx"
  End If

'Save Temporary Presentation
  SourcePres.SaveCopyAs TempFilePath & TempFileName & FileExtStr
  Set TargetPres = Presentations.Open(TempFilePath & TempFileName & FileExtStr)

'Create Instance of Outlook
  On Error Resume Next
    Set OutlookApp = GetObject(class:="Outlook.Application") 'Handles if Outlook is already open
  err.Clear
    If OutlookApp Is Nothing Then Set OutlookApp = CreateObject(class:="Outlook.Application") 'If not, open Outlook
    
'Create a new email message
  Set OutlookMessage = OutlookApp.CreateItem(0)

'Create Outlook email with attachment
    With OutlookMessage
     .To = ""
     .CC = ""
     .BCC = ""
     .Subject = TempFileName
     .Body = ""
     .Attachments.Add TargetPres.FullName
     .Display
    End With

'Close & Delete the temporary file
  With TargetPres
  .Close
  End With
  Kill TempFilePath & TempFileName & FileExtStr

'Clear Memory
  Set OutlookMessage = Nothing
  Set OutlookApp = Nothing
  
err:

  Exit Sub
  
End Sub

Public Sub MailPresPDF(control As IRibbonControl)
    'YYY If Not Init Then Exit Sub 'ZZZ

If (isMac()) Then
    MsgBox "As Outlook for Mac does not support VBA programming, this function is unfortunately unanvailable. We are sorry."
Else
    Call MailPresPDFWin
End If

End Sub

Private Sub MailPresPDFWin()

Dim SourcePres As Presentation
Dim TargetPres As Presentation
Dim OutlookApp As Object
Dim OutlookMessage As Object
Dim TempFileName As Variant
Dim ExternalLinks As Variant
Dim TempFilePath As String
Dim FileExtStr As String
Dim DefaultName As String
Dim UserAnswer As Long
Dim AttachIt As String

Set SourcePres = ActivePresentation

'Determine Temporary File Path - muss gegebenenfalls geändert werden
  TempFilePath = Environ("USERPROFILE") & "\Desktop\"
  
'Determine Default File Name for InputBox
  If SourcePres.Saved Then
    DefaultName = Left(SourcePres.Name, InStrRev(SourcePres.Name, ".") - 1)
  Else
    DefaultName = SourcePres.Name
  End If

'Ask user for a file name
On Error GoTo err
  TempFileName = InputBox("Please insert a name for your attachment and avoid using special characters", "Input box", Default:=DefaultName)
    
    If TempFileName = False Then Exit Sub 'Handle if user cancels
  
'Determine File Extension
  If SourcePres.Saved = True Then
    FileExtStr = "." & LCase(Right(SourcePres.Name, Len(SourcePres.Name) - InStrRev(SourcePres.Name, ".", , 1)))
  Else
    FileExtStr = ".pptx"
  End If

'Save Temporary Presentation
  SourcePres.SaveCopyAs TempFilePath & TempFileName & FileExtStr
  Set TargetPres = Presentations.Open(TempFilePath & TempFileName & FileExtStr)
  
  
  ActivePresentation.ExportAsFixedFormat ActivePresentation.Path & "\" & TempFileName & ".pdf", ppFixedFormatTypePDF, ppFixedFormatIntentPrint
  
  AttachIt = TempFilePath & TempFileName & ".pdf"

'Create Instance of Outlook
  On Error Resume Next
    Set OutlookApp = GetObject(class:="Outlook.Application") 'Handles if Outlook is already open
  err.Clear
    If OutlookApp Is Nothing Then Set OutlookApp = CreateObject(class:="Outlook.Application") 'If not, open Outlook
    
'Create a new email message
  Set OutlookMessage = OutlookApp.CreateItem(0)

'Create Outlook email with attachment
    With OutlookMessage
     .To = ""
     .CC = ""
     .BCC = ""
     .Subject = TempFileName
     .Body = ""
     .Attachments.Add AttachIt
     .Display
    End With

'Close & Delete the temporary file
  With TargetPres
  .Close
  End With
  Kill TempFilePath & TempFileName & FileExtStr
  Kill AttachIt

'Clear Memory
  Set OutlookMessage = Nothing
  Set OutlookApp = Nothing
  
err:
  
  Exit Sub
  
End Sub

Public Sub MailSel(control As IRibbonControl)
    'YYY If Not Init Then Exit Sub 'ZZZ

If (isMac()) Then
    MsgBox "As Outlook for Mac does not support VBA programming, this function is unfortunately unanvailable. We are sorry."
Else
    Call MailSelWin
End If

End Sub

Private Sub MailSelWin()

Dim OutlookApp As Object
Dim OutlookMessage As Object
Dim otemp As Presentation
Dim oPres As Presentation
Dim strName As String
Dim DefaultName As String
Dim TempFileName As Variant
Dim L As Long
Dim raySlides() As Long
Dim SlideNumbers() As Long

    On Error GoTo ErMsg

Set oPres = ActivePresentation

Call zaptags(oPres)

strName = CreateObject("Scripting.FileSystemObject").GetBaseName(oPres.Name) & " S_"
ReDim raySlides(1 To ActiveWindow.Selection.SlideRange.Count)
ReDim SlideNumbers(1 To ActiveWindow.Selection.SlideRange.Count)
For L = 1 To ActiveWindow.Selection.SlideRange.Count
SlideNumbers(L) = ActiveWindow.Selection.SlideRange(L).SlideIndex
ActiveWindow.Selection.SlideRange(L).Tags.Add "SELECTED", "YES"
Next L

Call QuickSort(SlideNumbers, LBound(SlideNumbers), UBound(SlideNumbers))

For L = LBound(SlideNumbers) To UBound(SlideNumbers)
If L <> UBound(SlideNumbers) Then
strName = strName & SlideNumbers(L) & "+"
Else
strName = strName & SlideNumbers(L)
End If
Next L

'Determine Default File Name for InputBox
DefaultName = strName

'Ask user for a file name
On Error GoTo err
TempFileName = InputBox("Please insert a name for your attachment", "Input box", Default:=DefaultName)
    
If TempFileName = False Then Exit Sub 'Handle if user cancels
  
' make a copy
oPres.SaveCopyAs Environ("TEMP") & "\" & TempFileName & ".pptx"
'open the copy
Set otemp = Presentations.Open(Environ("TEMP") & "\" & TempFileName & ".pptx")
'delete unwanted slides
For L = otemp.Slides.Count To 1 Step -1
Debug.Print otemp.Slides(L).Tags("SELECTED")
If otemp.Slides(L).Tags("SELECTED") <> "YES" Then otemp.Slides(L).Delete
Next L
otemp.Save
otemp.Close

On Error Resume Next
Set OutlookApp = GetObject(class:="Outlook.Application")
err.Clear
If OutlookApp Is Nothing Then Set OutlookApp = CreateObject(class:="Outlook.Application")
On Error GoTo 0
Set OutlookMessage = OutlookApp.CreateItem(0)

On Error Resume Next
With OutlookMessage
.To = "" 'Insert email address here!
.CC = ""
.Subject = TempFileName
.Body = ""
.Attachments.Add Environ("TEMP") & "\" & TempFileName & ".pptx"
.Display

End With

Exit Sub
    
ErMsg:
    MsgBox "Please do not place the cursor between two slides"

err:
  Exit Sub
 
End Sub

Public Sub MailSelPDF(control As IRibbonControl)
    'YYY If Not Init Then Exit Sub 'ZZZ

If (isMac()) Then
    MsgBox "As Outlook for Mac does not support VBA programming, this function is unfortunately unanvailable. We are sorry."
Else
    Call MailSelPDFWin
End If

End Sub

Private Sub MailSelPDFWin()

Dim OutlookApp As Object
Dim OutlookMessage As Object
Dim otemp As Presentation
Dim oPres As Presentation
Dim strName As String
Dim DefaultName As String
Dim TempFileName As Variant
Dim L As Long
Dim raySlides() As Long
Dim SlideNumbers() As Long

    On Error GoTo ErMsg

Set oPres = ActivePresentation

Call zaptags(oPres)

strName = CreateObject("Scripting.FileSystemObject").GetBaseName(oPres.Name) & " S_"
ReDim raySlides(1 To ActiveWindow.Selection.SlideRange.Count)
ReDim SlideNumbers(1 To ActiveWindow.Selection.SlideRange.Count)
For L = 1 To ActiveWindow.Selection.SlideRange.Count
SlideNumbers(L) = ActiveWindow.Selection.SlideRange(L).SlideIndex
ActiveWindow.Selection.SlideRange(L).Tags.Add "SELECTED", "YES"
Next L

Call QuickSort(SlideNumbers, LBound(SlideNumbers), UBound(SlideNumbers))

For L = LBound(SlideNumbers) To UBound(SlideNumbers)
If L <> UBound(SlideNumbers) Then
strName = strName & SlideNumbers(L) & "+"
Else
strName = strName & SlideNumbers(L)
End If
Next L

'Determine Default File Name for InputBox
DefaultName = strName

'Ask user for a file name
On Error GoTo err
TempFileName = InputBox("Please insert a name for your attachment", "Input box", Default:=DefaultName)
    
If TempFileName = False Then Exit Sub 'Handle if user cancels
  
' make a copy
oPres.SaveCopyAs Environ("TEMP") & "\" & TempFileName & ".pptx"
'open the copy
Set otemp = Presentations.Open(Environ("TEMP") & "\" & TempFileName & ".pptx")
'delete unwanted slides
For L = otemp.Slides.Count To 1 Step -1
Debug.Print otemp.Slides(L).Tags("SELECTED")
If otemp.Slides(L).Tags("SELECTED") <> "YES" Then otemp.Slides(L).Delete
Next L
otemp.Save
otemp.ExportAsFixedFormat Environ("TEMP") & "\" & TempFileName & ".pdf", ppFixedFormatTypePDF, ppFixedFormatIntentPrint
otemp.Close

On Error Resume Next
Set OutlookApp = GetObject(class:="Outlook.Application")
err.Clear
If OutlookApp Is Nothing Then Set OutlookApp = CreateObject(class:="Outlook.Application")
On Error GoTo 0
Set OutlookMessage = OutlookApp.CreateItem(0)

On Error Resume Next
With OutlookMessage
.To = "" 'Insert email address here!
.CC = ""
.Subject = TempFileName
.Body = ""
.Attachments.Add Environ("TEMP") & "\" & TempFileName & ".pdf"
.Display

End With

Kill Environ("TEMP") & "\" & TempFileName & ".pdf"

Exit Sub
    
ErMsg:
    MsgBox "Please do not place the cursor between two slides"

err:
  Exit Sub
 
End Sub

Public Sub MiniSlide(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
   
   'Dieses Macro ist nicht Mac-kompatibel ohne die Funktionen IsMac und GetUserNameMac
   
Dim Message As String
Dim Title As String
Dim Default As String
Dim myValue As String
Dim ScriptShell As Object
Dim strpath As String
Dim SW As Long
Dim SH As Long
Dim i As Integer
Dim Pfad As String

On Error Resume Next

Message = "Please enter a name for the new folder" & vbCrLf & "Already existing folders will not be overwritten - in this case nothing happens" & vbCrLf & " " & vbCrLf & "The new folder will be saved to your desktop"
Title = "Create folder name"
Default = "MiniSlides"
myValue = InputBox(Message, Title, Default)

'Handles if User cancels
On Error GoTo UserCancels

If (isMac()) Then
    strpath = "/Users/" & (GetUserNameMac) & "/Desktop/" & myValue & "/"
    MkDir (strpath)
    SW = ActivePresentation.PageSetup.SlideWidth
    SH = ActivePresentation.PageSetup.SlideHeight
    For i = 1 To ActivePresentation.Slides.Count
        ActivePresentation.Slides(i).Export strpath & "/" & "/MySlide" & CStr(i) & ".jpg", "JPG", SW / 2, SH / 2
    Next i
Else
    Set ScriptShell = CreateObject("WScript.Shell")
    strpath = ScriptShell.SpecialFolders("Desktop")
    MkDir strpath & "\" & myValue & "\"
    SW = ActivePresentation.PageSetup.SlideWidth
    SH = ActivePresentation.PageSetup.SlideHeight
    For i = 1 To ActivePresentation.Slides.Count
        ActivePresentation.Slides(i).Export strpath & "\" & myValue & "\" & "\MySlide" & CStr(i) & ".jpg", "JPG", SW / 2, SH / 2
    Next i
    Set ScriptShell = Nothing

'Den Teil bis vbMaximizedFocus nur aktivieren, wenn auch Öffnung des Ordners gewünscht
'On Error GoTo err
     
'     Pfad = strpath & "\" & myValue & "\" 'To be replaced by path
'     Shell "explorer.exe /e, " & Pfad, vbMaximizedFocus
End If

    Exit Sub

err:
    MsgBox "Folder not found"

UserCancels:
    Exit Sub
End Sub
 
Public Sub MoveBM(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

Dim phd As Shape
Dim shp As Shape

Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)

    On Error GoTo err
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Top = (phd.Top + phd.Height - shp.Height)
    Next shp
    Exit Sub
err:
    MsgBox "Please select at least one shape"

End Sub

Public Sub MoveLM(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

Dim phd As Shape
Dim shp As Shape

Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)

    On Error GoTo err
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Left = phd.Left
    Next shp
    Exit Sub
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub MoveRM(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

Dim phd As Shape
Dim shp As Shape

Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)

    On Error GoTo err
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Left = (phd.Left + phd.Width - shp.Width)
    Next shp
    Exit Sub
err:
    MsgBox "Please select at least one shape"

End Sub

Public Sub MoveTM(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

Dim phd As Shape
Dim shp As Shape

Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)

    On Error GoTo err
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Top = phd.Top
    Next shp
    Exit Sub
err:
    MsgBox "Please select at least one shape"

End Sub

Public Sub NFZoff(control As IRibbonControl)
    'YYY If Not Init Then Exit Sub 'ZZZ

    Dim L As Long
    For L = ActivePresentation.SlideMaster.Shapes.Count To 1 Step -1
        If ActivePresentation.SlideMaster.Shapes(L).Tags("NOFLYZONE") = "YES" Then
            ActivePresentation.SlideMaster.Shapes(L).Delete
        End If
    Next L
        Exit Sub
End Sub
 
'Die Funktionen flexWidth und flexHeight gehören zu NFZon und Summary und Footer und noch mehr. Besser nicht löschen!
 
Function flexWidth(oPres As Presentation) As Single
    flexWidth = oPres.PageSetup.SlideWidth
End Function

Function flexHeight(oPres As Presentation) As Single
    flexHeight = oPres.PageSetup.SlideHeight
End Function

Public Sub NFZon(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

'Name bei Ommax: Workspace. Farbe für OMMAX geändert, sonst original

'Flexible No-Fly-Zone
Dim phd As Shape
Dim tit As Shape
Dim shp1 As Shape
Dim shp2 As Shape
Dim shp3 As Shape
Dim shp4 As Shape
Dim TopBottom As Single
Dim LeftRight As Single

TopBottom = flexHeight(ActivePresentation)
LeftRight = flexWidth(ActivePresentation)

Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)
Set tit = Application.ActivePresentation.SlideMaster.Shapes.Title

'Left
Set shp1 = Application.ActivePresentation.SlideMaster.Shapes.AddShape(msoShapeRectangle, Left:=0, Top:=0, Width:=10, Height:=10)
With shp1
    .Fill.Visible = msoTrue
    .Line.Visible = msoTrue
    .Fill.ForeColor.RGB = RGB(207, 200, 5)
    .Line.ForeColor.RGB = RGB(207, 200, 5)
    .Line.Weight = 0.75
    .Tags.Add "NOFLYZONE", "YES"
    .Top = (tit.Top + tit.Height)
    .Width = phd.Left
    .Height = TopBottom - (tit.Top + tit.Height)
End With

'Right
Set shp2 = Application.ActivePresentation.SlideMaster.Shapes.AddShape(msoShapeRectangle, Left:=0, Top:=0, Width:=10, Height:=10)
With shp2
    .Fill.Visible = msoTrue
    .Line.Visible = msoTrue
    .Fill.ForeColor.RGB = RGB(207, 200, 5)
    .Line.ForeColor.RGB = RGB(207, 200, 5)
    .Line.Weight = 0.75
    .Tags.Add "NOFLYZONE", "YES"
    .Left = (phd.Left + phd.Width)
    .Top = (tit.Top + tit.Height)
    .Width = LeftRight - (phd.Left + phd.Width)
    .Height = TopBottom - (tit.Top + tit.Height)
End With

'Bottom
Set shp3 = Application.ActivePresentation.SlideMaster.Shapes.AddShape(msoShapeRectangle, Left:=0, Top:=0, Width:=10, Height:=10)
With shp3
    .Fill.Visible = msoTrue
    .Line.Visible = msoTrue
    .Fill.ForeColor.RGB = RGB(207, 200, 5)
    .Line.ForeColor.RGB = RGB(207, 200, 5)
    .Line.Weight = 0.75
    .Tags.Add "NOFLYZONE", "YES"
    .Top = (phd.Top + phd.Height)
    .Width = LeftRight
    .Height = TopBottom - (phd.Top + phd.Height)
End With

'Top
Set shp4 = Application.ActivePresentation.SlideMaster.Shapes.AddShape(msoShapeRectangle, Left:=0, Top:=0, Width:=10, Height:=10)
With shp4
    .Fill.Visible = msoTrue
    .Line.Visible = msoTrue
    .Fill.ForeColor.RGB = RGB(207, 200, 5)
    .Line.ForeColor.RGB = RGB(207, 200, 5)
    .Line.Weight = 0.75
    .Tags.Add "NOFLYZONE", "YES"
    .Top = (tit.Top + tit.Height)
    .Width = LeftRight
    .Height = TopBottom - (tit.Top + tit.Height + phd.Height + shp3.Height)
End With

End Sub

Public Sub NoFill(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim oshp As Shape
    Dim oTbl As Table
    Dim x As Long
    Dim y As Long
    On Error GoTo err
    
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange.Fill
        .Visible = msoFalse
 End With
    
    Else
    For Each oshp In ActiveWindow.Selection.ShapeRange
        If ActiveWindow.Selection.ShapeRange(1).HasTable Then
        Set oTbl = ActiveWindow.Selection.ShapeRange.Table
            For x = 1 To oTbl.Rows.Count
                For y = 1 To oTbl.Columns.Count
                    If oTbl.Cell(x, y).Selected Then
        oTbl.Cell(x, y).Shape.Fill.Visible = msoFalse
                    End If
                Next
            Next
    Else
        oshp.Fill.Visible = msoFalse
    End If
    Next oshp
    End If
    
Exit Sub
 
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub NoLine(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim oshp As Shape
    On Error GoTo err
    
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange.Line
        .Visible = msoFalse
 End With
    
    Else
    For Each oshp In ActiveWindow.Selection.ShapeRange
        oshp.Line.Visible = msoFalse
    Next oshp
    End If
    
Exit Sub
 
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub NoText(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim oshp As Shape
    Dim oTbl As Table
    Dim x As Long
    Dim y As Long
    On Error GoTo err
    
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange.TextFrame.TextRange
        .Delete
 End With
    
    Else
    For Each oshp In ActiveWindow.Selection.ShapeRange
        If ActiveWindow.Selection.ShapeRange(1).HasTable Then
        Set oTbl = ActiveWindow.Selection.ShapeRange.Table
            For x = 1 To oTbl.Rows.Count
                For y = 1 To oTbl.Columns.Count
                    If oTbl.Cell(x, y).Selected Then
        oTbl.Cell(x, y).Shape.TextFrame.TextRange.Delete
                    End If
                Next
            Next
        Else
        oshp.TextFrame.TextRange.Delete
        End If
    Next oshp
    End If
    
Exit Sub
 
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub NotesDel(control As IRibbonControl)
    'YYY If Not Init Then Exit Sub 'ZZZ
    Dim sld As Slide
    Dim L As Long
    'Extra für Ommax, eigentlich nichts anderes als CommDel
    If MsgBox("Do you want to delete ALL notes from the entire presentation?", vbYesNo) <> vbYes Then Exit Sub
    On Error Resume Next
    For Each sld In ActivePresentation.Slides
        For L = sld.Shapes.Count To 1 Step -1
            If sld.Shapes(L).Tags("COMMENT") = "YES" Then sld.Shapes(L).Delete
        Next L
    Next sld
End Sub

Public Sub NumSelShapes(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
Dim x As Long
Dim shp As Shape

If ActiveWindow.Selection.Type = ppSelectionText Then
    GoTo NumberTheShapes
ElseIf ActiveWindow.Selection.Type <> ppSelectionShapes Then
    MsgBox "Please select a shape"
    Exit Sub
End If

NumberTheShapes:
'Define the number to start with
x = 1

'Number the shapes
For Each shp In ActiveWindow.Selection.ShapeRange
    If shp.Type <> msoAutoShape And shp.Type <> msoTextBox Then
        MsgBox "Only shapes and text boxes can be numbered"
        'Exit Sub
    Else
        shp.TextFrame2.TextRange.InsertBefore (x & ". ") '(insert with dot and space and before text)
        x = x + 1 'go on counting up
    End If
Next shp

End Sub

Public Sub ParaZero(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    Dim oTbl As Table
    Dim x As Long
    Dim y As Long
      
On Error GoTo err
      
If ActiveWindow.Selection.Type = ppSelectionNone Then 'Nothing!

ElseIf ActiveWindow.Selection.Type = ppSelectionText Then
    With ActiveWindow.Selection.TextRange
    .ParagraphFormat.SpaceAfter = 0
    End With

Else
    For Each shp In ActiveWindow.Selection.ShapeRange
        If ActiveWindow.Selection.ShapeRange(1).HasTable Then
        Set oTbl = ActiveWindow.Selection.ShapeRange.Table
            For x = 1 To oTbl.Rows.Count
                For y = 1 To oTbl.Columns.Count
                    If oTbl.Cell(x, y).Selected Then
                        With oTbl.Cell(x, y).Shape.TextFrame2.TextRange
                            .ParagraphFormat.SpaceAfter = 0
                        End With
                    End If
                Next
            Next
        Else
            With shp.TextFrame2.TextRange
                .ParagraphFormat.SpaceAfter = 0
            End With
        End If
    Next shp
End If
Exit Sub

err:
    MsgBox "Please deselect tables from your range of objects and edit each of them separately"

End Sub

Public Sub PasteSeveral(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
Dim oSl As Slide
For Each oSl In ActiveWindow.Selection.SlideRange
    On Error Resume Next
    CommandBars.ExecuteMso "PasteSourceFormatting"
    If err.Number <> 0 Then
        oSl.Shapes.Paste
        DoEvents
    End If
    err.Clear
    On Error GoTo 0
Next

End Sub

Public Sub PNDel(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
   
    Dim sld As Slide
    Dim L As Long
    On Error Resume Next
    For Each sld In ActivePresentation.Slides
        For L = sld.Shapes.Count To 1 Step -1
            If sld.Shapes(L).Name = ("Slidexx") Then sld.Shapes(L).Delete
        Next L
    Next sld
End Sub

Public Sub PNHidAdjust(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
 'For v.2007 onwards only
 Dim osld As Slide
 Dim objSN As Shape
 Dim lngNum As Long
 'check all slides
 For Each osld In ActivePresentation.Slides
 'Is it hidden
 If osld.SlideShowTransition.Hidden Then
 osld.HeadersFooters.SlideNumber.Visible = False
 Else
 osld.HeadersFooters.SlideNumber.Visible = True
 Set objSN = getNumber(osld)
 lngNum = lngNum + 1
 If Not objSN Is Nothing Then ' there is a number placeholder
 objSN.TextFrame.TextRange = CStr(lngNum)
 End If
 End If
 Next osld
 End Sub
 
 'getNumber gehört zu PNHidAdjust

 Function getNumber(thisSlide As Slide) As Shape
 For Each getNumber In thisSlide.Shapes
 If getNumber.Type = msoPlaceholder Then
 If getNumber.PlaceholderFormat.Type = ppPlaceholderSlideNumber Then
 'it's the slide number
 Exit Function
 End If
 End If
 Next getNumber
 End Function
 
Public Sub PositionApply(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
If ActiveWindow.Selection.Type <> ppSelectionShapes Then
    GoTo err
End If
    
If ActiveWindow.Selection.ShapeRange.Count <> 1 Then
    GoTo err
End If

With ActiveWindow.Selection.ShapeRange(1)
    .Left = sngleft
    .Top = sngtop
End With

Exit Sub

err:
    MsgBox "Please select exactly one shape"
    Exit Sub

End Sub

Public Sub PositionPickup(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
If ActiveWindow.Selection.Type <> ppSelectionShapes Then
    GoTo err
End If
    
If ActiveWindow.Selection.ShapeRange.Count <> 1 Then
    GoTo err
End If

With ActiveWindow.Selection.ShapeRange(1)
    sngleft = .Left
    sngtop = .Top
End With
    
Exit Sub

err:
    MsgBox "Please select exactly one shape"
    Exit Sub

End Sub

Public Sub ReadGap(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
Dim shp1 As Shape
Dim shp2 As Shape
Dim HoriGap As Long
Dim VertGap As Long

On Error Resume Next
If ActiveWindow.Selection.ShapeRange.Count <> 2 Then
    MsgBox "Please select exactly two shapes"
    Exit Sub
End If
    
Set shp1 = ActiveWindow.Selection.ShapeRange(1)
Set shp2 = ActiveWindow.Selection.ShapeRange(2)
    
HoriGap = (shp2.Left - (shp1.Left + shp1.Width)) / 0.28346
VertGap = (shp2.Top - (shp1.Top + shp1.Height)) / 0.28346

MsgBox "Horizontal gap: " & HoriGap & vbCrLf & "Vertical gap: " & VertGap

End Sub

Public Sub SameHeight(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
Dim shp As Shape

    On Error GoTo err

For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Height = ActiveWindow.Selection.ShapeRange(1).Height
    Next
    
    Exit Sub
err:
    MsgBox "Please select at least two shapes"
End Sub

Public Sub SameWidth(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
Dim shp As Shape

    On Error GoTo err

For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Width = ActiveWindow.Selection.ShapeRange(1).Width
    Next
    
    Exit Sub
err:
    MsgBox "Please select at least two shapes"
End Sub

Private Sub ReplaceNumbers(sFindMe As String, sSwapme As String) 'gehört zu SanitizeNumbers

Dim osld As Slide
Dim oshp As Shape
Dim otemp As TextRange
Dim otext As TextRange
Dim Inewstart As Integer
Dim i As Long
Dim j As Long
Dim x As Long

For Each osld In ActiveWindow.Selection.SlideRange
    For Each oshp In osld.Shapes
        If oshp.HasTextFrame Then
            If oshp.TextFrame.HasText Then
                Set otext = oshp.TextFrame.TextRange
                Set otemp = otext.Replace(sFindMe, sSwapme, , msoFalse, msoFalse)
                Do While Not otemp Is Nothing
                Inewstart = otemp.Start + otemp.Length
                Set otemp = otext.Replace(sFindMe, sSwapme, Inewstart, msoFalse, msoFalse)
                Loop
            End If
        End If
        
        If oshp.HasTable Then
            For i = 1 To oshp.Table.Rows.Count
            For j = 1 To oshp.Table.Columns.Count
                Set otext = oshp.Table.Rows.Item(i).Cells(j).Shape.TextFrame.TextRange
                Set otemp = otext.Replace(sFindMe, sSwapme, , msoFalse, msoFalse)
                Do While Not otemp Is Nothing
                Inewstart = otemp.Start + otemp.Length
                Set otemp = otext.Replace(sFindMe, sSwapme, Inewstart, msoFalse, msoFalse)
                Loop
            Next j
            Next i
        End If
    Next oshp
Next osld
 
For Each osld In ActiveWindow.Selection.SlideRange
    For Each oshp In osld.Shapes
        With oshp
            Select Case .Type
                Case Is = msoGroup
                For x = 1 To .GroupItems.Count
                    If .GroupItems(x).HasTextFrame Then
                        If .GroupItems(x).TextFrame.HasText Then
                            Set otext = oshp.GroupItems(x).TextFrame.TextRange
                            Set otemp = otext.Replace(sFindMe, sSwapme, , msoFalse, msoFalse)
                            Do While Not otemp Is Nothing
                            Inewstart = otemp.Start + otemp.Length
                            Set otemp = otext.Replace(sFindMe, sSwapme, Inewstart, msoFalse, msoFalse)
                            Loop
                        End If
                    End If
                Next x
            End Select
        End With
    Next oshp
Next
 
End Sub

Public Sub SanitizeNumbers(control As IRibbonControl)
    'YYY If Not Init Then Exit Sub 'ZZZ
 Dim sFindMe As String
 Dim sSwapme As String
 
    On Error GoTo ErMsg
 
 sFindMe = "0"
 'change this to suit
 sSwapme = "x"
 Call ReplaceNumbers(sFindMe, sSwapme)
 sFindMe = "1"
 'change this to suit
 sSwapme = "x"
 Call ReplaceNumbers(sFindMe, sSwapme)
  sFindMe = "2"
 'change this to suit
 sSwapme = "x"
 Call ReplaceNumbers(sFindMe, sSwapme)
  sFindMe = "3"
 'change this to suit
 sSwapme = "x"
 Call ReplaceNumbers(sFindMe, sSwapme)
  sFindMe = "4"
 'change this to suit
 sSwapme = "x"
 Call ReplaceNumbers(sFindMe, sSwapme)
  sFindMe = "5"
 'change this to suit
 sSwapme = "x"
 Call ReplaceNumbers(sFindMe, sSwapme)
  sFindMe = "6"
 'change this to suit
 sSwapme = "x"
 Call ReplaceNumbers(sFindMe, sSwapme)
  sFindMe = "7"
 'change this to suit
 sSwapme = "x"
 Call ReplaceNumbers(sFindMe, sSwapme)
  sFindMe = "8"
 'change this to suit
 sSwapme = "x"
 Call ReplaceNumbers(sFindMe, sSwapme)
  sFindMe = "9"
 'change this to suit
 sSwapme = "x"
 Call ReplaceNumbers(sFindMe, sSwapme)
 sFindMe = "0"
 'change this to suit
 sSwapme = "x"
 Call ReplaceNumbers(sFindMe, sSwapme)
 sFindMe = "1"
 'change this to suit
 sSwapme = "x"
 Call ReplaceNumbers(sFindMe, sSwapme)
  sFindMe = "2"
 'change this to suit
 sSwapme = "x"
 Call ReplaceNumbers(sFindMe, sSwapme)
  sFindMe = "3"
 'change this to suit
 sSwapme = "x"
 Call ReplaceNumbers(sFindMe, sSwapme)
  sFindMe = "4"
 'change this to suit
 sSwapme = "x"
 Call ReplaceNumbers(sFindMe, sSwapme)
  sFindMe = "5"
 'change this to suit
 sSwapme = "x"
 Call ReplaceNumbers(sFindMe, sSwapme)
  sFindMe = "6"
 'change this to suit
 sSwapme = "x"
 Call ReplaceNumbers(sFindMe, sSwapme)
  sFindMe = "7"
 'change this to suit
 sSwapme = "x"
 Call ReplaceNumbers(sFindMe, sSwapme)
  sFindMe = "8"
 'change this to suit
 sSwapme = "x"
 Call ReplaceNumbers(sFindMe, sSwapme)
  sFindMe = "9"
 'change this to suit
 sSwapme = "x"
 Call ReplaceNumbers(sFindMe, sSwapme)
 
Exit Sub
    
ErMsg:
    MsgBox "Please do not place the cursor between two slides"
 
End Sub

Public Sub ScaleByP(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    Dim Message As String
    Dim Title As String
    Dim Default As Integer
    Dim myValue As Integer
    Dim x As Long
    
    On Error Resume Next
    
    If ActiveWindow.Selection.ShapeRange.Count < 1 Then
    MsgBox "Please select at least one shape"
        Exit Sub
    End If
    
    Message = "Please enter value in % (but without the %-sign)"
    Title = "Scale by percentage - Input box"
    Default = "100"
    myValue = InputBox(Message, Title, Default)
    
    'Handles if User cancels
    If myValue = False Then Exit Sub
    
    For Each shp In ActiveWindow.Selection.ShapeRange
    If shp.Type = msoGroup Then
        shp.Width = shp.Width / 100 * myValue
        shp.Height = shp.Height / 100 * myValue
            For x = 1 To shp.GroupItems.Count
                If shp.GroupItems(x).HasTextFrame Then
                    shp.GroupItems(x).TextFrame2.TextRange.Font.Size = shp.GroupItems(x).TextFrame2.TextRange.Font.Size / 100 * myValue
                End If
            Next x
    ElseIf shp.LockAspectRatio = msoTrue Then
        shp.LockAspectRatio = msoFalse
        shp.Width = shp.Width / 100 * myValue
        shp.Height = shp.Height / 100 * myValue
            If shp.HasTextFrame Then
                shp.TextFrame.TextRange.Font.Size = shp.TextFrame.TextRange.Font.Size / 100 * myValue
            End If
        shp.LockAspectRatio = msoTrue
    Else
        shp.Width = shp.Width / 100 * myValue
        shp.Height = shp.Height / 100 * myValue
            If shp.HasTextFrame Then
                shp.TextFrame.TextRange.Font.Size = shp.TextFrame.TextRange.Font.Size / 100 * myValue
             End If
    End If
    Next
Exit Sub

End Sub

Public Sub SelByCol(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
   
    Dim shSelected As Shape
    Dim shp As Shape
    Dim c As New Collection
    Dim i As Integer
   
    If ActiveWindow.Selection.Type <> ppSelectionShapes And ActiveWindow.Selection.Type <> ppSelectionText Then
        MsgBox "Please select a shape"
        Exit Sub
    End If
    
    If ActiveWindow.Selection.ShapeRange.Count > 1 Then
        MsgBox "Please select only one shape"
        Exit Sub
    End If

    Set shSelected = ActiveWindow.Selection.ShapeRange(1)

    If shSelected.Type = msoPicture Or shSelected.AutoShapeType = msoShapeMixed Then
        MsgBox "Sorry! For pictures and lines or connectors no fill color is defined"
        Exit Sub
    End If

    i = 1
    For Each shp In ActiveWindow.Selection.SlideRange.Shapes
        If shp.Type = shSelected.Type And shp.Type <> msoPicture Then
            If shp.Fill.ForeColor = shSelected.Fill.ForeColor Then
                c.Add i
            End If
        End If
        i = i + 1
    Next

    If c.Count = 1 Then
        MsgBox "No matching shape found"
    Else
        ActiveWindow.Selection.Unselect
        Dim a() As Integer
        ReDim a(c.Count)
        For i = 1 To c.Count
            a(i) = c(i)
        Next
        ActiveWindow.Selection.SlideRange(1).Shapes.Range(a).Select
    End If
    
End Sub

Public Sub SelByObj(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
   
    Dim shSelected As Shape
    Dim shp As Shape
    Dim c As New Collection
    Dim i As Integer
   
    If ActiveWindow.Selection.Type <> ppSelectionShapes And ActiveWindow.Selection.Type <> ppSelectionText Then
        MsgBox "Please select a shape"
        Exit Sub
    End If
    
    If ActiveWindow.Selection.ShapeRange.Count > 1 Then
        MsgBox "Please select only one shape"
        Exit Sub
    End If

    Set shSelected = ActiveWindow.Selection.ShapeRange(1)

    i = 1
    For Each shp In ActiveWindow.Selection.SlideRange.Shapes
        If shp.Type = shSelected.Type Then
            If shp.AutoShapeType = shSelected.AutoShapeType Then
                c.Add i
            End If
        End If
        i = i + 1
    Next

    If c.Count = 1 Then
        MsgBox "No matching shape found"
    Else
        ActiveWindow.Selection.Unselect
        Dim a() As Integer
        ReDim a(c.Count)
        For i = 1 To c.Count
            a(i) = c(i)
        Next
        ActiveWindow.Selection.SlideRange(1).Shapes.Range(a).Select
    End If
    
End Sub

Public Sub SelByObjCol(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
   
    Dim shSelected As Shape
    Dim shp As Shape
    Dim c As New Collection
    Dim i As Integer
   
    If ActiveWindow.Selection.Type <> ppSelectionShapes And ActiveWindow.Selection.Type <> ppSelectionText Then
        MsgBox "Please select a shape"
        Exit Sub
    End If
    
    If ActiveWindow.Selection.ShapeRange.Count > 1 Then
        MsgBox "Please select only one shape"
        Exit Sub
    End If

    Set shSelected = ActiveWindow.Selection.ShapeRange(1)

    If shSelected.Type = msoPicture Or shSelected.AutoShapeType = msoShapeMixed Then
        MsgBox "Sorry! For pictures and lines or connectors no fill color is defined"
        Exit Sub
    End If

    i = 1
    For Each shp In ActiveWindow.Selection.SlideRange.Shapes
        If shp.Type = shSelected.Type And shp.Type <> msoPicture Then
            If shp.Fill.ForeColor = shSelected.Fill.ForeColor Then
                If shp.AutoShapeType = shSelected.AutoShapeType Then
                    c.Add i
                End If
            End If
        End If
        i = i + 1
    Next

    If c.Count = 1 Then
        MsgBox "No matching shape found"
    Else
        ActiveWindow.Selection.Unselect
        Dim a() As Integer
        ReDim a(c.Count)
        For i = 1 To c.Count
            a(i) = c(i)
        Next
        ActiveWindow.Selection.SlideRange(1).Shapes.Range(a).Select
    End If
    
End Sub

Public Sub SetGapHeight(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim sngGap As Single
    Dim rayShapes() As Shape
    Dim L As Long
    Dim Message As String
    Dim Title As String
    Dim Default As Integer
    Dim myValue As Integer
    Dim Ret As String
    
    On Error Resume Next
    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please select at least two shapes"
        Exit Sub
    End If
    
    Message = "Please enter value in 1/10mm" & vbCrLf & "(Examples:" & vbCrLf & "If you want 1.23cm, type in 123." & vbCrLf & "If you want 0.01cm, type in 1)"
    Title = "Set gap height - Input box"
    Default = "100"
    Ret = InputBox(Message, Title, Default)
    
    'Handles if User cancels
    If Ret = "" Then Exit Sub
    myValue = CInt(Ret)
    
    On Error GoTo 0
    ReDim rayShapes(1 To ActiveWindow.Selection.ShapeRange.Count)
    sngGap = mm2Points(1 * myValue) ' whatever gap
    For L = 1 To ActiveWindow.Selection.ShapeRange.Count
        Set rayShapes(L) = ActiveWindow.Selection.ShapeRange(L)
    Next L
     ' make sure selected shapes are sorted by top value
    Call SortByTop(rayShapes)
     ' set the gap
    For L = 2 To UBound(rayShapes)
        Debug.Print rayShapes(L).Name
        rayShapes(L).Top = rayShapes(L - 1).Top + rayShapes(L - 1).Height + sngGap
    Next L
End Sub
 
Sub SortByTop(ArrayIn As Variant)
     ' sort the shapes based on their top value
    Dim b_Cont As Boolean
    Dim lngCount As Long
    Dim vSwap As Shape
    Do
        b_Cont = False
        For lngCount = LBound(ArrayIn) To UBound(ArrayIn) - 1
            Debug.Print ArrayIn(lngCount).Name
            If ArrayIn(lngCount).Top > ArrayIn(lngCount + 1).Top Then
                Set vSwap = ArrayIn(lngCount)
                Set ArrayIn(lngCount) = ArrayIn(lngCount + 1)
                Set ArrayIn(lngCount + 1) = vSwap
                b_Cont = True
            End If
        Next lngCount
    Loop Until Not b_Cont
     'release objects
    Set vSwap = Nothing
End Sub

Public Sub SetGapWidth(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim sngGap As Single
    Dim rayShapes() As Shape
    Dim L As Long
    Dim Message As String
    Dim Title As String
    Dim Default As Integer
    Dim myValue As Integer
    Dim Ret As String
    
    On Error Resume Next
    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please select at least two shapes"
        Exit Sub
    End If
    
    Message = "Please enter value in 1/10mm" & vbCrLf & "(Examples:" & vbCrLf & "If you want 1.23cm, type in 123." & vbCrLf & "If you want 0.01cm, type in 1)"
    Title = "Set gap width - Input box"
    Default = "100"
    Ret = InputBox(Message, Title, Default)
    
    'Handles if User cancels
    If Ret = "" Then Exit Sub
    myValue = CInt(Ret)
    
    On Error GoTo 0
    ReDim rayShapes(1 To ActiveWindow.Selection.ShapeRange.Count)
    sngGap = mm2Points(1 * myValue) ' whatever gap
    For L = 1 To ActiveWindow.Selection.ShapeRange.Count
        Set rayShapes(L) = ActiveWindow.Selection.ShapeRange(L)
    Next L
     ' make sure selected shapes are sorted by left value
    Call SortByLeft(rayShapes)
     ' set the gap
    For L = 2 To UBound(rayShapes)
        Debug.Print rayShapes(L).Name
        rayShapes(L).Left = rayShapes(L - 1).Left + rayShapes(L - 1).Width + sngGap
    Next L
End Sub
 
Sub SortByLeft(ArrayIn As Variant)
     ' sort the shapes based on their left value
    Dim b_Cont As Boolean
    Dim lngCount As Long
    Dim vSwap As Shape
    Do
        b_Cont = False
        For lngCount = LBound(ArrayIn) To UBound(ArrayIn) - 1
            Debug.Print ArrayIn(lngCount).Name
            If ArrayIn(lngCount).Left > ArrayIn(lngCount + 1).Left Then
                Set vSwap = ArrayIn(lngCount)
                Set ArrayIn(lngCount) = ArrayIn(lngCount + 1)
                Set ArrayIn(lngCount + 1) = vSwap
                b_Cont = True
            End If
        Next lngCount
    Loop Until Not b_Cont
     'release objects
    Set vSwap = Nothing
End Sub

'mm2Points gehört zu den beiden SetGap-Tools
 
Function mm2Points(inVal As Single) As Single
     'convert cm to points
    mm2Points = inVal * 0.28346
End Function

Public Sub SlideMasterCleanUp(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
Dim i As Integer
Dim j As Integer
Dim oPres As Presentation
Set oPres = ActivePresentation
    If MsgBox("Do you want to delete ALL unused master layouts from the presentation?", vbYesNo) <> vbYes Then Exit Sub
On Error Resume Next
With oPres
    For i = 1 To .Designs.Count
        For j = .Designs(i).SlideMaster.CustomLayouts.Count To 1 Step -1
            .Designs(i).SlideMaster.CustomLayouts(j).Delete
        Next
    Next i
End With

End Sub

Public Sub SortLR(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim oshpR As ShapeRange
    Dim L As Long
    Dim rayPOS() As Single

    On Error Resume Next
    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please select at least two shapes"
        Exit Sub
    End If

    Set oshpR = ActiveWindow.Selection.ShapeRange
    ReDim rayPOS(1 To oshpR.Count)
     'add to array
    For L = 1 To oshpR.Count
        rayPOS(L) = oshpR(L).Left
    Next L
     'sort
    Call sortray(rayPOS)
     'apply
    For L = 1 To oshpR.Count
        oshpR(L).Left = rayPOS(L)
    Next

End Sub

Public Sub SortTB(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim oshpR As ShapeRange
    Dim L As Long
    Dim rayPOS() As Single

    On Error Resume Next
    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please select at least two shapes"
        Exit Sub
    End If

    Set oshpR = ActiveWindow.Selection.ShapeRange
    ReDim rayPOS(1 To oshpR.Count)
     'add to array
    For L = 1 To oshpR.Count
        rayPOS(L) = oshpR(L).Top
    Next L
     'sort
    Call sortray(rayPOS)
     'apply
    For L = 1 To oshpR.Count
        oshpR(L).Top = rayPOS(L)
    Next

End Sub
 
Private Sub sortray(ArrayIn As Variant) 'gehört zu SortLR und SortTB

    Dim b_Cont As Boolean
    Dim lngCount As Long
    Dim vSwap As Long

    Do
        b_Cont = False
        For lngCount = LBound(ArrayIn) To UBound(ArrayIn) - 1
            If ArrayIn(lngCount) > ArrayIn(lngCount + 1) Then
                vSwap = ArrayIn(lngCount)
                ArrayIn(lngCount) = ArrayIn(lngCount + 1)
                ArrayIn(lngCount + 1) = vSwap
                b_Cont = True
            End If
        Next lngCount
    Loop Until Not b_Cont

End Sub

Public Sub SpelPresDE(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
 Dim osld As Slide
 Dim oshp As Shape
 Dim notesshp As Shape
 Dim i As Long
 Dim j As Long
 Dim x As Long
 
 For Each osld In ActivePresentation.Slides
 For Each oshp In osld.Shapes
 If oshp.HasTextFrame Then
 oshp.TextFrame2.TextRange.LanguageID = msoLanguageIDGerman
 End If
 If oshp.HasTable Then
             For i = 1 To oshp.Table.Rows.Count
                For j = 1 To oshp.Table.Columns.Count
 oshp.Table.Rows.Item(i).Cells(j).Shape.TextFrame2.TextRange.LanguageID = msoLanguageIDGerman
                 Next j
            Next i
 End If
 Next oshp

 For Each notesshp In osld.NotesPage.Shapes
 If notesshp.HasTextFrame Then
 notesshp.TextFrame2.TextRange.LanguageID = msoLanguageIDGerman
 End If
 Next notesshp
 Next osld
 
    For Each osld In ActivePresentation.Slides
    For Each oshp In osld.Shapes
    
        With oshp
            Select Case .Type
                Case Is = msoGroup
                For x = 1 To .GroupItems.Count
                    If .GroupItems(x).HasTextFrame Then
                         oshp.TextFrame2.TextRange.LanguageID = msoLanguageIDGerman
                    End If
                Next x
            End Select
        End With
    Next oshp
    Next
    
'Wenn Umstellung der Default Language auch gewünscht, dann:
'With ActivePresentation
'   .DefaultLanguageID = msoLanguageIDGerman
'End With

End Sub

Public Sub SpelPresES(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
 Dim osld As Slide
 Dim oshp As Shape
 Dim notesshp As Shape
 Dim i As Long
 Dim j As Long
 Dim x As Long
 
 For Each osld In ActivePresentation.Slides
 For Each oshp In osld.Shapes
 If oshp.HasTextFrame Then
 oshp.TextFrame2.TextRange.LanguageID = msoLanguageIDSpanish
 End If
 If oshp.HasTable Then
             For i = 1 To oshp.Table.Rows.Count
                For j = 1 To oshp.Table.Columns.Count
 oshp.Table.Rows.Item(i).Cells(j).Shape.TextFrame2.TextRange.LanguageID = msoLanguageIDSpanish
                 Next j
            Next i
 End If
 Next oshp

 For Each notesshp In osld.NotesPage.Shapes
 If notesshp.HasTextFrame Then
 notesshp.TextFrame2.TextRange.LanguageID = msoLanguageIDSpanish
 End If
 Next notesshp
 Next osld
 
    For Each osld In ActivePresentation.Slides
    For Each oshp In osld.Shapes
    
        With oshp
            Select Case .Type
                Case Is = msoGroup
                For x = 1 To .GroupItems.Count
                    If .GroupItems(x).HasTextFrame Then
                         oshp.TextFrame2.TextRange.LanguageID = msoLanguageIDSpanish
                    End If
                Next x
            End Select
        End With
    Next oshp
    Next
    
'Wenn Umstellung der Default Language auch gewünscht, dann:
'With ActivePresentation
'   .DefaultLanguageID = msoLanguageIDSpanish
'End With

End Sub

Public Sub SpelPresFR(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
 Dim osld As Slide
 Dim oshp As Shape
 Dim notesshp As Shape
 Dim i As Long
 Dim j As Long
 Dim x As Long
 
 For Each osld In ActivePresentation.Slides
 For Each oshp In osld.Shapes
 If oshp.HasTextFrame Then
 oshp.TextFrame2.TextRange.LanguageID = msoLanguageIDFrench
 End If
 If oshp.HasTable Then
             For i = 1 To oshp.Table.Rows.Count
                For j = 1 To oshp.Table.Columns.Count
 oshp.Table.Rows.Item(i).Cells(j).Shape.TextFrame2.TextRange.LanguageID = msoLanguageIDFrench
                 Next j
            Next i
 End If
 Next oshp

 For Each notesshp In osld.NotesPage.Shapes
 If notesshp.HasTextFrame Then
 notesshp.TextFrame2.TextRange.LanguageID = msoLanguageIDFrench
 End If
 Next notesshp
 Next osld
 
    For Each osld In ActivePresentation.Slides
    For Each oshp In osld.Shapes
    
        With oshp
            Select Case .Type
                Case Is = msoGroup
                For x = 1 To .GroupItems.Count
                    If .GroupItems(x).HasTextFrame Then
                         oshp.TextFrame2.TextRange.LanguageID = msoLanguageIDFrench
                    End If
                Next x
            End Select
        End With
    Next oshp
    Next
    
'Wenn Umstellung der Default Language auch gewünscht, dann:
'With ActivePresentation
'   .DefaultLanguageID = msoLanguageIDFrench
'End With

End Sub

Public Sub SpelPresIT(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
 Dim osld As Slide
 Dim oshp As Shape
 Dim notesshp As Shape
 Dim i As Long
 Dim j As Long
 Dim x As Long
 
 For Each osld In ActivePresentation.Slides
 For Each oshp In osld.Shapes
 If oshp.HasTextFrame Then
 oshp.TextFrame2.TextRange.LanguageID = msoLanguageIDItalian
 End If
 If oshp.HasTable Then
             For i = 1 To oshp.Table.Rows.Count
                For j = 1 To oshp.Table.Columns.Count
 oshp.Table.Rows.Item(i).Cells(j).Shape.TextFrame2.TextRange.LanguageID = msoLanguageIDItalian
                 Next j
            Next i
 End If
 Next oshp

 For Each notesshp In osld.NotesPage.Shapes
 If notesshp.HasTextFrame Then
 notesshp.TextFrame2.TextRange.LanguageID = msoLanguageIDItalian
 End If
 Next notesshp
 Next osld
 
    For Each osld In ActivePresentation.Slides
    For Each oshp In osld.Shapes
    
        With oshp
            Select Case .Type
                Case Is = msoGroup
                For x = 1 To .GroupItems.Count
                    If .GroupItems(x).HasTextFrame Then
                         oshp.TextFrame2.TextRange.LanguageID = msoLanguageIDItalian
                    End If
                Next x
            End Select
        End With
    Next oshp
    Next

'Wenn Umstellung der Default Language auch gewünscht, dann:
'With ActivePresentation
'   .DefaultLanguageID = msoLanguageIDItalian
'End With

End Sub

Public Sub SpelPresUK(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
 Dim osld As Slide
 Dim oshp As Shape
 Dim notesshp As Shape
 Dim i As Long
 Dim j As Long
 Dim x As Long
 
 For Each osld In ActivePresentation.Slides
 For Each oshp In osld.Shapes
 If oshp.HasTextFrame Then
 oshp.TextFrame2.TextRange.LanguageID = msoLanguageIDEnglishUK
 End If
 If oshp.HasTable Then
             For i = 1 To oshp.Table.Rows.Count
                For j = 1 To oshp.Table.Columns.Count
 oshp.Table.Rows.Item(i).Cells(j).Shape.TextFrame2.TextRange.LanguageID = msoLanguageIDEnglishUK
                 Next j
            Next i
 End If
 Next oshp

 For Each notesshp In osld.NotesPage.Shapes
 If notesshp.HasTextFrame Then
 notesshp.TextFrame2.TextRange.LanguageID = msoLanguageIDEnglishUK
 End If
 Next notesshp
 Next osld
 
    For Each osld In ActivePresentation.Slides
    For Each oshp In osld.Shapes
    
        With oshp
            Select Case .Type
                Case Is = msoGroup
                For x = 1 To .GroupItems.Count
                    If .GroupItems(x).HasTextFrame Then
                         oshp.TextFrame2.TextRange.LanguageID = msoLanguageIDEnglishUK
                    End If
                Next x
            End Select
        End With
    Next oshp
    Next
    
'Wenn Umstellung der Default Language auch gewünscht, dann:
'With ActivePresentation
'   .DefaultLanguageID = msoLanguageIDEnglishUK
'End With

End Sub

Public Sub SpelPresUS(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
 Dim osld As Slide
 Dim oshp As Shape
 Dim notesshp As Shape
 Dim i As Long
 Dim j As Long
 Dim x As Long
 
 For Each osld In ActivePresentation.Slides
 For Each oshp In osld.Shapes
 If oshp.HasTextFrame Then
 oshp.TextFrame2.TextRange.LanguageID = msoLanguageIDEnglishUS
 End If
 If oshp.HasTable Then
             For i = 1 To oshp.Table.Rows.Count
                For j = 1 To oshp.Table.Columns.Count
 oshp.Table.Rows.Item(i).Cells(j).Shape.TextFrame2.TextRange.LanguageID = msoLanguageIDEnglishUS
                 Next j
            Next i
 End If
 Next oshp

 For Each notesshp In osld.NotesPage.Shapes
 If notesshp.HasTextFrame Then
 notesshp.TextFrame2.TextRange.LanguageID = msoLanguageIDEnglishUS
 End If
 Next notesshp
 Next osld
 
    For Each osld In ActivePresentation.Slides
    For Each oshp In osld.Shapes
    
        With oshp
            Select Case .Type
                Case Is = msoGroup
                For x = 1 To .GroupItems.Count
                    If .GroupItems(x).HasTextFrame Then
                         oshp.TextFrame2.TextRange.LanguageID = msoLanguageIDEnglishUS
                    End If
                Next x
            End Select
        End With
    Next oshp
    Next

'Wenn Umstellung der Default Language auch gewünscht, dann:
'With ActivePresentation
'   .DefaultLanguageID = msoLanguageIDEnglishUS
'End With

End Sub

Public Sub SpelSelDE(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
 Dim oshp As Shape
 Dim i As Long
 Dim j As Long
 Dim x As Long
 
     On Error GoTo err
     
 For Each oshp In ActiveWindow.Selection.ShapeRange
 If oshp.HasTextFrame Then
 oshp.TextFrame2.TextRange.LanguageID = msoLanguageIDGerman
 End If
  If oshp.HasTable Then
             For i = 1 To oshp.Table.Rows.Count
                For j = 1 To oshp.Table.Columns.Count
 oshp.Table.Rows.Item(i).Cells(j).Shape.TextFrame2.TextRange.LanguageID = msoLanguageIDGerman
                 Next j
            Next i
 End If
 Next oshp
 
 For Each oshp In ActiveWindow.Selection.ShapeRange
        With oshp
            Select Case .Type
                Case Is = msoGroup
                For x = 1 To .GroupItems.Count
                    If .GroupItems(x).HasTextFrame Then
                         oshp.TextFrame2.TextRange.LanguageID = msoLanguageIDGerman
                    End If
                Next x
            End Select
        End With
    Next oshp
Exit Sub
 
err:
    MsgBox "Please select at least one shape"
    
 End Sub

Public Sub SpelSelES(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
 Dim oshp As Shape
 Dim i As Long
 Dim j As Long
 Dim x As Long
 
     On Error GoTo err
     
 For Each oshp In ActiveWindow.Selection.ShapeRange
 If oshp.HasTextFrame Then
 oshp.TextFrame2.TextRange.LanguageID = msoLanguageIDSpanish
 End If
  If oshp.HasTable Then
             For i = 1 To oshp.Table.Rows.Count
                For j = 1 To oshp.Table.Columns.Count
 oshp.Table.Rows.Item(i).Cells(j).Shape.TextFrame2.TextRange.LanguageID = msoLanguageIDSpanish
                 Next j
            Next i
 End If
 Next oshp
 
 For Each oshp In ActiveWindow.Selection.ShapeRange
        With oshp
            Select Case .Type
                Case Is = msoGroup
                For x = 1 To .GroupItems.Count
                    If .GroupItems(x).HasTextFrame Then
                         oshp.TextFrame2.TextRange.LanguageID = msoLanguageIDSpanish
                    End If
                Next x
            End Select
        End With
    Next oshp
Exit Sub
 
err:
    MsgBox "Please select at least one shape"
    
 End Sub

Public Sub SpelSelFR(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
 Dim oshp As Shape
 Dim i As Long
 Dim j As Long
 Dim x As Long
 
     On Error GoTo err
     
 For Each oshp In ActiveWindow.Selection.ShapeRange
 If oshp.HasTextFrame Then
 oshp.TextFrame2.TextRange.LanguageID = msoLanguageIDFrench
 End If
  If oshp.HasTable Then
             For i = 1 To oshp.Table.Rows.Count
                For j = 1 To oshp.Table.Columns.Count
 oshp.Table.Rows.Item(i).Cells(j).Shape.TextFrame2.TextRange.LanguageID = msoLanguageIDFrench
                 Next j
            Next i
 End If
 Next oshp
 
 For Each oshp In ActiveWindow.Selection.ShapeRange
        With oshp
            Select Case .Type
                Case Is = msoGroup
                For x = 1 To .GroupItems.Count
                    If .GroupItems(x).HasTextFrame Then
                         oshp.TextFrame2.TextRange.LanguageID = msoLanguageIDFrench
                    End If
                Next x
            End Select
        End With
    Next oshp
Exit Sub
 
err:
    MsgBox "Please select at least one shape"
    
 End Sub

Public Sub SpelSelIT(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
 Dim oshp As Shape
 Dim i As Long
 Dim j As Long
 Dim x As Long
 
     On Error GoTo err
     
 For Each oshp In ActiveWindow.Selection.ShapeRange
 If oshp.HasTextFrame Then
 oshp.TextFrame2.TextRange.LanguageID = msoLanguageIDItalian
 End If
  If oshp.HasTable Then
             For i = 1 To oshp.Table.Rows.Count
                For j = 1 To oshp.Table.Columns.Count
 oshp.Table.Rows.Item(i).Cells(j).Shape.TextFrame2.TextRange.LanguageID = msoLanguageIDItalian
                 Next j
            Next i
 End If
 Next oshp
 
 For Each oshp In ActiveWindow.Selection.ShapeRange
        With oshp
            Select Case .Type
                Case Is = msoGroup
                For x = 1 To .GroupItems.Count
                    If .GroupItems(x).HasTextFrame Then
                         oshp.TextFrame2.TextRange.LanguageID = msoLanguageIDItalian
                    End If
                Next x
            End Select
        End With
    Next oshp
Exit Sub
 
err:
    MsgBox "Please select at least one shape"
    
 End Sub

Public Sub SpelSelUK(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
 Dim oshp As Shape
 Dim i As Long
 Dim j As Long
 Dim x As Long
 
     On Error GoTo err
     
 For Each oshp In ActiveWindow.Selection.ShapeRange
 If oshp.HasTextFrame Then
 oshp.TextFrame2.TextRange.LanguageID = msoLanguageIDEnglishUK
 End If
  If oshp.HasTable Then
             For i = 1 To oshp.Table.Rows.Count
                For j = 1 To oshp.Table.Columns.Count
 oshp.Table.Rows.Item(i).Cells(j).Shape.TextFrame2.TextRange.LanguageID = msoLanguageIDEnglishUK
                 Next j
            Next i
 End If
 Next oshp
 
 For Each oshp In ActiveWindow.Selection.ShapeRange
        With oshp
            Select Case .Type
                Case Is = msoGroup
                For x = 1 To .GroupItems.Count
                    If .GroupItems(x).HasTextFrame Then
                         oshp.TextFrame2.TextRange.LanguageID = msoLanguageIDEnglishUK
                    End If
                Next x
            End Select
        End With
    Next oshp
Exit Sub
 
err:
    MsgBox "Please select at least one shape"
    
End Sub

Public Sub SpelSelUS(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
 Dim oshp As Shape
 Dim i As Long
 Dim j As Long
 Dim x As Long
 
     On Error GoTo err
     
 For Each oshp In ActiveWindow.Selection.ShapeRange
 If oshp.HasTextFrame Then
 oshp.TextFrame2.TextRange.LanguageID = msoLanguageIDEnglishUS
 End If
  If oshp.HasTable Then
             For i = 1 To oshp.Table.Rows.Count
                For j = 1 To oshp.Table.Columns.Count
 oshp.Table.Rows.Item(i).Cells(j).Shape.TextFrame2.TextRange.LanguageID = msoLanguageIDEnglishUS
                 Next j
            Next i
 End If
 Next oshp
 
 For Each oshp In ActiveWindow.Selection.ShapeRange
        With oshp
            Select Case .Type
                Case Is = msoGroup
                For x = 1 To .GroupItems.Count
                    If .GroupItems(x).HasTextFrame Then
                         oshp.TextFrame2.TextRange.LanguageID = msoLanguageIDEnglishUS
                    End If
                Next x
            End Select
        End With
    Next oshp
Exit Sub
 
err:
    MsgBox "Please select at least one shape"
    
End Sub

Public Sub StackLR(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim rayShapes() As Shape
    Dim L As Long
    
    On Error Resume Next
    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please select at least two shapes"
        Exit Sub
    End If

    On Error GoTo 0
    ReDim rayShapes(1 To ActiveWindow.Selection.ShapeRange.Count)
        For L = 1 To ActiveWindow.Selection.ShapeRange.Count
            Set rayShapes(L) = ActiveWindow.Selection.ShapeRange(L)
        Next L

    For L = 2 To UBound(rayShapes)
        Debug.Print rayShapes(L).Name
        rayShapes(L).Left = rayShapes(L - 1).Left + rayShapes(L - 1).Width
    Next L
End Sub

Public Sub StackTB(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim rayShapes() As Shape
    Dim L As Long
    
    On Error Resume Next
    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please select at least two shapes"
        Exit Sub
    End If

    On Error GoTo 0
    ReDim rayShapes(1 To ActiveWindow.Selection.ShapeRange.Count)
        For L = 1 To ActiveWindow.Selection.ShapeRange.Count
            Set rayShapes(L) = ActiveWindow.Selection.ShapeRange(L)
        Next L

    For L = 2 To UBound(rayShapes)
        Debug.Print rayShapes(L).Name
        rayShapes(L).Top = rayShapes(L - 1).Top + rayShapes(L - 1).Height
    Next L
End Sub

Public Sub StatusDel(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim sld As Slide
    Dim L As Long
    If MsgBox("Do you want to delete ALL status stickers from entire presentation?", vbYesNo) <> vbYes Then Exit Sub
    On Error Resume Next
    For Each sld In ActivePresentation.Slides
        For L = sld.Shapes.Count To 1 Step -1
            If sld.Shapes(L).Tags("STATUS") = "YES" Then sld.Shapes(L).Delete
        Next L
    Next sld
End Sub

Public Sub StretchLR(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

Dim phd As Shape
Dim shp As Shape

Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)

    On Error GoTo err
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Left = phd.Left
        shp.Width = phd.Width
    Next shp
    Exit Sub
err:
    MsgBox "Please select at least one shape"

End Sub

Public Sub StretchTB(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

Dim phd As Shape
Dim shp As Shape

Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)

    On Error GoTo err
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Top = phd.Top
        shp.Height = phd.Height
    Next shp
    Exit Sub
err:
    MsgBox "Please select at least one shape"

End Sub

Public Sub SwapPos(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp1 As Shape
    Dim shp2 As Shape

    On Error Resume Next

    If ActiveWindow.Selection.ShapeRange.Count <> 2 Then
        MsgBox "Please select exactly two shapes"
        Exit Sub
    End If
    
    Set shp1 = ActiveWindow.Selection.ShapeRange(1)
    Set shp2 = ActiveWindow.Selection.ShapeRange(2)
    
    If err Then Exit Sub
    
    With shp1.Duplicate
        .Left = shp2.Left
        .Top = shp2.Top
    End With
    DoEvents
    
    shp2.Left = shp1.Left
    shp2.Top = shp1.Top
    
    shp1.Delete
    
End Sub

Public Sub TitleReset(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
Dim osld As Slide
    Dim oshp As Shape
    Dim ocust As Shape
    For Each osld In ActiveWindow.Selection.SlideRange
        If osld.Shapes.HasTitle Then
            Set oshp = osld.Shapes.Title
            Set ocust = osld.CustomLayout.Shapes.Title
            With oshp
                .Left = ocust.Left
                .Top = ocust.Top
                .Height = ocust.Height
                .Width = ocust.Width
                .TextFrame2.TextRange.Font.Name = ocust.TextFrame2.TextRange.Font.Name
                .TextFrame2.TextRange.Font.Size = ocust.TextFrame2.TextRange.Font.Size
                .TextFrame2.TextRange.Font.Bold = ocust.TextFrame2.TextRange.Font.Bold
                .TextFrame2.TextRange.Font.Italic = ocust.TextFrame2.TextRange.Font.Italic
                .TextFrame2.TextRange.Font.UnderlineStyle = ocust.TextFrame2.TextRange.Font.UnderlineStyle
                .TextFrame2.TextRange.Font.Fill.ForeColor = ocust.TextFrame2.TextRange.Font.Fill.ForeColor
                .TextFrame2.VerticalAnchor = ocust.TextFrame2.VerticalAnchor
                .TextFrame2.TextRange.ParagraphFormat.Alignment = ocust.TextFrame2.TextRange.ParagraphFormat.Alignment
                .TextFrame2.MarginBottom = ocust.TextFrame2.MarginBottom
                .TextFrame2.MarginLeft = ocust.TextFrame2.MarginLeft
                .TextFrame2.MarginRight = ocust.TextFrame2.MarginRight
                .TextFrame2.MarginTop = ocust.TextFrame2.MarginTop
'Add or delete what you need to add or delete, e.g., Bold, LeftMargin, Alignment
            End With
        End If
    Next osld
End Sub

Public Sub WordWrap(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    On Error GoTo err
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.TextFrame.WordWrap = msoFalse = shp.TextFrame.WordWrap = msoTrue
    Next shp
    Exit Sub
err:
    MsgBox "Please select at least one shape"
End Sub

'CCC Anpassbare

Public Sub BackGreen(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
With ActiveWindow.Selection.SlideRange
     .FollowMasterBackground = msoFalse
     .Background.Fill.Solid
     .Background.Fill.ForeColor.RGB = RGB(153, 255, 153)
End With

End Sub

Public Sub BackRed(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
With ActiveWindow.Selection.SlideRange
     .FollowMasterBackground = msoFalse
     .Background.Fill.Solid
     .Background.Fill.ForeColor.RGB = RGB(255, 153, 153)
End With

End Sub

Public Sub BackYellow(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
With ActiveWindow.Selection.SlideRange
     .FollowMasterBackground = msoFalse
     .Background.Fill.Solid
     .Background.Fill.ForeColor.RGB = RGB(255, 255, 153)
End With

End Sub

Public Sub ByteCountAdd(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

If (isMac()) Then
    MsgBox "Mac does not support this function. We are sorry."
Else
    Call ByteCountAddWin
End If

End Sub

Private Sub ByteCountAddWin()

Dim oFS As Object
Dim oPres As Presentation
Dim ocopy As Presentation
Dim L As Long
Dim sngSize As Single
Dim strFileName As String
Dim folderpath As String
Dim sizes() As String

If MsgBox("Attention: Macro needs up to one minute for every 10 MB of presentation size (e.g., 7 minutes for 70 MB). Continue?", vbYesNo) <> vbYes Then Exit Sub

Set oPres = ActivePresentation
On Error Resume Next

folderpath = Environ("TEMP") & "\Slides\"
Kill folderpath & "*.*"
MkDir folderpath
oPres.SaveCopyAs folderpath & "copy.pptx"

Call killoldbytesizeshapes(oPres)

For L = 1 To oPres.Slides.Count
    Set ocopy = Presentations.Open(folderpath & "copy.pptx", WithWindow:=False)
    ocopy.Slides.Range(MyRange(L + 1, oPres.Slides.Count)).Delete
        If L > 1 Then ocopy.Slides.Range(MyRange(1, L - 1)).Delete
    ocopy.SaveAs folderpath & "Slide" & L & ".pptx"
    ocopy.Close
Next L

ReDim sizes(1 To oPres.Slides.Count)
Set oFS = CreateObject("Scripting.FileSystemObject")
    For L = 1 To oPres.Slides.Count
        strFileName = folderpath & "Slide" & CStr(L) & ".pptx"
        sngSize = oFS.GetFile(strFileName).Size
        sizes(L) = Format(Int(sngSize / 1024), "##,##") & "kb"
        With oPres.Slides(L).Shapes.AddShape(msoShapeRectangle, 0, 0, 200, 20)
            With .Fill
                .Visible = msoTrue
                .Transparency = 0
                .ForeColor.RGB = RGB(0, 0, 255)
            End With
            With .Line
                .Visible = msoTrue
                .ForeColor.RGB = RGB(255, 255, 255)
                .Weight = 0.75
            End With
            With .Tags
                .Add "BYTECOUNT", "YES"
            End With
            With .TextFrame2
                With .TextRange
                    .Text = "Slide size is " & sizes(L)
                    .Paragraphs.ParagraphFormat.Alignment = msoAlignCenter
                    With .Font
                        .Name = "Arial"
                        .Size = 12
                        .Fill.ForeColor.RGB = RGB(255, 255, 255)
                        .Bold = msoTrue
                        .Italic = msoFalse
                        .UnderlineStyle = msoNoUnderline
                    End With
                End With
                .VerticalAnchor = msoAnchorMiddle
                .Orientation = msoTextOrientationHorizontal
                .MarginBottom = 7.0866097
                .MarginLeft = 7.0866097
                .MarginRight = 7.0866097
                .MarginTop = 7.0866097
                .WordWrap = msoTrue
            End With
        End With
    Next

End Sub

Function MyRange(ByVal StartIndex As Long, ByVal EndIndex As Long) As Variant 'gehört zu ByteCountAdd
Dim Arr() As Long
Dim i As Long

ReDim Arr(StartIndex To EndIndex)
    For i = StartIndex To EndIndex: Arr(i) = i: Next
        MyRange = Arr
End Function

Sub killoldbytesizeshapes(oPres As Presentation) 'gehört zu ByteCountAdd
Dim osld As Slide
Dim L As Long

    For Each osld In oPres.Slides
        If osld.Shapes.Count > 0 Then
            For L = osld.Shapes.Count To 1 Step -1
                If osld.Shapes(L).Tags("BYTECOUNT") = "YES" Then osld.Shapes(L).Delete
            Next L
        End If
    Next osld

End Sub

Public Sub HeightDec(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    On Error GoTo err
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Height = shp.Height - 2.8346439
    Next shp
    Exit Sub
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub HeightInc(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    On Error GoTo err
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Height = shp.Height + 2.8346439
    Next shp
    Exit Sub
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub LineSpaceDec(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    Dim oTbl As Table
    Dim P As Long
    Dim x As Long
    Dim y As Long
    Dim oldPT As Single
    Dim trueSW As Single
    Dim factor As Single
      
On Error GoTo err
      
If ActiveWindow.Selection.Type = ppSelectionNone Then 'Nothing!

ElseIf ActiveWindow.Selection.Type = ppSelectionText Then
    With ActiveWindow.Selection.TextRange2
        For P = 1 To .Paragraphs.Count
            If .Paragraphs(P).ParagraphFormat.LineRuleWithin = msoFalse Then
                oldPT = .ParagraphFormat.SpaceWithin
                factor = .Paragraphs(P).Font.Size * 1.2
                trueSW = oldPT / factor
            With .Paragraphs(P).ParagraphFormat
                .LineRuleWithin = msoTrue
                .SpaceWithin = trueSW - 0.1
            End With
            Else
                trueSW = .Paragraphs(P).ParagraphFormat.SpaceWithin
            With .Paragraphs(P).ParagraphFormat
                .LineRuleWithin = msoTrue
                .SpaceWithin = trueSW - 0.1
            End With
            End If
        Next P
    End With
Else
    For Each shp In ActiveWindow.Selection.ShapeRange
        If ActiveWindow.Selection.ShapeRange(1).HasTable Then
        Set oTbl = ActiveWindow.Selection.ShapeRange.Table
            For x = 1 To oTbl.Rows.Count
                For y = 1 To oTbl.Columns.Count
                    If oTbl.Cell(x, y).Selected Then
                        With oTbl.Cell(x, y).Shape.TextFrame2.TextRange
                            For P = 1 To .Paragraphs.Count
                                If .Paragraphs(P).ParagraphFormat.LineRuleWithin = msoFalse Then
                                    oldPT = .ParagraphFormat.SpaceWithin
                                    factor = .Paragraphs(P).Font.Size * 1.2
                                    trueSW = oldPT / factor
                                With .Paragraphs(P).ParagraphFormat
                                    .LineRuleWithin = msoTrue
                                    .SpaceWithin = trueSW - 0.1
                                End With
                                Else
                                    trueSW = .Paragraphs(P).ParagraphFormat.SpaceWithin
                                With .Paragraphs(P).ParagraphFormat
                                    .LineRuleWithin = msoTrue
                                    .SpaceWithin = trueSW - 0.1
                                End With
                                End If
                            Next P
                        End With
                    End If
                Next
            Next
        Else
            With shp.TextFrame2.TextRange
                For P = 1 To .Paragraphs.Count
                    If .Paragraphs(P).ParagraphFormat.LineRuleWithin = msoFalse Then
                        oldPT = .ParagraphFormat.SpaceWithin
                        factor = .Paragraphs(P).Font.Size * 1.2
                        trueSW = oldPT / factor
                    With .Paragraphs(P).ParagraphFormat
                        .LineRuleWithin = msoTrue
                        .SpaceWithin = trueSW - 0.1
                    End With
                    Else
                        trueSW = .Paragraphs(P).ParagraphFormat.SpaceWithin
                    With .Paragraphs(P).ParagraphFormat
                        .LineRuleWithin = msoTrue
                        .SpaceWithin = trueSW - 0.1
                    End With
                    End If
                Next P
            End With
        End If
    Next shp
End If
Exit Sub

err:
    MsgBox "Please deselect tables from your range of objects and edit each of them separately or check if line spacing somewhere in the selected area is already smaller than 0.1"

End Sub

Public Sub LineSpaceInc(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    Dim oTbl As Table
    Dim P As Long
    Dim x As Long
    Dim y As Long
    Dim oldPT As Single
    Dim trueSW As Single
    Dim factor As Single
      
On Error GoTo err
      
If ActiveWindow.Selection.Type = ppSelectionNone Then 'Nothing!

ElseIf ActiveWindow.Selection.Type = ppSelectionText Then
    With ActiveWindow.Selection.TextRange2
        For P = 1 To .Paragraphs.Count
            If .Paragraphs(P).ParagraphFormat.LineRuleWithin = msoFalse Then
                oldPT = .ParagraphFormat.SpaceWithin
                factor = .Paragraphs(P).Font.Size * 1.2
                trueSW = oldPT / factor
            With .Paragraphs(P).ParagraphFormat
                .LineRuleWithin = msoTrue
                .SpaceWithin = trueSW + 0.1
            End With
            Else
                trueSW = .Paragraphs(P).ParagraphFormat.SpaceWithin
            With .Paragraphs(P).ParagraphFormat
                .LineRuleWithin = msoTrue
                .SpaceWithin = trueSW + 0.1
            End With
            End If
        Next P
    End With
Else
    For Each shp In ActiveWindow.Selection.ShapeRange
        If ActiveWindow.Selection.ShapeRange(1).HasTable Then
        Set oTbl = ActiveWindow.Selection.ShapeRange.Table
            For x = 1 To oTbl.Rows.Count
                For y = 1 To oTbl.Columns.Count
                    If oTbl.Cell(x, y).Selected Then
                        With oTbl.Cell(x, y).Shape.TextFrame2.TextRange
                            For P = 1 To .Paragraphs.Count
                                If .Paragraphs(P).ParagraphFormat.LineRuleWithin = msoFalse Then
                                    oldPT = .ParagraphFormat.SpaceWithin
                                    factor = .Paragraphs(P).Font.Size * 1.2
                                    trueSW = oldPT / factor
                                With .Paragraphs(P).ParagraphFormat
                                    .LineRuleWithin = msoTrue
                                    .SpaceWithin = trueSW + 0.1
                                End With
                                Else
                                    trueSW = .Paragraphs(P).ParagraphFormat.SpaceWithin
                                With .Paragraphs(P).ParagraphFormat
                                    .LineRuleWithin = msoTrue
                                    .SpaceWithin = trueSW + 0.1
                                End With
                                End If
                            Next P
                        End With
                    End If
                Next
            Next
        Else
            With shp.TextFrame2.TextRange
                For P = 1 To .Paragraphs.Count
                    If .Paragraphs(P).ParagraphFormat.LineRuleWithin = msoFalse Then
                        oldPT = .ParagraphFormat.SpaceWithin
                        factor = .Paragraphs(P).Font.Size * 1.2
                        trueSW = oldPT / factor
                    With .Paragraphs(P).ParagraphFormat
                        .LineRuleWithin = msoTrue
                        .SpaceWithin = trueSW + 0.1
                    End With
                    Else
                        trueSW = .Paragraphs(P).ParagraphFormat.SpaceWithin
                    With .Paragraphs(P).ParagraphFormat
                        .LineRuleWithin = msoTrue
                        .SpaceWithin = trueSW + 0.1
                    End With
                    End If
                Next P
            End With
        End If
    Next shp
End If
Exit Sub

err:
    MsgBox "Please deselect tables from your range of objects and edit each of them separately"

End Sub

Public Sub LineSpaceOne(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    Dim oTbl As Table
    Dim x As Long
    Dim y As Long
      
On Error GoTo err
      
If ActiveWindow.Selection.Type = ppSelectionNone Then 'Nothing!

ElseIf ActiveWindow.Selection.Type = ppSelectionText Then
    With ActiveWindow.Selection.TextRange
        If .ParagraphFormat.LineRuleWithin = msoFalse Then .ParagraphFormat.LineRuleWithin = msoTrue
    .ParagraphFormat.SpaceWithin = 1
    End With

Else
    For Each shp In ActiveWindow.Selection.ShapeRange
        If ActiveWindow.Selection.ShapeRange(1).HasTable Then
        Set oTbl = ActiveWindow.Selection.ShapeRange.Table
            For x = 1 To oTbl.Rows.Count
                For y = 1 To oTbl.Columns.Count
                    If oTbl.Cell(x, y).Selected Then
                        With oTbl.Cell(x, y).Shape.TextFrame2.TextRange
                            If .ParagraphFormat.LineRuleWithin = msoFalse Then .ParagraphFormat.LineRuleWithin = msoTrue
                        .ParagraphFormat.SpaceWithin = 1
                        End With
                    End If
                Next
            Next
        Else
            With shp.TextFrame2.TextRange
                If .ParagraphFormat.LineRuleWithin = msoFalse Then .ParagraphFormat.LineRuleWithin = msoTrue
            .ParagraphFormat.SpaceWithin = 1
            End With
        End If
    Next shp
End If
Exit Sub

err:
    MsgBox "Please deselect tables from your range of objects and edit each of them separately"

End Sub

Public Sub ParaInc(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    Dim oTbl As Table
    Dim P As Long
    Dim x As Long
    Dim y As Long
      
On Error GoTo err
      
If ActiveWindow.Selection.Type = ppSelectionNone Then 'Nothing!

ElseIf ActiveWindow.Selection.Type = ppSelectionText Then
    With ActiveWindow.Selection.TextRange
        If .ParagraphFormat.LineRuleAfter = msoTrue Then .ParagraphFormat.LineRuleAfter = msoFalse
            For P = 1 To .Paragraphs.Count
                .Paragraphs(P).ParagraphFormat.SpaceAfter = .Paragraphs(P).ParagraphFormat.SpaceAfter + 3
            Next
    End With

Else
    For Each shp In ActiveWindow.Selection.ShapeRange
        If ActiveWindow.Selection.ShapeRange(1).HasTable Then
        Set oTbl = ActiveWindow.Selection.ShapeRange.Table
            For x = 1 To oTbl.Rows.Count
                For y = 1 To oTbl.Columns.Count
                    If oTbl.Cell(x, y).Selected Then
                        With oTbl.Cell(x, y).Shape.TextFrame2.TextRange
                            If .ParagraphFormat.LineRuleAfter = msoTrue Then .ParagraphFormat.LineRuleAfter = msoFalse
                                For P = 1 To .Paragraphs.Count
                                    .Paragraphs(P).ParagraphFormat.SpaceAfter = .Paragraphs(P).ParagraphFormat.SpaceAfter + 3
                            Next
                        End With
                    End If
                Next
            Next
        Else
            With shp.TextFrame2.TextRange
                If .ParagraphFormat.LineRuleAfter = msoTrue Then .ParagraphFormat.LineRuleAfter = msoFalse
                    For P = 1 To .Paragraphs.Count
                        .Paragraphs(P).ParagraphFormat.SpaceAfter = .Paragraphs(P).ParagraphFormat.SpaceAfter + 3
                    Next
            End With
        End If
    Next shp
End If
Exit Sub

err:
    MsgBox "Please deselect tables from your range of objects and edit each of them separately"

End Sub

Function rightPos(oPres As Presentation) As Single 'gehört zu den Status-Stickern
    rightPos = oPres.PageSetup.SlideWidth - 215.7164
End Function

Public Sub StatusNew(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    Dim sld As Slide
    Dim placeMe As Single
    
    placeMe = rightPos(ActivePresentation)
    
On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else
    
    'New
    Set sld = Application.ActiveWindow.View.Slide
    Set shp = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=placeMe, Top:=28.913368, Width:=198.14161, Height:=35.149584)
    shp.Fill.Visible = msoTrue
    shp.Fill.ForeColor.RGB = RGB(255, 255, 0)
    shp.Fill.Transparency = 0
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(255, 255, 0)
    shp.Line.Weight = 0.75
    shp.Rotation = 10
    shp.Tags.Add "STATUS", "YES"
    
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    shp.TextFrame2.TextRange.Characters.Text = "NEW"
    shp.TextFrame2.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.Font.Size = 16
    shp.TextFrame2.TextRange.Font.Name = "Arial"
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.Font.Italic = msoFalse
    shp.TextFrame2.TextRange.Font.UnderlineStyle = msoNoUnderline
    shp.TextFrame2.Orientation = msoTextOrientationHorizontal
    shp.TextFrame2.MarginBottom = 7.0866097
    shp.TextFrame2.MarginLeft = 7.0866097
    shp.TextFrame2.MarginRight = 7.0866097
    shp.TextFrame2.MarginTop = 7.0866097
    shp.TextFrame2.WordWrap = msoTrue

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub StatusOut(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    Dim sld As Slide
    Dim placeMe As Single
    
    placeMe = rightPos(ActivePresentation)
    
On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else
    
    'Outdated
    Set sld = Application.ActiveWindow.View.Slide
    Set shp = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=placeMe, Top:=28.913368, Width:=198.14161, Height:=35.149584)
    shp.Fill.Visible = msoTrue
    shp.Fill.ForeColor.RGB = RGB(89, 89, 89)
    shp.Fill.Transparency = 0
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(89, 89, 89)
    shp.Line.Weight = 0.75
    shp.Rotation = 10
    shp.Tags.Add "STATUS", "YES"
    
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Characters.Text = "OUTDATED"
    shp.TextFrame2.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.Font.Size = 16
    shp.TextFrame2.TextRange.Font.Name = "Arial"
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.Font.Italic = msoFalse
    shp.TextFrame2.TextRange.Font.UnderlineStyle = msoNoUnderline
    shp.TextFrame2.Orientation = msoTextOrientationHorizontal
    shp.TextFrame2.MarginBottom = 7.0866097
    shp.TextFrame2.MarginLeft = 7.0866097
    shp.TextFrame2.MarginRight = 7.0866097
    shp.TextFrame2.MarginTop = 7.0866097
    shp.TextFrame2.WordWrap = msoTrue

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub StatusToDo(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    Dim sld As Slide
    Dim placeMe As Single
    
    placeMe = rightPos(ActivePresentation)
    
On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else
    
    'To do
    Set sld = Application.ActiveWindow.View.Slide
    Set shp = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=placeMe, Top:=28.913368, Width:=198.14161, Height:=35.149584)
    shp.Fill.Visible = msoTrue
    shp.Fill.ForeColor.RGB = RGB(255, 0, 0)
    shp.Fill.Transparency = 0
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(255, 0, 0)
    shp.Line.Weight = 0.75
    shp.Rotation = 10
    shp.Tags.Add "STATUS", "YES"
    
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Characters.Text = "TO DO"
    shp.TextFrame2.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.Font.Size = 16
    shp.TextFrame2.TextRange.Font.Name = "Arial"
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.Font.Italic = msoFalse
    shp.TextFrame2.TextRange.Font.UnderlineStyle = msoNoUnderline
    shp.TextFrame2.Orientation = msoTextOrientationHorizontal
    shp.TextFrame2.MarginBottom = 7.0866097
    shp.TextFrame2.MarginLeft = 7.0866097
    shp.TextFrame2.MarginRight = 7.0866097
    shp.TextFrame2.MarginTop = 7.0866097
    shp.TextFrame2.WordWrap = msoTrue

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub StatusUp(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    Dim sld As Slide
    Dim placeMe As Single
    
    placeMe = rightPos(ActivePresentation)
    
On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else
    
    'Updated
    Set sld = Application.ActiveWindow.View.Slide
    Set shp = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=placeMe, Top:=28.913368, Width:=198.14161, Height:=35.149584)
    shp.Fill.Visible = msoTrue
    shp.Fill.ForeColor.RGB = RGB(0, 176, 80)
    shp.Fill.Transparency = 0
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(0, 176, 80)
    shp.Line.Weight = 0.75
    shp.Rotation = 10
    shp.Tags.Add "STATUS", "YES"
    
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Characters.Text = "UPDATED"
    shp.TextFrame2.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.Font.Size = 16
    shp.TextFrame2.TextRange.Font.Name = "Arial"
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.Font.Italic = msoFalse
    shp.TextFrame2.TextRange.Font.UnderlineStyle = msoNoUnderline
    shp.TextFrame2.Orientation = msoTextOrientationHorizontal
    shp.TextFrame2.MarginBottom = 7.0866097
    shp.TextFrame2.MarginLeft = 7.0866097
    shp.TextFrame2.MarginRight = 7.0866097
    shp.TextFrame2.MarginTop = 7.0866097
    shp.TextFrame2.WordWrap = msoTrue

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub StatusWip(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    Dim sld As Slide
    Dim placeMe As Single
    
    placeMe = rightPos(ActivePresentation)
    
On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else
    
    'Work in progress
    Set sld = Application.ActiveWindow.View.Slide
    Set shp = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=placeMe, Top:=28.913368, Width:=198.14161, Height:=35.149584)
    shp.Fill.Visible = msoTrue
    shp.Fill.ForeColor.RGB = RGB(255, 192, 0)
    shp.Fill.Transparency = 0
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(255, 192, 0)
    shp.Line.Weight = 0.75
    shp.Rotation = 10
    shp.Tags.Add "STATUS", "YES"
    
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    shp.TextFrame2.TextRange.Characters.Text = "WORK IN PROGRESS"
    shp.TextFrame2.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.Font.Size = 16
    shp.TextFrame2.TextRange.Font.Name = "Arial"
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.Font.Italic = msoFalse
    shp.TextFrame2.TextRange.Font.UnderlineStyle = msoNoUnderline
    shp.TextFrame2.Orientation = msoTextOrientationHorizontal
    shp.TextFrame2.MarginBottom = 7.0866097
    shp.TextFrame2.MarginLeft = 7.0866097
    shp.TextFrame2.MarginRight = 7.0866097
    shp.TextFrame2.MarginTop = 7.0866097
    shp.TextFrame2.WordWrap = msoTrue

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub TextMarBDec(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    On Error GoTo err
    
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange
            .TextFrame.MarginBottom = .TextFrame.MarginBottom - 1.4173219
        End With
    
    Else
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.TextFrame.MarginBottom = shp.TextFrame.MarginBottom - 1.4173219
    Next shp
    End If
    Exit Sub
err:
    MsgBox "Either the margin is already smaller than 0.05 or you haven't selected at least one shape"
    
End Sub

Public Sub TextMarBInc(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    On Error GoTo err
    
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange
            .TextFrame.MarginBottom = .TextFrame.MarginBottom + 1.4173219
        End With
    
    Else
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.TextFrame.MarginBottom = shp.TextFrame.MarginBottom + 1.4173219
    Next shp
    End If
    Exit Sub
err:
    MsgBox "Please select at least one shape"
    
End Sub

Public Sub TextMarBZero(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    On Error GoTo err
    
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange
            .TextFrame.MarginBottom = 0
        End With
    
    Else
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.TextFrame.MarginBottom = 0
    Next shp
    End If
    Exit Sub
err:
    MsgBox "Please select at least one shape"
    
End Sub

Public Sub TextMarLRDec(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    On Error GoTo err

    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange
            .TextFrame.MarginLeft = .TextFrame.MarginLeft - 1.4173219
            .TextFrame.MarginRight = .TextFrame.MarginRight - 1.4173219
        End With
    
    Else
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.TextFrame.MarginLeft = shp.TextFrame.MarginLeft - 1.4173219
        shp.TextFrame.MarginRight = shp.TextFrame.MarginRight - 1.4173219
    Next shp
    End If
    Exit Sub
err:
    MsgBox "Either the margin is already smaller than 0.05 or you haven't selected at least one shape"
    
End Sub

Public Sub TextMarLRInc(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    On Error GoTo err
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange
            .TextFrame.MarginLeft = .TextFrame.MarginLeft + 1.4173219
            .TextFrame.MarginRight = .TextFrame.MarginRight + 1.4173219
        End With
    
    Else
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.TextFrame.MarginLeft = shp.TextFrame.MarginLeft + 1.4173219
        shp.TextFrame.MarginRight = shp.TextFrame.MarginRight + 1.4173219
    Next shp
    End If
    Exit Sub
err:
    MsgBox "Please select at least one shape"
    
End Sub

Public Sub TextMarLRZero(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    On Error GoTo err

    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange
            .TextFrame.MarginLeft = 0
            .TextFrame.MarginRight = 0
        End With
    
    Else
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.TextFrame.MarginLeft = 0
        shp.TextFrame.MarginRight = 0
    Next shp
    End If
    Exit Sub
err:
    MsgBox "Please select at least one shape"
    
End Sub

Public Sub TextMarLDec(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    On Error GoTo err

    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange
            .TextFrame.MarginLeft = .TextFrame.MarginLeft - 1.4173219
        End With
    
    Else
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.TextFrame.MarginLeft = shp.TextFrame.MarginLeft - 1.4173219
    Next shp
    End If
    Exit Sub
err:
    MsgBox "Either the margin is already smaller than 0.05 or you haven't selected at least one shape"
    
End Sub

Public Sub TextMarLInc(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    On Error GoTo err
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange
            .TextFrame.MarginLeft = .TextFrame.MarginLeft + 1.4173219
        End With
    
    Else
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.TextFrame.MarginLeft = shp.TextFrame.MarginLeft + 1.4173219
    Next shp
    End If
    Exit Sub
err:
    MsgBox "Please select at least one shape"
    
End Sub

Public Sub TextMarLZero(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    On Error GoTo err

    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange
            .TextFrame.MarginLeft = 0
        End With
    
    Else
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.TextFrame.MarginLeft = 0
    Next shp
    End If
    Exit Sub
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub TextMarRDec(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    On Error GoTo err

    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange
            .TextFrame.MarginRight = .TextFrame.MarginRight - 1.4173219
        End With
    
    Else
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.TextFrame.MarginRight = shp.TextFrame.MarginRight - 1.4173219
    Next shp
    End If
    Exit Sub
err:
    MsgBox "Either the margin is already smaller than 0.05 or you haven't selected at least one shape"
End Sub

Public Sub TextMarRInc(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    On Error GoTo err
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange
            .TextFrame.MarginRight = .TextFrame.MarginRight + 1.4173219
        End With
    
    Else
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.TextFrame.MarginRight = shp.TextFrame.MarginRight + 1.4173219
    Next shp
    End If
    Exit Sub
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub TextMarRZero(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    On Error GoTo err

    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange
            .TextFrame.MarginRight = 0
        End With
    
    Else
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.TextFrame.MarginRight = 0
    Next shp
    End If
    Exit Sub
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub TextMarTBDec(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    On Error GoTo err
    
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange
            .TextFrame.MarginTop = .TextFrame.MarginTop - 1.4173219
            .TextFrame.MarginBottom = .TextFrame.MarginBottom - 1.4173219
        End With
    
    Else
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.TextFrame.MarginTop = shp.TextFrame.MarginTop - 1.4173219
        shp.TextFrame.MarginBottom = shp.TextFrame.MarginBottom - 1.4173219
    Next shp
    End If
    Exit Sub
err:
    MsgBox "Either the margin is already smaller than 0.05 or you haven't selected at least one shape"
    
End Sub

Public Sub TextMarTBInc(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    On Error GoTo err
    
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange
            .TextFrame.MarginTop = .TextFrame.MarginTop + 1.4173219
            .TextFrame.MarginBottom = .TextFrame.MarginBottom + 1.4173219
        End With
    
    Else
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.TextFrame.MarginTop = shp.TextFrame.MarginTop + 1.4173219
        shp.TextFrame.MarginBottom = shp.TextFrame.MarginBottom + 1.4173219
    Next shp
    End If
    Exit Sub
err:
    MsgBox "Please select at least one shape"
    
End Sub

Public Sub TextMarTBZero(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    On Error GoTo err
    
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange
            .TextFrame.MarginTop = 0
            .TextFrame.MarginBottom = 0
        End With
    
    Else
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.TextFrame.MarginTop = 0
        shp.TextFrame.MarginBottom = 0
    Next shp
    End If
    Exit Sub
err:
    MsgBox "Please select at least one shape"
    
End Sub

Public Sub TextMarTDec(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    On Error GoTo err
    
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange
            .TextFrame.MarginTop = .TextFrame.MarginTop - 1.4173219
        End With
    
    Else
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.TextFrame.MarginTop = shp.TextFrame.MarginTop - 1.4173219
    Next shp
    End If
    Exit Sub
err:
    MsgBox "Either the margin is already smaller than 0.05 or you haven't selected at least one shape"
    
End Sub

Public Sub TextMarTInc(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    On Error GoTo err
    
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange
            .TextFrame.MarginTop = .TextFrame.MarginTop + 1.4173219
        End With
    
    Else
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.TextFrame.MarginTop = shp.TextFrame.MarginTop + 1.4173219
    Next shp
    End If
    Exit Sub
err:
    MsgBox "Please select at least one shape"
    
End Sub

Public Sub TextMarTZero(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    On Error GoTo err
    
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange
            .TextFrame.MarginTop = 0
        End With
    
    Else
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.TextFrame.MarginTop = 0
    Next shp
    End If
    Exit Sub
err:
    MsgBox "Please select at least one shape"
    
End Sub

Public Sub TrafficNewFull(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

    Dim phd As Shape
    Dim back As Shape
    Dim red As Shape
    Dim yell As Shape
    Dim green As Shape
    Dim sld As Slide
    'Traffic lights new and full

On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else

    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)
    Set sld = Application.ActiveWindow.View.Slide
    
    Set back = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=10.204718, Top:=132.0944, Width:=15.590541, Height:=42.803123)
    back.Fill.ForeColor.RGB = RGB(150, 150, 150)
    back.Line.Visible = msoTrue
    back.Line.ForeColor.RGB = RGB(150, 150, 150)
    back.Line.Weight = 0.75

    Set red = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=12.472433, Top:=135.21251, Width:=10.204718, Height:=10.204718)
    red.Fill.ForeColor.RGB = RGB(255, 0, 0)
    red.Line.Visible = msoTrue
    red.Line.ForeColor.RGB = RGB(255, 255, 255)
    red.Line.Weight = 0.75
    red.Tags.Add "AMPELO", "YES"

    Set yell = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=12.472433, Top:=148.25188, Width:=10.204718, Height:=10.204718)
    yell.Fill.ForeColor.RGB = RGB(255, 255, 0)
    yell.Line.Visible = msoTrue
    yell.Line.ForeColor.RGB = RGB(255, 255, 255)
    yell.Line.Weight = 0.75
    yell.Tags.Add "AMPELM", "YES"

    Set green = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=12.472433, Top:=161.5747, Width:=10.204718, Height:=10.204718)
    green.Fill.ForeColor.RGB = RGB(0, 176, 80)
    green.Line.Visible = msoTrue
    green.Line.ForeColor.RGB = RGB(255, 255, 255)
    green.Line.Weight = 0.75
    green.Tags.Add "AMPELU", "YES"

    back.Select (True)
    red.Select (False)
    yell.Select (False)
    green.Select (False)
    ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, msoFalse
    ActiveWindow.Selection.ShapeRange.Group.Select
    ActiveWindow.Selection.ShapeRange.Tags.Add "AMPEL", "YES"
    ActiveWindow.Selection.ShapeRange.LockAspectRatio = msoTrue
    ActiveWindow.Selection.ShapeRange.Top = phd.Top

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub TrafficNewRed(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

    Dim phd As Shape
    Dim back As Shape
    Dim red As Shape
    Dim yell As Shape
    Dim green As Shape
    Dim sld As Slide
    'Traffic lights new and red

On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else

    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)
    Set sld = Application.ActiveWindow.View.Slide

    Set back = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=10.204718, Top:=200.69279, Width:=15.590541, Height:=42.803123)
    back.Fill.ForeColor.RGB = RGB(150, 150, 150)
    back.Line.Visible = msoTrue
    back.Line.ForeColor.RGB = RGB(150, 150, 150)
    back.Line.Weight = 0.75

    Set red = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=12.472433, Top:=203.52743, Width:=10.204718, Height:=10.204718)
    red.Fill.ForeColor.RGB = RGB(255, 0, 0)
    red.Line.Visible = msoTrue
    red.Line.ForeColor.RGB = RGB(255, 255, 255)
    red.Line.Weight = 0.75
    red.Tags.Add "AMPELO", "YES"

    Set yell = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=12.472433, Top:=216.85026, Width:=10.204718, Height:=10.204718)
    yell.Fill.ForeColor.RGB = RGB(186, 186, 186)
    yell.Line.Visible = msoTrue
    yell.Line.ForeColor.RGB = RGB(186, 186, 186)
    yell.Line.Weight = 0.75
    yell.Tags.Add "AMPELM", "YES"

    Set green = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=12.472433, Top:=230.17308, Width:=10.204718, Height:=10.204718)
    green.Fill.ForeColor.RGB = RGB(186, 186, 186)
    green.Line.Visible = msoTrue
    green.Line.ForeColor.RGB = RGB(186, 186, 186)
    green.Line.Weight = 0.75
    green.Tags.Add "AMPELU", "YES"

    back.Select (True)
    red.Select (False)
    yell.Select (False)
    green.Select (False)
    ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, msoFalse
    ActiveWindow.Selection.ShapeRange.Group.Select
    ActiveWindow.Selection.ShapeRange.Tags.Add "AMPEL", "YES"
    ActiveWindow.Selection.ShapeRange.LockAspectRatio = msoTrue
    ActiveWindow.Selection.ShapeRange.Top = phd.Top + 68
    
    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub TrafficNewYellow(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

    Dim phd As Shape
    Dim back As Shape
    Dim red As Shape
    Dim yell As Shape
    Dim green As Shape
    Dim sld As Slide
    'Traffic lights new and yellow

On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else

    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)
    Set sld = Application.ActiveWindow.View.Slide

    Set back = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=10.204718, Top:=269.0077, Width:=15.590541, Height:=42.803123)
    back.Fill.ForeColor.RGB = RGB(150, 150, 150)
    back.Line.Visible = msoTrue
    back.Line.ForeColor.RGB = RGB(150, 150, 150)
    back.Line.Weight = 0.75

    Set red = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=12.472433, Top:=272.12581, Width:=10.204718, Height:=10.204718)
    red.Fill.ForeColor.RGB = RGB(186, 186, 186)
    red.Line.Visible = msoTrue
    red.Line.ForeColor.RGB = RGB(186, 186, 186)
    red.Line.Weight = 0.75
    red.Tags.Add "AMPELO", "YES"

    Set yell = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=12.472433, Top:=285.44864, Width:=10.204718, Height:=10.204718)
    yell.Fill.ForeColor.RGB = RGB(255, 255, 0)
    yell.Line.Visible = msoTrue
    yell.Line.ForeColor.RGB = RGB(255, 255, 255)
    yell.Line.Weight = 0.75
    yell.Tags.Add "AMPELM", "YES"

    Set green = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=12.472433, Top:=298.488, Width:=10.204718, Height:=10.204718)
    green.Fill.ForeColor.RGB = RGB(186, 186, 186)
    green.Line.Visible = msoTrue
    green.Line.ForeColor.RGB = RGB(186, 186, 186)
    green.Line.Weight = 0.75
    green.Tags.Add "AMPELU", "YES"
    
    back.Select (True)
    red.Select (False)
    yell.Select (False)
    green.Select (False)
    ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, msoFalse
    ActiveWindow.Selection.ShapeRange.Group.Select
    ActiveWindow.Selection.ShapeRange.Tags.Add "AMPEL", "YES"
    ActiveWindow.Selection.ShapeRange.LockAspectRatio = msoTrue
    ActiveWindow.Selection.ShapeRange.Top = phd.Top + 68 + 68

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub TrafficNewGreen(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

    Dim phd As Shape
    Dim back As Shape
    Dim red As Shape
    Dim yell As Shape
    Dim green As Shape
    Dim sld As Slide
    'Traffic lights new and green
    
On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else

    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)
    Set sld = Application.ActiveWindow.View.Slide

    Set back = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=10.204718, Top:=337.60609, Width:=15.590541, Height:=42.803123)
    back.Fill.ForeColor.RGB = RGB(150, 150, 150)
    back.Line.Visible = msoTrue
    back.Line.ForeColor.RGB = RGB(150, 150, 150)
    back.Line.Weight = 0.75

    Set red = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=12.472433, Top:=340.44073, Width:=10.204718, Height:=10.204718)
    red.Fill.ForeColor.RGB = RGB(186, 186, 186)
    red.Line.Visible = msoTrue
    red.Line.ForeColor.RGB = RGB(186, 186, 186)
    red.Line.Weight = 0.75
    red.Tags.Add "AMPELO", "YES"

    Set yell = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=12.472433, Top:=353.76356, Width:=10.204718, Height:=10.204718)
    yell.Fill.ForeColor.RGB = RGB(186, 186, 186)
    yell.Line.Visible = msoTrue
    yell.Line.ForeColor.RGB = RGB(186, 186, 186)
    yell.Line.Weight = 0.75
    yell.Tags.Add "AMPELM", "YES"

    Set green = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=12.472433, Top:=367.08638, Width:=10.204718, Height:=10.204718)
    green.Fill.ForeColor.RGB = RGB(0, 176, 80)
    green.Line.Visible = msoTrue
    green.Line.ForeColor.RGB = RGB(255, 255, 255)
    green.Line.Weight = 0.75
    green.Tags.Add "AMPELU", "YES"
    
    back.Select (True)
    red.Select (False)
    yell.Select (False)
    green.Select (False)
    ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, msoFalse
    ActiveWindow.Selection.ShapeRange.Group.Select
    ActiveWindow.Selection.ShapeRange.Tags.Add "AMPEL", "YES"
    ActiveWindow.Selection.ShapeRange.LockAspectRatio = msoTrue
    ActiveWindow.Selection.ShapeRange.Top = phd.Top + 68 + 68 + 68

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub TrafficSetFull(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

Dim shp1 As Shape
Dim shp2 As Shape
Dim shp3 As Shape
Dim shp4 As Shape

    If ActiveWindow.Selection.Type <> ppSelectionShapes And ActiveWindow.Selection.Type <> ppSelectionText Then
        MsgBox "Please select a traffic light to switch"
        Exit Sub
    ElseIf ActiveWindow.Selection.ShapeRange.Count > 1 Then
        MsgBox "Please select only one traffic light to switch"
        Exit Sub
    ElseIf ActiveWindow.Selection.ShapeRange.Count = 1 And ActiveWindow.Selection.ShapeRange(1).Tags("AMPEL") <> "YES" Then
        MsgBox "Please select a traffic light to switch"
        Exit Sub
    Else
        GoTo SwitchIt
    End If

SwitchIt:
    ActiveWindow.Selection.ShapeRange(1).Ungroup.Select
        Set shp1 = ActiveWindow.Selection.ShapeRange(1)
        Set shp2 = ActiveWindow.Selection.ShapeRange(2)
        Set shp3 = ActiveWindow.Selection.ShapeRange(3)
        Set shp4 = ActiveWindow.Selection.ShapeRange(4)
        
    If shp1.Tags("AMPELO") = "YES" Then
        shp1.Fill.ForeColor.RGB = RGB(255, 0, 0)
        shp1.Line.ForeColor.RGB = RGB(255, 255, 255)
    ElseIf shp2.Tags("AMPELO") = "YES" Then
        shp2.Fill.ForeColor.RGB = RGB(255, 0, 0)
        shp2.Line.ForeColor.RGB = RGB(255, 255, 255)
    ElseIf shp3.Tags("AMPELO") = "YES" Then
        shp3.Fill.ForeColor.RGB = RGB(255, 0, 0)
        shp3.Line.ForeColor.RGB = RGB(255, 255, 255)
    ElseIf shp4.Tags("AMPELO") = "YES" Then
        shp4.Fill.ForeColor.RGB = RGB(255, 0, 0)
        shp4.Line.ForeColor.RGB = RGB(255, 255, 255)
    Else
    End If
    
    If shp1.Tags("AMPELM") = "YES" Then
        shp1.Fill.ForeColor.RGB = RGB(255, 255, 0)
        shp1.Line.ForeColor.RGB = RGB(255, 255, 255)
    ElseIf shp2.Tags("AMPELM") = "YES" Then
        shp2.Fill.ForeColor.RGB = RGB(255, 255, 0)
        shp2.Line.ForeColor.RGB = RGB(255, 255, 255)
    ElseIf shp3.Tags("AMPELM") = "YES" Then
        shp3.Fill.ForeColor.RGB = RGB(255, 255, 0)
        shp3.Line.ForeColor.RGB = RGB(255, 255, 255)
    ElseIf shp4.Tags("AMPELM") = "YES" Then
        shp4.Fill.ForeColor.RGB = RGB(255, 255, 0)
        shp4.Line.ForeColor.RGB = RGB(255, 255, 255)
    Else
    End If
    
    If shp1.Tags("AMPELU") = "YES" Then
        shp1.Fill.ForeColor.RGB = RGB(0, 176, 80)
        shp1.Line.ForeColor.RGB = RGB(255, 255, 255)
    ElseIf shp2.Tags("AMPELU") = "YES" Then
        shp2.Fill.ForeColor.RGB = RGB(0, 176, 80)
        shp2.Line.ForeColor.RGB = RGB(255, 255, 255)
    ElseIf shp3.Tags("AMPELU") = "YES" Then
        shp3.Fill.ForeColor.RGB = RGB(0, 176, 80)
        shp3.Line.ForeColor.RGB = RGB(255, 255, 255)
    ElseIf shp4.Tags("AMPELU") = "YES" Then
        shp4.Fill.ForeColor.RGB = RGB(0, 176, 80)
        shp4.Line.ForeColor.RGB = RGB(255, 255, 255)
    Else
    End If
    
    shp1.Select (True)
    shp2.Select (False)
    shp3.Select (False)
    shp4.Select (False)
    ActiveWindow.Selection.ShapeRange.Group.Select
    ActiveWindow.Selection.ShapeRange.Tags.Add "AMPEL", "YES"
    ActiveWindow.Selection.ShapeRange.LockAspectRatio = msoTrue
    
End Sub

Public Sub TrafficSetRed(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

Dim shp1 As Shape
Dim shp2 As Shape
Dim shp3 As Shape
Dim shp4 As Shape

    If ActiveWindow.Selection.Type <> ppSelectionShapes And ActiveWindow.Selection.Type <> ppSelectionText Then
        MsgBox "Please select a traffic light to switch"
        Exit Sub
    ElseIf ActiveWindow.Selection.ShapeRange.Count > 1 Then
        MsgBox "Please select only one traffic light to switch"
        Exit Sub
    ElseIf ActiveWindow.Selection.ShapeRange.Count = 1 And ActiveWindow.Selection.ShapeRange(1).Tags("AMPEL") <> "YES" Then
        MsgBox "Please select a traffic light to switch"
        Exit Sub
    Else
        GoTo SwitchIt
    End If

SwitchIt:
    ActiveWindow.Selection.ShapeRange(1).Ungroup.Select
        Set shp1 = ActiveWindow.Selection.ShapeRange(1)
        Set shp2 = ActiveWindow.Selection.ShapeRange(2)
        Set shp3 = ActiveWindow.Selection.ShapeRange(3)
        Set shp4 = ActiveWindow.Selection.ShapeRange(4)
        
    If shp1.Tags("AMPELO") = "YES" Then
        shp1.Fill.ForeColor.RGB = RGB(255, 0, 0)
        shp1.Line.ForeColor.RGB = RGB(255, 255, 255)
    ElseIf shp2.Tags("AMPELO") = "YES" Then
        shp2.Fill.ForeColor.RGB = RGB(255, 0, 0)
        shp2.Line.ForeColor.RGB = RGB(255, 255, 255)
    ElseIf shp3.Tags("AMPELO") = "YES" Then
        shp3.Fill.ForeColor.RGB = RGB(255, 0, 0)
        shp3.Line.ForeColor.RGB = RGB(255, 255, 255)
    ElseIf shp4.Tags("AMPELO") = "YES" Then
        shp4.Fill.ForeColor.RGB = RGB(255, 0, 0)
        shp4.Line.ForeColor.RGB = RGB(255, 255, 255)
    Else
    End If
    
    If shp1.Tags("AMPELM") = "YES" Then
        shp1.Fill.ForeColor.RGB = RGB(186, 186, 186)
        shp1.Line.ForeColor.RGB = RGB(186, 186, 186)
    ElseIf shp2.Tags("AMPELM") = "YES" Then
        shp2.Fill.ForeColor.RGB = RGB(186, 186, 186)
        shp2.Line.ForeColor.RGB = RGB(186, 186, 186)
    ElseIf shp3.Tags("AMPELM") = "YES" Then
        shp3.Fill.ForeColor.RGB = RGB(186, 186, 186)
        shp3.Line.ForeColor.RGB = RGB(186, 186, 186)
    ElseIf shp4.Tags("AMPELM") = "YES" Then
        shp4.Fill.ForeColor.RGB = RGB(186, 186, 186)
        shp4.Line.ForeColor.RGB = RGB(186, 186, 186)
    Else
    End If
    
    If shp1.Tags("AMPELU") = "YES" Then
        shp1.Fill.ForeColor.RGB = RGB(186, 186, 186)
        shp1.Line.ForeColor.RGB = RGB(186, 186, 186)
    ElseIf shp2.Tags("AMPELU") = "YES" Then
        shp2.Fill.ForeColor.RGB = RGB(186, 186, 186)
        shp2.Line.ForeColor.RGB = RGB(186, 186, 186)
    ElseIf shp3.Tags("AMPELU") = "YES" Then
        shp3.Fill.ForeColor.RGB = RGB(186, 186, 186)
        shp3.Line.ForeColor.RGB = RGB(186, 186, 186)
    ElseIf shp4.Tags("AMPELU") = "YES" Then
        shp4.Fill.ForeColor.RGB = RGB(186, 186, 186)
        shp4.Line.ForeColor.RGB = RGB(186, 186, 186)
    Else
    End If
    
    shp1.Select (True)
    shp2.Select (False)
    shp3.Select (False)
    shp4.Select (False)
    ActiveWindow.Selection.ShapeRange.Group.Select
    ActiveWindow.Selection.ShapeRange.Tags.Add "AMPEL", "YES"
    ActiveWindow.Selection.ShapeRange.LockAspectRatio = msoTrue
    
End Sub

Public Sub TrafficSetYellow(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

Dim shp1 As Shape
Dim shp2 As Shape
Dim shp3 As Shape
Dim shp4 As Shape

    If ActiveWindow.Selection.Type <> ppSelectionShapes And ActiveWindow.Selection.Type <> ppSelectionText Then
        MsgBox "Please select a traffic light to switch"
        Exit Sub
    ElseIf ActiveWindow.Selection.ShapeRange.Count > 1 Then
        MsgBox "Please select only one traffic light to switch"
        Exit Sub
    ElseIf ActiveWindow.Selection.ShapeRange.Count = 1 And ActiveWindow.Selection.ShapeRange(1).Tags("AMPEL") <> "YES" Then
        MsgBox "Please select a traffic light to switch"
        Exit Sub
    Else
        GoTo SwitchIt
    End If

SwitchIt:
    ActiveWindow.Selection.ShapeRange(1).Ungroup.Select
        Set shp1 = ActiveWindow.Selection.ShapeRange(1)
        Set shp2 = ActiveWindow.Selection.ShapeRange(2)
        Set shp3 = ActiveWindow.Selection.ShapeRange(3)
        Set shp4 = ActiveWindow.Selection.ShapeRange(4)
        
    If shp1.Tags("AMPELO") = "YES" Then
        shp1.Fill.ForeColor.RGB = RGB(186, 186, 186)
        shp1.Line.ForeColor.RGB = RGB(186, 186, 186)
    ElseIf shp2.Tags("AMPELO") = "YES" Then
        shp2.Fill.ForeColor.RGB = RGB(186, 186, 186)
        shp2.Line.ForeColor.RGB = RGB(186, 186, 186)
    ElseIf shp3.Tags("AMPELO") = "YES" Then
        shp3.Fill.ForeColor.RGB = RGB(186, 186, 186)
        shp3.Line.ForeColor.RGB = RGB(186, 186, 186)
    ElseIf shp4.Tags("AMPELO") = "YES" Then
        shp4.Fill.ForeColor.RGB = RGB(186, 186, 186)
        shp4.Line.ForeColor.RGB = RGB(186, 186, 186)
    Else
    End If
    
    If shp1.Tags("AMPELM") = "YES" Then
        shp1.Fill.ForeColor.RGB = RGB(255, 255, 0)
        shp1.Line.ForeColor.RGB = RGB(255, 255, 255)
    ElseIf shp2.Tags("AMPELM") = "YES" Then
        shp2.Fill.ForeColor.RGB = RGB(255, 255, 0)
        shp2.Line.ForeColor.RGB = RGB(255, 255, 255)
    ElseIf shp3.Tags("AMPELM") = "YES" Then
        shp3.Fill.ForeColor.RGB = RGB(255, 255, 0)
        shp3.Line.ForeColor.RGB = RGB(255, 255, 255)
    ElseIf shp4.Tags("AMPELM") = "YES" Then
        shp4.Fill.ForeColor.RGB = RGB(255, 255, 0)
        shp4.Line.ForeColor.RGB = RGB(255, 255, 255)
    Else
    End If
    
    If shp1.Tags("AMPELU") = "YES" Then
        shp1.Fill.ForeColor.RGB = RGB(186, 186, 186)
        shp1.Line.ForeColor.RGB = RGB(186, 186, 186)
    ElseIf shp2.Tags("AMPELU") = "YES" Then
        shp2.Fill.ForeColor.RGB = RGB(186, 186, 186)
        shp2.Line.ForeColor.RGB = RGB(186, 186, 186)
    ElseIf shp3.Tags("AMPELU") = "YES" Then
        shp3.Fill.ForeColor.RGB = RGB(186, 186, 186)
        shp3.Line.ForeColor.RGB = RGB(186, 186, 186)
    ElseIf shp4.Tags("AMPELU") = "YES" Then
        shp4.Fill.ForeColor.RGB = RGB(186, 186, 186)
        shp4.Line.ForeColor.RGB = RGB(186, 186, 186)
    Else
    End If
    
    shp1.Select (True)
    shp2.Select (False)
    shp3.Select (False)
    shp4.Select (False)
    ActiveWindow.Selection.ShapeRange.Group.Select
    ActiveWindow.Selection.ShapeRange.Tags.Add "AMPEL", "YES"
    ActiveWindow.Selection.ShapeRange.LockAspectRatio = msoTrue
    
End Sub

Public Sub TrafficSetGreen(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

Dim shp1 As Shape
Dim shp2 As Shape
Dim shp3 As Shape
Dim shp4 As Shape

    If ActiveWindow.Selection.Type <> ppSelectionShapes And ActiveWindow.Selection.Type <> ppSelectionText Then
        MsgBox "Please select a traffic light to switch"
        Exit Sub
    ElseIf ActiveWindow.Selection.ShapeRange.Count > 1 Then
        MsgBox "Please select only one traffic light to switch"
        Exit Sub
    ElseIf ActiveWindow.Selection.ShapeRange.Count = 1 And ActiveWindow.Selection.ShapeRange(1).Tags("AMPEL") <> "YES" Then
        MsgBox "Please select a traffic light to switch"
        Exit Sub
    Else
        GoTo SwitchIt
    End If

SwitchIt:
    ActiveWindow.Selection.ShapeRange(1).Ungroup.Select
        Set shp1 = ActiveWindow.Selection.ShapeRange(1)
        Set shp2 = ActiveWindow.Selection.ShapeRange(2)
        Set shp3 = ActiveWindow.Selection.ShapeRange(3)
        Set shp4 = ActiveWindow.Selection.ShapeRange(4)
        
    If shp1.Tags("AMPELO") = "YES" Then
        shp1.Fill.ForeColor.RGB = RGB(186, 186, 186)
        shp1.Line.ForeColor.RGB = RGB(186, 186, 186)
    ElseIf shp2.Tags("AMPELO") = "YES" Then
        shp2.Fill.ForeColor.RGB = RGB(186, 186, 186)
        shp2.Line.ForeColor.RGB = RGB(186, 186, 186)
    ElseIf shp3.Tags("AMPELO") = "YES" Then
        shp3.Fill.ForeColor.RGB = RGB(186, 186, 186)
        shp3.Line.ForeColor.RGB = RGB(186, 186, 186)
    ElseIf shp4.Tags("AMPELO") = "YES" Then
        shp4.Fill.ForeColor.RGB = RGB(186, 186, 186)
        shp4.Line.ForeColor.RGB = RGB(186, 186, 186)
    Else
    End If
    
    If shp1.Tags("AMPELM") = "YES" Then
        shp1.Fill.ForeColor.RGB = RGB(186, 186, 186)
        shp1.Line.ForeColor.RGB = RGB(186, 186, 186)
    ElseIf shp2.Tags("AMPELM") = "YES" Then
        shp2.Fill.ForeColor.RGB = RGB(186, 186, 186)
        shp2.Line.ForeColor.RGB = RGB(186, 186, 186)
    ElseIf shp3.Tags("AMPELM") = "YES" Then
        shp3.Fill.ForeColor.RGB = RGB(186, 186, 186)
        shp3.Line.ForeColor.RGB = RGB(186, 186, 186)
    ElseIf shp4.Tags("AMPELM") = "YES" Then
        shp4.Fill.ForeColor.RGB = RGB(186, 186, 186)
        shp4.Line.ForeColor.RGB = RGB(186, 186, 186)
    Else
    End If
    
    If shp1.Tags("AMPELU") = "YES" Then
        shp1.Fill.ForeColor.RGB = RGB(0, 176, 80)
        shp1.Line.ForeColor.RGB = RGB(255, 255, 255)
    ElseIf shp2.Tags("AMPELU") = "YES" Then
        shp2.Fill.ForeColor.RGB = RGB(0, 176, 80)
        shp2.Line.ForeColor.RGB = RGB(255, 255, 255)
    ElseIf shp3.Tags("AMPELU") = "YES" Then
        shp3.Fill.ForeColor.RGB = RGB(0, 176, 80)
        shp3.Line.ForeColor.RGB = RGB(255, 255, 255)
    ElseIf shp4.Tags("AMPELU") = "YES" Then
        shp4.Fill.ForeColor.RGB = RGB(0, 176, 80)
        shp4.Line.ForeColor.RGB = RGB(255, 255, 255)
    Else
    End If
    
    shp1.Select (True)
    shp2.Select (False)
    shp3.Select (False)
    shp4.Select (False)
    ActiveWindow.Selection.ShapeRange.Group.Select
    ActiveWindow.Selection.ShapeRange.Tags.Add "AMPEL", "YES"
    ActiveWindow.Selection.ShapeRange.LockAspectRatio = msoTrue
    
End Sub

Public Sub WidthDec(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    On Error GoTo err
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Width = shp.Width - 2.8346439
    Next shp
    Exit Sub
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub WidthInc(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp As Shape
    On Error GoTo err
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Width = shp.Width + 2.8346439
    Next shp
    Exit Sub
err:
    MsgBox "Please select at least one shape"
End Sub

'DDD Shapes

Public Sub BoxHori03(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim phd As Shape
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim shp3 As Shape
    Dim shp11 As Shape
    Dim shp12 As Shape
    Dim shp13 As Shape
    Dim shp21 As Shape
    Dim sld As Slide

    'Three horizontal boxes
    
On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
    MsgBox "This function cannot be used for several slides at the same time"
    Exit Sub
Else
    
    Set sld = Application.ActiveWindow.View.Slide
    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)
    
If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=127.55897, Height:=98.929072)
    shp1.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=120.47237, Height:=73.984205)
    shp1.Select msoTrue
Else
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=120.47237, Height:=98.929072)
    shp1.Select msoTrue
End If

Call ParameterBoxHeaderShadow
    

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=127.55897, Height:=98.929072)
    shp2.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=120.47237, Height:=73.984205)
    shp2.Select msoTrue
Else
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=120.47237, Height:=98.929072)
    shp2.Select msoTrue
End If

Call ParameterBoxHeaderShadow


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=127.55897, Height:=98.929072)
    shp3.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=120.47237, Height:=73.984205)
    shp3.Select msoTrue
Else
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=120.47237, Height:=98.929072)
    shp3.Select msoTrue
End If

Call ParameterBoxHeaderShadow


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=574.01539, Height:=98.929072)
    shp11.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=542.55084, Height:=73.984205)
    shp11.Select msoTrue
Else
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=542.55084, Height:=98.929072)
    shp11.Select msoTrue
End If

Call ParameterBoxBodyShadow


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=574.01539, Height:=98.929072)
    shp12.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=542.55084, Height:=73.984205)
    shp12.Select msoTrue
Else
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=542.55084, Height:=98.929072)
    shp12.Select msoTrue
End If

Call ParameterBoxBodyShadow


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=574.01539, Height:=98.929072)
    shp13.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=542.55084, Height:=73.984205)
    shp13.Select msoTrue
Else
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=542.55084, Height:=98.929072)
    shp13.Select msoTrue
End If

Call ParameterBoxBodyShadow


shp1.Left = phd.Left
shp2.Left = phd.Left
shp3.Left = phd.Left
shp11.Left = shp1.Left + shp1.Width
shp12.Left = shp2.Left + shp2.Width
shp13.Left = shp3.Left + shp3.Width


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp21 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=10, Height:=10.771647)
Else
    Set shp21 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=10, Height:=14.456684)
End If

shp1.Top = phd.Top
shp11.Top = phd.Top
shp2.Top = shp1.Top + shp1.Height + shp21.Height
shp12.Top = shp1.Top + shp1.Height + shp21.Height
shp3.Top = shp2.Top + shp2.Height + shp21.Height
shp13.Top = shp12.Top + shp12.Height + shp21.Height

shp21.Delete

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    shp1.TextFrame2.TextRange.Font.Size = 14
    shp2.TextFrame2.TextRange.Font.Size = 14
    shp3.TextFrame2.TextRange.Font.Size = 14
    shp11.TextFrame2.TextRange.Font.Size = 12
    shp12.TextFrame2.TextRange.Font.Size = 12
    shp13.TextFrame2.TextRange.Font.Size = 12
End If

shp1.Select (msoTrue)
shp2.Select (msoFalse)
shp3.Select (msoFalse)
shp11.Select (msoFalse)
shp12.Select (msoFalse)
shp13.Select (msoFalse)

    ActiveWindow.Selection.ShapeRange.Group.Select
    ActiveWindow.Selection.ShapeRange(1).Width = phd.Width
    ActiveWindow.Selection.ShapeRange(1).Ungroup.Select

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub BoxHori04(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim phd As Shape
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim shp3 As Shape
    Dim shp4 As Shape
    Dim shp11 As Shape
    Dim shp12 As Shape
    Dim shp13 As Shape
    Dim shp14 As Shape
    Dim shp21 As Shape
    Dim sld As Slide

    'Four horizontal boxes
    
On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
    MsgBox "This function cannot be used for several slides at the same time"
    Exit Sub
Else
    
    Set sld = Application.ActiveWindow.View.Slide
    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)
    
If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=127.55897, Height:=70.866097)
    shp1.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=120.47237, Height:=53.007841)
    shp1.Select msoTrue
Else
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=120.47237, Height:=70.866097)
    shp1.Select msoTrue
End If

Call ParameterBoxHeaderShadow
    

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=127.55897, Height:=70.866097)
    shp2.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=120.47237, Height:=53.007841)
    shp2.Select msoTrue
Else
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=120.47237, Height:=70.866097)
    shp2.Select msoTrue
End If

Call ParameterBoxHeaderShadow


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=127.55897, Height:=70.866097)
    shp3.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=120.47237, Height:=53.007841)
    shp3.Select msoTrue
Else
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=120.47237, Height:=70.866097)
    shp3.Select msoTrue
End If

Call ParameterBoxHeaderShadow


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp4 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=127.55897, Height:=70.866097)
    shp4.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp4 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=120.47237, Height:=53.007841)
    shp4.Select msoTrue
Else
    Set shp4 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=120.47237, Height:=70.866097)
    shp4.Select msoTrue
End If

Call ParameterBoxHeaderShadow


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=574.01539, Height:=70.866097)
    shp11.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=542.55084, Height:=53.007841)
    shp11.Select msoTrue
Else
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=542.55084, Height:=70.866097)
    shp11.Select msoTrue
End If

Call ParameterBoxBodyShadow


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=574.01539, Height:=70.866097)
    shp12.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=542.55084, Height:=53.007841)
    shp12.Select msoTrue
Else
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=542.55084, Height:=70.866097)
    shp12.Select msoTrue
End If

Call ParameterBoxBodyShadow


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=574.01539, Height:=70.866097)
    shp13.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=542.55084, Height:=53.007841)
    shp13.Select msoTrue
Else
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=542.55084, Height:=70.866097)
    shp13.Select msoTrue
End If

Call ParameterBoxBodyShadow


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp14 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=574.01539, Height:=70.866097)
    shp14.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp14 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=542.55084, Height:=53.007841)
    shp14.Select msoTrue
Else
    Set shp14 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=542.55084, Height:=70.866097)
    shp14.Select msoTrue
End If

Call ParameterBoxBodyShadow


shp1.Left = phd.Left
shp2.Left = phd.Left
shp3.Left = phd.Left
shp4.Left = phd.Left
shp11.Left = shp1.Left + shp1.Width
shp12.Left = shp2.Left + shp2.Width
shp13.Left = shp3.Left + shp3.Width
shp14.Left = shp4.Left + shp4.Width

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp21 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=10, Height:=10.488182)
Else
    Set shp21 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=10, Height:=14.173219)
End If

shp1.Top = phd.Top
shp11.Top = phd.Top
shp2.Top = shp1.Top + shp1.Height + shp21.Height
shp12.Top = shp1.Top + shp1.Height + shp21.Height
shp3.Top = shp2.Top + shp2.Height + shp21.Height
shp13.Top = shp12.Top + shp12.Height + shp21.Height
shp4.Top = shp3.Top + shp3.Height + shp21.Height
shp14.Top = shp13.Top + shp13.Height + shp21.Height

shp21.Delete

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    shp1.TextFrame2.TextRange.Font.Size = 14
    shp2.TextFrame2.TextRange.Font.Size = 14
    shp3.TextFrame2.TextRange.Font.Size = 14
    shp4.TextFrame2.TextRange.Font.Size = 14
    shp11.TextFrame2.TextRange.Font.Size = 12
    shp12.TextFrame2.TextRange.Font.Size = 12
    shp13.TextFrame2.TextRange.Font.Size = 12
    shp14.TextFrame2.TextRange.Font.Size = 12
End If

shp1.Select (msoTrue)
shp2.Select (msoFalse)
shp3.Select (msoFalse)
shp4.Select (msoFalse)
shp11.Select (msoFalse)
shp12.Select (msoFalse)
shp13.Select (msoFalse)
shp14.Select (msoFalse)

    ActiveWindow.Selection.ShapeRange.Group.Select
    ActiveWindow.Selection.ShapeRange(1).Width = phd.Width
    ActiveWindow.Selection.ShapeRange(1).Ungroup.Select

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub BoxHori05(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim phd As Shape
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim shp3 As Shape
    Dim shp4 As Shape
    Dim shp5 As Shape
    Dim shp11 As Shape
    Dim shp12 As Shape
    Dim shp13 As Shape
    Dim shp14 As Shape
    Dim shp15 As Shape
    Dim shp21 As Shape
    Dim sld As Slide

    'Five horizontal boxes
    
On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
    MsgBox "This function cannot be used for several slides at the same time"
    Exit Sub
Else
    
    Set sld = Application.ActiveWindow.View.Slide
    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)
    
If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=127.55897, Height:=55.842485)
    shp1.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=120.47237, Height:=40.818872)
    shp1.Select msoTrue
Else
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=120.47237, Height:=55.842485)
    shp1.Select msoTrue
End If

Call ParameterBoxHeaderShadow
    

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=127.55897, Height:=55.842485)
    shp2.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=120.47237, Height:=40.818872)
    shp2.Select msoTrue
Else
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=120.47237, Height:=55.842485)
    shp2.Select msoTrue
End If

Call ParameterBoxHeaderShadow


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=127.55897, Height:=55.842485)
    shp3.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=120.47237, Height:=40.818872)
    shp3.Select msoTrue
Else
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=120.47237, Height:=55.842485)
    shp3.Select msoTrue
End If

Call ParameterBoxHeaderShadow


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp4 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=127.55897, Height:=55.842485)
    shp4.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp4 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=120.47237, Height:=40.818872)
    shp4.Select msoTrue
Else
    Set shp4 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=120.47237, Height:=55.842485)
    shp4.Select msoTrue
End If

Call ParameterBoxHeaderShadow


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp5 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=127.55897, Height:=55.842485)
    shp5.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp5 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=120.47237, Height:=40.818872)
    shp5.Select msoTrue
Else
    Set shp5 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=120.47237, Height:=55.842485)
    shp5.Select msoTrue
End If

Call ParameterBoxHeaderShadow


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=574.01539, Height:=55.842485)
    shp11.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=542.55084, Height:=40.818872)
    shp11.Select msoTrue
Else
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=542.55084, Height:=55.842485)
    shp11.Select msoTrue
End If

Call ParameterBoxBodyShadow


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=574.01539, Height:=55.842485)
    shp12.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=542.55084, Height:=40.818872)
    shp12.Select msoTrue
Else
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=542.55084, Height:=55.842485)
    shp12.Select msoTrue
End If

Call ParameterBoxBodyShadow


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=574.01539, Height:=55.842485)
    shp13.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=542.55084, Height:=40.818872)
    shp13.Select msoTrue
Else
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=542.55084, Height:=55.842485)
    shp13.Select msoTrue
End If

Call ParameterBoxBodyShadow


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp14 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=574.01539, Height:=55.842485)
    shp14.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp14 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=542.55084, Height:=40.818872)
    shp14.Select msoTrue
Else
    Set shp14 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=542.55084, Height:=55.842485)
    shp14.Select msoTrue
End If

Call ParameterBoxBodyShadow


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp15 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=574.01539, Height:=55.842485)
    shp15.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp15 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=542.55084, Height:=40.818872)
    shp15.Select msoTrue
Else
    Set shp15 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=542.55084, Height:=55.842485)
    shp15.Select msoTrue
End If

Call ParameterBoxBodyShadow


shp1.Left = phd.Left
shp2.Left = phd.Left
shp3.Left = phd.Left
shp4.Left = phd.Left
shp5.Left = phd.Left
shp11.Left = shp1.Left + shp1.Width
shp12.Left = shp2.Left + shp2.Width
shp13.Left = shp3.Left + shp3.Width
shp14.Left = shp4.Left + shp4.Width
shp15.Left = shp5.Left + shp5.Width

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp21 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=10, Height:=9.9212536)
Else
    Set shp21 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=10, Height:=11.62204)
End If

shp1.Top = phd.Top
shp11.Top = phd.Top
shp2.Top = shp1.Top + shp1.Height + shp21.Height
shp12.Top = shp1.Top + shp1.Height + shp21.Height
shp3.Top = shp2.Top + shp2.Height + shp21.Height
shp13.Top = shp12.Top + shp12.Height + shp21.Height
shp4.Top = shp3.Top + shp3.Height + shp21.Height
shp14.Top = shp13.Top + shp13.Height + shp21.Height
shp5.Top = shp4.Top + shp4.Height + shp21.Height
shp15.Top = shp14.Top + shp14.Height + shp21.Height

shp21.Delete

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    shp1.TextFrame2.TextRange.Font.Size = 10
    shp2.TextFrame2.TextRange.Font.Size = 10
    shp3.TextFrame2.TextRange.Font.Size = 10
    shp4.TextFrame2.TextRange.Font.Size = 10
    shp5.TextFrame2.TextRange.Font.Size = 10
    shp11.TextFrame2.TextRange.Font.Size = 9
    shp12.TextFrame2.TextRange.Font.Size = 9
    shp13.TextFrame2.TextRange.Font.Size = 9
    shp14.TextFrame2.TextRange.Font.Size = 9
    shp15.TextFrame2.TextRange.Font.Size = 9
Else
    shp1.TextFrame2.TextRange.Font.Size = 14
    shp2.TextFrame2.TextRange.Font.Size = 14
    shp3.TextFrame2.TextRange.Font.Size = 14
    shp4.TextFrame2.TextRange.Font.Size = 14
    shp5.TextFrame2.TextRange.Font.Size = 14
    shp11.TextFrame2.TextRange.Font.Size = 12
    shp12.TextFrame2.TextRange.Font.Size = 12
    shp13.TextFrame2.TextRange.Font.Size = 12
    shp14.TextFrame2.TextRange.Font.Size = 12
    shp15.TextFrame2.TextRange.Font.Size = 12
End If

shp1.Select (msoTrue)
shp2.Select (msoFalse)
shp3.Select (msoFalse)
shp4.Select (msoFalse)
shp5.Select (msoFalse)
shp11.Select (msoFalse)
shp12.Select (msoFalse)
shp13.Select (msoFalse)
shp14.Select (msoFalse)
shp15.Select (msoFalse)

    ActiveWindow.Selection.ShapeRange.Group.Select
    ActiveWindow.Selection.ShapeRange(1).Width = phd.Width
    ActiveWindow.Selection.ShapeRange(1).Ungroup.Select

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub BoxVert02(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim shp11 As Shape
    Dim shp12 As Shape
    Dim shp21 As Shape
    Dim sld As Slide

    'Two vertical boxes, edited for Ommax
    
On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else
    
    Set sld = Application.ActiveWindow.View.Slide
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=29.763761, Top:=82.204673, Width:=330.23601, Height:=19.842507)
    shp1.Select msoTrue

Call ParameterBoxHeaderFlat
    
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=359.99977, Top:=82.204673, Width:=330.23601, Height:=19.842507)
    shp2.Select msoTrue

Call ParameterBoxHeaderFlat

    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=29.763761, Top:=102.04718, Width:=330.23601, Height:=370.20449)
    shp11.Select msoTrue

Call ParameterColumnBodyStandard

    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=359.99977, Top:=102.04718, Width:=330.23601, Height:=370.20449)
    shp12.Select msoTrue

Call ParameterColumnBodyStandard

shp1.TextFrame2.TextRange.Characters.Text = "Header 1"
shp2.TextFrame2.TextRange.Characters.Text = "Header 2"
shp11.TextFrame2.TextRange.Characters.Text = "Text 1"
shp12.TextFrame2.TextRange.Characters.Text = "Text 2"

shp1.TextFrame2.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignLeft
shp2.TextFrame2.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignLeft

shp1.TextFrame2.TextRange.Font.Size = 10
shp2.TextFrame2.TextRange.Font.Size = 10

Set shp21 = sld.Shapes.AddLine(BeginX:=359.99977, BeginY:=82.204673, EndX:=359.99977, EndY:=474.23592)
With shp21
    .Line.Visible = msoTrue
    .Line.ForeColor.RGB = RGB(18, 126, 129)
    .Line.Weight = 0.75
End With

shp1.Select (msoTrue)
shp2.Select (msoFalse)
shp11.Select (msoFalse)
shp12.Select (msoFalse)
shp21.Select (msoFalse)

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub BoxVert03(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim phd As Shape
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim shp3 As Shape
    Dim shp11 As Shape
    Dim shp12 As Shape
    Dim shp13 As Shape
    Dim sld As Slide

    'Three vertical boxes
    
On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else
    
    Set sld = Application.ActiveWindow.View.Slide
    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=211.18097, Height:=34.015727)
    shp1.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=31.464547)
    shp1.Select msoTrue
Else
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=34.015727)
    shp1.Select msoTrue
End If

Call ParameterBoxHeaderFlat
    

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=211.18097, Height:=34.015727)
    shp2.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=31.464547)
    shp2.Select msoTrue
Else
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=34.015727)
    shp2.Select msoTrue
End If

Call ParameterBoxHeaderFlat


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=211.18097, Height:=34.015727)
    shp3.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=31.464547)
    shp3.Select msoTrue
Else
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=34.015727)
    shp3.Select msoTrue
End If

Call ParameterBoxHeaderFlat


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=211.18097, Height:=283.46439)
    shp11.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=252.85023)
    shp11.Select msoTrue
Else
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=283.46439)
    shp11.Select msoTrue
End If

Call ParameterColumnBodyStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=211.18097, Height:=283.46439)
    shp12.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=252.85023)
    shp12.Select msoTrue
Else
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=283.46439)
    shp12.Select msoTrue
End If

Call ParameterColumnBodyStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=211.18097, Height:=283.46439)
    shp13.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=252.85023)
    shp13.Select msoTrue
Else
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=283.46439)
    shp13.Select msoTrue
End If

Call ParameterColumnBodyStandard


shp1.Top = phd.Top
shp2.Top = phd.Top
shp3.Top = phd.Top
shp11.Top = shp1.Top + shp1.Height
shp12.Top = shp2.Top + shp2.Height
shp13.Top = shp3.Top + shp3.Height

shp1.Left = phd.Left
shp2.Left = phd.Left + 10
shp3.Left = (phd.Left + phd.Width) - shp3.Width
shp11.Left = phd.Left
shp12.Left = phd.Left + 10
shp13.Left = (phd.Left + phd.Width) - shp13.Width

shp1.Select (msoTrue)
shp2.Select (msoFalse)
shp3.Select (msoFalse)
ActiveWindow.Selection.ShapeRange.Distribute msoDistributeHorizontally, msoFalse
shp11.Select (msoTrue)
shp12.Select (msoFalse)
shp13.Select (msoFalse)
ActiveWindow.Selection.ShapeRange.Distribute msoDistributeHorizontally, msoFalse


    shp11.Fill.Visible = msoTrue
    shp11.Fill.ForeColor.RGB = RGB(221, 221, 221)
    shp11.Line.Visible = msoTrue
    shp11.Line.ForeColor.RGB = RGB(221, 221, 221)
    shp11.Line.Weight = 0.75
    
    shp12.Fill.Visible = msoTrue
    shp12.Fill.ForeColor.RGB = RGB(221, 221, 221)
    shp12.Line.Visible = msoTrue
    shp12.Line.ForeColor.RGB = RGB(221, 221, 221)
    shp12.Line.Weight = 0.75
    
    shp13.Fill.Visible = msoTrue
    shp13.Fill.ForeColor.RGB = RGB(221, 221, 221)
    shp13.Line.Visible = msoTrue
    shp13.Line.ForeColor.RGB = RGB(221, 221, 221)
    shp13.Line.Weight = 0.75
    

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    shp1.TextFrame2.TextRange.Font.Size = 14
    shp2.TextFrame2.TextRange.Font.Size = 14
    shp3.TextFrame2.TextRange.Font.Size = 14
    shp11.TextFrame2.TextRange.Font.Size = 12
    shp12.TextFrame2.TextRange.Font.Size = 12
    shp13.TextFrame2.TextRange.Font.Size = 12
End If

shp1.Select (msoTrue)
shp2.Select (msoFalse)
shp3.Select (msoFalse)
shp11.Select (msoFalse)
shp12.Select (msoFalse)
shp13.Select (msoFalse)

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub BoxVert04(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim phd As Shape
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim shp3 As Shape
    Dim shp4 As Shape
    Dim shp11 As Shape
    Dim shp12 As Shape
    Dim shp13 As Shape
    Dim shp14 As Shape
    Dim sld As Slide

    'Four vertical boxes
    
On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else
    
    Set sld = Application.ActiveWindow.View.Slide
    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=160.44084, Height:=34.015727)
    shp1.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=31.464547)
    shp1.Select msoTrue
Else
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=34.015727)
    shp1.Select msoTrue
End If

Call ParameterBoxHeaderFlat
    

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=160.44084, Height:=34.015727)
    shp2.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=31.464547)
    shp2.Select msoTrue
Else
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=34.015727)
    shp2.Select msoTrue
End If

Call ParameterBoxHeaderFlat


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=160.44084, Height:=34.015727)
    shp3.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=31.464547)
    shp3.Select msoTrue
Else
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=34.015727)
    shp3.Select msoTrue
End If

Call ParameterBoxHeaderFlat


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp4 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=160.44084, Height:=34.015727)
    shp4.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp4 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=31.464547)
    shp4.Select msoTrue
Else
    Set shp4 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=34.015727)
    shp4.Select msoTrue
End If

Call ParameterBoxHeaderFlat


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=160.44084, Height:=283.46439)
    shp11.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=252.85023)
    shp11.Select msoTrue
Else
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=283.46439)
    shp11.Select msoTrue
End If

Call ParameterColumnBodyStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=160.44084, Height:=283.46439)
    shp12.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=252.85023)
    shp12.Select msoTrue
Else
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=283.46439)
    shp12.Select msoTrue
End If

Call ParameterColumnBodyStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=160.44084, Height:=283.46439)
    shp13.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=252.85023)
    shp13.Select msoTrue
Else
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=283.46439)
    shp13.Select msoTrue
End If

Call ParameterColumnBodyStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp14 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=160.44084, Height:=283.46439)
    shp14.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp14 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=252.85023)
    shp14.Select msoTrue
Else
    Set shp14 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=283.46439)
    shp14.Select msoTrue
End If

Call ParameterColumnBodyStandard


shp1.Top = phd.Top
shp2.Top = phd.Top
shp3.Top = phd.Top
shp4.Top = phd.Top
shp11.Top = shp1.Top + shp1.Height
shp12.Top = shp2.Top + shp2.Height
shp13.Top = shp3.Top + shp3.Height
shp14.Top = shp4.Top + shp4.Height

shp1.Left = phd.Left
shp2.Left = phd.Left + 10
shp3.Left = phd.Left + 10 + 10
shp4.Left = (phd.Left + phd.Width) - shp4.Width
shp11.Left = phd.Left
shp12.Left = phd.Left + 10
shp13.Left = phd.Left + 10 + 10
shp14.Left = (phd.Left + phd.Width) - shp14.Width

shp1.Select (msoTrue)
shp2.Select (msoFalse)
shp3.Select (msoFalse)
shp4.Select (msoFalse)
ActiveWindow.Selection.ShapeRange.Distribute msoDistributeHorizontally, msoFalse
shp11.Select (msoTrue)
shp12.Select (msoFalse)
shp13.Select (msoFalse)
shp14.Select (msoFalse)
ActiveWindow.Selection.ShapeRange.Distribute msoDistributeHorizontally, msoFalse

shp11.Fill.Visible = msoTrue
    shp11.Fill.ForeColor.RGB = RGB(221, 221, 221)
    shp11.Line.Visible = msoTrue
    shp11.Line.ForeColor.RGB = RGB(221, 221, 221)
    shp11.Line.Weight = 0.75
    
    shp12.Fill.Visible = msoTrue
    shp12.Fill.ForeColor.RGB = RGB(221, 221, 221)
    shp12.Line.Visible = msoTrue
    shp12.Line.ForeColor.RGB = RGB(221, 221, 221)
    shp12.Line.Weight = 0.75
    
    shp13.Fill.Visible = msoTrue
    shp13.Fill.ForeColor.RGB = RGB(221, 221, 221)
    shp13.Line.Visible = msoTrue
    shp13.Line.ForeColor.RGB = RGB(221, 221, 221)
    shp13.Line.Weight = 0.75
    
    shp14.Fill.Visible = msoTrue
    shp14.Fill.ForeColor.RGB = RGB(221, 221, 221)
    shp14.Line.Visible = msoTrue
    shp14.Line.ForeColor.RGB = RGB(221, 221, 221)
    shp14.Line.Weight = 0.75


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    shp1.TextFrame2.TextRange.Font.Size = 14
    shp2.TextFrame2.TextRange.Font.Size = 14
    shp3.TextFrame2.TextRange.Font.Size = 14
    shp4.TextFrame2.TextRange.Font.Size = 14
    shp11.TextFrame2.TextRange.Font.Size = 12
    shp12.TextFrame2.TextRange.Font.Size = 12
    shp13.TextFrame2.TextRange.Font.Size = 12
    shp14.TextFrame2.TextRange.Font.Size = 12
End If

shp1.Select (msoTrue)
shp2.Select (msoFalse)
shp3.Select (msoFalse)
shp4.Select (msoFalse)
shp11.Select (msoFalse)
shp12.Select (msoFalse)
shp13.Select (msoFalse)
shp14.Select (msoFalse)

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub Column02(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim phd As Shape
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim shp11 As Shape
    Dim shp12 As Shape
    Dim sld As Slide

    'Two Columns

On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else

    Set sld = Application.ActiveWindow.View.Slide
    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)
    
If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=323.9998, Height:=34.015727)
    shp1.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=306.14154, Height:=31.464547)
    shp1.Select msoTrue
Else
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=306.14154, Height:=34.015727)
    shp1.Select msoTrue
End If

Call ParameterColumnHeader
    

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=323.9998, Height:=34.015727)
    shp2.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=306.14154, Height:=31.464547)
    shp2.Select msoTrue
Else
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=306.14154, Height:=34.015727)
    shp2.Select msoTrue
End If

Call ParameterColumnHeader


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=323.9998, Height:=283.46439)
    shp11.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=306.14154, Height:=252.85023)
    shp11.Select msoTrue
Else
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=306.14154, Height:=283.46439)
    shp11.Select msoTrue
End If

Call ParameterColumnBodyStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=323.9998, Height:=283.46439)
    shp12.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=306.14154, Height:=252.85023)
    shp12.Select msoTrue
Else
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=306.14154, Height:=283.46439)
    shp12.Select msoTrue
End If

Call ParameterColumnBodyStandard


shp1.Top = phd.Top
shp2.Top = phd.Top
shp11.Top = shp1.Top + shp1.Height
shp12.Top = shp2.Top + shp2.Height

shp1.Left = phd.Left
shp2.Left = (phd.Left + phd.Width) - shp2.Width
shp11.Left = phd.Left
shp12.Left = (phd.Left + phd.Width) - shp12.Width

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    shp1.TextFrame2.TextRange.Font.Size = 14
    shp2.TextFrame2.TextRange.Font.Size = 14
    shp11.TextFrame2.TextRange.Font.Size = 12
    shp12.TextFrame2.TextRange.Font.Size = 12
End If

shp1.Select (msoTrue)
shp2.Select (msoFalse)
shp11.Select (msoFalse)
shp12.Select (msoFalse)

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub Column03(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim phd As Shape
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim shp3 As Shape
    Dim shp11 As Shape
    Dim shp12 As Shape
    Dim shp13 As Shape
    Dim sld As Slide

    'Three Columns

On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else

    Set sld = Application.ActiveWindow.View.Slide
    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)
    
If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=211.18097, Height:=34.015727)
    shp1.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=31.464547)
    shp1.Select msoTrue
Else
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=34.015727)
    shp1.Select msoTrue
End If

Call ParameterColumnHeader
    

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=211.18097, Height:=34.015727)
    shp2.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=31.464547)
    shp2.Select msoTrue
Else
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=34.015727)
    shp2.Select msoTrue
End If

Call ParameterColumnHeader


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=211.18097, Height:=34.015727)
    shp3.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=31.464547)
    shp3.Select msoTrue
Else
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=34.015727)
    shp3.Select msoTrue
End If

Call ParameterColumnHeader


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=211.18097, Height:=283.46439)
    shp11.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=252.85023)
    shp11.Select msoTrue
Else
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=283.46439)
    shp11.Select msoTrue
End If

Call ParameterColumnBodyStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=211.18097, Height:=283.46439)
    shp12.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=252.85023)
    shp12.Select msoTrue
Else
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=283.46439)
    shp12.Select msoTrue
End If

Call ParameterColumnBodyStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=211.18097, Height:=283.46439)
    shp13.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=252.85023)
    shp13.Select msoTrue
Else
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=283.46439)
    shp13.Select msoTrue
End If

Call ParameterColumnBodyStandard

shp1.Top = phd.Top
shp2.Top = phd.Top
shp3.Top = phd.Top
shp11.Top = shp1.Top + shp1.Height
shp12.Top = shp2.Top + shp2.Height
shp13.Top = shp3.Top + shp3.Height

shp1.Left = phd.Left
shp2.Left = phd.Left + 10
shp3.Left = (phd.Left + phd.Width) - shp3.Width
shp11.Left = phd.Left
shp12.Left = phd.Left + 10
shp13.Left = (phd.Left + phd.Width) - shp13.Width

shp1.Select (msoTrue)
shp2.Select (msoFalse)
shp3.Select (msoFalse)
ActiveWindow.Selection.ShapeRange.Distribute msoDistributeHorizontally, msoFalse
shp11.Select (msoTrue)
shp12.Select (msoFalse)
shp13.Select (msoFalse)
ActiveWindow.Selection.ShapeRange.Distribute msoDistributeHorizontally, msoFalse

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    shp1.TextFrame2.TextRange.Font.Size = 14
    shp2.TextFrame2.TextRange.Font.Size = 14
    shp3.TextFrame2.TextRange.Font.Size = 14
    shp11.TextFrame2.TextRange.Font.Size = 12
    shp12.TextFrame2.TextRange.Font.Size = 12
    shp13.TextFrame2.TextRange.Font.Size = 12
End If

shp1.Select (msoTrue)
shp2.Select (msoFalse)
shp3.Select (msoFalse)
shp11.Select (msoFalse)
shp12.Select (msoFalse)
shp13.Select (msoFalse)

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub Column04(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim phd As Shape
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim shp3 As Shape
    Dim shp4 As Shape
    Dim shp11 As Shape
    Dim shp12 As Shape
    Dim shp13 As Shape
    Dim shp14 As Shape
    Dim sld As Slide

    'Four Columns

On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else

    Set sld = Application.ActiveWindow.View.Slide
    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)
    
If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=160.44084, Height:=34.015727)
    shp1.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=31.464547)
    shp1.Select msoTrue
Else
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=34.015727)
    shp1.Select msoTrue
End If

Call ParameterColumnHeader
    

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=160.44084, Height:=34.015727)
    shp2.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=31.464547)
    shp2.Select msoTrue
Else
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=34.015727)
    shp2.Select msoTrue
End If

Call ParameterColumnHeader


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=160.44084, Height:=34.015727)
    shp3.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=31.464547)
    shp3.Select msoTrue
Else
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=34.015727)
    shp3.Select msoTrue
End If

Call ParameterColumnHeader


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp4 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=160.44084, Height:=34.015727)
    shp4.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp4 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=31.464547)
    shp4.Select msoTrue
Else
    Set shp4 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=34.015727)
    shp4.Select msoTrue
End If

Call ParameterColumnHeader


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=160.44084, Height:=283.46439)
    shp11.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=252.85023)
    shp11.Select msoTrue
Else
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=283.46439)
    shp11.Select msoTrue
End If

Call ParameterColumnBodyStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=160.44084, Height:=283.46439)
    shp12.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=252.85023)
    shp12.Select msoTrue
Else
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=283.46439)
    shp12.Select msoTrue
End If

Call ParameterColumnBodyStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=160.44084, Height:=283.46439)
    shp13.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=252.85023)
    shp13.Select msoTrue
Else
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=283.46439)
    shp13.Select msoTrue
End If

Call ParameterColumnBodyStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp14 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=160.44084, Height:=283.46439)
    shp14.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp14 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=252.85023)
    shp14.Select msoTrue
Else
    Set shp14 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=151.65345, Height:=283.46439)
    shp14.Select msoTrue
End If

Call ParameterColumnBodyStandard


shp1.Top = phd.Top
shp2.Top = phd.Top
shp3.Top = phd.Top
shp4.Top = phd.Top
shp11.Top = shp1.Top + shp1.Height
shp12.Top = shp2.Top + shp2.Height
shp13.Top = shp3.Top + shp3.Height
shp14.Top = shp4.Top + shp4.Height

shp1.Left = phd.Left
shp2.Left = phd.Left + 10
shp3.Left = phd.Left + 10 + 10
shp4.Left = (phd.Left + phd.Width) - shp4.Width
shp11.Left = phd.Left
shp12.Left = phd.Left + 10
shp13.Left = phd.Left + 10 + 10
shp14.Left = (phd.Left + phd.Width) - shp14.Width

shp1.Select (msoTrue)
shp2.Select (msoFalse)
shp3.Select (msoFalse)
shp4.Select (msoFalse)
ActiveWindow.Selection.ShapeRange.Distribute msoDistributeHorizontally, msoFalse
shp11.Select (msoTrue)
shp12.Select (msoFalse)
shp13.Select (msoFalse)
shp14.Select (msoFalse)
ActiveWindow.Selection.ShapeRange.Distribute msoDistributeHorizontally, msoFalse

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    shp1.TextFrame2.TextRange.Font.Size = 14
    shp2.TextFrame2.TextRange.Font.Size = 14
    shp3.TextFrame2.TextRange.Font.Size = 14
    shp4.TextFrame2.TextRange.Font.Size = 14
    shp11.TextFrame2.TextRange.Font.Size = 12
    shp12.TextFrame2.TextRange.Font.Size = 12
    shp13.TextFrame2.TextRange.Font.Size = 12
    shp14.TextFrame2.TextRange.Font.Size = 12
End If

shp1.Select (msoTrue)
shp2.Select (msoFalse)
shp3.Select (msoFalse)
shp4.Select (msoFalse)
shp11.Select (msoFalse)
shp12.Select (msoFalse)
shp13.Select (msoFalse)
shp14.Select (msoFalse)

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub Column10(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim phd As Shape
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim shp3 As Shape
    Dim shp4 As Shape
    Dim shp5 As Shape
    Dim shp6 As Shape
    Dim shp7 As Shape
    Dim shp8 As Shape
    Dim shp9 As Shape
    Dim shp10 As Shape
    Dim shp11 As Shape
    Dim shp12 As Shape
    Dim shp13 As Shape
    Dim shp14 As Shape
    Dim shp15 As Shape
    Dim shp16 As Shape
    Dim shp17 As Shape
    Dim shp18 As Shape
    Dim shp19 As Shape
    Dim shp20 As Shape
    Dim sld As Slide

    'Ten Columns

On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else

    Set sld = Application.ActiveWindow.View.Slide
    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)
    
If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=59.527522, Height:=34.015727)
    shp1.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=31.464547)
    shp1.Select msoTrue
Else
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=34.015727)
    shp1.Select msoTrue
End If

Call ParameterColumnHeader
    

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=59.527522, Height:=34.015727)
    shp2.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=31.464547)
    shp2.Select msoTrue
Else
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=34.015727)
    shp2.Select msoTrue
End If

Call ParameterColumnHeader


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=59.527522, Height:=34.015727)
    shp3.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=31.464547)
    shp3.Select msoTrue
Else
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=34.015727)
    shp3.Select msoTrue
End If

Call ParameterColumnHeader


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp4 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=59.527522, Height:=34.015727)
    shp4.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp4 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=31.464547)
    shp4.Select msoTrue
Else
    Set shp4 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=34.015727)
    shp4.Select msoTrue
End If

Call ParameterColumnHeader


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp5 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=59.527522, Height:=34.015727)
    shp5.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp5 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=31.464547)
    shp5.Select msoTrue
Else
    Set shp5 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=34.015727)
    shp5.Select msoTrue
End If

Call ParameterColumnHeader


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp6 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=59.527522, Height:=34.015727)
    shp6.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp6 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=31.464547)
    shp6.Select msoTrue
Else
    Set shp6 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=34.015727)
    shp6.Select msoTrue
End If

Call ParameterColumnHeader


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp7 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=59.527522, Height:=34.015727)
    shp7.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp7 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=31.464547)
    shp7.Select msoTrue
Else
    Set shp7 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=34.015727)
    shp7.Select msoTrue
End If

Call ParameterColumnHeader


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp8 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=59.527522, Height:=34.015727)
    shp8.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp8 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=31.464547)
    shp8.Select msoTrue
Else
    Set shp8 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=34.015727)
    shp8.Select msoTrue
End If

Call ParameterColumnHeader


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp9 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=59.527522, Height:=34.015727)
    shp9.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp9 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=31.464547)
    shp9.Select msoTrue
Else
    Set shp9 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=34.015727)
    shp9.Select msoTrue
End If

Call ParameterColumnHeader


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp10 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=59.527522, Height:=34.015727)
    shp10.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp10 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=31.464547)
    shp10.Select msoTrue
Else
    Set shp10 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=34.015727)
    shp10.Select msoTrue
End If

Call ParameterColumnHeader


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=59.527522, Height:=283.46439)
    shp11.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=252.85023)
    shp11.Select msoTrue
Else
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=283.46439)
    shp11.Select msoTrue
End If

Call ParameterColumnBodyStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=59.527522, Height:=283.46439)
    shp12.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=252.85023)
    shp12.Select msoTrue
Else
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=283.46439)
    shp12.Select msoTrue
End If

Call ParameterColumnBodyStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=59.527522, Height:=283.46439)
    shp13.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=252.85023)
    shp13.Select msoTrue
Else
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=283.46439)
    shp13.Select msoTrue
End If

Call ParameterColumnBodyStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp14 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=59.527522, Height:=283.46439)
    shp14.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp14 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=252.85023)
    shp14.Select msoTrue
Else
    Set shp14 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=283.46439)
    shp14.Select msoTrue
End If

Call ParameterColumnBodyStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp15 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=59.527522, Height:=283.46439)
    shp15.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp15 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=252.85023)
    shp15.Select msoTrue
Else
    Set shp15 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=283.46439)
    shp15.Select msoTrue
End If

Call ParameterColumnBodyStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp16 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=59.527522, Height:=283.46439)
    shp16.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp16 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=252.85023)
    shp16.Select msoTrue
Else
    Set shp16 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=283.46439)
    shp16.Select msoTrue
End If

Call ParameterColumnBodyStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp17 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=59.527522, Height:=283.46439)
    shp17.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp17 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=252.85023)
    shp17.Select msoTrue
Else
    Set shp17 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=283.46439)
    shp17.Select msoTrue
End If

Call ParameterColumnBodyStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp18 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=59.527522, Height:=283.46439)
    shp18.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp18 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=252.85023)
    shp18.Select msoTrue
Else
    Set shp18 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=283.46439)
    shp18.Select msoTrue
End If

Call ParameterColumnBodyStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp19 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=59.527522, Height:=283.46439)
    shp19.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp19 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=252.85023)
    shp19.Select msoTrue
Else
    Set shp19 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=283.46439)
    shp19.Select msoTrue
End If

Call ParameterColumnBodyStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp20 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=59.527522, Height:=283.46439)
    shp20.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp20 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=252.85023)
    shp20.Select msoTrue
Else
    Set shp20 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=56.409413, Height:=283.46439)
    shp20.Select msoTrue
End If

Call ParameterColumnBodyStandard


shp1.Top = phd.Top
shp2.Top = phd.Top
shp3.Top = phd.Top
shp4.Top = phd.Top
shp5.Top = phd.Top
shp6.Top = phd.Top
shp7.Top = phd.Top
shp8.Top = phd.Top
shp9.Top = phd.Top
shp10.Top = phd.Top
shp11.Top = shp1.Top + shp1.Height
shp12.Top = shp2.Top + shp2.Height
shp13.Top = shp3.Top + shp3.Height
shp14.Top = shp4.Top + shp4.Height
shp15.Top = shp5.Top + shp5.Height
shp16.Top = shp6.Top + shp6.Height
shp17.Top = shp7.Top + shp7.Height
shp18.Top = shp8.Top + shp8.Height
shp19.Top = shp9.Top + shp9.Height
shp20.Top = shp10.Top + shp10.Height

shp1.Left = phd.Left
shp2.Left = phd.Left + 10
shp3.Left = phd.Left + 10 + 10
shp4.Left = phd.Left + 10 + 10 + 10
shp5.Left = phd.Left + 10 + 10 + 10 + 10
shp6.Left = phd.Left + 10 + 10 + 10 + 10 + 10
shp7.Left = phd.Left + 10 + 10 + 10 + 10 + 10 + 10
shp8.Left = phd.Left + 10 + 10 + 10 + 10 + 10 + 10 + 10
shp9.Left = phd.Left + 10 + 10 + 10 + 10 + 10 + 10 + 10 + 10
shp10.Left = (phd.Left + phd.Width) - shp10.Width
shp11.Left = phd.Left
shp12.Left = phd.Left + 10
shp13.Left = phd.Left + 10 + 10
shp14.Left = phd.Left + 10 + 10 + 10
shp15.Left = phd.Left + 10 + 10 + 10 + 10
shp16.Left = phd.Left + 10 + 10 + 10 + 10 + 10
shp17.Left = phd.Left + 10 + 10 + 10 + 10 + 10 + 10
shp18.Left = phd.Left + 10 + 10 + 10 + 10 + 10 + 10 + 10
shp19.Left = phd.Left + 10 + 10 + 10 + 10 + 10 + 10 + 10 + 10
shp20.Left = (phd.Left + phd.Width) - shp20.Width

shp1.Select (msoTrue)
shp2.Select (msoFalse)
shp3.Select (msoFalse)
shp4.Select (msoFalse)
shp5.Select (msoFalse)
shp6.Select (msoFalse)
shp7.Select (msoFalse)
shp8.Select (msoFalse)
shp9.Select (msoFalse)
shp10.Select (msoFalse)
ActiveWindow.Selection.ShapeRange.Distribute msoDistributeHorizontally, msoFalse
shp11.Select (msoTrue)
shp12.Select (msoFalse)
shp13.Select (msoFalse)
shp14.Select (msoFalse)
shp15.Select (msoFalse)
shp16.Select (msoFalse)
shp17.Select (msoFalse)
shp18.Select (msoFalse)
shp19.Select (msoFalse)
shp20.Select (msoFalse)
ActiveWindow.Selection.ShapeRange.Distribute msoDistributeHorizontally, msoFalse

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    shp1.TextFrame2.TextRange.Font.Size = 12
    shp2.TextFrame2.TextRange.Font.Size = 12
    shp3.TextFrame2.TextRange.Font.Size = 12
    shp4.TextFrame2.TextRange.Font.Size = 12
    shp5.TextFrame2.TextRange.Font.Size = 12
    shp6.TextFrame2.TextRange.Font.Size = 12
    shp7.TextFrame2.TextRange.Font.Size = 12
    shp8.TextFrame2.TextRange.Font.Size = 12
    shp9.TextFrame2.TextRange.Font.Size = 12
    shp10.TextFrame2.TextRange.Font.Size = 12
    shp11.TextFrame2.TextRange.Font.Size = 10
    shp12.TextFrame2.TextRange.Font.Size = 10
    shp13.TextFrame2.TextRange.Font.Size = 10
    shp14.TextFrame2.TextRange.Font.Size = 10
    shp15.TextFrame2.TextRange.Font.Size = 10
    shp16.TextFrame2.TextRange.Font.Size = 10
    shp17.TextFrame2.TextRange.Font.Size = 10
    shp18.TextFrame2.TextRange.Font.Size = 10
    shp19.TextFrame2.TextRange.Font.Size = 10
    shp20.TextFrame2.TextRange.Font.Size = 10
Else
    shp1.TextFrame2.TextRange.Font.Size = 14
    shp2.TextFrame2.TextRange.Font.Size = 14
    shp3.TextFrame2.TextRange.Font.Size = 14
    shp4.TextFrame2.TextRange.Font.Size = 14
    shp5.TextFrame2.TextRange.Font.Size = 14
    shp6.TextFrame2.TextRange.Font.Size = 14
    shp7.TextFrame2.TextRange.Font.Size = 14
    shp8.TextFrame2.TextRange.Font.Size = 14
    shp9.TextFrame2.TextRange.Font.Size = 14
    shp10.TextFrame2.TextRange.Font.Size = 14
    shp11.TextFrame2.TextRange.Font.Size = 12
    shp12.TextFrame2.TextRange.Font.Size = 12
    shp13.TextFrame2.TextRange.Font.Size = 12
    shp14.TextFrame2.TextRange.Font.Size = 12
    shp15.TextFrame2.TextRange.Font.Size = 12
    shp16.TextFrame2.TextRange.Font.Size = 12
    shp17.TextFrame2.TextRange.Font.Size = 12
    shp18.TextFrame2.TextRange.Font.Size = 12
    shp19.TextFrame2.TextRange.Font.Size = 12
    shp20.TextFrame2.TextRange.Font.Size = 12
End If

shp1.TextFrame2.TextRange.Text = "Head"
shp2.TextFrame2.TextRange.Text = "Head"
shp3.TextFrame2.TextRange.Text = "Head"
shp4.TextFrame2.TextRange.Text = "Head"
shp5.TextFrame2.TextRange.Text = "Head"
shp6.TextFrame2.TextRange.Text = "Head"
shp7.TextFrame2.TextRange.Text = "Head"
shp8.TextFrame2.TextRange.Text = "Head"
shp9.TextFrame2.TextRange.Text = "Head"
shp10.TextFrame2.TextRange.Text = "Head"
shp11.TextFrame2.TextRange.Text = "Text"
shp12.TextFrame2.TextRange.Text = "Text"
shp13.TextFrame2.TextRange.Text = "Text"
shp14.TextFrame2.TextRange.Text = "Text"
shp15.TextFrame2.TextRange.Text = "Text"
shp16.TextFrame2.TextRange.Text = "Text"
shp17.TextFrame2.TextRange.Text = "Text"
shp18.TextFrame2.TextRange.Text = "Text"
shp19.TextFrame2.TextRange.Text = "Text"
shp20.TextFrame2.TextRange.Text = "Text"

shp1.Select (msoTrue)
shp2.Select (msoFalse)
shp3.Select (msoFalse)
shp4.Select (msoFalse)
shp5.Select (msoFalse)
shp6.Select (msoFalse)
shp7.Select (msoFalse)
shp8.Select (msoFalse)
shp9.Select (msoFalse)
shp10.Select (msoFalse)
shp11.Select (msoFalse)
shp12.Select (msoFalse)
shp13.Select (msoFalse)
shp14.Select (msoFalse)
shp15.Select (msoFalse)
shp16.Select (msoFalse)
shp17.Select (msoFalse)
shp18.Select (msoFalse)
shp19.Select (msoFalse)
shp20.Select (msoFalse)

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub FlowTriangle01(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim phd As Shape
    Dim shp As Shape
    Dim sld As Slide
    'Flow Triangle

On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else
    
    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)
    Set sld = Application.ActiveWindow.View.Slide
    
If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp = sld.Shapes.AddShape(Type:=msoShapeIsoscelesTriangle, Left:=0, Top:=0, Width:=269.29117, Height:=21.259829)
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp = sld.Shapes.AddShape(Type:=msoShapeIsoscelesTriangle, Left:=0, Top:=0, Width:=240.66127, Height:=21.259829)
Else
    Set shp = sld.Shapes.AddShape(Type:=msoShapeIsoscelesTriangle, Left:=0, Top:=0, Width:=269.29117, Height:=21.259829)
End If
    
    shp.Fill.Visible = msoTrue
    shp.Fill.Transparency = 0
    shp.Fill.ForeColor.RGB = RGB(186, 186, 186)
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(186, 186, 186)
    shp.Line.Weight = 0.75
    shp.Rotation = 90
    
If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    shp.Top = phd.Top + 147.40148
Else
    shp.Top = phd.Top + 165.25974
End If
    
    shp.Left = (phd.Width / 2) + phd.Left - (shp.Width / 2)
    shp.Select (msoTrue)

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub FlowTriangle02(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim phd As Shape
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim shp21 As Shape
    Dim shp22 As Shape
    Dim shp23 As Shape
    Dim sld As Slide
    'Two Flow Triangles

On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else

    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)
    Set sld = Application.ActiveWindow.View.Slide
    
If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeIsoscelesTriangle, Left:=0, Top:=0, Width:=269.29117, Height:=21.259829)
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeIsoscelesTriangle, Left:=0, Top:=0, Width:=240.66127, Height:=21.259829)
Else
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeIsoscelesTriangle, Left:=0, Top:=0, Width:=269.29117, Height:=21.259829)
End If
    
If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeIsoscelesTriangle, Left:=0, Top:=0, Width:=269.29117, Height:=21.259829)
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeIsoscelesTriangle, Left:=0, Top:=0, Width:=240.66127, Height:=21.259829)
Else
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeIsoscelesTriangle, Left:=0, Top:=0, Width:=269.29117, Height:=21.259829)
End If
    
    shp1.Fill.Visible = msoTrue
    shp1.Fill.Transparency = 0
    shp1.Fill.ForeColor.RGB = RGB(186, 186, 186)
    shp1.Line.Visible = msoTrue
    shp1.Line.ForeColor.RGB = RGB(186, 186, 186)
    shp1.Line.Weight = 0.75
    shp1.Rotation = 90


    shp2.Fill.Visible = msoTrue
    shp2.Fill.Transparency = 0
    shp2.Fill.ForeColor.RGB = RGB(186, 186, 186)
    shp2.Line.Visible = msoTrue
    shp2.Line.ForeColor.RGB = RGB(186, 186, 186)
    shp2.Line.Weight = 0.75
    shp2.Rotation = 90

    
If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp21 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=211.18097, Height:=10)
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp21 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=10)
Else
    Set shp21 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=10)
End If

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp22 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=211.18097, Height:=10)
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp22 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=10)
Else
    Set shp22 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=10)
End If

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp23 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=211.18097, Height:=10)
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp23 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=10)
Else
    Set shp23 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=199.55893, Height:=10)
End If

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    shp1.Top = phd.Top + 147.40148
    shp2.Top = phd.Top + 147.40148
Else
    shp1.Top = phd.Top + 165.25974
    shp2.Top = phd.Top + 165.25974
End If

shp21.Left = phd.Left
shp1.Left = shp21.Left + 10
shp22.Left = phd.Left + (phd.Width / 2)
shp2.Left = shp22.Left + 10
shp23.Left = (phd.Left + phd.Width) - shp23.Width

shp1.Select (msoTrue)
shp2.Select (msoFalse)
shp21.Select (msoFalse)
shp22.Select (msoFalse)
shp23.Select (msoFalse)
ActiveWindow.Selection.ShapeRange.Distribute msoDistributeHorizontally, msoFalse

shp21.Delete
shp22.Delete
shp23.Delete

shp1.Select (msoTrue)
shp2.Select (msoFalse)

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub Footnote(control As IRibbonControl)

    Dim shp As Shape
    Dim sld As Slide
    'Footnote für Ommax

On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else
    Set sld = Application.ActiveWindow.View.Slide
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, Left:=29.763761, Top:=487.55875, Width:=559.27524, Height:=15.874006)

With shp
    .Fill.Visible = msoFalse
    .Line.Visible = msoFalse
    
With .TextFrame
    .AutoSize = ppAutoSizeNone
    .TextRange.Text = "Source: "
    .VerticalAnchor = msoAnchorTop
    .MarginBottom = 3.685037
    .MarginLeft = 0
    .MarginRight = 0
    .MarginTop = 3.685037
    .WordWrap = msoTrue
    
With .TextRange
    .Font.Size = 7
    .Font.Name = "Gotham Light"
    .Font.Color.RGB = RGB(18, 126, 129)
    .ParagraphFormat.Alignment = ppAlignLeft
    .Select

End With
End With
End With

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub NumberBalls(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim shp3 As Shape
    Dim shp4 As Shape
    Dim shp5 As Shape
    Dim shp6 As Shape
    Dim shp7 As Shape
    Dim shp8 As Shape
    Dim shp9 As Shape
    Dim shp10 As Shape
    Dim sld As Slide
    'Number balls edited for Ommax
    
On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else
    
    Set sld = Application.ActiveWindow.View.Slide
    
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=0, Top:=200, Width:=14.173219, Height:=14.173219)
    shp1.Select msoTrue
    
Call ParameterBoxHeaderFlat
Call ParameterNumberBall


    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=0, Top:=210, Width:=14.173219, Height:=14.173219)
    shp2.Select msoTrue
    
Call ParameterBoxHeaderFlat
Call ParameterNumberBall


    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=0, Top:=220, Width:=14.173219, Height:=14.173219)
    shp3.Select msoTrue
    
Call ParameterBoxHeaderFlat
Call ParameterNumberBall


    Set shp4 = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=0, Top:=230, Width:=14.173219, Height:=14.173219)
    shp4.Select msoTrue
    
Call ParameterBoxHeaderFlat
Call ParameterNumberBall


    Set shp5 = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=0, Top:=240, Width:=14.173219, Height:=14.173219)
    shp5.Select msoTrue
    
Call ParameterBoxHeaderFlat
Call ParameterNumberBall


    Set shp6 = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=0, Top:=250, Width:=14.173219, Height:=14.173219)
    shp6.Select msoTrue
    
Call ParameterBoxHeaderFlat
Call ParameterNumberBall


    Set shp7 = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=0, Top:=260, Width:=14.173219, Height:=14.173219)
    shp7.Select msoTrue
    
Call ParameterBoxHeaderFlat
Call ParameterNumberBall


    Set shp8 = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=0, Top:=270, Width:=14.173219, Height:=14.173219)
    shp8.Select msoTrue
    
Call ParameterBoxHeaderFlat
Call ParameterNumberBall


    Set shp9 = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=0, Top:=280, Width:=14.173219, Height:=14.173219)
    shp9.Select msoTrue
    
Call ParameterBoxHeaderFlat
Call ParameterNumberBall

    Set shp10 = sld.Shapes.AddShape(Type:=msoShapeOval, Left:=0, Top:=290, Width:=14.173219, Height:=14.173219)
    shp10.Select msoTrue
    
Call ParameterBoxHeaderFlat
Call ParameterNumberBall


shp1.TextFrame2.TextRange.Text = "1"
shp2.TextFrame2.TextRange.Text = "2"
shp3.TextFrame2.TextRange.Text = "3"
shp4.TextFrame2.TextRange.Text = "4"
shp5.TextFrame2.TextRange.Text = "5"
shp6.TextFrame2.TextRange.Text = "6"
shp7.TextFrame2.TextRange.Text = "7"
shp8.TextFrame2.TextRange.Text = "8"
shp9.TextFrame2.TextRange.Text = "9"
shp10.TextFrame2.TextRange.Text = "10"

shp1.Top = 117.92119
shp10.Top = 375.59028

shp1.Select (msoFalse)
shp2.Select (msoFalse)
shp3.Select (msoFalse)
shp4.Select (msoFalse)
shp5.Select (msoFalse)
shp6.Select (msoFalse)
shp7.Select (msoFalse)
shp8.Select (msoFalse)
shp9.Select (msoFalse)
shp10.Select (msoFalse)

ActiveWindow.Selection.ShapeRange.Distribute msoDistributeVertically, msoFalse

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub NumberBigSolo(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim phd As Shape
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim shp3 As Shape
    Dim shp4 As Shape
    Dim shp5 As Shape
    Dim shp6 As Shape
    Dim shp7 As Shape
    Dim shp8 As Shape
    Dim shp9 As Shape
    Dim shp10 As Shape
    Dim sld As Slide
    'Big numbers
    
On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else
    
    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)
    Set sld = Application.ActiveWindow.View.Slide
    
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=200, Width:=34.015727, Height:=34.015727)
    shp1.Select msoTrue
    
Call ParameterColumnBodyStandard
Call ParameterNumberBigSolo


    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=210, Width:=34.015727, Height:=34.015727)
    shp2.Select msoTrue
    
Call ParameterColumnBodyStandard
Call ParameterNumberBigSolo


    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=220, Width:=34.015727, Height:=34.015727)
    shp3.Select msoTrue
    
Call ParameterColumnBodyStandard
Call ParameterNumberBigSolo


    Set shp4 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=230, Width:=34.015727, Height:=34.015727)
    shp4.Select msoTrue
    
Call ParameterColumnBodyStandard
Call ParameterNumberBigSolo


    Set shp5 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=240, Width:=34.015727, Height:=34.015727)
    shp5.Select msoTrue
    
Call ParameterColumnBodyStandard
Call ParameterNumberBigSolo


    Set shp6 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=250, Width:=34.015727, Height:=34.015727)
    shp6.Select msoTrue
    
Call ParameterColumnBodyStandard
Call ParameterNumberBigSolo


    Set shp7 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=260, Width:=34.015727, Height:=34.015727)
    shp7.Select msoTrue
    
Call ParameterColumnBodyStandard
Call ParameterNumberBigSolo


    Set shp8 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=270, Width:=34.015727, Height:=34.015727)
    shp8.Select msoTrue
    
Call ParameterColumnBodyStandard
Call ParameterNumberBigSolo


    Set shp9 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=280, Width:=34.015727, Height:=34.015727)
    shp9.Select msoTrue
    
Call ParameterColumnBodyStandard
Call ParameterNumberBigSolo

    Set shp10 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=290, Width:=34.015727, Height:=34.015727)
    shp10.Select msoTrue
    
Call ParameterColumnBodyStandard
Call ParameterNumberBigSolo


shp1.TextFrame2.TextRange.Text = "1"
shp2.TextFrame2.TextRange.Text = "2"
shp3.TextFrame2.TextRange.Text = "3"
shp4.TextFrame2.TextRange.Text = "4"
shp5.TextFrame2.TextRange.Text = "5"
shp6.TextFrame2.TextRange.Text = "6"
shp7.TextFrame2.TextRange.Text = "7"
shp8.TextFrame2.TextRange.Text = "8"
shp9.TextFrame2.TextRange.Text = "9"
shp10.TextFrame2.TextRange.Text = "10"

shp1.Left = phd.Left
shp2.Left = phd.Left
shp3.Left = phd.Left
shp4.Left = phd.Left
shp5.Left = phd.Left
shp6.Left = phd.Left
shp7.Left = phd.Left
shp8.Left = phd.Left
shp9.Left = phd.Left
shp10.Left = phd.Left

shp1.Top = phd.Top
shp10.Top = (phd.Top + phd.Height) - shp10.Height

shp1.Select (msoFalse)
shp2.Select (msoFalse)
shp3.Select (msoFalse)
shp4.Select (msoFalse)
shp5.Select (msoFalse)
shp6.Select (msoFalse)
shp7.Select (msoFalse)
shp8.Select (msoFalse)
shp9.Select (msoFalse)
shp10.Select (msoFalse)

ActiveWindow.Selection.ShapeRange.Distribute msoDistributeVertically, msoFalse

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub NumberSquares(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim shp3 As Shape
    Dim shp4 As Shape
    Dim shp5 As Shape
    Dim shp6 As Shape
    Dim shp7 As Shape
    Dim shp8 As Shape
    Dim shp9 As Shape
    Dim shp10 As Shape
    Dim sld As Slide
    'Number squares edited for Ommax
    
On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else
    
    Set sld = Application.ActiveWindow.View.Slide
    
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=200, Width:=25.511795, Height:=14.173219)
    shp1.Select msoTrue
    
Call ParameterBoxHeaderFlat
Call ParameterNumberBall


    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=210, Width:=25.511795, Height:=14.173219)
    shp2.Select msoTrue
    
Call ParameterBoxHeaderFlat
Call ParameterNumberBall


    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=220, Width:=25.511795, Height:=14.173219)
    shp3.Select msoTrue
    
Call ParameterBoxHeaderFlat
Call ParameterNumberBall


    Set shp4 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=230, Width:=25.511795, Height:=14.173219)
    shp4.Select msoTrue
    
Call ParameterBoxHeaderFlat
Call ParameterNumberBall


    Set shp5 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=240, Width:=25.511795, Height:=14.173219)
    shp5.Select msoTrue
    
Call ParameterBoxHeaderFlat
Call ParameterNumberBall


    Set shp6 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=250, Width:=25.511795, Height:=14.173219)
    shp6.Select msoTrue
    
Call ParameterBoxHeaderFlat
Call ParameterNumberBall


    Set shp7 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=260, Width:=25.511795, Height:=14.173219)
    shp7.Select msoTrue
    
Call ParameterBoxHeaderFlat
Call ParameterNumberBall


    Set shp8 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=270, Width:=25.511795, Height:=14.173219)
    shp8.Select msoTrue
    
Call ParameterBoxHeaderFlat
Call ParameterNumberBall


    Set shp9 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=280, Width:=25.511795, Height:=14.173219)
    shp9.Select msoTrue
    
Call ParameterBoxHeaderFlat
Call ParameterNumberBall

    Set shp10 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=290, Width:=25.511795, Height:=14.173219)
    shp10.Select msoTrue
    
Call ParameterBoxHeaderFlat
Call ParameterNumberBall


shp1.TextFrame2.TextRange.Text = "1"
shp2.TextFrame2.TextRange.Text = "2"
shp3.TextFrame2.TextRange.Text = "3"
shp4.TextFrame2.TextRange.Text = "4"
shp5.TextFrame2.TextRange.Text = "5"
shp6.TextFrame2.TextRange.Text = "6"
shp7.TextFrame2.TextRange.Text = "7"
shp8.TextFrame2.TextRange.Text = "8"
shp9.TextFrame2.TextRange.Text = "9"
shp10.TextFrame2.TextRange.Text = "10"

shp1.Top = 117.92119
shp10.Top = 375.59028

shp1.Select (msoFalse)
shp2.Select (msoFalse)
shp3.Select (msoFalse)
shp4.Select (msoFalse)
shp5.Select (msoFalse)
shp6.Select (msoFalse)
shp7.Select (msoFalse)
shp8.Select (msoFalse)
shp9.Select (msoFalse)
shp10.Select (msoFalse)

ActiveWindow.Selection.ShapeRange.Distribute msoDistributeVertically, msoFalse

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub OrgChart(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim phd As Shape
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim shp3 As Shape
    Dim shp4 As Shape
    Dim shp5 As Shape
    Dim shp6 As Shape
    Dim shp7 As Shape
    Dim sld As Slide
    'Org Chart Starter
    
On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else
    
    Set sld = Application.ActiveWindow.View.Slide
    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)
    
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=318.04704, Top:=118.48811, Width:=144.28337, Height:=49.606268)
    shp1.Select msoTrue
    
Call ParameterShapeStandard


    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=138.33062, Top:=245.76362, Width:=22.960615, Height:=22.960615)
    shp2.Fill.ForeColor.RGB = RGB(89, 171, 244)
    shp2.Line.ForeColor.RGB = RGB(255, 255, 255)
    shp2.Line.Weight = 0.75
    
    
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=497.48, Top:=245.76362, Width:=22.960615, Height:=22.960615)
    shp3.Fill.ForeColor.RGB = RGB(89, 171, 244)
    shp3.Line.ForeColor.RGB = RGB(255, 255, 255)
    shp3.Line.Weight = 0.75
    
    
    Set shp4 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=138.33062, Top:=219.11797, Width:=144.28337, Height:=49.606268)
    shp4.Select msoTrue
    
Call ParameterShapeStandard


    Set shp5 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=497.48, Top:=219.11797, Width:=144.28337, Height:=49.606268)
    shp5.Select msoTrue
    
Call ParameterShapeStandard


    Set shp6 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=161.29124, Top:=282.33053, Width:=121.32276, Height:=49.606268)
    shp6.Select msoTrue
    
Call ParameterShapeStandard


    Set shp7 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=161.29124, Top:=348.37773, Width:=121.32276, Height:=49.606268)
    shp7.Select msoTrue
    
Call ParameterShapeStandard


shp1.TextFrame2.TextRange.Font.Size = 12
shp4.TextFrame2.TextRange.Font.Size = 12
shp5.TextFrame2.TextRange.Font.Size = 12
shp6.TextFrame2.TextRange.Font.Size = 12
shp7.TextFrame2.TextRange.Font.Size = 12

shp1.TextFrame2.TextRange.Font.Bold = msoTrue
shp4.TextFrame2.TextRange.Font.Bold = msoTrue
shp5.TextFrame2.TextRange.Font.Bold = msoTrue

With sld.Shapes.AddConnector(Type:=msoConnectorElbow, BeginX:=0, BeginY:=0, EndX:=100, EndY:=100)
 .Line.ForeColor.RGB = RGB(150, 150, 150)
 .ConnectorFormat.BeginConnect ConnectedShape:=shp1, ConnectionSite:=3
 .ConnectorFormat.EndConnect ConnectedShape:=shp4, ConnectionSite:=1
 .Select (msoTrue)
End With

With sld.Shapes.AddConnector(Type:=msoConnectorElbow, BeginX:=0, BeginY:=0, EndX:=100, EndY:=100)
 .Line.ForeColor.RGB = RGB(150, 150, 150)
 .ConnectorFormat.BeginConnect ConnectedShape:=shp1, ConnectionSite:=3
 .ConnectorFormat.EndConnect ConnectedShape:=shp5, ConnectionSite:=1
 .Select (msoFalse)
End With

With sld.Shapes.AddConnector(Type:=msoConnectorElbow, BeginX:=0, BeginY:=0, EndX:=100, EndY:=100)
 .Line.ForeColor.RGB = RGB(150, 150, 150)
 .ConnectorFormat.BeginConnect ConnectedShape:=shp2, ConnectionSite:=3
 .ConnectorFormat.EndConnect ConnectedShape:=shp6, ConnectionSite:=2
 .Select (msoFalse)
End With

With sld.Shapes.AddConnector(Type:=msoConnectorElbow, BeginX:=0, BeginY:=0, EndX:=100, EndY:=100)
 .Line.ForeColor.RGB = RGB(150, 150, 150)
 .ConnectorFormat.BeginConnect ConnectedShape:=shp2, ConnectionSite:=3
 .ConnectorFormat.EndConnect ConnectedShape:=shp7, ConnectionSite:=2
 .Select (msoFalse)
End With

shp1.Select (msoFalse)
shp2.Select (msoFalse)
shp3.Select (msoFalse)
shp4.Select (msoFalse)
shp5.Select (msoFalse)
shp6.Select (msoFalse)
shp7.Select (msoFalse)

ActiveWindow.Selection.ShapeRange.Group.Select
ActiveWindow.Selection.ShapeRange(1).Left = (phd.Left + (phd.Width / 2)) - (ActiveWindow.Selection.ShapeRange(1).Width / 2)
ActiveWindow.Selection.ShapeRange(1).Top = phd.Top
ActiveWindow.Selection.ShapeRange(1).Ungroup.Select

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub Pyramids03(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim phd As Shape
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim shp3 As Shape
    Dim sld As Slide

    'Pyramid with three elements
    
On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else
    
    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)
    Set sld = Application.ActiveWindow.View.Slide

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeIsoscelesTriangle, Left:=0, Top:=0, Width:=149.95266, Height:=102.04718)
    shp1.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeIsoscelesTriangle, Left:=0, Top:=0, Width:=138.89755, Height:=94.393641)
    shp1.Select msoTrue
Else
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeIsoscelesTriangle, Left:=0, Top:=0, Width:=149.95266, Height:=102.04718)
    shp1.Select msoTrue
End If

Call ParameterShapeStandard
    

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeTrapezoid, Left:=0, Top:=0, Width:=297.07058, Height:=102.04718)
    shp2.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeTrapezoid, Left:=0, Top:=0, Width:=277.7951, Height:=94.393641)
    shp2.Select msoTrue
Else
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeTrapezoid, Left:=0, Top:=0, Width:=297.07058, Height:=102.04718)
    shp2.Select msoTrue
End If

Call ParameterShapeStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeTrapezoid, Left:=0, Top:=0, Width:=443.90523, Height:=102.04718)
    shp3.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeTrapezoid, Left:=0, Top:=0, Width:=411.02336, Height:=94.393641)
    shp3.Select msoTrue
Else
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeTrapezoid, Left:=0, Top:=0, Width:=443.90523, Height:=102.04718)
    shp3.Select msoTrue
End If

Call ParameterShapeStandard


shp1.Left = (phd.Width / 2) + phd.Left - (shp1.Width / 2)
shp2.Left = (phd.Width / 2) + phd.Left - (shp2.Width / 2)
shp3.Left = (phd.Width / 2) + phd.Left - (shp3.Width / 2)

shp1.Top = phd.Top
shp2.Top = shp1.Top + shp1.Height
shp3.Top = shp2.Top + shp2.Height

shp1.Line.ForeColor.RGB = RGB(255, 255, 255)
shp2.Line.ForeColor.RGB = RGB(255, 255, 255)
shp3.Line.ForeColor.RGB = RGB(255, 255, 255)

shp2.Adjustments(1) = 0.725
shp3.Adjustments(1) = 0.725

shp1.Select (msoTrue)
shp2.Select (msoFalse)
shp3.Select (msoFalse)

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub Pyramids04(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim phd As Shape
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim shp3 As Shape
    Dim shp4 As Shape
    Dim sld As Slide

    'Pyramid with four elements

On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else

    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)
    Set sld = Application.ActiveWindow.View.Slide

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeIsoscelesTriangle, Left:=0, Top:=0, Width:=104.88182, Height:=71.999955)
    shp1.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeIsoscelesTriangle, Left:=0, Top:=0, Width:=97.228285, Height:=66.614131)
    shp1.Select msoTrue
Else
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeIsoscelesTriangle, Left:=0, Top:=0, Width:=104.88182, Height:=71.999955)
    shp1.Select msoTrue
End If

Call ParameterShapeStandard
    

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeTrapezoid, Left:=0, Top:=0, Width:=217.41719, Height:=77.952707)
    shp2.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeTrapezoid, Left:=0, Top:=0, Width:=201.25972, Height:=72.283419)
    shp2.Select msoTrue
Else
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeTrapezoid, Left:=0, Top:=0, Width:=217.41719, Height:=77.952707)
    shp2.Select msoTrue
End If

Call ParameterShapeStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeTrapezoid, Left:=0, Top:=0, Width:=330.23601, Height:=77.952707)
    shp3.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeTrapezoid, Left:=0, Top:=0, Width:=307.55886, Height:=72.283419)
    shp3.Select msoTrue
Else
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeTrapezoid, Left:=0, Top:=0, Width:=330.23601, Height:=77.952707)
    shp3.Select msoTrue
End If

Call ParameterShapeStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp4 = sld.Shapes.AddShape(Type:=msoShapeTrapezoid, Left:=0, Top:=0, Width:=443.90523, Height:=77.952707)
    shp4.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp4 = sld.Shapes.AddShape(Type:=msoShapeTrapezoid, Left:=0, Top:=0, Width:=411.02336, Height:=72.283419)
    shp4.Select msoTrue
Else
    Set shp4 = sld.Shapes.AddShape(Type:=msoShapeTrapezoid, Left:=0, Top:=0, Width:=443.90523, Height:=77.952707)
    shp4.Select msoTrue
End If

Call ParameterShapeStandard


shp1.Left = (phd.Width / 2) + phd.Left - (shp1.Width / 2)
shp2.Left = (phd.Width / 2) + phd.Left - (shp2.Width / 2)
shp3.Left = (phd.Width / 2) + phd.Left - (shp3.Width / 2)
shp4.Left = (phd.Width / 2) + phd.Left - (shp4.Width / 2)

shp1.Top = phd.Top
shp2.Top = shp1.Top + shp1.Height
shp3.Top = shp2.Top + shp2.Height
shp4.Top = shp3.Top + shp3.Height

shp1.Line.ForeColor.RGB = RGB(255, 255, 255)
shp2.Line.ForeColor.RGB = RGB(255, 255, 255)
shp3.Line.ForeColor.RGB = RGB(255, 255, 255)
shp4.Line.ForeColor.RGB = RGB(255, 255, 255)

shp2.Adjustments(1) = 0.725
shp3.Adjustments(1) = 0.725
shp4.Adjustments(1) = 0.725

shp1.Select (msoTrue)
shp2.Select (msoFalse)
shp3.Select (msoFalse)
shp4.Select (msoFalse)

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub Pyramids05(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim phd As Shape
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim shp3 As Shape
    Dim shp4 As Shape
    Dim shp5 As Shape
    Dim sld As Slide

    'Pyramid with five elements

On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else

    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)
    Set sld = Application.ActiveWindow.View.Slide

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeIsoscelesTriangle, Left:=0, Top:=0, Width:=96.094428, Height:=66.047202)
    shp1.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeIsoscelesTriangle, Left:=0, Top:=0, Width:=89.007818, Height:=61.228308)
    shp1.Select msoTrue
Else
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapeIsoscelesTriangle, Left:=0, Top:=0, Width:=96.094428, Height:=66.047202)
    shp1.Select msoTrue
End If

Call ParameterShapeStandard
    

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeTrapezoid, Left:=0, Top:=0, Width:=183.11799, Height:=60.09445)
    shp2.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeTrapezoid, Left:=0, Top:=0, Width:=169.5117, Height:=55.55902)
    shp2.Select msoTrue
Else
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeTrapezoid, Left:=0, Top:=0, Width:=183.11799, Height:=60.09445)
    shp2.Select msoTrue
End If

Call ParameterShapeStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeTrapezoid, Left:=0, Top:=0, Width:=269.57463, Height:=60.09445)
    shp3.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeTrapezoid, Left:=0, Top:=0, Width:=250.86598, Height:=55.55902)
    shp3.Select msoTrue
Else
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeTrapezoid, Left:=0, Top:=0, Width:=270.70849, Height:=60.09445)
    shp3.Select msoTrue
End If

Call ParameterShapeStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp4 = sld.Shapes.AddShape(Type:=msoShapeTrapezoid, Left:=0, Top:=0, Width:=356.88166, Height:=60.09445)
    shp4.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp4 = sld.Shapes.AddShape(Type:=msoShapeTrapezoid, Left:=0, Top:=0, Width:=330.80294, Height:=55.55902)
    shp4.Select msoTrue
Else
    Set shp4 = sld.Shapes.AddShape(Type:=msoShapeTrapezoid, Left:=0, Top:=0, Width:=356.88166, Height:=60.09445)
    shp4.Select msoTrue
End If

Call ParameterShapeStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp5 = sld.Shapes.AddShape(Type:=msoShapeTrapezoid, Left:=0, Top:=0, Width:=443.90523, Height:=60.09445)
    shp5.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp5 = sld.Shapes.AddShape(Type:=msoShapeTrapezoid, Left:=0, Top:=0, Width:=411.30683, Height:=55.55902)
    shp5.Select msoTrue
Else
    Set shp5 = sld.Shapes.AddShape(Type:=msoShapeTrapezoid, Left:=0, Top:=0, Width:=443.90523, Height:=60.09445)
    shp5.Select msoTrue
End If

Call ParameterShapeStandard


shp1.Left = (phd.Width / 2) + phd.Left - (shp1.Width / 2)
shp2.Left = (phd.Width / 2) + phd.Left - (shp2.Width / 2)
shp3.Left = (phd.Width / 2) + phd.Left - (shp3.Width / 2)
shp4.Left = (phd.Width / 2) + phd.Left - (shp4.Width / 2)
shp5.Left = (phd.Width / 2) + phd.Left - (shp5.Width / 2)

shp1.Top = phd.Top
shp2.Top = shp1.Top + shp1.Height
shp3.Top = shp2.Top + shp2.Height
shp4.Top = shp3.Top + shp3.Height
shp5.Top = shp4.Top + shp4.Height

shp1.Line.ForeColor.RGB = RGB(255, 255, 255)
shp2.Line.ForeColor.RGB = RGB(255, 255, 255)
shp3.Line.ForeColor.RGB = RGB(255, 255, 255)
shp4.Line.ForeColor.RGB = RGB(255, 255, 255)
shp5.Line.ForeColor.RGB = RGB(255, 255, 255)

shp2.Adjustments(1) = 0.725
shp3.Adjustments(1) = 0.725
shp4.Adjustments(1) = 0.725
shp5.Adjustments(1) = 0.725

shp1.Select (msoTrue)
shp2.Select (msoFalse)
shp3.Select (msoFalse)
shp4.Select (msoFalse)
shp5.Select (msoFalse)

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub StampBackup(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

    Dim shp As Shape
    Dim sld As Slide
    'Backup stamp edited for Ommax
    
On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else
    
    Set sld = Application.ActiveWindow.View.Slide
    
    Set shp = sld.Shapes.AddShape(msoShapeRectangle, Left:=335.62184, Top:=0, Width:=49.039339, Height:=15.590541)

    shp.TextFrame2.TextRange.Text = "Backup"
    shp.TextFrame2.TextRange.Font.Name = "Gotham Light"
    shp.TextFrame2.TextRange.Font.Size = 10
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    shp.TextFrame2.TextRange.Font.Bold = msoFalse
    shp.TextFrame2.TextRange.Font.Italic = msoFalse
    shp.TextFrame2.TextRange.Font.UnderlineStyle = msoNoUnderline
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.MarginBottom = 5.6692878
    shp.TextFrame2.MarginLeft = 5.6692878
    shp.TextFrame2.MarginRight = 5.6692878
    shp.TextFrame2.MarginTop = 5.6692878
    shp.TextFrame2.AutoSize = msoAutoSizeNone
    shp.TextFrame2.WordWrap = msoFalse
    

    shp.Fill.Visible = msoFalse
    shp.Fill.Transparency = 0
    shp.Line.Visible = msoFalse
    shp.Select (msoTrue)

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub StampChapter(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim tit As Shape
    Dim shp As Shape
    Dim sld As Slide
    'Chapter stamp

On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else

    Set sld = Application.ActiveWindow.View.Slide
    Set tit = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(1)
    
If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=43.93698, Height:=19.559043)
Else
    Set shp = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=53.007841, Height:=22.110222)
End If

    shp.TextFrame2.TextRange.Text = "Chapter 1"
    shp.TextFrame2.TextRange.Font.Name = "Arial"
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(225, 144, 12)
    shp.TextFrame2.TextRange.Font.Bold = msoFalse
    shp.TextFrame2.TextRange.Font.Italic = msoFalse
    shp.TextFrame2.TextRange.Font.UnderlineStyle = msoNoUnderline
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.MarginBottom = 3.685037
    shp.TextFrame2.MarginLeft = 0
    shp.TextFrame2.MarginRight = 0
    shp.TextFrame2.MarginTop = 3.685037
    shp.TextFrame2.WordWrap = msoFalse
    
If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    shp.TextFrame2.TextRange.Font.Size = 10
Else
    shp.TextFrame2.TextRange.Font.Size = 12
End If

    shp.Fill.Visible = msoFalse
    shp.Fill.Transparency = 0
    shp.Line.Visible = msoFalse
    shp.Left = (tit.Left) + (tit.Width / 2) - (shp.Width / 2)
    shp.Top = (tit.Top / 2) - (shp.Height / 2)
    shp.Select (msoTrue)

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub StampDraft(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ

    Dim shp As Shape
    Dim sld As Slide
    'clickable draft stamp for OMMAX
    
On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else
    Set sld = Application.ActiveWindow.View.Slide
    Set shp = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=594.14136, Top:=20.125972, Width:=94.677106, Height:=27.77951)
    shp.Line.Visible = msoTrue
    shp.Fill.ForeColor.RGB = RGB(16, 202, 207)
    shp.Line.ForeColor.RGB = RGB(16, 202, 207)
    shp.Line.Weight = 0.75
    shp.Rotation = 5
    
    shp.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    shp.TextFrame.TextRange.Characters.Text = "DRAFT"
    shp.TextFrame.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
    shp.TextFrame.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame.TextRange.Font.Size = 14
    shp.TextFrame.TextRange.Font.Name = "Gotham Medium"
    shp.TextFrame.TextRange.Font.Bold = msoFalse
    shp.TextFrame.TextRange.Font.Italic = msoFalse
    shp.TextFrame.TextRange.Font.Underline = msoFalse
    shp.TextFrame.Orientation = msoTextOrientationHorizontal
    shp.TextFrame.MarginBottom = 5.6692878
    shp.TextFrame.MarginLeft = 5.6692878
    shp.TextFrame.MarginRight = 5.6692878
    shp.TextFrame.MarginTop = 5.6692878
    shp.TextFrame.WordWrap = msoTrue
    shp.Select (msoTrue)
    End If

Exit Sub

ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub StampIllu(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim phd As Shape
    Dim shp As Shape
    Dim sld As Slide
    'Illustrative Stamp
    
On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else
    
    Set sld = Application.ActiveWindow.View.Slide
    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)
    Set shp = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=132.66133, Height:=27.496046)
    shp.Fill.Visible = msoTrue
    shp.Fill.ForeColor.RGB = RGB(112, 48, 160)
    shp.Fill.Transparency = 0
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(112, 48, 160)
    shp.Line.Weight = 0.75
    shp.Rotation = 10
    
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Characters.Text = "ILLUSTRATIVE"
    shp.TextFrame2.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.Font.Size = 16
    shp.TextFrame2.TextRange.Font.Name = "Arial"
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.Font.Italic = msoFalse
    shp.TextFrame2.TextRange.Font.UnderlineStyle = msoNoUnderline
    shp.TextFrame2.Orientation = msoTextOrientationHorizontal
    shp.TextFrame2.MarginBottom = 7.0866097
    shp.TextFrame2.MarginLeft = 7.0866097
    shp.TextFrame2.MarginRight = 7.0866097
    shp.TextFrame2.MarginTop = 7.0866097
    shp.TextFrame2.WordWrap = msoTrue

    shp.Left = (phd.Left + phd.Width) - shp.Width
    shp.Top = phd.Top
    shp.Select (msoTrue)

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub StampTracker(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim phd As Shape
    Dim tit As Shape
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim shp3 As Shape
    Dim sld As Slide
    'Value Chain Tracker
    
On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else
    
    Set sld = Application.ActiveWindow.View.Slide
    Set tit = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(1)
    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)
    
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapePentagon, Left:=354.33049, Top:=9.9212536, Width:=24.377937, Height:=13.606291)
    shp1.Fill.Visible = msoTrue
    shp1.Fill.ForeColor.RGB = RGB(89, 171, 244)
    shp1.Fill.Transparency = 0
    shp1.Line.Visible = msoTrue
    shp1.Line.ForeColor.RGB = RGB(255, 255, 255)
    shp1.Line.Weight = 0.75
    shp1.Adjustments(1) = 0.25
    shp1.Select (msoTrue)

    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeChevron, Left:=377.85803, Top:=9.9212536, Width:=24.377937, Height:=13.606291)
    shp2.Fill.Visible = msoTrue
    shp2.Fill.ForeColor.RGB = RGB(9, 91, 164)
    shp2.Fill.Transparency = 0
    shp2.Line.Visible = msoTrue
    shp2.Line.ForeColor.RGB = RGB(255, 255, 255)
    shp2.Line.Weight = 0.75
    shp2.Adjustments(1) = 0.25
    shp2.Select (msoFalse)

    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeChevron, Left:=401.38557, Top:=9.9212536, Width:=24.377937, Height:=13.606291)
    shp3.Fill.Visible = msoTrue
    shp3.Fill.ForeColor.RGB = RGB(89, 171, 244)
    shp3.Fill.Transparency = 0
    shp3.Line.Visible = msoTrue
    shp3.Line.ForeColor.RGB = RGB(255, 255, 255)
    shp3.Line.Weight = 0.75
    shp3.Adjustments(1) = 0.25
    shp3.Select (msoFalse)
    
    ActiveWindow.Selection.ShapeRange.Group.Select
    ActiveWindow.Selection.ShapeRange(1).Left = (phd.Left + (phd.Width / 2)) - (ActiveWindow.Selection.ShapeRange(1).Width / 2)
    ActiveWindow.Selection.ShapeRange(1).Top = (tit.Top / 2) - (ActiveWindow.Selection.ShapeRange(1).Height / 2)
    ActiveWindow.Selection.ShapeRange(1).Ungroup.Select

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub Summary(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim phd As Shape
    Dim shp As Shape
    Dim sld As Slide
    Dim LeftRight As Single

    'Summary box

On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else

    LeftRight = flexWidth(ActivePresentation)

    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)
    Set sld = Application.ActiveWindow.View.Slide
    
If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=10, Height:=41.952729)
    shp.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=10, Height:=31.748011)
    shp.Select msoTrue
Else
    Set shp = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=10, Height:=41.952729)
    shp.Select msoTrue
End If

Call ParameterBoxHeaderFlat
    
    
If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    shp.TextFrame2.TextRange.Characters.Text = "Summary box: 16 pt, bold, one line, do not change color, size, position or font size"
Else
    shp.TextFrame2.TextRange.Characters.Text = "Summary box: 18 pt, bold, one line, do not change color, size, position or font size"
    shp.TextFrame2.TextRange.Font.Size = 18
End If
    
    
    shp.Top = (phd.Top + phd.Height) - shp.Height
    shp.Width = LeftRight
    shp.Select (msoTrue)

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub Takeaway(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim phd As Shape
    Dim shp As Shape
    Dim sld As Slide
    
    'Takeaway box

On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else

    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)
    Set sld = Application.ActiveWindow.View.Slide
    
If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=467.99971, Height:=41.952729)
    shp.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=467.99971, Height:=31.748011)
    shp.Select msoTrue
Else
    Set shp = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=467.99971, Height:=41.952729)
    shp.Select msoTrue
End If

Call ParameterBoxHeaderFlat


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    shp.TextFrame2.TextRange.Characters.Text = "Takeaway box, 14pt, bold, two lines max., do not change color, size, position or font size"
    shp.TextFrame2.TextRange.Font.Size = 14
Else
    shp.TextFrame2.TextRange.Characters.Text = "Takeaway box, 16pt, bold, two lines max., do not change color, size, position or font size"
End If
    
    shp.Left = (phd.Width / 2) + phd.Left - (shp.Width / 2)
    shp.Top = (phd.Top + phd.Height) - shp.Height
    shp.Select (msoTrue)

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub ValueChain02(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    
    Dim phd As Shape
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim shp11 As Shape
    Dim shp12 As Shape
    Dim sld As Slide

    'Value Chain (2)
    
On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else
    
    Set sld = Application.ActiveWindow.View.Slide
    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapePentagon, Left:=38.834621, Top:=0, Width:=360.5667, Height:=34.015727)
    shp1.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapePentagon, Left:=28.346439, Top:=0, Width:=340.44073, Height:=31.464547)
    shp1.Select msoTrue
Else
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapePentagon, Left:=28.346439, Top:=0, Width:=340.44073, Height:=34.015727)
    shp1.Select msoTrue
End If

Call ParameterBoxHeaderFlat


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeChevron, Left:=390.33046, Top:=0, Width:=360.5667, Height:=34.015727)
    shp2.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeChevron, Left:=359.99977, Top:=0, Width:=340.44073, Height:=31.464547)
    shp2.Select msoTrue
Else
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeChevron, Left:=359.99977, Top:=0, Width:=340.44073, Height:=34.015727)
    shp2.Select msoTrue
End If

Call ParameterBoxHeaderFlat


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=38.834621, Top:=0, Width:=350.64545, Height:=283.46439)
    shp11.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=28.346439, Top:=0, Width:=331.08641, Height:=252.85023)
    shp11.Select msoTrue
Else
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=28.346439, Top:=0, Width:=331.08641, Height:=283.46439)
    shp11.Select msoTrue
End If

Call ParameterColumnBodyStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=390.33046, Top:=0, Width:=350.64545, Height:=283.46439)
    shp12.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=360.5667, Top:=0, Width:=331.08641, Height:=252.85023)
    shp12.Select msoTrue
Else
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=360.5667, Top:=0, Width:=331.08641, Height:=283.46439)
    shp12.Select msoTrue
End If

Call ParameterColumnBodyStandard


shp1.Top = phd.Top
shp2.Top = phd.Top
shp11.Top = shp1.Top + shp1.Height
shp12.Top = shp2.Top + shp2.Height

shp1.Line.ForeColor.RGB = RGB(255, 255, 255)
shp2.Line.ForeColor.RGB = RGB(255, 255, 255)

shp1.Adjustments(1) = 0.25
shp2.Adjustments(1) = 0.25

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    shp1.TextFrame2.TextRange.Font.Size = 14
    shp2.TextFrame2.TextRange.Font.Size = 14
    shp11.TextFrame2.TextRange.Font.Size = 12
    shp12.TextFrame2.TextRange.Font.Size = 12
End If

shp1.Select (msoTrue)
shp2.Select (msoFalse)
shp11.Select (msoFalse)
shp12.Select (msoFalse)

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub ValueChain03(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    
    Dim phd As Shape
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim shp3 As Shape
    Dim shp11 As Shape
    Dim shp12 As Shape
    Dim shp13 As Shape
    Dim sld As Slide

    'Value Chain (3)

On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else

    Set sld = Application.ActiveWindow.View.Slide
    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapePentagon, Left:=38.834621, Top:=0, Width:=243.77246, Height:=34.015727)
    shp1.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapePentagon, Left:=28.346439, Top:=0, Width:=230.17308, Height:=31.464547)
    shp1.Select msoTrue
Else
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapePentagon, Left:=28.346439, Top:=0, Width:=230.17308, Height:=34.015727)
    shp1.Select msoTrue
End If

Call ParameterBoxHeaderFlat


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeChevron, Left:=272.96847, Top:=0, Width:=243.77246, Height:=34.015727)
    shp2.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeChevron, Left:=249.44866, Top:=0, Width:=230.17308, Height:=31.464547)
    shp2.Select msoTrue
Else
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeChevron, Left:=249.44866, Top:=0, Width:=230.17308, Height:=34.015727)
    shp2.Select msoTrue
End If

Call ParameterBoxHeaderFlat


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeChevron, Left:=506.83433, Top:=0, Width:=243.77246, Height:=34.015727)
    shp3.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeChevron, Left:=470.55088, Top:=0, Width:=230.17308, Height:=31.464547)
    shp3.Select msoTrue
Else
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeChevron, Left:=470.55088, Top:=0, Width:=230.17308, Height:=34.015727)
    shp3.Select msoTrue
End If

Call ParameterBoxHeaderFlat


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=38.834621, Top:=0, Width:=234.14158, Height:=283.46439)
    shp11.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=28.346439, Top:=0, Width:=221.10222, Height:=252.85023)
    shp11.Select msoTrue
Else
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=28.346439, Top:=0, Width:=221.10222, Height:=283.46439)
    shp11.Select msoTrue
End If

Call ParameterColumnBodyStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=272.96847, Top:=0, Width:=234.14158, Height:=283.46439)
    shp12.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=249.44866, Top:=0, Width:=221.10222, Height:=252.85023)
    shp12.Select msoTrue
Else
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=249.44866, Top:=0, Width:=221.10222, Height:=283.46439)
    shp12.Select msoTrue
End If

Call ParameterColumnBodyStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=506.83433, Top:=0, Width:=234.14158, Height:=283.46439)
    shp13.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=470.55088, Top:=0, Width:=221.10222, Height:=252.85023)
    shp13.Select msoTrue
Else
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=470.55088, Top:=0, Width:=221.10222, Height:=283.46439)
    shp13.Select msoTrue
End If

Call ParameterColumnBodyStandard


shp1.Top = phd.Top
shp2.Top = phd.Top
shp3.Top = phd.Top
shp11.Top = shp1.Top + shp1.Height
shp12.Top = shp2.Top + shp2.Height
shp13.Top = shp3.Top + shp3.Height

shp1.Line.ForeColor.RGB = RGB(255, 255, 255)
shp2.Line.ForeColor.RGB = RGB(255, 255, 255)
shp3.Line.ForeColor.RGB = RGB(255, 255, 255)

shp1.Adjustments(1) = 0.25
shp2.Adjustments(1) = 0.25
shp3.Adjustments(1) = 0.25

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    shp1.TextFrame2.TextRange.Font.Size = 14
    shp2.TextFrame2.TextRange.Font.Size = 14
    shp3.TextFrame2.TextRange.Font.Size = 14
    shp11.TextFrame2.TextRange.Font.Size = 12
    shp12.TextFrame2.TextRange.Font.Size = 12
    shp13.TextFrame2.TextRange.Font.Size = 12
End If

shp1.Select (msoTrue)
shp2.Select (msoFalse)
shp3.Select (msoFalse)
shp11.Select (msoFalse)
shp12.Select (msoFalse)
shp13.Select (msoFalse)

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

Public Sub ValueChain04(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    
    Dim phd As Shape
    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim shp3 As Shape
    Dim shp4 As Shape
    Dim shp11 As Shape
    Dim shp12 As Shape
    Dim shp13 As Shape
    Dim shp14 As Shape
    Dim sld As Slide

    'Value Chain (4)
    
On Error GoTo ErrMsg

If ActiveWindow.Selection.SlideRange.Count <> 1 Then
        MsgBox "This function cannot be used for several slides at the same time"
        Exit Sub
    Else
    
    Set sld = Application.ActiveWindow.View.Slide
    Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapePentagon, Left:=38.834621, Top:=0, Width:=185.38571, Height:=34.015727)
    shp1.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapePentagon, Left:=28.346439, Top:=0, Width:=174.89753, Height:=31.464547)
    shp1.Select msoTrue
Else
    Set shp1 = sld.Shapes.AddShape(Type:=msoShapePentagon, Left:=28.346439, Top:=0, Width:=174.89753, Height:=34.015727)
    shp1.Select msoTrue
End If

Call ParameterBoxHeaderFlat


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeChevron, Left:=214.29908, Top:=0, Width:=185.38571, Height:=34.015727)
    shp2.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeChevron, Left:=194.17311, Top:=0, Width:=174.89753, Height:=31.464547)
    shp2.Select msoTrue
Else
    Set shp2 = sld.Shapes.AddShape(Type:=msoShapeChevron, Left:=194.17311, Top:=0, Width:=174.89753, Height:=34.015727)
    shp2.Select msoTrue
End If

Call ParameterBoxHeaderFlat


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeChevron, Left:=389.76353, Top:=0, Width:=185.38571, Height:=34.015727)
    shp3.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeChevron, Left:=359.71631, Top:=0, Width:=174.89753, Height:=31.464547)
    shp3.Select msoTrue
Else
    Set shp3 = sld.Shapes.AddShape(Type:=msoShapeChevron, Left:=359.71631, Top:=0, Width:=174.89753, Height:=34.015727)
    shp3.Select msoTrue
End If

Call ParameterBoxHeaderFlat


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp4 = sld.Shapes.AddShape(Type:=msoShapeChevron, Left:=565.51145, Top:=0, Width:=185.38571, Height:=34.015727)
    shp4.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp4 = sld.Shapes.AddShape(Type:=msoShapeChevron, Left:=525.25951, Top:=0, Width:=174.89753, Height:=31.464547)
    shp4.Select msoTrue
Else
    Set shp4 = sld.Shapes.AddShape(Type:=msoShapeChevron, Left:=525.25951, Top:=0, Width:=174.89753, Height:=34.015727)
    shp4.Select msoTrue
End If

Call ParameterBoxHeaderFlat


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=38.834621, Top:=0, Width:=175.74792, Height:=283.46439)
    shp11.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=28.346439, Top:=0, Width:=165.82667, Height:=252.85023)
    shp11.Select msoTrue
Else
    Set shp11 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=28.346439, Top:=0, Width:=165.82667, Height:=283.46439)
    shp11.Select msoTrue
End If

Call ParameterColumnBodyStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=214.29908, Top:=0, Width:=175.74792, Height:=283.46439)
    shp12.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=194.17311, Top:=0, Width:=165.82667, Height:=252.85023)
    shp12.Select msoTrue
Else
    Set shp12 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=194.17311, Top:=0, Width:=165.82667, Height:=283.46439)
    shp12.Select msoTrue
End If

Call ParameterColumnBodyStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=389.76353, Top:=0, Width:=175.74792, Height:=283.46439)
    shp13.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=359.71631, Top:=0, Width:=165.82667, Height:=252.85023)
    shp13.Select msoTrue
Else
    Set shp13 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=359.71631, Top:=0, Width:=165.82667, Height:=283.46439)
    shp13.Select msoTrue
End If

Call ParameterColumnBodyStandard


If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeA4Paper Then
    Set shp14 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=565.51145, Top:=0, Width:=175.74792, Height:=283.46439)
    shp14.Select msoTrue
ElseIf ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    Set shp14 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=525.25951, Top:=0, Width:=165.82667, Height:=252.85023)
    shp14.Select msoTrue
Else
    Set shp14 = sld.Shapes.AddShape(Type:=msoShapeRectangle, Left:=525.25951, Top:=0, Width:=165.82667, Height:=283.46439)
    shp14.Select msoTrue
End If

Call ParameterColumnBodyStandard


shp1.Top = phd.Top
shp2.Top = phd.Top
shp3.Top = phd.Top
shp4.Top = phd.Top
shp11.Top = shp1.Top + shp1.Height
shp12.Top = shp2.Top + shp2.Height
shp13.Top = shp3.Top + shp3.Height
shp14.Top = shp4.Top + shp4.Height

shp1.Line.ForeColor.RGB = RGB(255, 255, 255)
shp2.Line.ForeColor.RGB = RGB(255, 255, 255)
shp3.Line.ForeColor.RGB = RGB(255, 255, 255)
shp4.Line.ForeColor.RGB = RGB(255, 255, 255)

shp1.Adjustments(1) = 0.25
shp2.Adjustments(1) = 0.25
shp3.Adjustments(1) = 0.25
shp4.Adjustments(1) = 0.25

If ActiveWindow.Presentation.PageSetup.SlideSize = ppSlideSizeOnScreen16x9 Then
    shp1.TextFrame2.TextRange.Font.Size = 14
    shp2.TextFrame2.TextRange.Font.Size = 14
    shp3.TextFrame2.TextRange.Font.Size = 14
    shp4.TextFrame2.TextRange.Font.Size = 14
    shp11.TextFrame2.TextRange.Font.Size = 12
    shp12.TextFrame2.TextRange.Font.Size = 12
    shp13.TextFrame2.TextRange.Font.Size = 12
    shp14.TextFrame2.TextRange.Font.Size = 12
End If

shp1.Select (msoTrue)
shp2.Select (msoFalse)
shp3.Select (msoFalse)
shp4.Select (msoFalse)
shp11.Select (msoFalse)
shp12.Select (msoFalse)
shp13.Select (msoFalse)
shp14.Select (msoFalse)

    End If
Exit Sub
    
ErrMsg:
    MsgBox "Please select a slide"
End Sub

'EEE Farben

Public Sub F003003003(control As IRibbonControl)

    Dim oshp As Shape
    Dim oTbl As Table
    Dim x As Long
    Dim y As Long
    On Error GoTo err
    
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
        End With
    
    Else
    For Each oshp In ActiveWindow.Selection.ShapeRange
    If ActiveWindow.Selection.ShapeRange(1).HasTable Then
        Set oTbl = ActiveWindow.Selection.ShapeRange.Table
            For x = 1 To oTbl.Rows.Count
                For y = 1 To oTbl.Columns.Count
                    If oTbl.Cell(x, y).Selected Then
        oTbl.Cell(x, y).Shape.Fill.Visible = msoTrue
        oTbl.Cell(x, y).Shape.Fill.ForeColor.RGB = RGB(255, 255, 255)
                    End If
                Next
            Next
    Else
        oshp.Fill.Visible = msoTrue
        oshp.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End If
    Next oshp
    End If
    
Exit Sub
 
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub L003003003(control As IRibbonControl)

    Dim oshp As Shape
    On Error GoTo err
    
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 255)
 End With
    
    Else
    For Each oshp In ActiveWindow.Selection.ShapeRange
        oshp.Line.Visible = msoTrue
        oshp.Line.ForeColor.RGB = RGB(255, 255, 255)
    Next oshp
    End If
    
Exit Sub
 
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub TextBlack(control As IRibbonControl)

    Dim oshp As Shape
    Dim oTbl As Table
    Dim x As Long
    Dim y As Long
    On Error GoTo err
    
    If ActiveWindow.Selection.Type = ppSelectionText Then
    With ActiveWindow.Selection.TextRange
    .Font.Color.RGB = RGB(3, 3, 3)
        End With
    
    Else
    For Each oshp In ActiveWindow.Selection.ShapeRange
        If ActiveWindow.Selection.ShapeRange(1).HasTable Then
        Set oTbl = ActiveWindow.Selection.ShapeRange.Table
            For x = 1 To oTbl.Rows.Count
                For y = 1 To oTbl.Columns.Count
                    If oTbl.Cell(x, y).Selected Then
        oTbl.Cell(x, y).Shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(3, 3, 3)
                    End If
                Next
            Next
        Else
        oshp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(3, 3, 3)
        End If
    Next oshp
    
    End If
    Exit Sub
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub T003003003(control As IRibbonControl)

    Dim oshp As Shape
    Dim oTbl As Table
    Dim x As Long
    Dim y As Long
    On Error GoTo err
    
    If ActiveWindow.Selection.Type = ppSelectionText Then
    With ActiveWindow.Selection.TextRange
    .Font.Color.RGB = RGB(255, 255, 255)
        End With
    
    Else
    For Each oshp In ActiveWindow.Selection.ShapeRange
        If ActiveWindow.Selection.ShapeRange(1).HasTable Then
        Set oTbl = ActiveWindow.Selection.ShapeRange.Table
            For x = 1 To oTbl.Rows.Count
                For y = 1 To oTbl.Columns.Count
                    If oTbl.Cell(x, y).Selected Then
        oTbl.Cell(x, y).Shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                    End If
                Next
            Next
        Else
        oshp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        End If
    Next oshp
    
    End If
    Exit Sub
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub F255255255(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim oshp As Shape
    Dim oTbl As Table
    Dim x As Long
    Dim y As Long
    On Error GoTo err
    
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(154, 154, 154)
        End With
    
    Else
    For Each oshp In ActiveWindow.Selection.ShapeRange
    If ActiveWindow.Selection.ShapeRange(1).HasTable Then
        Set oTbl = ActiveWindow.Selection.ShapeRange.Table
            For x = 1 To oTbl.Rows.Count
                For y = 1 To oTbl.Columns.Count
                    If oTbl.Cell(x, y).Selected Then
        oTbl.Cell(x, y).Shape.Fill.Visible = msoTrue
        oTbl.Cell(x, y).Shape.Fill.ForeColor.RGB = RGB(154, 154, 154)
                    End If
                Next
            Next
    Else
        oshp.Fill.Visible = msoTrue
        oshp.Fill.ForeColor.RGB = RGB(154, 154, 154)
    End If
    Next oshp
    End If
    
Exit Sub
 
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub L255255255(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim oshp As Shape
    On Error GoTo err
    
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(154, 154, 154)
 End With
    
    Else
    For Each oshp In ActiveWindow.Selection.ShapeRange
        oshp.Line.Visible = msoTrue
        oshp.Line.ForeColor.RGB = RGB(154, 154, 154)
    Next oshp
    End If
    
Exit Sub
 
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub T255255255(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim oshp As Shape
    Dim oTbl As Table
    Dim x As Long
    Dim y As Long
    On Error GoTo err
    
    If ActiveWindow.Selection.Type = ppSelectionText Then
    With ActiveWindow.Selection.TextRange
    .Font.Color.RGB = RGB(154, 154, 154)
        End With
    
    Else
    For Each oshp In ActiveWindow.Selection.ShapeRange
        If ActiveWindow.Selection.ShapeRange(1).HasTable Then
        Set oTbl = ActiveWindow.Selection.ShapeRange.Table
            For x = 1 To oTbl.Rows.Count
                For y = 1 To oTbl.Columns.Count
                    If oTbl.Cell(x, y).Selected Then
        oTbl.Cell(x, y).Shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(154, 154, 154)
                    End If
                Next
            Next
        Else
        oshp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(154, 154, 154)
        End If
    Next oshp
    
    End If
    Exit Sub
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub F000015045(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim oshp As Shape
    Dim oTbl As Table
    Dim x As Long
    Dim y As Long
    On Error GoTo err
    
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 15, 80)
        End With
    
    Else
    For Each oshp In ActiveWindow.Selection.ShapeRange
    If ActiveWindow.Selection.ShapeRange(1).HasTable Then
        Set oTbl = ActiveWindow.Selection.ShapeRange.Table
            For x = 1 To oTbl.Rows.Count
                For y = 1 To oTbl.Columns.Count
                    If oTbl.Cell(x, y).Selected Then
        oTbl.Cell(x, y).Shape.Fill.Visible = msoTrue
        oTbl.Cell(x, y).Shape.Fill.ForeColor.RGB = RGB(0, 15, 80)
                    End If
                Next
            Next
    Else
        oshp.Fill.Visible = msoTrue
        oshp.Fill.ForeColor.RGB = RGB(0, 15, 80)
    End If
    Next oshp
    End If
    
Exit Sub
 
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub L000015045(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim oshp As Shape
    On Error GoTo err
    
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 15, 80)
 End With
    
    Else
    For Each oshp In ActiveWindow.Selection.ShapeRange
        oshp.Line.Visible = msoTrue
        oshp.Line.ForeColor.RGB = RGB(0, 15, 80)
    Next oshp
    End If
    
Exit Sub
 
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub T000015045(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim oshp As Shape
    Dim oTbl As Table
    Dim x As Long
    Dim y As Long
    On Error GoTo err
    
    If ActiveWindow.Selection.Type = ppSelectionText Then
    With ActiveWindow.Selection.TextRange
    .Font.Color.RGB = RGB(0, 15, 80)
        End With
    
    Else
    For Each oshp In ActiveWindow.Selection.ShapeRange
        If ActiveWindow.Selection.ShapeRange(1).HasTable Then
        Set oTbl = ActiveWindow.Selection.ShapeRange.Table
            For x = 1 To oTbl.Rows.Count
                For y = 1 To oTbl.Columns.Count
                    If oTbl.Cell(x, y).Selected Then
        oTbl.Cell(x, y).Shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 15, 80)
                    End If
                Next
            Next
        Else
        oshp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 15, 80)
        End If
    Next oshp
    
    End If
    Exit Sub
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub F002226202(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
   Dim oshp As Shape
    Dim oTbl As Table
    Dim x As Long
    Dim y As Long
    On Error GoTo err
    
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(29, 8, 146)
        End With
    
    Else
    For Each oshp In ActiveWindow.Selection.ShapeRange
    If ActiveWindow.Selection.ShapeRange(1).HasTable Then
        Set oTbl = ActiveWindow.Selection.ShapeRange.Table
            For x = 1 To oTbl.Rows.Count
                For y = 1 To oTbl.Columns.Count
                    If oTbl.Cell(x, y).Selected Then
        oTbl.Cell(x, y).Shape.Fill.Visible = msoTrue
        oTbl.Cell(x, y).Shape.Fill.ForeColor.RGB = RGB(29, 8, 146)
                    End If
                Next
            Next
    Else
        oshp.Fill.Visible = msoTrue
        oshp.Fill.ForeColor.RGB = RGB(29, 8, 146)
    End If
    Next oshp
    End If
    
Exit Sub
 
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub L002226202(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim oshp As Shape
    On Error GoTo err
    
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(29, 8, 146)
 End With
    
    Else
    For Each oshp In ActiveWindow.Selection.ShapeRange
        oshp.Line.Visible = msoTrue
        oshp.Line.ForeColor.RGB = RGB(29, 8, 146)
    Next oshp
    End If
    
Exit Sub
 
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub T002226202(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim oshp As Shape
    Dim oTbl As Table
    Dim x As Long
    Dim y As Long
    On Error GoTo err
    
    If ActiveWindow.Selection.Type = ppSelectionText Then
    With ActiveWindow.Selection.TextRange
    .Font.Color.RGB = RGB(29, 8, 146)
        End With
    
    Else
    For Each oshp In ActiveWindow.Selection.ShapeRange
        If ActiveWindow.Selection.ShapeRange(1).HasTable Then
        Set oTbl = ActiveWindow.Selection.ShapeRange.Table
            For x = 1 To oTbl.Rows.Count
                For y = 1 To oTbl.Columns.Count
                    If oTbl.Cell(x, y).Selected Then
        oTbl.Cell(x, y).Shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(29, 8, 146)
                    End If
                Next
            Next
        Else
        oshp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(29, 8, 146)
        End If
    Next oshp
    
    End If
    Exit Sub
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub F228228228(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim oshp As Shape
    Dim oTbl As Table
    Dim x As Long
    Dim y As Long
    On Error GoTo err
    
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(17, 186, 221)
        End With
    
    Else
    For Each oshp In ActiveWindow.Selection.ShapeRange
    If ActiveWindow.Selection.ShapeRange(1).HasTable Then
        Set oTbl = ActiveWindow.Selection.ShapeRange.Table
            For x = 1 To oTbl.Rows.Count
                For y = 1 To oTbl.Columns.Count
                    If oTbl.Cell(x, y).Selected Then
        oTbl.Cell(x, y).Shape.Fill.Visible = msoTrue
        oTbl.Cell(x, y).Shape.Fill.ForeColor.RGB = RGB(17, 186, 221)
                    End If
                Next
            Next
    Else
        oshp.Fill.Visible = msoTrue
        oshp.Fill.ForeColor.RGB = RGB(17, 186, 221)
    End If
    Next oshp
    End If
    
Exit Sub
 
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub L228228228(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim oshp As Shape
    On Error GoTo err
    
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(17, 186, 221)
 End With
    
    Else
    For Each oshp In ActiveWindow.Selection.ShapeRange
        oshp.Line.Visible = msoTrue
        oshp.Line.ForeColor.RGB = RGB(17, 186, 221)
    Next oshp
    End If
    
Exit Sub
 
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub T228228228(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim oshp As Shape
    Dim oTbl As Table
    Dim x As Long
    Dim y As Long
    On Error GoTo err
    
    If ActiveWindow.Selection.Type = ppSelectionText Then
    With ActiveWindow.Selection.TextRange
    .Font.Color.RGB = RGB(17, 186, 221)
        End With
    
    Else
    For Each oshp In ActiveWindow.Selection.ShapeRange
        If ActiveWindow.Selection.ShapeRange(1).HasTable Then
        Set oTbl = ActiveWindow.Selection.ShapeRange.Table
            For x = 1 To oTbl.Rows.Count
                For y = 1 To oTbl.Columns.Count
                    If oTbl.Cell(x, y).Selected Then
        oTbl.Cell(x, y).Shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(17, 186, 221)
                    End If
                Next
            Next
        Else
        oshp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(17, 186, 221)
        End If
    Next oshp
    
    End If
    Exit Sub
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub F206247233(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim oshp As Shape
    Dim oTbl As Table
    Dim x As Long
    Dim y As Long
    On Error GoTo err
    
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(9, 220, 223)
        End With
    
    Else
    For Each oshp In ActiveWindow.Selection.ShapeRange
    If ActiveWindow.Selection.ShapeRange(1).HasTable Then
        Set oTbl = ActiveWindow.Selection.ShapeRange.Table
            For x = 1 To oTbl.Rows.Count
                For y = 1 To oTbl.Columns.Count
                    If oTbl.Cell(x, y).Selected Then
        oTbl.Cell(x, y).Shape.Fill.Visible = msoTrue
        oTbl.Cell(x, y).Shape.Fill.ForeColor.RGB = RGB(9, 220, 223)
                    End If
                Next
            Next
    Else
        oshp.Fill.Visible = msoTrue
        oshp.Fill.ForeColor.RGB = RGB(9, 220, 223)
    End If
    Next oshp
    End If
    
Exit Sub
 
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub L206247233(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim oshp As Shape
    On Error GoTo err
    
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(9, 220, 223)
 End With
    
    Else
    For Each oshp In ActiveWindow.Selection.ShapeRange
        oshp.Line.Visible = msoTrue
        oshp.Line.ForeColor.RGB = RGB(9, 220, 223)
    Next oshp
    End If
    
Exit Sub
 
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub T206247233(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim oshp As Shape
    Dim oTbl As Table
    Dim x As Long
    Dim y As Long
    On Error GoTo err
    
    If ActiveWindow.Selection.Type = ppSelectionText Then
    With ActiveWindow.Selection.TextRange
    .Font.Color.RGB = RGB(9, 220, 223)
        End With
    
    Else
    For Each oshp In ActiveWindow.Selection.ShapeRange
        If ActiveWindow.Selection.ShapeRange(1).HasTable Then
        Set oTbl = ActiveWindow.Selection.ShapeRange.Table
            For x = 1 To oTbl.Rows.Count
                For y = 1 To oTbl.Columns.Count
                    If oTbl.Cell(x, y).Selected Then
        oTbl.Cell(x, y).Shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(9, 220, 223)
                    End If
                Next
            Next
        Else
        oshp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(9, 220, 223)
        End If
    Next oshp
    
    End If
    Exit Sub
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub F041245130(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim oshp As Shape
    Dim oTbl As Table
    Dim x As Long
    Dim y As Long
    On Error GoTo err
    
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(228, 228, 228)
        End With
    
    Else
    For Each oshp In ActiveWindow.Selection.ShapeRange
    If ActiveWindow.Selection.ShapeRange(1).HasTable Then
        Set oTbl = ActiveWindow.Selection.ShapeRange.Table
            For x = 1 To oTbl.Rows.Count
                For y = 1 To oTbl.Columns.Count
                    If oTbl.Cell(x, y).Selected Then
        oTbl.Cell(x, y).Shape.Fill.Visible = msoTrue
        oTbl.Cell(x, y).Shape.Fill.ForeColor.RGB = RGB(228, 228, 228)
                    End If
                Next
            Next
    Else
        oshp.Fill.Visible = msoTrue
        oshp.Fill.ForeColor.RGB = RGB(228, 228, 228)
    End If
    Next oshp
    End If
    
Exit Sub
 
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub L041245130(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim oshp As Shape
    On Error GoTo err
    
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(228, 228, 228)
 End With
    
    Else
    For Each oshp In ActiveWindow.Selection.ShapeRange
        oshp.Line.Visible = msoTrue
        oshp.Line.ForeColor.RGB = RGB(228, 228, 228)
    Next oshp
    End If
    
Exit Sub
 
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub T041245130(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim oshp As Shape
    Dim oTbl As Table
    Dim x As Long
    Dim y As Long
    On Error GoTo err
    
    If ActiveWindow.Selection.Type = ppSelectionText Then
    With ActiveWindow.Selection.TextRange
    .Font.Color.RGB = RGB(228, 228, 228)
        End With
    
    Else
    For Each oshp In ActiveWindow.Selection.ShapeRange
        If ActiveWindow.Selection.ShapeRange(1).HasTable Then
        Set oTbl = ActiveWindow.Selection.ShapeRange.Table
            For x = 1 To oTbl.Rows.Count
                For y = 1 To oTbl.Columns.Count
                    If oTbl.Cell(x, y).Selected Then
        oTbl.Cell(x, y).Shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(228, 228, 228)
                    End If
                Next
            Next
        Else
        oshp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(228, 228, 228)
        End If
    Next oshp
    
    End If
    Exit Sub
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub F021111119(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim oshp As Shape
    Dim oTbl As Table
    Dim x As Long
    Dim y As Long
    On Error GoTo err
    
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(32, 119, 218)
        End With
    
    Else
    For Each oshp In ActiveWindow.Selection.ShapeRange
    If ActiveWindow.Selection.ShapeRange(1).HasTable Then
        Set oTbl = ActiveWindow.Selection.ShapeRange.Table
            For x = 1 To oTbl.Rows.Count
                For y = 1 To oTbl.Columns.Count
                    If oTbl.Cell(x, y).Selected Then
        oTbl.Cell(x, y).Shape.Fill.Visible = msoTrue
        oTbl.Cell(x, y).Shape.Fill.ForeColor.RGB = RGB(32, 119, 218)
                    End If
                Next
            Next
    Else
        oshp.Fill.Visible = msoTrue
        oshp.Fill.ForeColor.RGB = RGB(32, 119, 218)
    End If
    Next oshp
    End If
    
Exit Sub
 
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub L021111119(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim oshp As Shape
    On Error GoTo err
    
    If ActiveWindow.Selection.HasChildShapeRange Then
        With ActiveWindow.Selection.ChildShapeRange.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(32, 119, 218)
 End With
    
    Else
    For Each oshp In ActiveWindow.Selection.ShapeRange
        oshp.Line.Visible = msoTrue
        oshp.Line.ForeColor.RGB = RGB(32, 119, 218)
    Next oshp
    End If
    
Exit Sub
 
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub T021111119(control As IRibbonControl)
   'YYY If Not Init Then Exit Sub 'ZZZ
    Dim oshp As Shape
    Dim oTbl As Table
    Dim x As Long
    Dim y As Long
    On Error GoTo err
    
    If ActiveWindow.Selection.Type = ppSelectionText Then
    With ActiveWindow.Selection.TextRange
    .Font.Color.RGB = RGB(32, 119, 218)
        End With
    
    Else
    For Each oshp In ActiveWindow.Selection.ShapeRange
        If ActiveWindow.Selection.ShapeRange(1).HasTable Then
        Set oTbl = ActiveWindow.Selection.ShapeRange.Table
            For x = 1 To oTbl.Rows.Count
                For y = 1 To oTbl.Columns.Count
                    If oTbl.Cell(x, y).Selected Then
        oTbl.Cell(x, y).Shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(32, 119, 218)
                    End If
                Next
            Next
        Else
        oshp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(32, 119, 218)
        End If
    Next oshp
    
    End If
    Exit Sub
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub MoveTL(control As IRibbonControl)
'Spezialanfertigung für OMMAX

Dim phd As Shape
Dim shp As Shape

Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)

    On Error GoTo err
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Top = phd.Top
        shp.Left = phd.Left
    Next shp
    Exit Sub
err:
    MsgBox "Please select at least one shape"

End Sub

Public Sub MoveBL(control As IRibbonControl)
'Spezialanfertigung für OMMAX
Dim phd As Shape
Dim shp As Shape

Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)

    On Error GoTo err
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Top = (phd.Top + phd.Height - shp.Height)
        shp.Left = phd.Left
    Next shp
    Exit Sub
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub MoveTR(control As IRibbonControl)
'Spezialanfertigung für OMMAX
Dim phd As Shape
Dim shp As Shape

Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)

    On Error GoTo err
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Top = phd.Top
        shp.Left = (phd.Left + phd.Width - shp.Width)
    Next shp
    Exit Sub
err:
    MsgBox "Please select at least one shape"
End Sub

Public Sub MoveBR(control As IRibbonControl)
'Spezialanfertigung für OMMAX
Dim phd As Shape
Dim shp As Shape

Set phd = Application.ActivePresentation.SlideMaster.Shapes.Placeholders(2)

    On Error GoTo err
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Top = (phd.Top + phd.Height - shp.Height)
        shp.Left = (phd.Left + phd.Width - shp.Width)
    Next shp
    Exit Sub
err:
    MsgBox "Please select at least one shape"
End Sub

