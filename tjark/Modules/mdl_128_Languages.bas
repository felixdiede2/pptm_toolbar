Attribute VB_Name = "mdl_128_Languages"
Option Explicit

Public Sub Lang_DE_Default()
    On Error Resume Next
       
    ActivePresentation.DefaultLanguageID = msoLanguageIDGerman

End Sub

Public Sub Lang_EN_Default()
    On Error Resume Next
       
    ActivePresentation.DefaultLanguageID = msoLanguageIDEnglishUS

End Sub

Sub Lang_DE_Presentation()
    On Error Resume Next
    
    Dim sCount As Integer
    Dim fCount As Integer
    Dim i, j, k As Integer
    Dim Sprache As Variant
    
    Sprache = msoLanguageIDGerman

    For j = 1 To ActivePresentation.Slides.Count
        For k = 1 To ActivePresentation.Slides(j).Shapes.Count
            If ActivePresentation.Slides(j).Shapes(k).HasTextFrame Then
                ActivePresentation.Slides(j).Shapes(k).TextFrame.TextRange.LanguageID = Sprache
            End If
            If ActivePresentation.Slides(j).Shapes(k).Type = msoGroup Then
                For i = 1 To ActivePresentation.Slides(j).Shapes(k).GroupItems.Count
                    If ActivePresentation.Slides(j).Shapes(k).GroupItems(i).HasTextFrame Then
                        ActivePresentation.Slides(j).Shapes(k).GroupItems(i).TextFrame.TextRange.LanguageID = Sprache
                    End If
                Next i
            End If
        Next k
    Next j
End Sub

Sub Lang_EN_Presentation()
    On Error Resume Next
    
    Dim sCount As Integer
    Dim fCount As Integer
    Dim i, j, k As Integer
    Dim Sprache As Variant
    
    Sprache = msoLanguageIDEnglishUS

    For j = 1 To ActivePresentation.Slides.Count
        For k = 1 To ActivePresentation.Slides(j).Shapes.Count
            If ActivePresentation.Slides(j).Shapes(k).HasTextFrame Then
                ActivePresentation.Slides(j).Shapes(k).TextFrame.TextRange.LanguageID = Sprache
            End If
            If ActivePresentation.Slides(j).Shapes(k).Type = msoGroup Then
                For i = 1 To ActivePresentation.Slides(j).Shapes(k).GroupItems.Count
                    If ActivePresentation.Slides(j).Shapes(k).GroupItems(i).HasTextFrame Then
                        ActivePresentation.Slides(j).Shapes(k).GroupItems(i).TextFrame.TextRange.LanguageID = Sprache
                    End If
                Next i
            End If
        Next k
    Next j
End Sub

Sub Lang_DE_Selection()
    On Error Resume Next
    
    Dim S As Shape
    Dim i As Integer
    Dim Sprache As Variant
    
    Sprache = msoLanguageIDGerman
    
    For Each S In ActiveWindow.Selection.ShapeRange
        If S.HasTextFrame Then
            S.TextFrame.TextRange.LanguageID = Sprache
        End If
        If S.Type = msoGroup Then
            For i = 1 To S.GroupItems.Count
                If S.GroupItems(i).HasTextFrame Then
                    S.GroupItems(i).TextFrame.TextRange.LanguageID = Sprache
                End If
            Next i
        End If
    Next S
End Sub

Sub Lang_EN_Selection()
    On Error Resume Next
    
    Dim S As Shape
    Dim i As Integer
    Dim Sprache As Variant
    
    Sprache = msoLanguageIDEnglishUS
    
    For Each S In ActiveWindow.Selection.ShapeRange
        If S.HasTextFrame Then
            S.TextFrame.TextRange.LanguageID = Sprache
        End If
        If S.Type = msoGroup Then
            For i = 1 To S.GroupItems.Count
                If S.GroupItems(i).HasTextFrame Then
                    S.GroupItems(i).TextFrame.TextRange.LanguageID = Sprache
                End If
            Next i
        End If
    Next S
End Sub

