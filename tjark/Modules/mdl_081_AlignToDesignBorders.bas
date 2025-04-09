Attribute VB_Name = "mdl_081_AlignToDesignBorders"
Option Explicit

'Input Lars Lewandowitz

Dim seitenbreite As Single
Dim seitenhoehe As Single
Dim anzahlElemente As Single

Sub AusrichtenGestaltungsrahmenOben()
    On Error Resume Next
    
    seitenbreite = ActiveWindow.Presentation.PageSetup.SlideWidth
    seitenhoehe = ActiveWindow.Presentation.PageSetup.SlideHeight
    
    ActiveWindow.Selection.ShapeRange.Top = (0.5 * seitenhoehe) - (5.6 * cm2pt)
End Sub

Sub AusrichtenGestaltungsrahmenUnten()
    On Error Resume Next
    
    Dim i As Integer
    seitenbreite = ActiveWindow.Presentation.PageSetup.SlideWidth
    seitenhoehe = ActiveWindow.Presentation.PageSetup.SlideHeight
    
    anzahlElemente = ActiveWindow.Selection.ShapeRange.Count
    
    For i = 1 To anzahlElemente
        ActiveWindow.Selection.ShapeRange(i).Top = (0.5 * seitenhoehe) + (7.3 * cm2pt) - ActiveWindow.Selection.ShapeRange(i).Height
    Next
                  
End Sub

Sub AusrichtenGestaltungsrahmenLinks()
    On Error Resume Next
    
    seitenbreite = ActiveWindow.Presentation.PageSetup.SlideWidth
    seitenhoehe = ActiveWindow.Presentation.PageSetup.SlideHeight
    
    ActiveWindow.Selection.ShapeRange.Left = (0.5 * seitenbreite) - (15.5 * cm2pt)
       
End Sub

Sub AusrichtenGestaltungsrahmenRechts()
    On Error Resume Next
    
    Dim i As Integer
    seitenbreite = ActiveWindow.Presentation.PageSetup.SlideWidth
    seitenhoehe = ActiveWindow.Presentation.PageSetup.SlideHeight
    
    anzahlElemente = ActiveWindow.Selection.ShapeRange.Count
    
    For i = 1 To anzahlElemente
        ActiveWindow.Selection.ShapeRange(i).Left = (0.5 * seitenbreite) + (15.5 * cm2pt) - ActiveWindow.Selection.ShapeRange(i).Width
    Next

End Sub

Sub StretchBreiteGestaltungsrahmen()
    On Error Resume Next
    
    Dim i As Integer
    seitenbreite = ActiveWindow.Presentation.PageSetup.SlideWidth
    anzahlElemente = ActiveWindow.Selection.ShapeRange.Count
    
    For i = 1 To anzahlElemente
        ActiveWindow.Selection.ShapeRange(i).Width = (2 * 15.5 * cm2pt)
    Next
    ActiveWindow.Selection.ShapeRange.Left = (0.5 * seitenbreite) - (15.5 * cm2pt)
End Sub

Sub StretchHoeheGestaltungsrahmen()
    On Error Resume Next
    Dim i As Integer
    
    seitenhoehe = ActiveWindow.Presentation.PageSetup.SlideHeight
    anzahlElemente = ActiveWindow.Selection.ShapeRange.Count
    
    For i = 1 To anzahlElemente
        ActiveWindow.Selection.ShapeRange(i).Height = ((5.6 + 7.3) * cm2pt)
    Next
    ActiveWindow.Selection.ShapeRange.Top = (0.5 * seitenhoehe) - (5.6 * cm2pt)

End Sub
