Attribute VB_Name = "mdl_123_Magic_Grid"
Option Explicit

Public Sub Auswahl_auf_Raster_Var()
    On Error Resume Next
    Dim Eingabe As Variant
    Dim Raster As Single
    
    
    'Eingabe = InputBox("Bitte Raster in mm eingeben", "Am Raster ausrichten", "2")
    
    'If Eingabe <> "" Then
          
        'Raster = Val(Eingabe)
      
        'If Raster < 0 Or Raster > 25 Then
        '    Raster = 2
        'End If
        
        'Elemente_auf_Raster ActiveWindow.Selection.ShapeRange, Raster
       
    'End If
    
    Elemente_auf_Raster ActiveWindow.Selection.ShapeRange, ActivePresentation.GridDistance * pt2mm
    
End Sub

Public Sub Elemente_auf_Raster(Elemente As ShapeRange, Raster As Single)

    Dim bar As Variant
    Dim xu, yu As Single
    
    xu = 0.5 * ActiveWindow.Presentation.PageSetup.SlideWidth
    yu = 0.5 * ActiveWindow.Presentation.PageSetup.SlideHeight
    
    For Each bar In Elemente
        
        bar.Left = xu + Runden_auf_Wert(bar.Left - xu, Raster)
        bar.Top = yu + Runden_auf_Wert(bar.Top - yu, Raster)
        bar.Width = Runden_auf_Wert(bar.Width, Raster)
        bar.Height = Runden_auf_Wert(bar.Height, Raster)
        
    Next bar

End Sub

Public Function Runden_auf_Wert(Unrund As Single, Wert As Single) As Single
    
    Dim foo As Single
    foo = Unrund * pt2mm ' Points in mm
    foo = Wert * Round(foo / Wert, 0) ' Auf 2mm runden
    Runden_auf_Wert = foo * mm2pt

End Function
