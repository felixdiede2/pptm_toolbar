VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MoonChange 
   Caption         =   "Harvey Balls ändern ..."
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   OleObjectBlob   =   "MoonChange.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MoonChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'
' PowerPoint 2010/2013 VBA Macro
' Porsche CO
' Copyright PaCE Graphic GbR
' September 2015
' April 2016
'

    '
    'Variablen vordefinieren
    '
    ' Allgemeine Variablen
    Dim ViewVersion As String               ' Aktuelle Ansicht speichern
    
    ' Variablen für Mondform ändern
    Dim MondAuswahl As Object               ' Variable für gewählte Objekte
    Dim AuswahlObjekt(1 To 1000) As Object  ' Variable für Objekte aus Auswahl
    Dim CirCou As Variant                   ' Kreissegment-Größe
    Dim ArcWidth As Variant                 ' Kreissegment-Breite
    Dim i As Integer                        ' prozent-kreis-zähler
    Dim MaxMondAuswahlZähler As Integer     ' Zähler für verbleibende Objecte
    
    ' Variablen für Mondnamen-Routine
    Dim GesamtAuswahl As Object             ' Variable für gesamte Auswahl
    Dim ZählerBenutzt As Boolean            ' Flag für vorhandenen Namen
    Dim ObjektZähler As Integer             ' Objektzähler
    Dim MondName As String                  ' Namenstest
    Dim MondNameBuchstaben As TextRange
    Dim MondZähler As Integer               ' Mond-Variable
    


' **********************************************************************
' *************************  Start Definition  *************************
' **********************************************************************
'
' Initalize Starting Values
'
Private Sub UserForm_Initialize()
    
    ViewVersion = ActiveWindow.ViewType
    ObjektZähler = 1
    OptionHalb.Value = True
    MenuImaging
    
End Sub



' **********************************************************************
' ***************************  Input changes  **************************
' **********************************************************************

' Klick auf Beschriftung
'
Private Sub LabelLeer_Click()
    OptionLeer.Value = True
End Sub

Private Sub LabelEinViertel_Click()
    OptionEinViertel.Value = True
End Sub

Private Sub LabelHalb_Click()
    OptionHalb.Value = True
End Sub

Private Sub LabelDreiViertel_Click()
    OptionDreiViertel.Value = True
End Sub

Private Sub LabelVoll_Click()
    OptionVoll.Value = True
End Sub

Private Sub LabelProzent_Click()
    OptionProzent.Value = True
End Sub


' Änderung/Klick auf Eingabe ******************************************
'
Private Sub TextBoxProzent_Enter()
    OptionProzent.Value = True
End Sub

Private Sub TextBoxProzent_Change()
    'Wert außerhalb zulässigem Bereich
    If Val(TextBoxProzent.Text) < 1 Or Val(TextBoxProzent.Text) > 100 Then
        TextBoxProzent.Text = "50"
        MsgBox "Dieses Tool unterstützt nur Zahleneingaben zwischen 1 und 100, für Serien nur Zahleneingaben zwischen 4 und 50!", vbInformation
    End If
End Sub

' Änderung/Klick auf Option
'
Private Sub OptionLeer_Change()
    If OptionLeer.Value = True Then
        MenuImaging
    End If
End Sub

Private Sub OptionEinViertel_Change()
    If OptionEinViertel.Value = True Then
        MenuImaging
    End If
End Sub

Private Sub OptionHalb_Change()
    If OptionHalb.Value = True Then
        MenuImaging
    End If
End Sub

Private Sub OptionDreiViertel_Change()
    If OptionDreiViertel.Value = True Then
        MenuImaging
    End If
End Sub

Private Sub OptionVoll_Change()
    If OptionVoll.Value = True Then
        MenuImaging
    End If
End Sub

Private Sub OptionProzent_Change()
    If OptionProzent.Value = True Then
        MenuImaging
    End If
End Sub
    
    
' **********************************************************************
' ****************************  Mainprogram  ***************************
' **********************************************************************
    
Private Sub MoonChangeStart_Click()

    Set MondAuswahl = ActiveWindow.Selection.ShapeRange
    MondZähler = 1
'    CheckMondAuswahl
    MondAuswahl.Select
    With ActiveWindow.Selection
        MaxMondAuswahlZähler = MondAuswahl.Count
        While MaxMondAuswahlZähler > 0
            Set AuswahlObjekt(MaxMondAuswahlZähler) = MondAuswahl(MaxMondAuswahlZähler)
            MondName = AuswahlObjekt(MaxMondAuswahlZähler).Name
            If MondName Like "Moon*" Then
                If OptionLeer.Value = True Then
                    ' Segment ändern
                    CirCou = 1 / 4
                    KreisSegment
                ElseIf OptionEinViertel.Value = True Then
                    ' Segment ändern
                    CirCou = 1 / 4
                    KreisSegment
                ElseIf OptionHalb.Value = True Then
                    ' Segment ändern
                    CirCou = 1 / 2
                    KreisSegment
                ElseIf OptionDreiViertel.Value = True Then
                    ' Segment ändern
                    CirCou = 3 / 4
                    KreisSegment
                ElseIf OptionVoll.Value = True Then
                    ' Segment ändern
                    CirCou = 4 / 4
                    KreisSegment
                ElseIf OptionProzent.Value = True Then
                    ' Segment ändern
                    CirCou = (Val(TextBoxProzent.Text) / 100)
                    KreisSegment
                End If
                BasisKreis
            End If
            MaxMondAuswahlZähler = MaxMondAuswahlZähler - 1
        Wend
    End With
        
    ActiveWindow.Selection.Unselect
    ActiveWindow.ViewType = ViewVersion
    MoonChange.Hide
    Unload Me
    
End Sub
    
    
' **********************************************************************
' ***************************  Cancel program  *************************
' **********************************************************************

Private Sub MoonChangeCancel_Click()
    ActiveWindow.ViewType = ViewVersion
    MoonChange.Hide
    Unload Me
End Sub


'
' Mainprogram Routinen
'
'Private Sub GetMoonNumber()
'    ActiveWindow.Selection.SlideRange.Shapes.SelectAll
'    Set GesamtAuswahl = ActiveWindow.Selection
'    With GesamtAuswahl
'        ZählerBenutzt = True
'        While ZählerBenutzt = True
'            ZählerBenutzt = False
'            For ObjektZähler = 1 To GesamtAuswahl.ShapeRange.Count
'                Set AuswahlObjekt(ObjektZähler) = .ShapeRange(ObjektZähler)
'                MondName = AuswahlObjekt(ObjektZähler).Name
'                If MondName = "Moon" & (MondZähler) Then
'                    ZählerBenutzt = True
'                End If
'            Next
'            If ZählerBenutzt = True Then
'                MondZähler = MondZähler + 1
'            End If
'        Wend
'    End With
'End Sub


'Private Sub CheckMondAuswahl()
'    For i = 1 To MondAuswahl.Count
'        For j = i + 1 To MondAuswahl.Count
'            If MondAuswahl(i).Name = MondAuswahl(j).Name Then
'                GetMoonNumber
'                Oldnumber = CVar(Mid(MondAuswahl(i).Name, 5, (Len(MondName) - 4)))
'                MondAuswahl(i).Name = "Moon" & MondZähler
'                With ActiveWindow.Selection.SlideRange.Shapes("Moon" & MondZähler)
'                    .Select
'                    .GroupItems("MoonArc" & Oldnumber).Name = "MoonArc" & MondZähler
'                    .GroupItems("MoonFrame" & Oldnumber).Name = "MoonFrame" & MondZähler
'                End With
'            End If
'        Next
'    Next
'End Sub


Private Sub GetCurrentMoonNumber()
    MondZähler = CVar(Mid(MondName, 5, (Len(MondName) - 4)))
End Sub


Private Sub KreisSegment()
    
    MondAuswahl(MaxMondAuswahlZähler).GroupItems.Item(2).Adjustments.Item(2) = ((CirCou * 359.99) - 90)
    If OptionLeer.Value = True Then
        With MondAuswahl(MaxMondAuswahlZähler).GroupItems.Item(2)
            .Fill.ForeColor.RGB = RGB(255, 255, 255) 'White - Background color
            .Line.Visible = msoTrue
            .Line.ForeColor.RGB = RGB(0, 0, 0) 'Black - Accent 3 color
            .Line.Weight = 0.75
        End With
    Else
        With MondAuswahl(MaxMondAuswahlZähler).GroupItems.Item(2)
            .Fill.Visible = msoTrue
            .Fill.ForeColor.RGB = RGB(0, 0, 0) 'Black
            '.Fill.ForeColor.RGB = RGB(55, 55, 55) 'Dark Grey
            .Fill.Solid
            .Fill.Transparency = 0#
            .Line.Visible = msoTrue
            .Line.ForeColor.RGB = RGB(0, 0, 0) 'Black
            .Line.Weight = 0.75
        End With
    End If
    
End Sub

Private Sub BasisKreis()

    With MondAuswahl(MaxMondAuswahlZähler).GroupItems.Item(1)
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.Transparency = 0#
        .Line.Weight = 0.75
        .Fill.ForeColor.RGB = RGB(255, 255, 255) 'White - Background color
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(0, 0, 0) 'Black
    End With
    
End Sub


'
' Image-Routine
'
Private Sub MenuImaging()
    
    ' all images off
    '
    WhiteImageLeer.Visible = False
    WhiteImageEinViertel.Visible = False
    WhiteImageHalb.Visible = False
    WhiteImageDreiViertel.Visible = False
    WhiteImageVoll.Visible = False
    WhiteImageProzent.Visible = False
'
    If OptionEinViertel.Value = True Then
        WhiteImageEinViertel.Visible = True
    ElseIf OptionLeer.Value = True Then
        WhiteImageLeer.Visible = True
    ElseIf OptionHalb.Value = True Then
        WhiteImageHalb.Visible = True
    ElseIf OptionDreiViertel.Value = True Then
        WhiteImageDreiViertel.Visible = True
    ElseIf OptionVoll.Value = True Then
        WhiteImageVoll.Visible = True
    Else
        WhiteImageProzent.Visible = True
    End If
    
End Sub
