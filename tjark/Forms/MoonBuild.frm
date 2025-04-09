VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MoonBuild 
   Caption         =   "Harvey Balls erstellen ..."
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "MoonBuild.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MoonBuild"
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
    
    ' Variablen für Mondgenerierung
    Dim GesamtAuswahl As Object             ' Variable für gesamte Auswahl
    Dim AuswahlObjekt(1 To 1000) As Object  ' Variable für Objekte
    Dim MondName As String                  ' Namenstest
    Dim ObjektZähler As Integer             ' Objektzähler
    Dim MondZähler As Integer               ' Mond-Variable
    Dim ZählerBenutzt As Boolean            ' Flag für vorhandenen Namen
    Dim Ypos As Variant                     ' Start Position für Moon
    Dim CirCou As Variant                   ' Kreissegment-Größe
    Dim ArcWidth As Variant                 ' Kreissegment-Breite
    Dim i As Integer                        ' prozent-kreis-zähler




' **********************************************************************
' *************************  Start Definition  *************************
' **********************************************************************
'
' Initalize Starting Values
'
Private Sub UserForm_Initialize()
    
    ViewVersion = ActiveWindow.ViewType
    ObjektZähler = 1
    CheckBoxHalb.Value = True
    MenuImaging
    
End Sub


' **********************************************************************
' ***************************  Input changes  **************************
' **********************************************************************


' Klick auf Beschriftung ***********************************************
'
' Mond-Teile
'
Private Sub LabelLeer_Click()
    If CheckBoxLeer.Value = True Then
        CheckBoxLeer.Value = False
    ElseIf CheckBoxLeer.Value = False Then
        CheckBoxLeer.Value = True
    End If
End Sub

Private Sub LabelEinViertel_Click()
    If CheckBoxEinViertel.Value = True Then
        CheckBoxEinViertel.Value = False
    ElseIf CheckBoxEinViertel.Value = False Then
        CheckBoxEinViertel.Value = True
    End If
End Sub

Private Sub LabelHalb_Click()
    If CheckBoxHalb.Value = True Then
        CheckBoxHalb.Value = False
    ElseIf CheckBoxHalb.Value = False Then
        CheckBoxHalb.Value = True
    End If
End Sub

Private Sub LabelDreiViertel_Click()
    If CheckBoxDreiViertel.Value = True Then
        CheckBoxDreiViertel.Value = False
    ElseIf CheckBoxDreiViertel.Value = False Then
        CheckBoxDreiViertel.Value = True
    End If
End Sub

Private Sub LabelVoll_Click()
    If CheckBoxVoll.Value = True Then
        CheckBoxVoll.Value = False
    ElseIf CheckBoxVoll.Value = False Then
        CheckBoxVoll.Value = True
    End If
End Sub

' Mond in Prozent
'
Private Sub LabelProzent_Click()
    If CheckBoxProzent.Value = True Then
        CheckBoxProzent.Value = False
    ElseIf CheckBoxProzent.Value = False Then
        CheckBoxProzent.Value = True
    End If
End Sub

' Serie an/aus
'
Private Sub LabelSerie_Click()
    If CheckBoxSerie.Value = True Then
        CheckBoxSerie.Value = False
    ElseIf CheckBoxSerie.Value = False Then
        CheckBoxSerie.Value = True
    End If
End Sub


' Änderung/Klick auf Eingabe ******************************************
'
Private Sub TextBoxProzent_Enter()
    CheckBoxProzent.Value = True
End Sub

Private Sub TextBoxProzent_Change()
    'Wert außerhalb zulässigem Bereich
    If Val(TextBoxProzent.Text) < 1 Or Val(TextBoxProzent.Text) > 100 Then
        TextBoxProzent.Text = "50"
        MsgBox "Dieses Tool unterstützt nur Zahleneingaben zwischen 1 und 100, für Serien nur Zahleneingaben zwischen 4 und 50!", vbInformation
    End If
End Sub


' Änderung/Klick auf Checkbox ******************************************
'
' Mond-Teile
'
Private Sub CheckBoxLeer_Change()
    If CheckBoxLeer.Value = True Then
        CheckBoxProzent.Value = False
        If CheckBoxSerie.Value = True Then
            PartsOn
        End If
        SerieCheck
    ElseIf CheckBoxLeer.Value = False And CheckBoxProzent.Value = False Then
        CheckBoxSerie.Value = False
    End If
    TestAllOff
    MenuImaging
End Sub

Private Sub CheckBoxEinViertel_Change()
    If CheckBoxEinViertel.Value = True Then
        CheckBoxProzent.Value = False
        If CheckBoxSerie.Value = True Then
            PartsOn
        End If
        SerieCheck
    ElseIf CheckBoxEinViertel.Value = False And CheckBoxProzent.Value = False Then
        CheckBoxSerie.Value = False
    End If
    TestAllOff
    MenuImaging
End Sub

Private Sub CheckBoxHalb_Change()
    If CheckBoxHalb.Value = True Then
        CheckBoxProzent.Value = False
        If CheckBoxSerie.Value = True Then
            PartsOn
        End If
        SerieCheck
    ElseIf CheckBoxHalb.Value = False And CheckBoxProzent.Value = False Then
        CheckBoxSerie.Value = False
    End If
    TestAllOff
    MenuImaging
End Sub

Private Sub CheckBoxDreiViertel_Change()
    If CheckBoxDreiViertel.Value = True Then
        CheckBoxProzent.Value = False
        If CheckBoxSerie.Value = True Then
            PartsOn
        End If
        SerieCheck
    ElseIf CheckBoxDreiViertel.Value = False And CheckBoxProzent.Value = False Then
        CheckBoxSerie.Value = False
    End If
    TestAllOff
    MenuImaging
End Sub

Private Sub CheckBoxVoll_Change()
    If CheckBoxVoll.Value = True Then
        CheckBoxProzent.Value = False
        If CheckBoxSerie.Value = True Then
            PartsOn
        End If
        SerieCheck
    ElseIf CheckBoxVoll.Value = False And CheckBoxProzent.Value = False Then
        CheckBoxSerie.Value = False
    End If
    TestAllOff
    MenuImaging
End Sub

' Mond in Prozent
'
Private Sub CheckBoxProzent_Change()
    If CheckBoxProzent.Value = True Then
        PartsOff
    ElseIf CheckBoxProzent.Value = False Then
        If CheckBoxLeer.Value = False And CheckBoxEinViertel.Value = False And _
        CheckBoxHalb.Value = False And CheckBoxDreiViertel.Value = False And _
        CheckBoxVoll.Value = False Then
            CheckBoxHalb.Value = True
            CheckBoxSerie.Value = False
        End If
    End If
    MenuImaging
End Sub

' Serie an/aus
'
Private Sub CheckBoxSerie_Change()
    If CheckBoxSerie.Value = True Then
        If CheckBoxLeer.Value = True Or CheckBoxEinViertel.Value = True Or _
        CheckBoxHalb.Value = True Or CheckBoxDreiViertel.Value = True Or _
        CheckBoxVoll.Value = True Then
            CheckBoxLeer.Value = True
            CheckBoxEinViertel.Value = True
            CheckBoxHalb.Value = True
            CheckBoxDreiViertel.Value = True
            CheckBoxVoll.Value = True
        End If
    ElseIf CheckBoxSerie.Value = False And CheckBoxProzent.Value = False And _
    CheckBoxLeer.Value = True And CheckBoxEinViertel.Value = True And _
    CheckBoxHalb.Value = True And CheckBoxDreiViertel.Value = True And _
    CheckBoxVoll.Value = True Then
        CheckBoxLeer.Value = False
        CheckBoxEinViertel.Value = False
        CheckBoxHalb.Value = True
        CheckBoxDreiViertel.Value = False
        CheckBoxVoll.Value = False
    End If
    MenuImaging
End Sub



    
' **********************************************************************
' ****************************  Mainprogram  ***************************
' **********************************************************************
    
Private Sub MoonBuildStart_Click()
    
    MondZähler = 1
    Ypos = 165
    If CheckBoxLeer.Value = True Then
        'feststellen welche Mondnamen benutzt
        GetMoonNumber
        ' BasisKreis
        BasisKreis
        ' Segment erstellen
        CirCou = 1 / 4
        KreisSegment
        ' Segment unsichtbar schalten
        ActiveWindow.Selection.SlideRange.Shapes("MoonArc" & MondZähler).Select
        With ActiveWindow.Selection.ShapeRange
            .Fill.Visible = msoFalse
            .Line.Visible = msoFalse
        End With
        ' Benennung Moon
        BenennungMond
    End If
    If CheckBoxEinViertel.Value = True Then
        'feststellen welche Mondnamen benutzt
        GetMoonNumber
        ' BasisKreis
        BasisKreis
        ' Segment erstellen
        CirCou = 1 / 4
        KreisSegment
        ' Benennung Moon
        BenennungMond
    End If
    If CheckBoxHalb.Value = True Then
        'feststellen welche Mondnamen benutzt
        GetMoonNumber
        ' BasisKreis
        BasisKreis
        ' Segment erstellen
        CirCou = 1 / 2
        KreisSegment
        ' Benennung Moon
        BenennungMond
    End If
    If CheckBoxDreiViertel.Value = True Then
        'feststellen welche Mondnamen benutzt
        GetMoonNumber
        ' BasisKreis
        BasisKreis
        ' Segment erstellen
        CirCou = 3 / 4
        KreisSegment
        ' Benennung Moon
        BenennungMond
    End If
    If CheckBoxVoll.Value = True Then
        'feststellen welche Mondnamen benutzt
        GetMoonNumber
        ' BasisKreis
        BasisKreis
        ' Segment erstellen
        CirCou = 4 / 4
        KreisSegment
        ' Benennung Moon
        BenennungMond
    End If
    If CheckBoxProzent.Value = True Then
        If CheckBoxSerie.Value = True And Val(TextBoxProzent.Text) > 50 Then
            CheckBoxSerie.Value = False
            MsgBox "Für Prozentwerte, die größer 50 sind, ist die Serienoption nicht möglich!", vbInformation
        ElseIf CheckBoxSerie.Value = True And Val(TextBoxProzent.Text) < 4 Then
            CheckBoxSerie.Value = False
            MsgBox "Für Prozentwerte, die kleiner 4 sind, ist die Serienoption nicht möglich!", vbInformation
        End If
        If CheckBoxSerie.Value = True Then
            'feststellen welche Mondnamen benutzt
            GetMoonNumber
            ' BasisKreis
            BasisKreis
            ' Segment erstellen
            CirCou = 1 / 4
            KreisSegment
            ' Segment unsichtbar schalten
            ActiveWindow.Selection.SlideRange.Shapes("MoonArc" & MondZähler).Select
            With ActiveWindow.Selection.ShapeRange
                .Fill.Visible = msoFalse
                .Line.Visible = msoFalse
            End With
            ' Benennung Moon
            BenennungMond
            For i = 1 To (100 / Val(TextBoxProzent.Text))
                'feststellen welche Mondnamen benutzt
                GetMoonNumber
                ' BasisKreis
                BasisKreis
                ' Segment erstellen
                CirCou = (Val(TextBoxProzent.Text) / 100 * i)
                If CirCou > 1 Then
                    CirCou = 4 / 4
                End If
                If (1 - (CirCou)) < (Val(TextBoxProzent.Text) / 100 * 0.5) Then
                    CirCou = 4 / 4
                End If
                KreisSegment
                ' Benennung Moon
                BenennungMond
            Next
        ElseIf CheckBoxSerie.Value = False Then
            'feststellen welche Mondnamen benutzt
            GetMoonNumber
            ' BasisKreis
            BasisKreis
            ' Segment erstellen
            CirCou = (Val(TextBoxProzent.Text) / 100)
            KreisSegment
            ' Benennung Moon
            BenennungMond
        End If
    End If
    ActiveWindow.Selection.Unselect
    ActiveWindow.ViewType = ViewVersion
    MoonBuild.Hide
    Unload Me
    
End Sub
    
    
' **********************************************************************
' ***************************  Cancel program  *************************
' **********************************************************************

Private Sub MoonBuildCancel_Click()
    ActiveWindow.ViewType = ViewVersion
    MoonBuild.Hide
    Unload Me
End Sub

' **********************************************************************
' ******************************  Routinen  ****************************
' **********************************************************************

Private Sub PartsOn()                      ' *** Option alle Parts an
    CheckBoxLeer.Value = True
    CheckBoxEinViertel.Value = True
    CheckBoxHalb.Value = True
    CheckBoxDreiViertel.Value = True
    CheckBoxVoll.Value = True
End Sub


Private Sub PartsOff()                      ' *** Option Parts aus
    CheckBoxLeer.Value = False
    CheckBoxEinViertel.Value = False
    CheckBoxHalb.Value = False
    CheckBoxDreiViertel.Value = False
    CheckBoxVoll.Value = False
End Sub

Private Sub SerieCheck()                    ' *** Option Serie Komlett
    If CheckBoxLeer.Value = True And CheckBoxEinViertel.Value = True And _
    CheckBoxHalb.Value = True And CheckBoxDreiViertel.Value = True And _
    CheckBoxVoll.Value = True Then
        CheckBoxSerie.Value = True
    End If
End Sub

Private Sub TestAllOff()
    If CheckBoxLeer.Value = False And CheckBoxEinViertel.Value = False And _
    CheckBoxHalb.Value = False And CheckBoxDreiViertel.Value = False And _
    CheckBoxVoll.Value = False And CheckBoxProzent.Value = False Then
        CheckBoxLeer.Value = False
        CheckBoxEinViertel.Value = False
        CheckBoxHalb.Value = True
        CheckBoxDreiViertel.Value = False
        CheckBoxVoll.Value = False
    End If
End Sub

'
' Mainprogram Routinen
'
Private Sub GetMoonNumber()

    If ActiveWindow.Selection.SlideRange.Shapes.Count > 0 Then
        ActiveWindow.Selection.SlideRange.Shapes.SelectAll
        Set GesamtAuswahl = ActiveWindow.Selection
        With GesamtAuswahl
            ZählerBenutzt = True
            While ZählerBenutzt = True
                ZählerBenutzt = False
                For ObjektZähler = 1 To GesamtAuswahl.ShapeRange.Count
                    Set AuswahlObjekt(ObjektZähler) = .ShapeRange(ObjektZähler)
                    MondName = AuswahlObjekt(ObjektZähler).Name
                    If MondName = "Moon" & (MondZähler) Then
                        ZählerBenutzt = True
                    End If
                Next
                If ZählerBenutzt = True Then
                    MondZähler = MondZähler + 1
                End If
            Wend
        End With
    End If
End Sub

Private Sub BasisKreis()
    ActiveWindow.Selection.SlideRange.Shapes.AddShape(msoShapeOval, 150, Ypos, 34, 34).Select
    ActiveWindow.Selection.ShapeRange.Name = "MoonFrame" & MondZähler
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.Transparency = 0#
        .Fill.ForeColor.RGB = RGB(255, 255, 255) 'White - Background color
        .Line.Weight = 0.75
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(0, 0, 0) 'Black
    End With
End Sub

Private Sub KreisSegment()
    ArcWidth = (ActiveWindow.Selection.SlideRange.Shapes("MoonFrame" & MondZähler).Width / 2)
    ActiveWindow.Selection.SlideRange.Shapes.AddShape(msoShapeArc, (150 + ArcWidth), Ypos, 17, 17).Select
    ActiveWindow.Selection.ShapeRange.Name = "MoonArc" & MondZähler
    With ActiveWindow.Selection.ShapeRange
        .Fill.Visible = msoTrue
        '.Fill.ForeColor.RGB = RGB(55, 55, 55) 'Dark Grey
        .Fill.ForeColor.RGB = RGB(0, 0, 0) 'Black
        .Fill.Solid
        .Fill.Transparency = 0#
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(0, 0, 0) 'Black
        .Line.Weight = 0.75
    End With
    ActiveWindow.Selection.ShapeRange.Adjustments.Item(2) = ((CirCou * 359.99) - 90)
    Ypos = Ypos + 57
End Sub

Private Sub BenennungMond()
    ActiveWindow.Selection.SlideRange.Shapes.Range(Array("MoonFrame" & MondZähler, "MoonArc" & MondZähler)).Select
    ActiveWindow.Selection.ShapeRange.Group.Name = "Moon" & MondZähler
    ActiveWindow.Selection.SlideRange.Shapes("Moon" & MondZähler).Select
    ActiveWindow.Selection.ShapeRange.LockAspectRatio = msoTrue
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
    If CheckBoxEinViertel.Value = True Then
        WhiteImageEinViertel.Visible = True
    ElseIf CheckBoxLeer.Value = True Then
        WhiteImageLeer.Visible = True
    ElseIf CheckBoxHalb.Value = True Then
        WhiteImageHalb.Visible = True
    ElseIf CheckBoxDreiViertel.Value = True Then
        WhiteImageDreiViertel.Visible = True
    ElseIf CheckBoxVoll.Value = True Then
        WhiteImageVoll.Visible = True
    Else
        WhiteImageProzent.Visible = True
    End If
        
End Sub
