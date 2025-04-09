Attribute VB_Name = "mdl_002_CO_DocumentIdentifier"
' PPT2010 Macro
' Programmed by Thomas Breuer
' Copyright PaCE Graphic GbR
' Germany - April 2012
'


Sub CO_DocID() 'ByVal control As IRibbonControl)
    ' **********************************************************************
    ' Variables
    '
    Dim ViewVersion As String 'Saving current View Type
    Dim FooterLeft, FooterTop, FooterWidth As String
    Dim FooterPHID, PHCounter As Integer
    
    
    ' **********************************************************************
    ' Abfangen keine Pr�sentation offen
    '
    If Application.Presentations.Count < 1 Then
        MsgBox "Keine ge�ffnete Pr�sentation! Bitte Pr�sentation �ffnen und Tool neu starten.", vbInformation, "Keine offene Pr�sentation"
        Exit Sub
    End If


    ' **********************************************************************
    ' Aktuelle Ansicht merken
    '
    ViewVersion = ActiveWindow.ViewType

' **********************************************************************
' Abfangen Vorlagen mit gel�schtem Footer �ber invertierte Fehlermeldung
'
    On Error GoTo ExistingFooter
    ActiveWindow.ViewType = ppViewSlideMaster
    ActivePresentation.SlideMaster.Shapes.AddPlaceholder Type:=ppPlaceholderFooter

' **********************************************************************
' Abfangen unbekannter Fehler
'
    On Error GoTo GeneralError
    For PHCounter = 1 To ActivePresentation.SlideMaster.Shapes.Placeholders.Count
        ActivePresentation.SlideMaster.Shapes.Placeholders(PHCounter).Select
        If ActiveWindow.Selection.ShapeRange.PlaceholderFormat.Type = ppPlaceholderFooter Then
            FooterPHID = PHCounter
        End If
    Next
    
    ActivePresentation.SlideMaster.Shapes.Placeholders(FooterPHID).Select
    ActiveWindow.Selection.ShapeRange.Delete
    
    ActiveWindow.ViewType = ViewVersion
    MsgBox "Kein Platzhalter f�r eine Fu�note in dieser Pr�sentation vorhanden!", vbInformation, "Defektes Template"
    
Exit Sub

' **********************************************************************
' **************************  Error Routines  **************************
' **********************************************************************

' **********************************************************************
' Fu�zeile vorhanden, Eingabemaske aufrufen
'
ExistingFooter:
    ActiveWindow.ViewType = ViewVersion
    CODocID.Show
Exit Sub

' **********************************************************************
' Unerwarteter Fehler
'
GeneralError:
    ActiveWindow.ViewType = ViewVersion
    MsgBox "Undefinierbarer Fehler! Tool gestoppt.", vbInformation
Exit Sub


End Sub

