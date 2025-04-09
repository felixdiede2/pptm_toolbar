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
    ' Abfangen keine Präsentation offen
    '
    If Application.Presentations.Count < 1 Then
        MsgBox "Keine geöffnete Präsentation! Bitte Präsentation öffnen und Tool neu starten.", vbInformation, "Keine offene Präsentation"
        Exit Sub
    End If


    ' **********************************************************************
    ' Aktuelle Ansicht merken
    '
    ViewVersion = ActiveWindow.ViewType

' **********************************************************************
' Abfangen Vorlagen mit gelöschtem Footer über invertierte Fehlermeldung
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
    MsgBox "Kein Platzhalter für eine Fußnote in dieser Präsentation vorhanden!", vbInformation, "Defektes Template"
    
Exit Sub

' **********************************************************************
' **************************  Error Routines  **************************
' **********************************************************************

' **********************************************************************
' Fußzeile vorhanden, Eingabemaske aufrufen
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

