Attribute VB_Name = "mdl_024_PCO_Moons"
'
' PowerPoint 2010 VBA Macro
' Porsche CO
' Copyright PaCE Graphic GbR
' September 2015
'

Sub StatusMoons()

' Variablendefinition
'
    Dim Selection As Object                 ' Variable für gesamte Auswahl
    Dim SelectedObject(1 To 250) As Object  ' Variable für Objekte
    Dim intCounter As Integer               ' Zähler für Objekte
    Dim ExistingMoon As Boolean             ' Merker für max. Anzahl Objekte
    Dim ViewVersion As String               ' Variable für aktuelle Ansicht
    Dim CurrentSlide As Integer
    Dim MoonName As String                  ' Textvariable für Objekte
    
    
ErrorExit = False

CheckErrorsSlideSelection
If ErrorExit = True Then
    Exit Sub
End If
 
'
' Variablen-Startwertzuweisung
'
    intCounter = 1
    ExistingMoon = False
    ViewVersion = ActiveWindow.ViewType

        
'
' Überprüfung ob Objekte markiert
'
    On Error GoTo NoObject

'
' Einlesen/Anpassen markierte(s) Objekt(e)
'
    Set Selection = ActiveWindow.Selection
    With Selection
        Set SelectedObject(intCounter) = .ShapeRange(intCounter)
        While ExistingMoon = False
            MoonName = SelectedObject(intCounter).Name
            If MoonName Like "Moon*" Then
                ExistingMoon = True
                intCounter = intCounter + 1
                On Error GoTo MaxObject
                Set SelectedObject(intCounter) = .ShapeRange(intCounter)
            Else
                intCounter = intCounter + 1
                On Error GoTo MaxObject
                Set SelectedObject(intCounter) = .ShapeRange(intCounter)
            End If
        Wend
    End With
'
' Aufrufen der entsprechenden Auswahl
'
    ActiveWindow.ViewType = ViewVersion
    If ExistingMoon = True Then
        MoonChange.Show
    End If
Exit Sub

'
' Fehlerroutinen
'
' Keine Objekte markiert
NoObject:
    ActiveWindow.ViewType = ViewVersion
    MoonBuild.Show
Exit Sub
    
' Letztes Objekt erreicht und keine Ampel in der Auswahl
MaxObject:
    ActiveWindow.ViewType = ViewVersion
    If ExistingMoon = True Then
        MoonChange.Show
    ElseIf ExistingMoon = False Then
        MsgBox "Das Tool ist nur zum Bearbeiten bestehender bzw. Erstellen neuer Harvey Balls.", vbInformation, "Falsche Auswahl!"
        MoonBuild.Show
    End If
Exit Sub


End Sub
