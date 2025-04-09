Attribute VB_Name = "mdl_127_Save_Version"
Option Explicit

Sub Save_as_Version()

Dim DateiName As String
Dim Kurzname As String
Dim Zeitstempel_alt As String
Dim Zeitstempel_neu As String
Dim Jetzt As String
Dim ZS_Ziffern As String
Dim ZS_Sonderzeichen As String
Dim Erweiterung As String
Dim PunktPos As Integer
Dim Monat, Tag, Stunde, now_Minute As String

'Prüfen, ob die Datei überhaupt schon einmal gespeichert wurde
DateiName = ActivePresentation.Name
'Debug.Print DateiName

PunktPos = InStrRev(DateiName, ".")

If ActivePresentation.Path = "" Or PunktPos < 1 Or PunktPos > Len(DateiName) Then
    MsgBox "Bitte speichern Sie diese neue Präsentation erstmalig ohne die Versionsfunkion. Anschließend können Versionen gespeichert werden.", vbOKOnly, "Neue Präsentation"
    Exit Sub
End If

'Prüfen, ob die Datei schon einmal mit Zeitstempel gespeichert wurde

If PunktPos < 14 Then
    'Kann gar kein gültiger Zeitstempel sein
    Zeitstempel_alt = ""
    Kurzname = Left(DateiName, PunktPos - 1)
Else
    'Zeitstempel auslesen
    Zeitstempel_alt = Mid(DateiName, PunktPos - 13, 13)
    'timestamp (_YYMMDD_HHhmm)
    ZS_Ziffern = Mid(Zeitstempel_alt, 2, 6) & Mid(Zeitstempel_alt, 9, 2) & Mid(Zeitstempel_alt, 12, 2)
    ZS_Sonderzeichen = Left(Zeitstempel_alt, 1) & Mid(Zeitstempel_alt, 8, 1) & Mid(Zeitstempel_alt, 11, 1)
    
    'Debug.Print "ZS_ZIFFERN:" & ZS_Ziffern
    'Debug.Print "ZS_SONDERZ:" & ZS_Sonderzeichen
    
    If IsNumeric(ZS_Ziffern) And ZS_Sonderzeichen = "__h" Then
        'Gültigen Zeitstempel gefunden
        Kurzname = Left(DateiName, PunktPos - 14)
    Else
        'Keinen gültigen Zeitstempel gefunden
        Zeitstempel_alt = ""
        Kurzname = Left(DateiName, PunktPos - 1)
    End If
End If

'Neuen Zeitstempel bauen
Monat = Month(Now())
If Len(Monat) = 1 Then
    Monat = "0" & Monat
End If

Tag = Day(Now())
If Len(Tag) = 1 Then
    Tag = "0" & Tag
End If

Stunde = Hour(Now())
If Len(Stunde) = 1 Then
    Stunde = "0" & Stunde
End If

now_Minute = Minute(Now())
If Len(now_Minute) = 1 Then
    now_Minute = "0" & now_Minute
End If

Zeitstempel_neu = "_" & Right(Year(Now()), 2) & Monat & Tag & "_" & Stunde & "h" & now_Minute

If Zeitstempel_neu = Zeitstempel_alt Then
    'Wurde eben erst gespeichert
    If MsgBox("Die Präsentation wurde zuletzt vor weniger als einer Minute gespeichert. Wollen Sie diese Version überschreiben?", vbExclamation + vbYesNo, "Save Version") = vbYes Then
        ActivePresentation.Save
        Exit Sub
    Else
        Exit Sub
    End If
End If

'Debug.Print ActivePresentation.Path
'Debug.Print "Punktpos: " & PunktPos
'Debug.Print "KN: " & Kurzname
'Debug.Print "ZS ALT: " & Zeitstempel_alt
'Debug.Print "ZS NEU: " & Zeitstempel_neu

Erweiterung = Right(DateiName, Len(DateiName) - PunktPos + 1)
Debug.Print "ERW: " & Erweiterung

DateiName = Kurzname & Zeitstempel_neu & Erweiterung
Debug.Print DateiName

If MsgBox("Speichern einer neue Version der aktuellen Präsentation als " & DateiName & " ?", vbExclamation + vbYesNo, "Save Version") = 6 Then
    'ActiveWorkbook.SaveAs DateiName
    ActivePresentation.SaveAs ActivePresentation.Path & "\" & DateiName
End If

End Sub


Sub AAA_Save_as_AddIn()
    On Error Resume Next
    Dim Name As String
    
    'Nur zum internen Aufruf
    'Speichert Toolbar als PPAM-Addin
        
    Name = ActivePresentation.Path & "\" & PPAFilename
    
    ActivePresentation.SaveCopyAs Name, ppSaveAsOpenXMLAddin
End Sub
