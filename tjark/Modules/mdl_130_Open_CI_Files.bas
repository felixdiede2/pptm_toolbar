Attribute VB_Name = "mdl_130_Open_CI_Files"
Option Explicit

' Prozeduren zum Öffnen von Dateien
' ---------------------------------
' Angewendet auf Chartbibliothek und Template
' Öffnen der Datei vom Netzlaufwerk
' Bei Laufzeitfehler: Einblenden einer Fehlermeldung und Öffnen der Datei aus dem Programmverzeichnis

Dim Meldung As String
Dim Ergebnis As Integer
Dim Datei As String
Dim Modus As Integer
Dim sMldg As String

Sub OC_openfile_Manual()
    'Wird immer lokal geladen
    On Error Resume Next

    Datei = GetSetting(PPAName, "Setup", "LocalManual")
    
    If Datei = "" Then
        Meldung = "Verweis nicht gefunden." & Chr(10) & Chr(10) & "Bitte Toolbar erneut installieren."
        Ergebnis = MsgBox(Meldung, Title:="Fehler")
        GoTo Ende
    Else
        On Error GoTo Lokaler_Fehler
        Presentations.Open FileName:=Datei, ReadOnly:=msoTrue
        GoTo Ende
    End If
    
Lokaler_Fehler:
    On Error Resume Next
    Meldung = "Die lokale Anleitung wurde nicht gefunden." & Chr(10) & Chr(10) & "Bitte Toolbar erneut installieren"
    Ergebnis = MsgBox(Meldung, Title:="Dateifehler")
    
Ende:
    
End Sub


Sub OC_openfile_Chartbib()
    '1. Anlauf: Online-Version laden - bei Fehler mit lokaler Version versuchen
    On Error Resume Next
    If GetSetting(PPAName, "Setup", "FileMode") = "Offline" Then GoTo Lokal_Laden
    If GetSetting(PPAName, "Setup", "NetBase") = "" Then GoTo Registry_Fehler
    If GetSetting(PPAName, "Setup", "NetChartBib") = "" Then GoTo Registry_Fehler
    Datei = GetSetting(PPAName, "Setup", "NetBase") & GetSetting(PPAName, "Setup", "NetChartBib")
    
    On Error GoTo Netz_Fehler
    Presentations.Open FileName:=Datei, ReadOnly:=msoTrue
    GoTo Ende 'Zum Ende wenn alles geklappt hat
    
Netz_Fehler:
    'Netz-Ladefehler aufgetreten
    On Error Resume Next
    Meldung = "Die Chartbibliothek wurde nicht gefunden." & Chr(10) & Chr(10) & _
                "Bitte überprüfen sie die Netzwerkverbindung." & Chr(10) & Chr(10) & _
                "Die lokale Version wird geöffnet, diese ist eventuell nicht aktuell."
    Ergebnis = MsgBox(Meldung, Title:="Netzwerkfehler")
    
Lokal_Laden:
    '2. Anlauf: Lokale Version laden
    On Error Resume Next
    If GetSetting(PPAName, "Setup", "LocalChartBib") = "" Then GoTo Registry_Fehler
    Datei = GetSetting(PPAName, "Setup", "LocalChartBib")
    
    On Error GoTo Lokaler_Fehler
    Presentations.Open FileName:=Datei, ReadOnly:=msoTrue
    GoTo Ende
    
Lokaler_Fehler:
    On Error Resume Next
    Meldung = "Die lokale Chartbibliothek wurde nicht gefunden." & Chr(10) & Chr(10) & "Bitte Toolbar erneut installieren."
    Ergebnis = MsgBox(Meldung, Title:="Dateifehler")
    GoTo Ende
    
Registry_Fehler:
    On Error Resume Next
    Meldung = "Verweis nicht gefunden." & Chr(10) & Chr(10) & "Bitte Toolbar erneut installieren."
    Ergebnis = MsgBox(Meldung, Title:="Fehler")
    
Ende:

End Sub

Sub OC_openfile_Template()
    '1. Anlauf: Online-Version laden - bei Fehler mit lokaler Version versuchen
    On Error Resume Next
    If GetSetting(PPAName, "Setup", "FileMode") = "Offline" Then GoTo Lokal_Laden
    If GetSetting(PPAName, "Setup", "NetBase") = "" Then GoTo Registry_Fehler
    If GetSetting(PPAName, "Setup", "NetMaster") = "" Then GoTo Registry_Fehler
    Datei = GetSetting(PPAName, "Setup", "NetBase") & GetSetting(PPAName, "Setup", "NetMaster")
    
    On Error GoTo Netz_Fehler
    Presentations.Open FileName:=Datei, ReadOnly:=msoTrue
    GoTo Ende 'Zum Ende wenn alles geklappt hat
    
Netz_Fehler:
    'Netz-Ladefehler aufgetreten
    Meldung = "Das Template wurde nicht gefunden." & Chr(10) & Chr(10) & _
                "Bitte überprüfen sie die Netzwerkverbindung." & Chr(10) & Chr(10) & _
                "Die lokale Version wird geöffnet, diese ist eventuell nicht aktuell."
    Ergebnis = MsgBox(Meldung, Title:="Netzwerkfehler")
    
Lokal_Laden:
    '2. Anlauf: Lokale Version laden
    On Error Resume Next
    If GetSetting(PPAName, "Setup", "LocalMaster") = "" Then GoTo Registry_Fehler
    Datei = GetSetting(PPAName, "Setup", "LocalMaster")
    
    On Error GoTo Lokaler_Fehler
    Presentations.Open FileName:=Datei, ReadOnly:=msoTrue
    GoTo Ende
    
Lokaler_Fehler:
    On Error Resume Next
    Meldung = "Das lokale Template wurde nicht gefunden." & Chr(10) & Chr(10) & "Bitte Toolbar erneut installieren."
    Ergebnis = MsgBox(Meldung, Title:="Dateifehler")
    GoTo Ende
    
Registry_Fehler:
    On Error Resume Next
    Meldung = "Verweis nicht gefunden." & Chr(10) & Chr(10) & "Bitte Toolbar erneut installieren."
    Ergebnis = MsgBox(Meldung, Title:="Fehler")
    
Ende:

End Sub

Sub OC_openfile_Styleguide()
    '1. Anlauf: Online-Version laden - bei Fehler mit lokaler Version versuchen
    On Error Resume Next
    If GetSetting(PPAName, "Setup", "FileMode") = "Offline" Then GoTo Lokal_Laden
    If GetSetting(PPAName, "Setup", "NetBase") = "" Then GoTo Registry_Fehler
    If GetSetting(PPAName, "Setup", "NetStyleGuide") = "" Then GoTo Registry_Fehler
    Datei = GetSetting(PPAName, "Setup", "NetBase") & GetSetting(PPAName, "Setup", "NetStyleGuide")
    
    On Error GoTo Netz_Fehler
    Presentations.Open FileName:=Datei, ReadOnly:=msoTrue
    GoTo Ende 'Zum Ende wenn alles geklappt hat
    
Netz_Fehler:
    'Netz-Ladefehler aufgetreten
    Meldung = "Der Styleguide wurde nicht gefunden." & Chr(10) & Chr(10) & _
                "Bitte überprüfen sie die Netzwerkverbindung." & Chr(10) & Chr(10) & _
                "Die lokale Version wird geöffnet, diese ist eventuell nicht aktuell."
    Ergebnis = MsgBox(Meldung, Title:="Netzwerkfehler")
    
Lokal_Laden:
    '2. Anlauf: Lokale Version laden
    On Error Resume Next
    If GetSetting(PPAName, "Setup", "LocalStyleGuide") = "" Then GoTo Registry_Fehler
    Datei = GetSetting(PPAName, "Setup", "LocalStyleGuide")
    
    On Error GoTo Lokaler_Fehler
    Presentations.Open FileName:=Datei, ReadOnly:=msoTrue
    GoTo Ende
    
Lokaler_Fehler:
    On Error Resume Next
    Meldung = "Der lokale Styleguide wurde nicht gefunden." & Chr(10) & Chr(10) & "Bitte Toolbar erneut installieren."
    Ergebnis = MsgBox(Meldung, Title:="Dateifehler")
    GoTo Ende
    
Registry_Fehler:
    On Error Resume Next
    Meldung = "Verweis nicht gefunden." & Chr(10) & Chr(10) & "Bitte Toolbar erneut installieren."
    Ergebnis = MsgBox(Meldung, Title:="Fehler")
    
Ende:

End Sub

Sub OC_Open_CI_Folder()
    On Error GoTo Shell_Fehler
    Dim filExplorerOpen As Double
        
    If GetSetting(PPAName, "Setup", "FileMode") = "" Then GoTo Registry_Fehler
    If GetSetting(PPAName, "Setup", "LocalCIFolder") = "" Then GoTo Registry_Fehler
    If GetSetting(PPAName, "Setup", "FileMode") = "Offline" Then GoTo Lokal_Laden
    If GetSetting(PPAName, "Setup", "NetBase") = "" Then GoTo Registry_Fehler
    ' If GetSetting(PPAName, "Setup", "NetCIFolder") = "" Then GoTo Registry_Fehler
        
    ' Verzeichnis im Netz öffnen
    Datei = "explorer.exe /e, " & GetSetting(PPAName, "Setup", "NetBase") & GetSetting(PPAName, "Setup", "NetCIFolder")
    filExplorerOpen = Shell(Datei, vbNormalFocus)
    GoTo Ende
    
Lokal_Laden:
    On Error GoTo Lokaler_Fehler
    Datei = "explorer.exe /e, " & GetSetting(PPAName, "Setup", "LocalCIFolder")
    filExplorerOpen = Shell(Datei, vbNormalFocus)
    GoTo Ende

Shell_Fehler:
    On Error Resume Next
        
    Meldung = "Beim Öffnen des CI-Ordners im Netzwerk ist ein Fehler aufgetreten." & Chr(10) & Chr(10) & _
                "Bitte überprüfen sie die Netzwerkverbindung." & Chr(10) & Chr(10) & _
                "Wenn der Fehler weiter besteht bitte Toolbar erneut installieren."
    Ergebnis = MsgBox(Meldung, Title:="Netzwerkfehler")
    
    GoTo Ende
    
Lokaler_Fehler:
    On Error Resume Next
    Meldung = "Beim Öffnen des lokalen CI-Ordner ist ein Fehler aufgetreten." & Chr(10) & Chr(10) & "Bitte Toolbar erneut installieren."
    Ergebnis = MsgBox(Meldung, Title:="Ordnerfehler")
    GoTo Ende
    
Registry_Fehler:
    On Error Resume Next
    Meldung = "Verweis nicht gefunden." & Chr(10) & Chr(10) & "Bitte Toolbar erneut installieren."
    Ergebnis = MsgBox(Meldung, Title:="Fehler")
    
Ende:
    
End Sub
