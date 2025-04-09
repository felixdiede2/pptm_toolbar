Attribute VB_Name = "mdl_122_Magic_Table"
Option Explicit

Public Sub Table_Magic_Var()
    On Error Resume Next
    Dim Eingabe As Variant
    Dim Abstand As Single
    
    Eingabe = InputBox("Bitte Abstand in mm eingeben", "Table Magic", "2")
    
    If Eingabe <> "" Then
          
        Abstand = Val(Eingabe)
      
        If Abstand < 0 Or Abstand > 25 Then
            Abstand = 2
        End If
        
    Else
    
        Abstand = 2
        
    End If
    
        Fix_Table_Columns ActiveWindow.Selection.ShapeRange, Abstand
        Fix_Table_Rows ActiveWindow.Selection.ShapeRange, Abstand
    
End Sub


Public Sub Fix_Table_Columns(Elemente As ShapeRange, Abstand As Single)
    On Error Resume Next

    Dim Anz_Elemente, Anz_Spalten As Integer
    Dim Rand_Links, Rand_Rechts As Single
    Dim Such_Links, Akt_Links, Neu_Links, Akt_Breite As Single
    Dim Element_Links, Element_Rechts As Integer
        
    Dim i, j, P, n As Integer
    Dim foo, bar As Integer
    Dim delta As Single
    
    Dim Spalten_Breite() As Single
    Dim Spalten_Links() As Single
    Dim Spalten_Rechts() As Single
    Dim Skalierfaktor As Single
    
    Dim Reihenfolge() As Integer
    Dim Breitenfolge() As Integer
    Dim Sp_Links() As Integer
    Dim Sp_Rechts() As Integer
    
    Dim vertauscht As Boolean
            
    Anz_Elemente = Elemente.Count
    
    ' Auﬂenmaﬂe bestimmen
    
    ' Initialisierung mit dem ersten Element der Auswahl
    Rand_Links = Elemente(1).Left
    Rand_Rechts = Rand_Links + Elemente(1).Width
    Element_Links = 1
    Element_Rechts = 1
    
    ' Abarbeiten der weiteren Elemente
    For i = 2 To Anz_Elemente
        If Elemente(i).Left < Rand_Links Then
            Rand_Links = Elemente(i).Left
            Element_Links = i
        End If
        If Elemente(i).Left + Elemente(i).Width > Rand_Rechts Then
            Rand_Rechts = Elemente(i).Left + Elemente(i).Width
            Element_Rechts = i
        End If
    Next
    
    'Phase 1: Reihenfolge der Elemente finden
    
    'Vorbereiten
    ReDim Reihenfolge(Anz_Elemente)
    For i = 1 To Anz_Elemente
        Reihenfolge(i) = i
    Next i
    
    'Sortieren mit Bubblesort
    n = Anz_Elemente
    Do
        vertauscht = False
        For i = 1 To n - 1
            If Elemente(Reihenfolge(i)).Left > Elemente(Reihenfolge(i + 1)).Left Then
                foo = Reihenfolge(i)
                Reihenfolge(i) = Reihenfolge(i + 1)
                Reihenfolge(i + 1) = foo
                vertauscht = True
            End If
        Next i
    Loop While vertauscht = True And n >= 1
        
    'Phase 2: Elemente links ausrichten
    
    'Vorbereiten
    ReDim Sp_Links(Anz_Elemente)
    ReDim Sp_Rechts(Anz_Elemente)
    Such_Links = Rand_Links
    Akt_Links = Rand_Links
    Anz_Spalten = 0
                
    'Spaltenweise Linksausrichtung
    For i = 1 To Anz_Elemente
        Akt_Breite = Rand_Rechts - Rand_Links
        ' Alle Elemente mit ‰hnlichem linkem Rand suchen
        ' Dabei Breite der aktuellen Spalte bestimmen
        For j = i To Anz_Elemente
            If Elemente(Reihenfolge(j)).Left < Such_Links + 3 * mm2pt Then
                If Elemente(Reihenfolge(j)).Width < Akt_Breite Then
                    Akt_Breite = Elemente(Reihenfolge(j)).Width
                End If
            End If
        Next j
        
        Anz_Spalten = Anz_Spalten + 1
        
        For j = i To Anz_Elemente
            If Elemente(Reihenfolge(j)).Left < Such_Links + 0.4 * Akt_Breite Then
                Elemente(Reihenfolge(j)).Left = Akt_Links
                Sp_Links(j) = Anz_Spalten
                i = j
            End If
        Next j
        
        ReDim Preserve Spalten_Breite(Anz_Spalten)
        Spalten_Breite(Anz_Spalten) = Akt_Breite
        Akt_Links = Akt_Links + Akt_Breite
        If i < Anz_Elemente Then
            Such_Links = Elemente(Reihenfolge(i + 1)).Left
        End If
    Next i
    
    'Phase 3: Rechte Seite der Elemente einsortieren
    
    ReDim Spalten_Links(Anz_Spalten)
    ReDim Spalten_Rechts(Anz_Spalten)
    foo = Rand_Links
    
    For i = 1 To Anz_Spalten
        Spalten_Links(i) = foo
        foo = foo + Spalten_Breite(i)
        Spalten_Rechts(i) = foo
    Next i
    
    For i = 1 To Anz_Elemente
        foo = Rand_Rechts - Rand_Links
        bar = Sp_Links(i)
        For j = 1 To Anz_Spalten
            delta = Abs(Elemente(Reihenfolge(i)).Left + Elemente(Reihenfolge(i)).Width - Spalten_Rechts(j))
            If delta < foo Then
                foo = delta
                bar = j
            End If
        Next j
        Elemente(Reihenfolge(i)).Width = Spalten_Rechts(bar) - Elemente(Reihenfolge(i)).Left
        Sp_Rechts(i) = bar
    Next i
    
    'Phase 4: Spaltenbreiten angleichen
    
    'Vorbereiten
    ReDim Breitenfolge(Anz_Spalten)
    For i = 1 To Anz_Spalten
        Breitenfolge(i) = i
    Next i
    
    'Sortieren mit Bubblesort
    n = Anz_Spalten
    Do
        vertauscht = False
        For i = 1 To n - 1
            If Spalten_Breite(Breitenfolge(i)) < Spalten_Breite(Breitenfolge(i + 1)) Then
                foo = Breitenfolge(i)
                Breitenfolge(i) = Breitenfolge(i + 1)
                Breitenfolge(i + 1) = foo
                vertauscht = True
            End If
        Next i
    Loop While vertauscht = True And n >= 1
    
    For i = 2 To Anz_Spalten
        If Spalten_Breite(Breitenfolge(i)) > 0.9 * Spalten_Breite(Breitenfolge(i - 1)) Then
            Spalten_Breite(Breitenfolge(i)) = Spalten_Breite(Breitenfolge(i - 1))
        End If
    Next i

    'Phase 5: Breite wieder skalieren
    
    Akt_Breite = 0
    For i = 1 To Anz_Spalten
        Akt_Breite = Akt_Breite + Spalten_Breite(i)
    Next i
           
    Skalierfaktor = (Rand_Rechts - Rand_Links - ((Anz_Spalten - 1) * Abstand * mm2pt)) / Akt_Breite
    Neu_Links = Rand_Links
    
    For i = 1 To Anz_Spalten
        Spalten_Links(i) = Neu_Links
        Spalten_Breite(i) = Spalten_Breite(i) * Skalierfaktor
        Spalten_Rechts(i) = Spalten_Links(i) + Spalten_Breite(i)
        Neu_Links = Neu_Links + Spalten_Breite(i) + Abstand * mm2pt
    Next i
    
    For i = 1 To Anz_Elemente
        Elemente(Reihenfolge(i)).Left = Spalten_Links(Sp_Links(i))
        Elemente(Reihenfolge(i)).Width = Spalten_Rechts(Sp_Rechts(i)) - Elemente(Reihenfolge(i)).Left
    Next i
    
End Sub



Public Sub Fix_Table_Rows(Elemente As ShapeRange, Abstand As Single)
    On Error Resume Next

    Dim Anz_Elemente, Anz_Zeilen As Integer
    Dim Rand_Oben, Rand_Unten As Single
    Dim Such_Oben, Akt_Oben, Neu_Oben, Akt_Hoehe As Single
    Dim Element_Oben, Element_Unten As Integer
        
    Dim i, j, P, n As Integer
    Dim foo, bar As Integer
    Dim delta As Single
    
    Dim Zeilen_Hoehe() As Single
    Dim Zeilen_Oben() As Single
    Dim Zeilen_Unten() As Single
    Dim Skalierfaktor As Single
    
    Dim Reihenfolge() As Integer
    Dim Hoehenfolge() As Integer
    Dim Sp_Oben() As Integer
    Dim Sp_Unten() As Integer
    
    
    Dim vertauscht As Boolean
            
    Anz_Elemente = Elemente.Count
    
    ' Auﬂenmaﬂe bestimmen
    
    ' Initialisierung mit dem ersten Element der Auswahl
    Rand_Oben = Elemente(1).Top
    Rand_Unten = Rand_Oben + Elemente(1).Height
    
    Element_Oben = 1
    Element_Unten = 1
    
    ' Abarbeiten der weiteren Elemente
    For i = 2 To Anz_Elemente

        If Elemente(i).Top < Rand_Oben Then
            Rand_Oben = Elemente(i).Top
            Element_Oben = i
        End If
        
        If Elemente(i).Top + Elemente(i).Height > Rand_Unten Then
            Rand_Unten = Elemente(i).Top + Elemente(i).Height
            Element_Unten = i
        End If
                
    Next
    
    'Phase 1: Reihenfolge der Elemente finden
    
    'Vorbereiten
    ReDim Reihenfolge(Anz_Elemente)
    For i = 1 To Anz_Elemente
        Reihenfolge(i) = i
    Next i
    
    'Sortieren mit Bubblesort
    n = Anz_Elemente
    Do
        vertauscht = False
        For i = 1 To n - 1
            If Elemente(Reihenfolge(i)).Top > Elemente(Reihenfolge(i + 1)).Top Then
                foo = Reihenfolge(i)
                Reihenfolge(i) = Reihenfolge(i + 1)
                Reihenfolge(i + 1) = foo
                vertauscht = True
            End If
        Next i
    Loop While vertauscht = True And n >= 1
        
    'Phase 2: Elemente Oben ausrichten
    
    'Vorbereiten
    ReDim Sp_Oben(Anz_Elemente)
    ReDim Sp_Unten(Anz_Elemente)
    Such_Oben = Rand_Oben
    Akt_Oben = Rand_Oben
    Anz_Zeilen = 0
                
    'Zeilenweise Obenausrichtung
    For i = 1 To Anz_Elemente
        
        Akt_Hoehe = Rand_Unten - Rand_Oben
        ' Alle Elemente mit ‰hnlichem oberen Rand suchen
        ' Dabei Hoehe der aktuellen Zeile bestimmen
        For j = i To Anz_Elemente
            If Elemente(Reihenfolge(j)).Top < Such_Oben + 3 * mm2pt Then
                If Elemente(Reihenfolge(j)).Height < Akt_Hoehe Then
                    Akt_Hoehe = Elemente(Reihenfolge(j)).Height
                End If
            End If
        Next j
        
        Anz_Zeilen = Anz_Zeilen + 1
        
        For j = i To Anz_Elemente
            If Elemente(Reihenfolge(j)).Top < Such_Oben + 0.4 * Akt_Hoehe Then
                Elemente(Reihenfolge(j)).Top = Akt_Oben
                Sp_Oben(j) = Anz_Zeilen
                i = j
            End If
        Next j
        
        ReDim Preserve Zeilen_Hoehe(Anz_Zeilen)
        Zeilen_Hoehe(Anz_Zeilen) = Akt_Hoehe
        Akt_Oben = Akt_Oben + Akt_Hoehe
        If i < Anz_Elemente Then
            Such_Oben = Elemente(Reihenfolge(i + 1)).Top
        End If
    Next i
    
    'Phase 3: Rechte Seite der Elemente einsortieren
    
    ReDim Zeilen_Oben(Anz_Zeilen)
    ReDim Zeilen_Unten(Anz_Zeilen)
    foo = Rand_Oben
    
    For i = 1 To Anz_Zeilen
        Zeilen_Oben(i) = foo
        foo = foo + Zeilen_Hoehe(i)
        Zeilen_Unten(i) = foo
    Next i
    
    For i = 1 To Anz_Elemente
        foo = Rand_Unten - Rand_Oben
        bar = Sp_Oben(i)
        For j = 1 To Anz_Zeilen
            delta = Abs(Elemente(Reihenfolge(i)).Top + Elemente(Reihenfolge(i)).Height - Zeilen_Unten(j))
            If delta < foo Then
                foo = delta
                bar = j
            End If
        Next j
        Elemente(Reihenfolge(i)).Height = Zeilen_Unten(bar) - Elemente(Reihenfolge(i)).Top
        Sp_Unten(i) = bar
    Next i
    
    'Phase 4: Zeilenhoehen angleichen
    
    'Vorbereiten
    ReDim Hoehenfolge(Anz_Zeilen)
    For i = 1 To Anz_Zeilen
        Hoehenfolge(i) = i
    Next i
    
    'Sortieren mit Bubblesort
    n = Anz_Zeilen
    Do
        vertauscht = False
        For i = 1 To n - 1
            If Zeilen_Hoehe(Hoehenfolge(i)) < Zeilen_Hoehe(Hoehenfolge(i + 1)) Then
                foo = Hoehenfolge(i)
                Hoehenfolge(i) = Hoehenfolge(i + 1)
                Hoehenfolge(i + 1) = foo
                vertauscht = True
            End If
        Next i
    Loop While vertauscht = True And n >= 1
    
    For i = 2 To Anz_Zeilen
        If Zeilen_Hoehe(Hoehenfolge(i)) > 0.9 * Zeilen_Hoehe(Hoehenfolge(i - 1)) Then
            Zeilen_Hoehe(Hoehenfolge(i)) = Zeilen_Hoehe(Hoehenfolge(i - 1))
        End If
    Next i

    'Phase 5: Hoehe wieder skalieren
    
    'foo = 0
    'For i = 1 To Anz_Zeilen
    '    foo = foo + Zeilen_Hoehe(i)
    'Next i
    
    Akt_Hoehe = 0
    For i = 1 To Anz_Zeilen
        Akt_Hoehe = Akt_Hoehe + Zeilen_Hoehe(i)
    Next i
           
    Skalierfaktor = (Rand_Unten - Rand_Oben - ((Anz_Zeilen - 1) * Abstand * mm2pt)) / Akt_Hoehe
        
    Neu_Oben = Rand_Oben
    
    For i = 1 To Anz_Zeilen
        Zeilen_Oben(i) = Neu_Oben
        Zeilen_Hoehe(i) = Zeilen_Hoehe(i) * Skalierfaktor
        Zeilen_Unten(i) = Zeilen_Oben(i) + Zeilen_Hoehe(i)
        Neu_Oben = Neu_Oben + Zeilen_Hoehe(i) + Abstand * mm2pt
    Next i
    
    For i = 1 To Anz_Elemente
        Elemente(Reihenfolge(i)).Top = Zeilen_Oben(Sp_Oben(i))
    Next i
    
    For i = 1 To Anz_Elemente
        Elemente(Reihenfolge(i)).Height = Zeilen_Unten(Sp_Unten(i)) - Elemente(Reihenfolge(i)).Top
    Next i
    
End Sub

