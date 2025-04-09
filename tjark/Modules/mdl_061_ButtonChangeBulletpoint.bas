Attribute VB_Name = "mdl_061_ButtonChangeBulletpoint"
Option Explicit

' Prozedur zum Ändern des Bullet Points vor dem angewählten Text oder im angewählten Shape
' ----------------------------------------------------------------------------------------
' Primäre Schleife:
' Wenn aktueller Bulletpoint Pfeil, dann umwandeln zu Bullet
' Wenn aktueller Bulletpoint Bullet, dann umwandeln zu Bindestrich
' Wenn aktueller Bulletpoint Bindestrich, dann umwandeln zu Pfeil
'
' Eintrittsvektoren in die Schleife:
' Wenn Text ohne Anstrich, dann Anstich aktivieren
' Wenn anderer Anstrich als Pfeil, Quadrat oder Bindestrich, dann in Pfeil umwandeln


Sub OC_change_bullet()
    On Error Resume Next
    
    With ActiveWindow.Selection.TextRange.ParagraphFormat.Bullet
        Select Case .Character
            Case Is = 8226                'Bullet zu Anstrich
                If .Visible = msoTrue Then
                .Visible = msoTrue
                .UseTextColor = msoTrue
                .Font.Name = "Arial"
                .Character = 45        ' Anstrich
                Else: .Visible = msoTrue
                End If
            Case Is = 45                ' Anstrich zu Pfeil
                If .Visible = msoTrue Then
                .Visible = msoTrue
                .UseTextColor = msoTrue
                .Font.Name = "Wingdings"
                .Character = 167         ' Quadrat
                Else: .Visible = msoTrue
                End If
            Case Is = 167               ' Quadrat zu Bullet
                If .Visible = msoTrue Then
                .Visible = msoTrue
                .UseTextColor = msoTrue
                .Font.Name = "Arial"
                .Character = 8226         ' Bullet
                Else: .Visible = msoTrue
                End If
            Case Else                   ' Kein Bullet zu Bullet
                .Visible = msoTrue
                .UseTextColor = msoTrue
                .Font.Name = "Arial"
                .Character = 8226
        End Select
    End With
    
End Sub
