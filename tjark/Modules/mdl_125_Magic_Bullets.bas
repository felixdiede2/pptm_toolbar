Attribute VB_Name = "mdl_125_Magic_Bullets"
Option Explicit

Public Sub Bullets_E1()
    On Error Resume Next
    
    Dim S As Shape
    Dim P As Variant
    Dim foo As Variant
        
    For Each S In ActiveWindow.Selection.ShapeRange
    
        If S.HasTextFrame Then
        
            For Each P In S.TextFrame.TextRange.Paragraphs
                        
                foo = P.indentLevel
                
                If foo = 1 Then
                    'Ebene 1: Bullet
                    With P.ParagraphFormat
                        .Bullet.Visible = msoTrue
                        .Bullet.Font.Name = "Arial"
                        .Bullet.Character = 8226
                        .Bullet.UseTextColor = msoTrue
                    End With
                Else
                    If foo = 2 Then
                        'Ebene 2: Anstrich
                        With P.ParagraphFormat
                            .Bullet.Visible = msoTrue
                            .Bullet.Font.Name = "Arial"
                            .Bullet.Character = 45
                            .Bullet.UseTextColor = msoTrue
                        End With
                    Else
                        'Ab Ebene 3: Quadrat
                        With P.ParagraphFormat
                            .Bullet.Visible = msoTrue
                            .Bullet.Font.Name = "Wingdings"
                            .Bullet.Character = 167
                            .Bullet.UseTextColor = msoTrue
                        End With
                    End If
                End If
            
            Next P
        
        End If
        
    Next S
    
End Sub

Public Sub Bullets_E2()
    On Error Resume Next
    
    Dim S As Shape
    Dim P As Variant
    Dim foo As Variant
        
    For Each S In ActiveWindow.Selection.ShapeRange
    
        If S.HasTextFrame Then
        
            For Each P In S.TextFrame.TextRange.Paragraphs
                        
                foo = P.indentLevel
                
                If foo = 1 Then
                    'Ebene 1: Gar nix
                    With P.ParagraphFormat
                        .Bullet.Visible = msoFalse
                    End With
                    
                Else
                    If foo = 2 Then
                        'Ebene 2: Bullet
                        With P.ParagraphFormat
                            .Bullet.Visible = msoTrue
                            .Bullet.Font.Name = "Arial"
                            .Bullet.Character = 8226
                            .Bullet.UseTextColor = msoTrue
                        End With
                    Else
                        If foo = 3 Then
                            'Ebene 3: Anstrich
                            With P.ParagraphFormat
                                .Bullet.Visible = msoTrue
                                .Bullet.Font.Name = "Arial"
                                .Bullet.Character = 45
                                .Bullet.UseTextColor = msoTrue
                            End With
                        Else
                            'Ab Ebene 4: Quadrat
                            With P.ParagraphFormat
                                .Bullet.Visible = msoTrue
                                .Bullet.Font.Name = "Wingdings"
                                .Bullet.Character = 167
                                .Bullet.UseTextColor = msoTrue
                            End With
                        End If
                    End If
                End If
            Next P
        End If
    Next S
End Sub

