Attribute VB_Name = "mdl_072_MenuTab"
Option Explicit

' Prozeduren zum Erzeugen oder Ver‰ndern von Tabellen
' ---------------------------------------------------

Sub changeto_tab_horizontal()
    On Error Resume Next
    
    Dim sRightObjectPosition As Single
    Dim sTabOverallWidth, sTabObjectWidth, sTabSpaces, sSpaceBetweenObjects, sLeftObjectPosition, sLeftObjectWidth, sRightObjectWidth, sObjNum, i As Single
    Dim sRightBorderPosition, sRightObjectIndex As Single

    sSpaceBetweenObjects = 5.66                                                                                 ' Abstand zwischen Tabellenspalten
    sObjNum = ActiveWindow.Selection.ShapeRange.Count                                                           ' Gesamtzahl der angew‰‰hlten Objekte bestimmen
    sLeftObjectPosition = 3000                                                                                  ' Linkes Objekt der Tabelle finden
    
    For i = sObjNum To 1 Step -1
        If ActiveWindow.Selection.ShapeRange(i).Left < sLeftObjectPosition Then
        sLeftObjectPosition = ActiveWindow.Selection.ShapeRange(i).Left
        End If
    Next i
    
    sRightObjectPosition = 0                                                                                    ' Rechtes Objekt der Tabelle finden
    sRightObjectWidth = 0
    
    For i = sObjNum To 1 Step -1
        If (ActiveWindow.Selection.ShapeRange(i).Left + ActiveWindow.Selection.ShapeRange(i).Width) > (sRightObjectPosition + sRightObjectWidth) Then
        sRightObjectPosition = ActiveWindow.Selection.ShapeRange(i).Left
        sRightObjectWidth = ActiveWindow.Selection.ShapeRange(i).Width
        sRightBorderPosition = sRightObjectPosition + sRightObjectWidth
        sRightObjectIndex = i
        End If
    Next i
    
    sTabOverallWidth = (sRightObjectPosition + sRightObjectWidth) - sLeftObjectPosition                         ' Gesamtbreite der Tabelle
    sTabSpaces = ((sObjNum - 1) * sSpaceBetweenObjects)
    sTabObjectWidth = (sTabOverallWidth - sTabSpaces) / sObjNum
    ActiveWindow.Selection.ShapeRange.Width = sTabObjectWidth
    
    ActiveWindow.Selection.ShapeRange(sRightObjectIndex).Left = sRightBorderPosition - sTabObjectWidth          ' Position des Rechten Elementes nach Resize wiederherstellen
    ActiveWindow.Selection.ShapeRange.Align msoAlignTops, msoFalse                                              ' Ausrichten
    ActiveWindow.Selection.ShapeRange.Distribute msoDistributeHorizontally, msoFalse                            ' Verteilen

End Sub

Sub changeto_tab_vertical()
    On Error Resume Next
    
    Dim sLowerObjectPosition, sLowerObjectHeight As Single
    Dim sTabOverallHeight, sTabObjectHeight, sTabSpaces, sSpaceBetweenObjects, sUpperObjectPosition, sUpperObjectheight, sLowerObjectHeighth, sObjNum, i As Single
    Dim sLowerBorderPosition, sLowerObjectIndex As Single

    sSpaceBetweenObjects = 5.66                                                                                 ' Abstand zwischen Tabellenspalten
    sObjNum = ActiveWindow.Selection.ShapeRange.Count                                                          ' Gesamtzahl der angew‰‰hlten Objekte bestimmen
    sUpperObjectPosition = 3000                                                                                  ' Linkes Objekt der Tabelle finden
    
    For i = sObjNum To 1 Step -1
        If ActiveWindow.Selection.ShapeRange(i).Top < sUpperObjectPosition Then
        sUpperObjectPosition = ActiveWindow.Selection.ShapeRange(i).Top
        End If
    Next i
    
    sLowerObjectPosition = 0                                                                                    ' Rechtes Objekt der Tabelle finden
    sLowerObjectHeight = 0
    
    For i = sObjNum To 1 Step -1
        If (ActiveWindow.Selection.ShapeRange(i).Top + ActiveWindow.Selection.ShapeRange(i).Height) > (sLowerObjectPosition + sLowerObjectHeight) Then
        sLowerObjectPosition = ActiveWindow.Selection.ShapeRange(i).Top
        sLowerObjectHeight = ActiveWindow.Selection.ShapeRange(i).Height
        sLowerBorderPosition = sLowerObjectPosition + sLowerObjectHeight
        sLowerObjectIndex = i
        End If
    Next i
    
    sTabOverallHeight = (sLowerObjectPosition + sLowerObjectHeight) - sUpperObjectPosition                         ' Gesamtbreite der Tabelle
    sTabSpaces = ((sObjNum - 1) * sSpaceBetweenObjects)
    sTabObjectHeight = (sTabOverallHeight - sTabSpaces) / sObjNum
    ActiveWindow.Selection.ShapeRange.Height = sTabObjectHeight
    
    ActiveWindow.Selection.ShapeRange(sLowerObjectIndex).Top = sLowerBorderPosition - sTabObjectHeight          ' Position des Rechten Elementes nach Resize wiederherstellen
    ActiveWindow.Selection.ShapeRange.Align msoAlignLefts, msoFalse                                             ' Ausrichten
    ActiveWindow.Selection.ShapeRange.Distribute msoDistributeVertically, msoFalse                            ' Verteilen

End Sub

Sub changeto_tab()
    On Error GoTo 9
    
    If ActiveWindow.Selection.ShapeRange.AutoShapeType = msoShapeRectangle Then
        usrTabCreate.Show
    Else
        GoTo 9
    End If

    If 1 = 0 Then
9:      MsgBox ("Bitte ein Rechteck mit den Abmessungen " & Chr(10) & Chr(10) & _
            "der zu erstellenden Tabelle ausw‰hlen.")
    End If
End Sub

