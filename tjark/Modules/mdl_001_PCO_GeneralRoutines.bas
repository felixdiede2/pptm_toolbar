Attribute VB_Name = "mdl_001_PCO_GeneralRoutines"
'
' PowerPoint 2010 VBA Macro
' Porsche CO
' Copyright PaCE Graphic GbR
' Germany - April 2012
' Update September 2015
'

' Variables for complete Project
Public TitleIndex As Integer
Public AgendaIndex As Integer
Public LastPageIndex As Integer
Public LayoutIndex As Integer

Public ErrorExit As Boolean             'Variable for Error-Exit (several Tools)

Public TopPosition As Variant           'Variable for top position (Position Copy)
Public LeftPosition As Variant          'Variable for left position (Position Copy)
Public BottomPosition As Variant        'Variable for top position (Position Copy)
Public RightPosition As Variant         'Variable for left position (Position Copy)

' ** Subroutine – checking for errors in slide selection
'
Sub CheckErrorsSlideSelection()

' **** Check if presentation open
'
    If Application.Presentations.Count < 1 Then
        MsgBox "No open presentation! Please, open a presenation and restart this tool.", vbInformation, "No presentation"
        ErrorExit = True
        Exit Sub
    End If
    
' **** Check if presentation has slides
'
    If ActivePresentation.Slides.Count < 1 Then
        MsgBox "No slides in current presentation! Please, add a slide and restart this tool.", vbInformation, "Missing slides"
        ErrorExit = True
        Exit Sub
    End If

' **** Check if right view version
'
    If ActiveWindow.ViewType <> ppViewNormal And ActiveWindow.ViewType <> ppViewSlide Then
        MsgBox "Please change to Slide View or Normal View and restart this tool.", vbInformation, "Wrong view type!"
        ErrorExit = True
        Exit Sub
    End If
    
End Sub


' ** Subroutine – checking for errors in view selection
'
Sub CheckErrorViewWrong()
       
' **** Check if slide frame selected in view version "Normal"
'
    If ActiveWindow.ViewType = ppViewNormal Then
        If ActiveWindow.ViewType <> ppViewSlide And ActiveWindow.Panes.Item(2).Active = False Then
            ActiveWindow.Panes.Item(2).Activate
        End If
    End If
    
End Sub


' ** Subroutine – checking for errors in view selection
'
Sub CheckErrorViewSelection1()
       
' **** Check if slide frame selected in view version "Normal"
'
    If ActiveWindow.ViewType = ppViewNormal Then
        If ActiveWindow.ViewType <> ppViewSlide And ActiveWindow.Panes.Item(2).Active = False Then
            MsgBox "At least 1 object must be selected for this tool. Please, select one or more objects and restart tool.", _
                vbInformation, "No selection!"
            ErrorExit = True
            Exit Sub
        End If
    End If
    
End Sub



' ** Subroutine – checking for errors in view selection
'
Sub CheckErrorViewSelection()
       
' **** Check if slide frame selected in view version "Normal"
'
    If ActiveWindow.ViewType = ppViewNormal Then
        If ActiveWindow.ViewType <> ppViewSlide And ActiveWindow.Panes.Item(2).Active = False Then
            MsgBox "At least 1 object must be selected for this tool. Please, select one or more objects and restart tool.", _
                vbInformation, "No selection!"
            ErrorExit = True
            Exit Sub
        End If
    End If
    
End Sub


' ** Subroutine – checking for errors in view selection
'
Sub CheckErrorViewSelection2()
       
' **** Check if slide frame selected in view version "Normal"
'
    If ActiveWindow.ViewType = ppViewNormal Then
        If ActiveWindow.ViewType <> ppViewSlide And ActiveWindow.Panes.Item(2).Active = False Then
            MsgBox "At least 2 objects must be selected for this tool. Please, select two or more objects and restart tool.", _
                    vbInformation, "No selection!"
            ErrorExit = True
            Exit Sub
        End If
    End If
    
End Sub


' ** Subroutine – check if exactly one object on slide selected
'
Sub CheckIfOnly1ObjectSelected()

' **** Check if objects selected
'
    If ActiveWindow.Selection.Type = ppSelectionNone Then
        MsgBox "At least 1 object must be selected for this tool. Please, select one object and restart tool.", _
                vbInformation, "No selection!"
        ErrorExit = True
        Exit Sub
    ElseIf ActiveWindow.Selection.ShapeRange.Count > 1 Then
        MsgBox "At least only 1 object must be selected for this tool. Please, select only one object and restart tool.", _
               vbInformation, "Wrong selection!"
        ErrorExit = True
        Exit Sub
    End If

End Sub


' ** Subroutine – check if one or more objects on slide selected
'
Sub CheckIf1ObjectSelected()

' **** Check if objects selected
'
    If ActiveWindow.Selection.Type = ppSelectionNone Then
        MsgBox "At least 1 object must be selected for this tool. Please, select one or more objects and restart tool.", _
                    vbInformation, "No selection!"
        ErrorExit = True
        Exit Sub
    End If

End Sub



' ** Subroutine – check if two or more objects on slide selected
'
Sub CheckIf2ObjectsSelected()

' **** Check if objects selected
'
    If ActiveWindow.Selection.Type = ppSelectionNone Then
        MsgBox "At least 2 objects must be selected for this tool. Please, select two or more objects and restart tool.", _
                vbInformation, "No selection!"
        ErrorExit = True
        Exit Sub
    ElseIf ActiveWindow.Selection.ShapeRange.Count < 2 Then
         MsgBox "At least 2 object must be selected for this tool. Please, select two or more objects and restart tool.", _
                vbInformation, "Wrong selection!"
        ErrorExit = True
        Exit Sub
    End If
    
End Sub


Sub GetTitleIndex()

    Dim i As Integer
    
    With ActivePresentation.SlideMaster
        For i = 1 To .CustomLayouts.Count
            If .CustomLayouts(i).Name = "Titelfolie" Or .CustomLayouts(i).Name = "Cover page" Then
                TitleIndex = i
                Exit For
            End If
        Next
    End With

End Sub


Sub GetAgendaIndex()

    Dim i As Integer
    
    With ActivePresentation.SlideMaster
        For i = 1 To .CustomLayouts.Count
            If .CustomLayouts(i).Name = "Agenda" Then
                AgendaIndex = i
                Exit For
            End If
        Next
    End With

End Sub


Sub GetLastPageIndex()

    Dim i As Integer
    
    With ActivePresentation.SlideMaster
        For i = 1 To .CustomLayouts.Count
            If .CustomLayouts(i).Name = "Abschlussfolie" Or .CustomLayouts(i).Name = "Final slide" Then
                LastPageIndex = i
                Exit For
            End If
        Next
    End With

End Sub



Sub GetLayoutIndex()

    Dim i As Integer
    
    With ActivePresentation.SlideMaster
        For i = 1 To .CustomLayouts.Count
            If .CustomLayouts(i).Name = "NICHT VERWENDEN" Or .CustomLayouts(i).Name = "NEVER USE THIS LAYOUT" Then
                LayoutIndex = i
                Exit For
            End If
        Next
    End With
                
End Sub


