VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CODocID 
   Caption         =   "Fußnote definieren ..."
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   OleObjectBlob   =   "CODocID.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CODocID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




' PPT2010 Macro
' Programmed by Thomas Breuer
' Copyright PaCE Graphic GbR
' Germany - April 2012
'

'
' Dimensioning of needed variables
'
Dim DocIDFilename As String     ' Variable for Filename
Dim DocIDExisting As String     ' Variable for Existing
Dim DocIDPosX As Variant        ' X-Pos for DocID
Dim DocIDPosY As Variant        ' Y-Pos for DocID
Dim DocIDSize As Variant        ' Länge DocID
Dim ViewVersion As String       ' Saving current View Type
Dim OptionCounter As Integer    ' variable für Optionstyp
Dim SlideCounter As Integer                ' Seitenzähler


' ********************** Definition Start **********************
'
' Initalize Starting Values
'
Private Sub UserForm_Initialize()
    
    ViewVersion = ActiveWindow.ViewType
    DocIDFilename = ActivePresentation.Name
    DocIDContent.Text = DocIDFilename
    GetExistingDocID
    GetTitleIndex
    GetLastPageIndex
    GetLayoutIndex
    If DocIDExisting = "                " Then
        OptionButton2.Enabled = False
        Label2.Enabled = False
    End If
    If OptionCounter = 1 Then
        OptionButton1.Value = True
        DocIDContent.Text = DocIDFilename
    ElseIf OptionCounter = 2 Then
        OptionButton2.Value = True
        DocIDContent.Text = DocIDExisting
    ElseIf OptionCounter = 3 Then
        OptionButton3.Value = 3
        DocIDContent.Text = ""
    End If
    
End Sub
Private Sub GetExistingDocID()

' Inhalt existierende DocID abfragen
    ActiveWindow.ViewType = ppViewSlideMaster
    DocIDExisting = ActivePresentation.SlideMaster.HeadersFooters.Footer.Text
' zurück in Ausgangsposition vor macro
    ActiveWindow.ViewType = ViewVersion

End Sub


' ****************** Klick-Ereignisse Auswahl ********************
'
' Clicks on Labels
'
Private Sub Labelonoff_Click()
    If CheckBoxonoff.Value = False Then
        CheckBoxonoff.Value = True
    ElseIf CheckBoxonoff.Value = True Then
        CheckBoxonoff.Value = False
    End If
End Sub
Private Sub CheckBoxonoff_Click()
    If CheckBoxonoff.Value = True Then
        LabelDocIDType.Enabled = True
        OptionButton1.Enabled = True
        Label1.Enabled = True
        OptionButton2.Enabled = True
        Label2.Enabled = True
        OptionButton3.Enabled = True
        Label3.Enabled = True
        Label4.Enabled = True
        DocIDContent.Enabled = True
    ElseIf CheckBoxonoff.Value = False Then
        LabelDocIDType.Enabled = False
        OptionButton1.Enabled = False
        Label1.Enabled = False
        OptionButton2.Enabled = False
        Label2.Enabled = False
        OptionButton3.Enabled = False
        Label3.Enabled = False
        Label4.Enabled = False
        DocIDContent.Enabled = False
    End If
End Sub
Private Sub Label1_Click()
    OptionButton1.Value = True
End Sub
Private Sub Label2_Click()
    OptionButton2.Value = True
End Sub
Private Sub Label3_Click()
    OptionButton3.Value = True
End Sub
Private Sub OptionButton1_Click()
    DocIDContent.Text = DocIDFilename
End Sub
Private Sub OptionButton2_Click()
    DocIDContent.Text = DocIDExisting
End Sub
'Private Sub OptionButton3_Click()
'    DocIDContent.Text
'End Sub
Private Sub DocIDContent_KeyPress(ByVal KeyAscii As _
    MSForms.ReturnInteger)
    OptionButton3.Value = True
End Sub


' ****************** Exit program ********************
'
' Exit Program without any action
'
Private Sub CancelAction_Click()
    ActiveWindow.ViewType = ViewVersion
    CODocID.Hide
    Unload Me
End Sub


' ****************** Main program ********************
'
' Exit Program with Disclaimer-action
'
Private Sub StartAction_Click()
    
' Disclaimer off
    If CheckBoxonoff.Value = False Then
        ActiveWindow.ViewType = ppViewSlideMaster
        With ActivePresentation.SlideMaster.HeadersFooters.Footer
            .Visible = msoFalse
        End With
        For SlideCounter = 1 To ActivePresentation.Slides.Count
            If ActivePresentation.Slides(SlideCounter).CustomLayout.Index <> TitleIndex _
            And ActivePresentation.Slides(SlideCounter).CustomLayout.Index <> LastPageIndex _
            And ActivePresentation.Slides(SlideCounter).CustomLayout.Index <> LayoutIndex Then
                With ActivePresentation.Slides(SlideCounter).HeadersFooters.Footer
                    .Visible = msoFalse
                End With
            End If
        Next
' Disclaimer on
    ElseIf CheckBoxonoff.Value = True Then
        With ActivePresentation.SlideMaster.HeadersFooters.Footer
            .Text = DocIDContent.Text
            .Visible = msoTrue
        End With
        For SlideCounter = 1 To ActivePresentation.Slides.Count
            If ActivePresentation.Slides(SlideCounter).CustomLayout.Index <> TitleIndex _
            And ActivePresentation.Slides(SlideCounter).CustomLayout.Index <> LastPageIndex _
            And ActivePresentation.Slides(SlideCounter).CustomLayout.Index <> LayoutIndex Then
                With ActivePresentation.Slides(SlideCounter).HeadersFooters.Footer
                    .Visible = msoTrue
                    .Text = DocIDContent.Text
                End With
            End If
        Next
    End If

    ActiveWindow.ViewType = ViewVersion
    CODocID.Hide
    Unload Me
End Sub

