VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usrScaleFactor 
   Caption         =   "Bitte eingeben..."
   ClientHeight    =   1710
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   2670
   OleObjectBlob   =   "usrScaleFactor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usrScaleFactor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' Prozedur zum Skalieren einer Form inkl. des Inhalts um einen selbstgewählten Faktors
' ------------------------------------------------------------------------------------

Private Sub OK_Click()

On Error GoTo 7

Dim sFactor As Single
Dim newHeight, newWidth As Single

        If (usrScaleFactor.sScaleFactor.Value > 0) And (usrScaleFactor.sScaleFactor.Value < 501) Then
            sFactor = usrScaleFactor.sScaleFactor.Value / 100
        Else
            MsgBox ("Nur Werte zwischen 1 und 500% zulässig.")
            GoTo 8
        End If

            For Each S In ActiveWindow.Selection.ShapeRange
                Select Case S.Type
                Case msoEmbeddedOLEObject, _
                        msoLinkedOLEObject, _
                        msoOLEControlObject, _
                        msoLinkedPicture, msoPicture
                    S.ScaleWidth sFactor, msoTrue, msoScaleFromMiddle
                    S.ScaleHeight sFactor, msoTrue, msoScaleFromMiddle
                    ActiveWindow.Selection.TextRange.Font.Size = sFactor * ActiveWindow.Selection.TextRange.Font.Size
                Case Else
                    S.ScaleWidth sFactor, msoFalse, msoScaleFromMiddle
                    S.ScaleHeight sFactor, msoFalse, msoScaleFromMiddle
                    ActiveWindow.Selection.TextRange.Font.Size = sFactor * ActiveWindow.Selection.TextRange.Font.Size
                End Select
            Next
            
    If 1 = 0 Then
7:      MsgBox ("Bitte eine Form zum Skalieren auswählen.")
    End If
8:      Unload Me

End Sub

Private Sub UserForm_Initialize()

usrScaleFactor.sScaleFactor.Value = 100

End Sub
