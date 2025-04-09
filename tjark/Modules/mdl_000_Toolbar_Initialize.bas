Attribute VB_Name = "mdl_000_Toolbar_Initialize"
Option Explicit

Public Const PPAName As String = "Otto Group Consulting Toolbar"      ' Interner Name der Menü-Toolbar
Public Const PPAFilename As String = "INTOC_STE_Otto Group Consulting Office 365 Toolbar v1p04"  ' Dateiname der Toolbar, auch für AddIn-Dialog (ohne .ppt/.ppa)

Public Const mm2pt = 72 / 25.4
Public Const cm2pt = 72 / 2.54
Public Const pt2mm = 25.4 / 72
Public Const pt2cm = 2.54 / 72

Public Sub Auto_Open()
    
    On Error Resume Next
    'Wird nicht mehr verwendet - hier wurde unter Office 2003 noch die Menüleiste erzeugt
       
End Sub


Public Sub Auto_Close()

    On Error Resume Next
    'Wird nicht mehr verwendet - hier wurde unter Office 2003 noch die Menüleiste entfernt
    
End Sub

Public Function PPA_Path(ThePPAName As String) As String
' Returns the path to the named add-in (no trailing backslash)

    PPA_Path = GetSetting(PPAName, "Setup", "Path", "foo")
    
    If PPA_Path = "foo" Then
        ' Not found, so we are probably in IDE mode instead of a PPA:
        PPA_Path = ActivePresentation.Path
        SaveSetting PPAName, "Setup", "Path", PPA_Path
        
    End If
    
    'MsgBox (PPA_Path)

End Function
