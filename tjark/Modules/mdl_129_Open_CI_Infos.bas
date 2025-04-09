Attribute VB_Name = "mdl_129_Open_CI_Infos"
' Prozedur zum Aufruf der CI-Info
' -------------------------------
' Aufruf des benutzerdefinierten Formulars "CI-Info"

Option Explicit

Sub OC_open_CI_Info()

usrCI_Info.Show

End Sub



' Prozedur zum Aufruf der Hilfe
' -----------------------------
' Aufruf des benutzerdefinierten Formulars "Hilfe"
'
' Erweiterungsideen
' -------------------
' Beständiges Update der Hilfe
' Hilfe als Fenster mit mehreren Reitern anzeigen, ein Reiter für jedes Menü der CI-Toolbar

Sub open_help()

    On Error Resume Next
    
'    usrHelp.label_Package_Name.Caption = "Package für " & GetSetting(PPAName, "Setup", "PackageName")
    
'    usrHelp.label_Version_Date.Caption = "Build  " & GetSetting(PPAName, "Setup", "BuildDate")
    
'    If GetSetting(PPAName, "Setup", "FileMode") = "Offline" Then
'        usrHelp.Online_CheckBox.Value = False
'    Else
'        usrHelp.Online_CheckBox.Value = True
'        usrHelp.Online_CheckBox.Enabled = True

'    End If

    usrHelp.Show

End Sub
