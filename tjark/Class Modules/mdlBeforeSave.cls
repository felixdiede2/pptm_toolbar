VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mdlBeforeSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents DeleteDesignBorders As Application
Attribute DeleteDesignBorders.VB_VarHelpID = -1

Private Sub DeleteDesignBorders_PresentationBeforeSave(ByVal Pres As Presentation, Cancel As Boolean)

    On Error Resume Next

    Set Pres = ActivePresentation
    Call designborders_delete

End Sub
