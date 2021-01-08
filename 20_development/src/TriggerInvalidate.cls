VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TriggerInvalidate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents PPTEvent As Application
Attribute PPTEvent.VB_VarHelpID = -1

Public Ribbon As IRibbonUI


Private Sub PPTEvent_SlideSelectionChanged(ByVal sldRange As SlideRange)
    On Error Resume Next
    Ribbon.Invalidate
End Sub

Private Sub PPTEvent_WindowActivate(ByVal pres As Presentation, ByVal Wn As DocumentWindow)
    On Error Resume Next
    Ribbon.Invalidate
End Sub

Private Sub PPTEvent_WindowSelectionChange(ByVal Sel As selection)
    On Error Resume Next
    Ribbon.Invalidate
'    ' Immer zum Toolbox Tab wechseln
'    If Not IsAltKeyDown Then
'        SwitchTab "Toolbox"
'        'SendKeys "%Y1%"
'    Else
'        SwitchTab "Toolbox p.2"
'        'SendKeys "%Y2%"
'    End If
End Sub
