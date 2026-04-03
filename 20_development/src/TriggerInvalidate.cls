VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TriggerInvalidate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents PPTEvent As Application
Attribute PPTEvent.VB_VarHelpID = -1

Public Ribbon As IRibbonUI


Private Sub PPTEvent_SlideSelectionChanged(ByVal sldRange As SlideRange)
    If Ribbon Is Nothing Then Exit Sub
    Ribbon.Invalidate
End Sub

Private Sub PPTEvent_WindowActivate(ByVal pres As Presentation, ByVal Wn As DocumentWindow)
    If Ribbon Is Nothing Then Exit Sub
    Ribbon.Invalidate
End Sub

Private Sub PPTEvent_WindowSelectionChange(ByVal sel As Selection)
    If Ribbon Is Nothing Then Exit Sub
    Ribbon.Invalidate
End Sub
