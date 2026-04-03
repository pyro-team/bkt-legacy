Attribute VB_Name = "ToolboxLibraryRibbon"
Option Explicit

Private mLibrary As ToolboxLibrary


Private Function GetToolboxLibrary() As ToolboxLibrary
    If mLibrary Is Nothing Then
        Set mLibrary = New ToolboxLibrary
    End If
    Set GetToolboxLibrary = mLibrary
End Function

Public Sub GetLibraryMenuContent(control As IRibbonControl, ByRef xmlStr)
    GetToolboxLibrary.GetMenuContent control, xmlStr
End Sub

Public Sub OpenLibraryPresentation(control As IRibbonControl)
    GetToolboxLibrary.OpenPresentation control
End Sub
