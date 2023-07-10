VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VorlagenForm 
   Caption         =   "Templatefolie einfügen"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   OleObjectBlob   =   "VorlagenForm.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "VorlagenForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Public selectedTemplate As Integer


Private Sub btn_cancel_Click()
    selectedTemplate = -1
    
    VorlagenForm.Hide
End Sub

Private Sub btn_selectTemplate_Click()
    selectedTemplate = VorlagenForm.list_Vorlagen.ListIndex
    
    VorlagenForm.Hide
End Sub

