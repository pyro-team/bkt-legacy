Attribute VB_Name = "Slides"
Option Explicit

Private Const TEMPLATE_FILE = "Templates.pptx"


Public Sub ShowTemplateDialog()
    Dim activePresentation As Presentation
    Set activePresentation = Application.Presentations(1)
    
    
    ' Get path and filename of template file
    Dim templateFilePath As String
    Dim addin As addin
    For Each addin In Application.AddIns
        'If addin.Name = "BKT-Legacy" Then
        If addin.Loaded And InStr(1, addin.Name, "BKT-Legacy") = 1 Then
            #If Mac Then
                templateFilePath = addin.Path & "/" & TEMPLATE_FILE
            #Else
                templateFilePath = addin.Path & "\" & TEMPLATE_FILE
            #End If
        End If
    Next
    
    If templateFilePath = "" Then
        MsgBox "Addin-Ordner konnte nicht identifiziert werden. Eventuell wurde die Addin-Datei umbenannt.", vbCritical
        Exit Sub
    End If
' Doenst work on Mac:
'    If Dir(templateFilePath) = "" Then
'        MsgBox "Keine Datei Templates.pptx im Addin-Ordner gefunden." & vbCrLf & templateFilePath, vbInformation
'        Exit Sub
'    End If

    ' Template-Foliensatz ˆffnen
    Dim templatePresentation As Presentation
    On Error GoTo ErrOpen
    #If Mac Then
        'mac does not support windowless mode
        ' GrantAccessToMultipleFiles Array(templateFilePath)
        Set templatePresentation = Presentations.Open(templateFilePath, msoTrue, msoFalse, msoTrue)
    #Else
        Set templatePresentation = Presentations.Open(templateFilePath, msoTrue, msoFalse, msoFalse)
    #End If
    On Error GoTo 0
    
    ' Anzahl Templates ermitteln
    Dim numOfTemplates As Integer
    numOfTemplates = templatePresentation.Slides.Count
        
    ' Array Templates erstellen
    Dim Templates() As String
    
    ' Array bef¸llen
    Dim curTemplateName As String
    
    Dim slideCounter As Integer
    
    
    For slideCounter = 1 To numOfTemplates
        ReDim Preserve Templates(slideCounter) As String
    
        If templatePresentation.Slides(slideCounter).shapes.HasTitle Then
            curTemplateName = templatePresentation.Slides(slideCounter).shapes.Title.TextFrame.TextRange.text
        Else
            curTemplateName = "Slide " & slideCounter
        End If
                                        
        Templates(slideCounter - 1) = curTemplateName
    Next
    
    
    ' Build template form
    With VorlagenForm
        With .list_Vorlagen
            .Clear
            .List = Templates
        End With
        
        .Show
    End With
    
    
    ' Get selectes template
    Dim selectedTemplate As Integer
    selectedTemplate = VorlagenForm.selectedTemplate
    
    If selectedTemplate >= 0 Then
        ' Copy selected template
        templatePresentation.Slides.Item(selectedTemplate + 1).Copy
        
        
        Dim slidePosition As Integer
        slidePosition = 0
        
        On Error GoTo Err1
        slidePosition = ActiveWindow.View.Slide.SlideIndex
Err1:
        activePresentation.Slides.Paste (slidePosition + 1)
    End If
    
    
    ' Template-Foliensatz schlieﬂen
    With templatePresentation
        .Saved = True
        .Close
    End With
Exit Sub

ErrOpen:
    MsgBox "Fehler beim ÷ffnen der Datei Templates.pptx im Addin-Ordner. Eventuell ist die Datei nicht vorhanden?" & vbCrLf & templateFilePath, vbCritical
    Exit Sub
End Sub


