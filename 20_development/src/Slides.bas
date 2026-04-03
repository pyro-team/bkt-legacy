Attribute VB_Name = "Slides"
Option Explicit

Private Const TEMPLATE_FILE = "Templates.pptx"


Public Sub ShowTemplateDialog()
    Dim activePresentation As Presentation
    Dim library As ToolboxLibrary
    Dim templatePresentation As Presentation
    Dim templateFilePath As String
    Dim selectedTemplate As Integer

    Set activePresentation = Application.Presentations(1)
    Set library = New ToolboxLibrary

    templateFilePath = library.BuildAddinFilePath(TEMPLATE_FILE)
    If templateFilePath = "" Then
        MsgBox "Addin-Ordner konnte nicht identifiziert werden. Eventuell wurde die Addin-Datei umbenannt.", vbCritical
        Exit Sub
    End If

    If Not library.DoesFileExist(templateFilePath) Then
        MsgBox "Keine Datei Templates.pptx im Addin-Ordner gefunden." & vbCrLf & templateFilePath, vbInformation
        Exit Sub
    End If

    On Error GoTo ErrOpen
    Set templatePresentation = library.OpenPresentationFile(templateFilePath, msoTrue, msoFalse, True)
    On Error GoTo 0

    PopulateTemplateList templatePresentation
    VorlagenForm.Show

    selectedTemplate = VorlagenForm.selectedTemplate
    If selectedTemplate >= 0 Then
        templatePresentation.Slides.Item(selectedTemplate + 1).Copy
        activePresentation.Slides.Paste (GetInsertSlidePosition() + 1)
    End If

    With templatePresentation
        .Saved = True
        .Close
    End With
    Exit Sub

ErrOpen:
    MsgBox "Fehler beim Oeffnen der Datei Templates.pptx im Addin-Ordner. Eventuell ist die Datei nicht vorhanden?" & vbCrLf & templateFilePath, vbCritical
End Sub

Private Sub PopulateTemplateList(ByVal templatePresentation As Presentation)
    Dim templateNames() As String
    Dim slideCounter As Integer

    For slideCounter = 1 To templatePresentation.Slides.Count
        ReDim Preserve templateNames(slideCounter) As String
        templateNames(slideCounter - 1) = GetTemplateName(templatePresentation.Slides(slideCounter), slideCounter)
    Next slideCounter

    With VorlagenForm.list_Vorlagen
        .Clear
        .List = templateNames
    End With
End Sub

Private Function GetTemplateName(ByVal templateSlide As Slide, ByVal slideNumber As Integer) As String
    If templateSlide.Shapes.HasTitle Then
        GetTemplateName = templateSlide.Shapes.Title.TextFrame.TextRange.Text
    Else
        GetTemplateName = "Slide " & slideNumber
    End If
End Function

Private Function GetInsertSlidePosition() As Integer
    On Error GoTo ErrHandler
    GetInsertSlidePosition = ActiveWindow.View.Slide.SlideIndex
    Exit Function

ErrHandler:
    GetInsertSlidePosition = 0
End Function
