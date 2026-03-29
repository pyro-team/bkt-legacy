Attribute VB_Name = "VBAImportExport"
Option Explicit


Sub btnExport(control As IRibbonControl)
    Export
End Sub

Sub btnImport(control As IRibbonControl)
    Import
End Sub

Public Sub Export()
    On Error GoTo Err_Handler
    ExportModules
    
    Exit Sub
Err_Handler:
    MsgBox Err.Description, vbCritical, "Fehler!"
End Sub

Public Sub Import()
    On Error GoTo Err_Handler
    ImportModules
    
    Exit Sub
Err_Handler:
    MsgBox Err.Description, vbCritical, "Fehler!"
End Sub

Private Sub ExportModules()
    Dim bExport As Boolean
    Dim source As PowerPoint.Presentation
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As Object ' VBIDE.VBComponent

    If Len(ActivePresentation.Path) = 0 Then
        MsgBox "Please save first"
        Exit Sub
    End If

    Set source = ActivePresentation

    If source.VBProject.Protection = 1 Then
        MsgBox "The VBA in this workbook is protected," & _
            "not possible to export the code"
        Exit Sub
    End If

    szExportPath = FolderWithVBAProjectFiles()
    DeleteExistingSourceFiles szExportPath

    For Each cmpComponent In source.VBProject.VBComponents
        bExport = True
        szFileName = cmpComponent.Name

        Select Case cmpComponent.Type
            Case 2 ' Class
                szFileName = szFileName & ".cls"
            Case 3 ' Form
                szFileName = szFileName & ".frm"
            Case 1 ' Module
                szFileName = szFileName & ".bas"
            Case Else
                bExport = False
        End Select

        If bExport Then
            cmpComponent.Export BuildPath(szExportPath, szFileName)
        End If
    Next cmpComponent

    MsgBox "Export is ready"
End Sub


Private Sub ImportModules()
    Dim target As PowerPoint.Presentation
    Dim szImportPath As String
    Dim cmpComponents As Variant 'VBIDE.VBComponents

    Set target = ActivePresentation

    If target.VBProject.Protection = 1 Then
        MsgBox "The VBA in this workbook is protected," & _
            "not possible to Import the code"
        Exit Sub
    End If

    szImportPath = FolderWithVBAProjectFiles()

    If Not HasImportFiles(szImportPath) Then
        MsgBox "There are no files to import"
        Exit Sub
    End If

    Call DeleteVBAModulesAndUserForms

    Set cmpComponents = target.VBProject.VBComponents

    ImportFilesWithExtension cmpComponents, szImportPath, "bas"
    ImportFilesWithExtension cmpComponents, szImportPath, "cls"
    ImportFilesWithExtension cmpComponents, szImportPath, "frm"

    MsgBox "Import is ready"
End Sub

Private Function FolderWithVBAProjectFiles() As String
    Dim folder As String
    Dim srcFolder As String

    If Len(ActivePresentation.Path) = 0 Then
        Err.Raise 0, , "Please save the presentation first"
    End If

    folder = ActivePresentation.Path
    srcFolder = BuildPath(folder, "src")

    If FolderExists(srcFolder) Then
        folder = srcFolder
    Else
        folder = BuildPath(folder, GetPresentationBaseName() & "-src")
    End If

    If Not FolderExists(folder) Then
        MkDir folder
    End If

    If FolderExists(folder) Then
        FolderWithVBAProjectFiles = folder
    Else
        Err.Raise 0, , "Folder for VBA ProjectFiles could not be created"
    End If
End Function

Private Sub DeleteExistingSourceFiles(ByVal folderPath As String)
    DeleteFilesWithExtension folderPath, "frm"
    DeleteFilesWithExtension folderPath, "bas"
    DeleteFilesWithExtension folderPath, "cls"
End Sub

Private Sub DeleteFilesWithExtension(ByVal folderPath As String, ByVal extension As String)
    Dim fileName As String
    Dim fullPath As String

    fileName = Dir(BuildPath(folderPath, "*." & extension))
    Do While Len(fileName) > 0
        fullPath = BuildPath(folderPath, fileName)
        Kill fullPath
        fileName = Dir()
    Loop
End Sub

Private Function HasImportFiles(ByVal folderPath As String) As Boolean
    HasImportFiles = HasFilesWithExtension(folderPath, "bas") Or _
        HasFilesWithExtension(folderPath, "cls") Or _
        HasFilesWithExtension(folderPath, "frm")
End Function

Private Function HasFilesWithExtension(ByVal folderPath As String, ByVal extension As String) As Boolean
    HasFilesWithExtension = Len(Dir(BuildPath(folderPath, "*." & extension))) > 0
End Function

Private Sub ImportFilesWithExtension(ByVal cmpComponents As Variant, ByVal folderPath As String, ByVal extension As String)
    Dim fileName As String

    fileName = Dir(BuildPath(folderPath, "*." & extension))
    Do While Len(fileName) > 0
        cmpComponents.Import BuildPath(folderPath, fileName)
        fileName = Dir()
    Loop
End Sub

Private Function FolderExists(ByVal folderPath As String) As Boolean
    On Error Resume Next
    FolderExists = (GetAttr(folderPath) And vbDirectory) = vbDirectory
    If Err.Number <> 0 Then
        FolderExists = False
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function BuildPath(ByVal folderPath As String, ByVal fileName As String) As String
    If Right$(folderPath, 1) = GetPathSeparator() Then
        BuildPath = folderPath & fileName
    Else
        BuildPath = folderPath & GetPathSeparator() & fileName
    End If
End Function

Private Function GetPathSeparator() As String
    Dim presentationPath As String
    Dim fullName As String
    Dim nameLength As Long

    presentationPath = ActivePresentation.Path
    fullName = ActivePresentation.fullName
    nameLength = Len(ActivePresentation.Name)

    If Len(presentationPath) > 0 And Len(fullName) > Len(presentationPath) + nameLength Then
        GetPathSeparator = Mid$(fullName, Len(presentationPath) + 1, Len(fullName) - Len(presentationPath) - nameLength)
    ElseIf InStr(fullName, "/") > 0 Then
        GetPathSeparator = "/"
    Else
        GetPathSeparator = "\"
    End If
End Function

Private Function GetPresentationBaseName() As String
    Dim fileName As String
    Dim dotPosition As Long

    fileName = ActivePresentation.Name
    dotPosition = InStrRev(fileName, ".")

    If dotPosition > 0 Then
        GetPresentationBaseName = Left$(fileName, dotPosition - 1)
    Else
        GetPresentationBaseName = fileName
    End If
End Function


Private Function DeleteVBAModulesAndUserForms()
    Dim VBProj As Object ' VBIDE.VBProject
    Dim VBComp As Object ' VBIDE.VBComponent
    Dim index As Long

    Set VBProj = ActivePresentation.VBProject

    For index = VBProj.VBComponents.Count To 1 Step -1
        Set VBComp = VBProj.VBComponents(index)
        Select Case VBComp.Type
            Case 1, 2, 3
                ' 1-Module, 2-Class, 3-Form
                VBProj.VBComponents.Remove VBComp
            Case Else
                ' Do Nothing
        End Select
    Next index
End Function


