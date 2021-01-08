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
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As Object ' VBIDE.VBComponent
    Dim FSO As Object

    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.

    Set FSO = CreateObject("scripting.filesystemobject")

    If FSO.fileExists(ActivePresentation.FullName) = False Then
        MsgBox "Please save first"
        Exit Sub
    End If

    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If

    On Error Resume Next
    Kill FolderWithVBAProjectFiles & "\*.frm"
    Kill FolderWithVBAProjectFiles & "\*.bas"
    Kill FolderWithVBAProjectFiles & "\*.cls"
    On Error GoTo 0

    ''' NOTE: This workbook must be open in Excel.
    Set source = ActivePresentation

    If source.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the code"
    Exit Sub
    End If

    szExportPath = FolderWithVBAProjectFiles & "\"

    For Each cmpComponent In source.VBProject.VBComponents

        bExport = True
        szFileName = cmpComponent.Name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case 2 ' Class
                szFileName = szFileName & ".cls"
            Case 3 ' From
                szFileName = szFileName & ".frm"
            Case 1 ' Module
                szFileName = szFileName & ".bas"
            Case Else
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                bExport = False
        End Select

        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & szFileName

        ''' remove it from the project if you want
        '''wkbSource.VBProject.VBComponents.Remove cmpComponent

        End If

    Next cmpComponent

    MsgBox "Export is ready"
End Sub


Private Sub ImportModules()
    Dim target As PowerPoint.Presentation
    Dim objFSO As Object ' Scripting.FileSystemObject
    Dim objFile As Object ' Scripting.File
    Dim szTargetWorkbook As String
    Dim szImportPath As String
    Dim szFileName As String
    Dim cmpComponents As Variant 'VBIDE.VBComponents



'    If ActivePresentation = ThisWorkbook.Name Then
'        MsgBox "Select another destination workbook" & _
'        "Not possible to import in this workbook "
'        Exit Sub
'    End If

    'Get the path to the folder with modules
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If

    ''' NOTE: This workbook must be open in Excel.
    Set target = ActivePresentation

    If target.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to Import the code"
    Exit Sub
    End If

    ''' NOTE: Path where the code modules are located.
    szImportPath = FolderWithVBAProjectFiles & "\"

    Set objFSO = CreateObject("scripting.filesystemobject")
    If objFSO.GetFolder(szImportPath).Files.Count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If

    'Delete all modules/Userforms from the ActiveWorkbook
    Call DeleteVBAModulesAndUserForms

    Set cmpComponents = target.VBProject.VBComponents

    ''' Import all the code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(szImportPath).Files

        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            cmpComponents.Import objFile.path
        End If

    Next objFile

    MsgBox "Import is ready"
End Sub

Private Function FolderWithVBAProjectFiles() As String
    Dim WshShell As Object
    Dim FSO As Scripting.FileSystemObject
    Dim SpecialPath As String
    Dim folder As String

    Set WshShell = CreateObject("WScript.Shell")
    Set FSO = CreateObject("scripting.filesystemobject")
    
    folder = FSO.GetParentFolderName(ActivePresentation.FullName)
    If FSO.FolderExists(folder & "\src") Then
        folder = folder & "\src"
    Else
        folder = folder & FSO.GetBaseName(ActivePresentation.FullName) & "-src"
    End If
    
    If FSO.FolderExists(folder) = False Then
        On Error Resume Next
        MkDir folder
        On Error GoTo 0
    End If

    If FSO.FolderExists(folder) = True Then
        FolderWithVBAProjectFiles = folder
    Else
        Err.Raise 0, , "Folder for VBA ProjectFiles could not be created"
    End If
End Function


Private Function DeleteVBAModulesAndUserForms()
        Dim VBProj As Object ' VBIDE.VBProject
        Dim VBComp As Object ' VBIDE.VBComponent

        Set VBProj = ActivePresentation.VBProject

        For Each VBComp In VBProj.VBComponents
            Select Case VBComp.Type
            Case 1, 2, 3
                ' 1-Module, 2-Class, 3-Form
                VBProj.VBComponents.Remove VBComp
            Case Else
                ' Do Nothing
            End Select
        Next VBComp
End Function

