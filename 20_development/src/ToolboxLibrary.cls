VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ToolboxLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CUSTOM_UI_NAMESPACE As String = "http://schemas.microsoft.com/office/2009/07/customui"
Private Const LIBRARY_FOLDER_NAME As String = "Library"
Private Const LIBRARY_CONTROL_PREFIX As String = "toolbox-library-open-"
Private Const LIBRARY_MENU_PREFIX As String = "toolbox-library-menu-"
Private Const LIBRARY_SEPARATOR_PREFIX As String = "toolbox-library-separator-"


Public Sub GetMenuContent(control As IRibbonControl, ByRef xmlStr)
    Dim libraryFolder As String
    Dim menuContent As String

    On Error GoTo ErrHandler

    libraryFolder = GetLibraryFolderPath()
    If Len(libraryFolder) = 0 Then
        xmlStr = BuildRootMenuXml(BuildDisabledButtonXml("toolbox-library-missing-addin", "Add-in folder not found"))
        Exit Sub
    End If

    If Not FolderExists(libraryFolder) Then
        xmlStr = BuildRootMenuXml(BuildDisabledButtonXml("toolbox-library-missing-folder", "Library folder not found"))
        Exit Sub
    End If

    menuContent = BuildFolderContentXml(libraryFolder)
    If Len(menuContent) = 0 Then
        xmlStr = BuildRootMenuXml(BuildDisabledButtonXml("toolbox-library-empty", "Library folder is empty"))
    Else
        xmlStr = BuildRootMenuXml(menuContent)
    End If
    Exit Sub

ErrHandler:
    xmlStr = BuildRootMenuXml(BuildDisabledButtonXml("toolbox-library-error", "Library could not be loaded"))
End Sub

Public Sub OpenPresentation(control As IRibbonControl)
    Dim filePath As String
    Dim openedPresentation As Presentation

    On Error GoTo ErrHandler

    filePath = DecodePathFromControlId(control.Id)
    If Len(filePath) = 0 Then
        MsgBox "The selected library entry is invalid.", vbExclamation
        Exit Sub
    End If

    If Not FileExists(filePath) Then
        MsgBox "The selected library presentation could not be found:" & vbCrLf & filePath, vbExclamation
        Exit Sub
    End If

    Set openedPresentation = OpenPresentationFile(filePath)
    Exit Sub

ErrHandler:
    MsgBox "The library presentation could not be opened:" & vbCrLf & filePath, vbExclamation
End Sub

Public Function GetAddinRootFolderPath() As String
    GetAddinRootFolderPath = GetAddinRootPath()
End Function

Public Function BuildAddinFilePath(ByVal fileName As String) As String
    Dim addinRoot As String

    addinRoot = GetAddinRootPath()
    If Len(addinRoot) = 0 Then Exit Function

    BuildAddinFilePath = BuildPath(addinRoot, fileName)
End Function

Public Function DoesFileExist(ByVal filePath As String) As Boolean
    DoesFileExist = FileExists(filePath)
End Function

Public Function OpenPresentationFile(ByVal filePath As String, Optional ByVal readOnly As MsoTriState = msoFalse, Optional ByVal untitled As MsoTriState = msoFalse, Optional ByVal allowWindowless As Boolean = False) As Presentation
#If Mac Then
    Set OpenPresentationFile = Presentations.Open(filePath, readOnly, untitled, msoTrue)
#Else
    Set OpenPresentationFile = Presentations.Open(filePath, readOnly, untitled, IIf(allowWindowless, msoFalse, msoTrue))
#End If
End Function


Private Function BuildRootMenuXml(ByVal innerXml As String) As String
    BuildRootMenuXml = "<menu xmlns=""" & CUSTOM_UI_NAMESPACE & """>" & innerXml & "</menu>"
End Function

Private Function BuildFolderContentXml(ByVal folderPath As String) As String
    Dim folderNames() As String
    Dim fileNames() As String
    Dim folderCount As Long
    Dim fileCount As Long
    Dim i As Long
    Dim subFolderPath As String
    Dim subFolderContent As String
    Dim xmlParts As String

    CollectFolderEntries folderPath, folderNames, folderCount, fileNames, fileCount

    For i = 1 To folderCount
        subFolderPath = BuildPath(folderPath, folderNames(i))
        subFolderContent = BuildFolderContentXml(subFolderPath)
        If Len(subFolderContent) > 0 Then
            xmlParts = xmlParts & BuildSubMenuXml(subFolderPath, folderNames(i), subFolderContent)
        End If
    Next i

    If folderCount > 0 And fileCount > 0 Then
        xmlParts = xmlParts & BuildSeparatorXml(folderPath)
    End If

    For i = 1 To fileCount
        xmlParts = xmlParts & BuildFileButtonXml(BuildPath(folderPath, fileNames(i)), fileNames(i))
    Next i

    BuildFolderContentXml = xmlParts
End Function

Private Sub CollectFolderEntries(ByVal folderPath As String, ByRef folderNames() As String, ByRef folderCount As Long, ByRef fileNames() As String, ByRef fileCount As Long)
    Dim entryName As String
    Dim entryNames() As String
    Dim entryCount As Long
    Dim i As Long
    Dim entryPath As String

    entryName = Dir(BuildPath(folderPath, "*"), vbDirectory)
    Do While Len(entryName) > 0
        If entryName <> "." And entryName <> ".." Then
            AppendString entryNames, entryCount, entryName
        End If
        entryName = Dir()
    Loop

    RequestAccessForFolderLevel folderPath, entryNames, entryCount

    For i = 1 To entryCount
        entryPath = BuildPath(folderPath, entryNames(i))
        If FolderExists(entryPath) Then
            AppendString folderNames, folderCount, entryNames(i)
        ElseIf IsPptxFile(entryNames(i)) Then
            AppendString fileNames, fileCount, entryNames(i)
        End If
    Next i

    SortStringArray folderNames, folderCount
    SortStringArray fileNames, fileCount
End Sub

Private Function BuildSubMenuXml(ByVal folderPath As String, ByVal folderName As String, ByVal innerXml As String) As String
    BuildSubMenuXml = "<menu id=""" & LIBRARY_MENU_PREFIX & EncodePath(folderPath) & """ label=""" & EscapeXml(folderName) & """>" & innerXml & "</menu>"
End Function

Private Function BuildFileButtonXml(ByVal filePath As String, ByVal fileName As String) As String
    Dim labelText As String

    labelText = StripPptxExtension(fileName)
    BuildFileButtonXml = "<button id=""" & LIBRARY_CONTROL_PREFIX & EncodePath(filePath) & """ label=""" & EscapeXml(labelText) & """ screentip=""" & EscapeXml(fileName) & """ supertip=""" & EscapeXml(filePath) & """ onAction=""OpenLibraryPresentation"" />"
End Function

Private Function BuildSeparatorXml(ByVal folderPath As String) As String
    BuildSeparatorXml = "<menuSeparator id=""" & LIBRARY_SEPARATOR_PREFIX & EncodePath(folderPath) & """ />"
End Function

Private Function BuildDisabledButtonXml(ByVal buttonId As String, ByVal labelText As String) As String
    BuildDisabledButtonXml = "<button id=""" & buttonId & """ label=""" & EscapeXml(labelText) & """ enabled=""false"" />"
End Function

Private Function GetLibraryFolderPath() As String
    Dim addinRoot As String

    addinRoot = GetAddinRootPath()
    If Len(addinRoot) = 0 Then Exit Function

    GetLibraryFolderPath = BuildPath(addinRoot, LIBRARY_FOLDER_NAME)
End Function

Private Function GetAddinRootPath() As String
    Dim currentAddin As addin

    For Each currentAddin In Application.AddIns
        If currentAddin.Loaded And InStr(1, currentAddin.Name, "BKT-Legacy") = 1 Then
            GetAddinRootPath = currentAddin.Path
            Exit Function
        End If
    Next currentAddin
End Function

Private Function BuildPath(ByVal folderPath As String, ByVal entryName As String) As String
    If Right$(folderPath, 1) = GetPathSeparator() Then
        BuildPath = folderPath & entryName
    Else
        BuildPath = folderPath & GetPathSeparator() & entryName
    End If
End Function

Private Function GetPathSeparator() As String
#If Mac Then
    GetPathSeparator = "/"
#Else
    GetPathSeparator = "\"
#End If
End Function

Private Function FolderExists(ByVal folderPath As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(folderPath) And vbDirectory) = vbDirectory)
    If Err.Number <> 0 Then
        FolderExists = False
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Len(filePath) > 0) And ((GetAttr(filePath) And vbDirectory) <> vbDirectory)
    If Err.Number <> 0 Then
        FileExists = False
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function IsPptxFile(ByVal fileName As String) As Boolean
    IsPptxFile = (LCase$(Right$(fileName, 5)) = ".pptx")
End Function

Private Function StripPptxExtension(ByVal fileName As String) As String
    If IsPptxFile(fileName) Then
        StripPptxExtension = Left$(fileName, Len(fileName) - 5)
    Else
        StripPptxExtension = fileName
    End If
End Function

Private Sub AppendString(ByRef values() As String, ByRef valueCount As Long, ByVal newValue As String)
    valueCount = valueCount + 1
    ReDim Preserve values(1 To valueCount)
    values(valueCount) = newValue
End Sub

Private Sub SortStringArray(ByRef values() As String, ByVal valueCount As Long)
    Dim i As Long
    Dim j As Long
    Dim tempValue As String

    If valueCount <= 1 Then Exit Sub

    For i = 1 To valueCount - 1
        For j = i + 1 To valueCount
            If StrComp(values(i), values(j), vbTextCompare) > 0 Then
                tempValue = values(i)
                values(i) = values(j)
                values(j) = tempValue
            End If
        Next j
    Next i
End Sub

Private Function EscapeXml(ByVal value As String) As String
    value = Replace(value, "&", "&amp;")
    value = Replace(value, """", "&quot;")
    value = Replace(value, "<", "&lt;")
    value = Replace(value, ">", "&gt;")
    EscapeXml = value
End Function

Private Function EncodePath(ByVal pathValue As String) As String
    Dim i As Long
    Dim charCode As Long
    Dim encodedValue As String

    For i = 1 To Len(pathValue)
        charCode = AscW(Mid$(pathValue, i, 1))
        If charCode < 0 Then charCode = charCode + 65536
        encodedValue = encodedValue & Right$("0000" & Hex$(charCode), 4)
    Next i

    EncodePath = encodedValue
End Function

Private Function DecodePathFromControlId(ByVal controlId As String) As String
    Dim encodedValue As String
    Dim i As Long
    Dim charCode As Long
    Dim decodedValue As String

    If Left$(controlId, Len(LIBRARY_CONTROL_PREFIX)) <> LIBRARY_CONTROL_PREFIX Then Exit Function

    encodedValue = Mid$(controlId, Len(LIBRARY_CONTROL_PREFIX) + 1)
    If Len(encodedValue) Mod 4 <> 0 Then Exit Function

    For i = 1 To Len(encodedValue) Step 4
        charCode = CLng("&H" & Mid$(encodedValue, i, 4))
        decodedValue = decodedValue & ChrW$(charCode)
    Next i

    DecodePathFromControlId = decodedValue
End Function

Private Sub RequestAccessForFolderLevel(ByVal folderPath As String, ByRef entryNames() As String, ByVal entryCount As Long)
#If Mac Then
    Dim permissionCandidates() As Variant
    Dim candidateCount As Long
    Dim i As Long
    Dim accessGranted As Boolean

    AppendPermissionCandidate permissionCandidates, candidateCount, folderPath

    For i = 1 To entryCount
        AppendPermissionCandidate permissionCandidates, candidateCount, BuildPath(folderPath, entryNames(i))
    Next i

    If candidateCount > 0 Then
        On Error Resume Next
        accessGranted = GrantAccessToMultipleFiles(permissionCandidates)
        On Error GoTo 0
    End If
#End If
End Sub

Private Sub AppendPermissionCandidate(ByRef permissionCandidates() As Variant, ByRef candidateCount As Long, ByVal candidatePath As String)
    candidateCount = candidateCount + 1
    ReDim Preserve permissionCandidates(0 To candidateCount - 1)
    permissionCandidates(candidateCount - 1) = candidatePath
End Sub
