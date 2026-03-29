Attribute VB_Name = "SelectionContext"
Option Explicit

Public Function GetActiveShapeRange() As ShapeRange
    Dim sel As Selection

    On Error GoTo ErrHandler

    Set sel = ActiveWindow.Selection
    If sel Is Nothing Then Exit Function
    
    If sel.HasChildShapeRange Then
        Set GetActiveShapeRange = sel.ChildShapeRange
    ElseIf sel.Type = ppSelectionShapes Or sel.Type = ppSelectionText Then
        Set GetActiveShapeRange = sel.ShapeRange
    End If
    
'    Set GetActiveShapeRange = GetSelectionShapeRange(sel)
'    If GetActiveShapeRange Is Nothing Then
'        Set GetActiveShapeRange = GetTextSelectionShapeRange(sel)
'    End If
'    Exit Function

ErrHandler:
End Function

Public Function GetActiveShapeCount() As Long
    Dim shpRange As ShapeRange

    Set shpRange = GetActiveShapeRange()
    If shpRange Is Nothing Then Exit Function

    GetActiveShapeCount = shpRange.Count
End Function

Public Function GetActiveFirstShape() As Shape
    Dim shpRange As ShapeRange

    Set shpRange = GetActiveShapeRange()
    If shpRange Is Nothing Then Exit Function
    If shpRange.Count = 0 Then Exit Function

    Set GetActiveFirstShape = shpRange(1)
End Function

Public Function GetActiveSlideRange() As SlideRange
    Dim sel As Selection

    On Error GoTo ErrHandler

    Set sel = ActiveWindow.Selection
    If sel Is Nothing Then Exit Function
    If sel.Type = ppSelectionSlides Then
        Set GetActiveSlideRange = sel.SlideRange
        Exit Function
    End If

    If Not ActiveWindow.View Is Nothing Then
        Set GetActiveSlideRange = ActiveWindow.View.Slide.Range
    End If
    Exit Function

ErrHandler:
End Function

Public Function SelectionContainsTextFrame(Optional ByVal shpRange As ShapeRange = Nothing) As Boolean
    Dim shp As Shape

    On Error GoTo ErrHandler

    If shpRange Is Nothing Then
        Set shpRange = GetActiveShapeRange()
    End If
    If shpRange Is Nothing Then Exit Function

    If shpRange.HasTextFrame <> msoFalse Or shpRange.HasTable <> msoFalse Then
        SelectionContainsTextFrame = True
        Exit Function
    End If

    For Each shp In shpRange
        If ShapeContainsTextFrame(shp) Then
            SelectionContainsTextFrame = True
            Exit Function
        End If
    Next shp
    Exit Function

ErrHandler:
    SelectionContainsTextFrame = False
End Function

Public Function ShapeRangeSortedByTop(ByVal shpRange As ShapeRange) As Shape()
    ShapeRangeSortedByTop = SortShapeRange(shpRange, True)
End Function

Public Function ShapeRangeSortedByLeft(ByVal shpRange As ShapeRange) As Shape()
    ShapeRangeSortedByLeft = SortShapeRange(shpRange, False)
End Function

Private Function SortShapeRange(ByVal shpRange As ShapeRange, ByVal sortByTop As Boolean) As Shape()
    Dim Shapes() As Variant
    Dim result() As Shape
    Dim i As Long
    Dim shapeCount As Long

    If shpRange Is Nothing Then Exit Function

    shapeCount = shpRange.Count
    If shapeCount = 0 Then Exit Function

    ReDim Shapes(1 To shapeCount, 1 To 2)
    For i = 1 To shapeCount
        If sortByTop Then
            Shapes(i, 1) = shpRange(i).Top
        Else
            Shapes(i, 1) = shpRange(i).Left
        End If
        Set Shapes(i, 2) = shpRange(i)
    Next i

    QuickSortM Shapes, 1, shapeCount

    ReDim result(1 To shapeCount)
    For i = 1 To shapeCount
        Set result(i) = Shapes(i, 2)
    Next i

    SortShapeRange = result
End Function

'Private Function GetTextSelectionShapeRange(ByVal sel As Selection) As ShapeRange
'    Dim candidateShape As Shape
'    Dim shapeNames(0 To 0) As String
'
'    On Error GoTo ErrHandler
'
'    If sel Is Nothing Then Exit Function
'    If sel.Type <> ppSelectionText Then Exit Function
'
'    If sel.HasChildShapeRange Then
'        Set GetTextSelectionShapeRange = sel.ChildShapeRange
'        Exit Function
'    End If
'
'    Set candidateShape = TryGetShapeFromTextParent(sel.TextRange2)
'    If candidateShape Is Nothing Then
'        Set candidateShape = TryGetShapeFromTextParent(sel.TextRange)
'    End If
'    If candidateShape Is Nothing Then
'        Set candidateShape = TryGetSelectionShape(sel)
'    End If
'
'    If candidateShape Is Nothing Then Exit Function
'
'    shapeNames(0) = candidateShape.Name
'    Set GetTextSelectionShapeRange = ActiveWindow.View.Slide.Shapes.Range(shapeNames)
'    Exit Function
'
'ErrHandler:
'End Function

'Private Function TryGetSelectionShape(ByVal sel As Selection) As Shape
'    On Error GoTo ErrHandler
'
'    If sel Is Nothing Then Exit Function
'    Set TryGetSelectionShape = sel.ShapeRange(1)
'    Exit Function
'
'ErrHandler:
'End Function

'Private Function TryGetShapeFromTextParent(ByVal sourceObject As Object) As Shape
'    Dim parentObject As Object
'
'    On Error GoTo ErrHandler
'
'    If sourceObject Is Nothing Then Exit Function
'
'    Set parentObject = sourceObject.Parent
'    If parentObject Is Nothing Then Exit Function
'
'    Select Case TypeName(parentObject)
'    Case "Shape"
'        Set TryGetShapeFromTextParent = parentObject
'    Case "TextFrame"
'        Set TryGetShapeFromTextParent = parentObject.Parent
'    Case "TextFrame2"
'        Set TryGetShapeFromTextParent = parentObject.Parent
'    End Select
'    Exit Function
'
'ErrHandler:
'End Function

Private Function ShapeContainsTextFrame(ByVal shp As Shape) As Boolean
    Dim groupItem As Shape

    On Error GoTo ErrHandler

    If shp Is Nothing Then Exit Function

    If shp.HasTextFrame <> msoFalse Or shp.HasTable <> msoFalse Then
        ShapeContainsTextFrame = True
        Exit Function
    End If

    Select Case shp.Type
    Case msoGroup
        For Each groupItem In shp.GroupItems
            If ShapeContainsTextFrame(groupItem) Then
                ShapeContainsTextFrame = True
                Exit Function
            End If
        Next groupItem
    Case msoSmartArt
        ShapeContainsTextFrame = True
    End Select
    Exit Function

ErrHandler:
    ShapeContainsTextFrame = False
End Function

