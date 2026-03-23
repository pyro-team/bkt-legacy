Attribute VB_Name = "SelectionContext"
Option Explicit

Public Function GetActiveShapeRange() As ShapeRange
    Dim sel As Selection

    On Error GoTo ErrHandler

    Set sel = ActiveWindow.Selection
    Set GetActiveShapeRange = GetSelectionShapeRange(sel)
    Exit Function

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

Public Function ShapeRangeSortedByTop(ByVal shpRange As ShapeRange) As Shape()
    ShapeRangeSortedByTop = SortShapeRange(shpRange, True)
End Function

Public Function ShapeRangeSortedByLeft(ByVal shpRange As ShapeRange) As Shape()
    ShapeRangeSortedByLeft = SortShapeRange(shpRange, False)
End Function

Private Function SortShapeRange(ByVal shpRange As ShapeRange, ByVal sortByTop As Boolean) As Shape()
    Dim shapes() As Variant
    Dim result() As Shape
    Dim i As Long
    Dim shapeCount As Long

    If shpRange Is Nothing Then Exit Function

    shapeCount = shpRange.Count
    If shapeCount = 0 Then Exit Function

    ReDim shapes(1 To shapeCount, 1 To 2)
    For i = 1 To shapeCount
        If sortByTop Then
            shapes(i, 1) = shpRange(i).Top
        Else
            shapes(i, 1) = shpRange(i).Left
        End If
        Set shapes(i, 2) = shpRange(i)
    Next i

    QuickSortM shapes, 1, shapeCount

    ReDim result(1 To shapeCount)
    For i = 1 To shapeCount
        Set result(i) = shapes(i, 2)
    Next i

    SortShapeRange = result
End Function
