Attribute VB_Name = "ToolboxSelections"
Option Explicit

' Liefert ein Array mit Shape-Objekten, sortiert nach Top-Wert (Abstand vom oberen Rand)
Public Function ActiveWindowSelectionSortedByTop() As Shape()
    Dim shapes() As Variant
    Dim result() As Shape
    Dim i As Integer
    Dim Count As Integer
    
    ' zwei-dimensionales Array aus Top-Werten und Shape-Objekten befüllen
    ReDim shapes(1 To ActiveWindow.selection.ShapeRange.Count, 1 To 2)
    Count = ActiveWindow.selection.ShapeRange.Count
    For i = 1 To Count
        shapes(i, 1) = ActiveWindow.selection.ShapeRange(i).Top
        Set shapes(i, 2) = ActiveWindow.selection.ShapeRange(i)
    Next
    QuickSortM shapes, 1, ActiveWindow.selection.ShapeRange.Count
    
    ReDim result(1 To ActiveWindow.selection.ShapeRange.Count)
    For i = 1 To ActiveWindow.selection.ShapeRange.Count
        Set result(i) = shapes(i, 2)
    Next
    
    ActiveWindowSelectionSortedByTop = result
End Function

' Liefert ein Array mit Shape-Objekten, sortiert nach Left-Wert (Abstand vom linken Rand)
Public Function ActiveWindowSelectionSortedByLeft() As Shape()
    Dim shapes() As Variant
    Dim result() As Shape
    Dim i As Integer
    Dim Count As Integer
    
    ' zwei-dimensionales Array aus Left-Werten und Shape-Objekten befüllen
    ReDim shapes(1 To ActiveWindow.selection.ShapeRange.Count, 1 To 2)
    Count = ActiveWindow.selection.ShapeRange.Count
    For i = 1 To Count
        ' Array umgekehrt befüllen, damit Reihenfolge von Shapes mit geichen Werten erhalten bleibt
        shapes(i, 1) = ActiveWindow.selection.ShapeRange(i).Left
        Set shapes(i, 2) = ActiveWindow.selection.ShapeRange(i)
    Next
    QuickSortM shapes, 1, ActiveWindow.selection.ShapeRange.Count
    
    ReDim result(1 To ActiveWindow.selection.ShapeRange.Count)
    For i = 1 To ActiveWindow.selection.ShapeRange.Count
        Set result(i) = shapes(i, 2)
    Next
    
    ActiveWindowSelectionSortedByLeft = result
End Function





' Alle Shapes auswählen, bei denen Shape-Typ mit aktuellen Shpape-Typ übereinstimmt
Public Sub SelectShapesByShapeType()
    Dim shp As Shape
    Dim shpMaster As Shape
    
    On Error GoTo Err_Handler
    If ActiveWindow.selection.Type <> ppSelectionShapes Then Exit Sub
    
    Set shpMaster = ActiveWindow.selection.ShapeRange(1)
    ActiveWindow.selection.Unselect
    
    For Each shp In ActiveWindow.View.Slide.shapes
        If shp.Visible And shp.Type = shpMaster.Type And shp.AutoShapeType = shpMaster.AutoShapeType Then
            shp.Select Replace:=False
        End If
    Next
    
    Exit Sub
Err_Handler:
    MsgBox "Fehler " & Err.Number & ":" & vbNewLine & Err.Description
End Sub

' Alle Shapes auswählen, bei denen Shape-Typ mit aktuellen Shpape-Hintergrund übereinstimmt
Public Sub SelectShapesByFillType()
    Dim shp As Shape
    Dim shpMaster As Shape
    
    On Error GoTo Err_Handler
    If ActiveWindow.selection.Type <> ppSelectionShapes Then Exit Sub
    
    Set shpMaster = ActiveWindow.selection.ShapeRange(1)
    ActiveWindow.selection.Unselect
    
    For Each shp In ActiveWindow.View.Slide.shapes
        If shp.Visible And shp.Fill.Visible = shpMaster.Fill.Visible And shp.Fill.Type = shpMaster.Fill.Type And shp.Fill.ForeColor.RGB = shpMaster.Fill.ForeColor.RGB Then
            shp.Select Replace:=False
        End If
    Next
    
    Exit Sub
Err_Handler:
    MsgBox "Fehler " & Err.Number & ":" & vbNewLine & Err.Description
End Sub

' Alle Shapes auswählen, bei denen Shape-Typ mit aktuellen Shpape-Rahmen übereinstimmt
Public Sub SelectShapesByLineType()
    Dim shp As Shape
    Dim shpMaster As Shape
    
    On Error GoTo Err_Handler
    If ActiveWindow.selection.Type <> ppSelectionShapes Then Exit Sub
    
    Set shpMaster = ActiveWindow.selection.ShapeRange(1)
    ActiveWindow.selection.Unselect
    
    For Each shp In ActiveWindow.View.Slide.shapes
        If shp.Visible And shp.Line.Visible = shpMaster.Line.Visible And shp.Line.Weight = shpMaster.Line.Weight And shp.Line.DashStyle = shpMaster.Line.DashStyle And shp.Line.ForeColor.RGB = shpMaster.Line.ForeColor.RGB Then
            shp.Select Replace:=False
        End If
    Next
    
    Exit Sub
Err_Handler:
    MsgBox "Fehler " & Err.Number & ":" & vbNewLine & Err.Description
End Sub


' SlideRangeIdentifier(activewindow.Presentation.Slides.Range(Array(2,1))) = "1,2"
' SlideRangeIdentifier(activewindow.Presentation.Slides.Range(Array(2,1,3))) = "1-3"
Public Function SlideRangeIdentifier(ByVal sldRange As SlideRange)
    SlideRangeIdentifier = RangeIdentifier(SlideRangeIndices(sldRange))
End Function

Private Function SlideRangeIndices(ByVal sldRange As SlideRange)
    Dim sldIndices() As Long
    Dim idx As Long
    
    ReDim sldIndices(1 To sldRange.Count)
    For idx = 1 To sldRange.Count
        sldIndices(idx) = sldRange(idx).SlideIndex
    Next
    QuickSort sldIndices, 1, sldRange.Count
    SlideRangeIndices = sldIndices
End Function

' RangeIdentifier(Array(1, 2)) = "1,2"
' RangeIdentifier(Array(1, 3)) = "1,3"
' RangeIdentifier(Array(1, 2, 3)) = "1-3"
' RangeIdentifier(Array(2, 3, 4, 6)) = "2-4,6"
' RangeIdentifier(Array(1, 2, 4, 5, 6)) = "1,2,4-6"
Private Function RangeIdentifier(vArray As Variant)
    Dim sldRangeId As String
    Dim idx As Long
    Dim lastIndex As Long
    Dim curIndex As Long
    Dim rangeFrom As Long
    Dim inRange As Boolean
    
    On Error Resume Next
    
    For idx = LBound(vArray) To UBound(vArray)
        curIndex = vArray(idx)
        If sldRangeId = "" Then
            sldRangeId = curIndex
        ElseIf Not inRange Then
            If curIndex = lastIndex + 1 Then
                rangeFrom = lastIndex
                inRange = True
            Else
                sldRangeId = sldRangeId & "," & curIndex
            End If
        Else
            If curIndex = lastIndex + 1 Then
                ' im Range-Modus bleiben
            Else
                inRange = False
                If lastIndex - rangeFrom = 1 Then
                    ' Pseudo-Range
                    sldRangeId = sldRangeId & "," & lastIndex & "," & curIndex
                Else
                    sldRangeId = sldRangeId & "-" & lastIndex & "," & curIndex
                End If
            End If
        End If
        
        lastIndex = curIndex
    Next
    
    If inRange Then
        If lastIndex - rangeFrom = 1 Then
            ' Pseudo-Range
            sldRangeId = sldRangeId & "," & lastIndex
        Else
            sldRangeId = sldRangeId & "-" & lastIndex
        End If
    End If
    
    
    RangeIdentifier = sldRangeId
End Function





