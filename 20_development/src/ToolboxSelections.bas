Attribute VB_Name = "ToolboxSelections"
Option Explicit


' Alle Shapes auswaehlen, bei denen Shape-Typ mit aktuellen Shpape-Typ uebereinstimmt
Public Sub SelectShapesByShapeType()
    SelectShapesByProperty "shape"
End Sub

' Alle Shapes auswaehlen, bei denen Shape-Typ mit aktuellen Shpape-Hintergrund uebereinstimmt
Public Sub SelectShapesByFillType()
    SelectShapesByProperty "fill"
End Sub

' Alle Shapes auswaehlen, bei denen Shape-Typ mit aktuellen Shpape-Rahmen uebereinstimmt
Public Sub SelectShapesByLineType()
    SelectShapesByProperty "line"
End Sub

Private Sub SelectShapesByProperty(ByVal compareMode As String)
    Dim shp As Shape
    Dim masters As Collection
    Dim shpRange As ShapeRange

    On Error GoTo Err_Handler
    Set shpRange = GetActiveShapeRange()
    If shpRange Is Nothing Then Exit Sub

    Set masters = CreateShapeCollection(shpRange)
    ActiveWindow.Selection.Unselect

    For Each shp In ActiveWindow.View.Slide.Shapes
        If ShapeMatchesAnyMaster(shp, masters, compareMode) Then
            shp.Select Replace:=False
        End If
    Next

    Exit Sub
Err_Handler:
    MsgBox "Fehler " & Err.Number & ":" & vbNewLine & Err.Description
End Sub

Private Function CreateShapeCollection(ByVal shpRange As ShapeRange) As Collection
    Dim shp As Shape
    Dim result As Collection

    Set result = New Collection
    For Each shp In shpRange
        result.Add shp
    Next

    Set CreateShapeCollection = result
End Function

Private Function ShapeMatchesAnyMaster(ByVal shp As Shape, ByVal shpMasters As Collection, ByVal compareMode As String) As Boolean
    Dim shpMaster As Shape

    On Error GoTo ErrHandler
    If shp Is Nothing Then Exit Function
    If Not shp.Visible Then Exit Function

    For Each shpMaster In shpMasters
        If ShapeMatchesMaster(shp, shpMaster, compareMode) Then
            ShapeMatchesAnyMaster = True
            Exit Function
        End If
    Next
    Exit Function

ErrHandler:
    ShapeMatchesAnyMaster = False
End Function

Private Function ShapeMatchesMaster(ByVal shp As Shape, ByVal shpMaster As Shape, ByVal compareMode As String) As Boolean
    On Error GoTo ErrHandler

    Select Case compareMode
    Case "shape"
        ShapeMatchesMaster = shp.Type = shpMaster.Type And shp.AutoShapeType = shpMaster.AutoShapeType
    Case "fill"
        ShapeMatchesMaster = shp.Fill.Visible = shpMaster.Fill.Visible And shp.Fill.Type = shpMaster.Fill.Type And shp.Fill.ForeColor.RGB = shpMaster.Fill.ForeColor.RGB
    Case "line"
        ShapeMatchesMaster = shp.Line.Visible = shpMaster.Line.Visible And shp.Line.Weight = shpMaster.Line.Weight And shp.Line.DashStyle = shpMaster.Line.DashStyle And shp.Line.ForeColor.RGB = shpMaster.Line.ForeColor.RGB
    End Select
    Exit Function

ErrHandler:
    ShapeMatchesMaster = False
End Function


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
    Dim lowerBound As Long
    Dim upperBound As Long
    
    If Not IsArray(vArray) Then Exit Function
    On Error GoTo ErrHandler
    lowerBound = LBound(vArray)
    upperBound = UBound(vArray)
    On Error GoTo 0
    
    For idx = lowerBound To upperBound
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
    Exit Function

ErrHandler:
End Function





