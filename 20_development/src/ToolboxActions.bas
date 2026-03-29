Attribute VB_Name = "ToolboxActions"
Option Explicit

Private Const SLIDENUMBERING = "Toolbox-SlideNumbering"


Public Sub AddSticker()
    Dim shp As Shape
    Dim sld As Slide
    Dim fontcol As ColorFormat
    
    Set sld = ActiveWindow.View.Slide
    
    ' Shape rechts oben auf slide erstellen
    'Set shp = sld.shapes.AddShape(msoShapeRectangle, 0, 60, 100, 20)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 60, 100, 20)
    ' Shape-Typ ist links-rechts-Pfeil, weil es die passenden Connector-Ecken hat
    shp.AutoShapeType = msoShapeLeftRightArrow
    ' Shape-Anpassung, so dass es wie ein Rechteck aussieht
    shp.Adjustments(1) = 1
    shp.Adjustments(2) = 0
    ' Shape-Stil
    'shp.Line.Weight = 0.75
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
    ' Text-Stil
    'shp.TextFrame.TextRange.Font.Color.RGB = 0
    shp.TextFrame.TextRange.Font.Size = 14
    shp.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignRight
    shp.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = False
    ' Autosize / Text nicht umbrechen
    shp.TextFrame.WordWrap = msoFalse
    shp.TextFrame.AutoSize = ppAutoSizeShapeToFitText
    ' Innenabstand
    shp.TextFrame.MarginBottom = 0
    shp.TextFrame.MarginTop = 0
    shp.TextFrame.MarginLeft = 0
    shp.TextFrame.MarginRight = 0
    ' Text
    shp.TextFrame.TextRange.text = "tbd"
    shp.Left = activePresentation.PageSetup.SlideWidth - shp.Width - 15
    
    Set fontcol = shp.TextFrame.TextRange.Font.Color
    
    ' Connectoren erstellen und mit Connector-Ecken des Shapes verbinden
    With sld.Shapes.AddConnector(Type:=msoConnectorStraight, BeginX:=0, _
            BeginY:=0, EndX:=100, EndY:=100)
        .ConnectorFormat.BeginConnect ConnectedShape:=shp, ConnectionSite:=1
        .ConnectorFormat.EndConnect ConnectedShape:=shp, ConnectionSite:=3
        .Line.ForeColor.RGB = 0
        .Line.Weight = 0.75
        If fontcol.Type = msoColorTypeScheme Then
            .Line.ForeColor.ObjectThemeColor = fontcol.ObjectThemeColor
            .Line.ForeColor.Brightness = fontcol.Brightness
        Else
            .Line.ForeColor.RGB = fontcol.RGB
        End If
    End With
    With sld.Shapes.AddConnector(Type:=msoConnectorStraight, BeginX:=0, _
            BeginY:=0, EndX:=100, EndY:=100)
        .ConnectorFormat.BeginConnect ConnectedShape:=shp, ConnectionSite:=5
        .ConnectorFormat.EndConnect ConnectedShape:=shp, ConnectionSite:=7
        .Line.ForeColor.RGB = 0
        .Line.Weight = 0.75
        If fontcol.Type = msoColorTypeScheme Then
            .Line.ForeColor.ObjectThemeColor = fontcol.ObjectThemeColor
            .Line.ForeColor.Brightness = fontcol.Brightness
        Else
            .Line.ForeColor.RGB = fontcol.RGB
        End If
    End With
    
    ' Text ausw?hlen
    shp.Select
    shp.TextFrame.TextRange.Select
End Sub



Public Sub AddProcessShapes(Optional numSteps As Long = 3)
    Dim sld As Slide
    Dim shp As Shape
    Dim i As Long
    Dim baseLeft As Single
    Dim baseTop As Single
    Dim shapeWidth As Single
    Dim shapeHeight As Single
    Dim minDimension As Single
    Dim AdjustmentValue As Single
    Dim shapeNames() As Variant

    If numSteps < 1 Then Exit Sub

    Set sld = ActiveWindow.View.Slide

    baseLeft = 100
    baseTop = 100
    shapeWidth = 120
    shapeHeight = 60

    minDimension = shapeWidth
    If shapeHeight < minDimension Then minDimension = shapeHeight
    AdjustmentValue = CentimetersToPoints(0.5) / minDimension

    ReDim shapeNames(1 To numSteps)

    For i = 1 To numSteps
        If i = 1 Then
            Set shp = sld.Shapes.AddShape(msoShapePentagon, baseLeft + (i - 1) * shapeWidth, baseTop, shapeWidth, shapeHeight)
        Else
            Set shp = sld.Shapes.AddShape(msoShapeChevron, baseLeft + (i - 1) * shapeWidth, baseTop, shapeWidth, shapeHeight)
        End If

        shp.Adjustments(1) = AdjustmentValue
        shapeNames(i) = shp.Name
    Next i

    sld.Shapes.Range(shapeNames).Select
End Sub


Public Sub AddConnectorShapeLeftRight()
    Dim shpLeft As Shape
    Dim shpRight As Shape
    Dim shpConnector As Shape
    Dim Shapes As Variant
    
    On Error GoTo Err_Handler
    
    Shapes = ActiveWindowSelectionSortedByLeft()
    
    Set shpLeft = Shapes(1)
    Set shpRight = Shapes(2)
    
    Set shpConnector = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeRectangle, shpLeft.Left + shpLeft.Width, shpLeft.Top, shpRight.Left - shpLeft.Left - shpLeft.Width, shpLeft.Height)
    
    ' node 2: top right
    shpConnector.Nodes.SetPosition 2, shpRight.Left, shpRight.Top
    ' node 3: bottom right
    shpConnector.Nodes.SetPosition 3, shpRight.Left, shpRight.Top + shpRight.Height
    shpConnector.Fill.ForeColor.RGB = 12566463 '193
    shpConnector.Line.ForeColor.RGB = 8355711 ' 127 127 127
    shpConnector.Line.Weight = 0.75
Err_Handler:
End Sub

Public Sub AddConnectorShapeTopBottom()
    Dim shpTop As Shape
    Dim shpBottom As Shape
    Dim shpConnector As Shape
    Dim Shapes As Variant
    
    On Error GoTo Err_Handler
    
    Shapes = ActiveWindowSelectionSortedByTop()
    
    Set shpTop = Shapes(1)
    Set shpBottom = Shapes(2)
    
    Set shpConnector = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeRectangle, shpTop.Left, shpTop.Top + shpTop.Height, shpTop.Width, shpBottom.Top - shpTop.Top - shpTop.Height)
    
    ' node 3: bottom right
    shpConnector.Nodes.SetPosition 3, shpBottom.Left + shpBottom.Width, shpBottom.Top
    ' node 4: bottom left
    shpConnector.Nodes.SetPosition 4, shpBottom.Left, shpBottom.Top
    shpConnector.Fill.ForeColor.RGB = 12566463 '193
    shpConnector.Line.ForeColor.RGB = 8355711 ' 127 127 127
    shpConnector.Line.Weight = 0.75
Err_Handler:
End Sub

Public Sub AddProtectedHyphen()
    InsertSpecialCharacter ChrW$(&HAD)
End Sub

Public Sub AddProtectedSpace()
    InsertSpecialCharacter ChrW$(&HA0)
End Sub

Public Sub AddProtectedNarrowSpace()
    InsertSpecialCharacter ChrW$(&H202F)
End Sub

Private Sub InsertSpecialCharacter(ByVal character As String)
    On Error GoTo Err_Handler
    
    If ActiveWindow.Selection.Type <> ppSelectionText Then Exit Sub
    ActiveWindow.Selection.TextRange2.text = character
    
Exit Sub
Err_Handler:
End Sub


Public Sub SetSameHeight(Optional func As String = "Max")
    Dim shp As Shape
    Dim Height As Single
    Dim currentSize As Single
    Dim factor As Single
    Dim shpRange As ShapeRange
    
    Set shpRange = GetActiveShapeRange()
    If shpRange Is Nothing Then Exit Sub

    Height = GetSelectedSizeValue(func, True)
    For Each shp In shpRange
        currentSize = GetNormalizedShapeSize(shp, True)
        If currentSize = 0 Then
            If ShapeUsesSwappedSize(shp) Then
                shp.Width = Height
            Else
                shp.Height = Height
            End If
            GoTo NextShape
        End If

        factor = Height / currentSize
        If ShapeUsesSwappedSize(shp) Then
            shp.ScaleWidth factor, msoFalse, msoScaleFromTopLeft
        Else
            shp.ScaleHeight factor, msoFalse, msoScaleFromTopLeft
        End If
NextShape:
    Next
End Sub

Public Sub SetSameWidth(Optional func As String = "Max")
    Dim shp As Shape
    Dim Width As Single
    Dim currentSize As Single
    Dim factor As Single
    Dim shpRange As ShapeRange
    
    Set shpRange = GetActiveShapeRange()
    If shpRange Is Nothing Then Exit Sub

    Width = GetSelectedSizeValue(func, False)
    For Each shp In shpRange
        currentSize = GetNormalizedShapeSize(shp, False)
        If currentSize = 0 Then
            If ShapeUsesSwappedSize(shp) Then
                shp.Height = Width
            Else
                shp.Width = Width
            End If
            GoTo NextShape
        End If

        factor = Width / currentSize
        If ShapeUsesSwappedSize(shp) Then
            shp.ScaleHeight factor, msoFalse, msoScaleFromTopLeft
        Else
            shp.ScaleWidth factor, msoFalse, msoScaleFromTopLeft
        End If
NextShape:
    Next
End Sub

Private Function GetSelectedSizeValue(ByVal func As String, ByVal useHeight As Boolean) As Single
    Dim shpRange As ShapeRange
    Dim shapeCount As Long
    Dim i As Long
    Dim currentValue As Single
    Dim valueSum As Double
    
    Set shpRange = GetActiveShapeRange()
    If shpRange Is Nothing Then Exit Function
    shapeCount = shpRange.Count
    If shapeCount = 0 Then Exit Function
    
    GetSelectedSizeValue = GetNormalizedShapeSize(shpRange(1), useHeight)
    
    Select Case UCase$(func)
        Case "LAST"
            GetSelectedSizeValue = GetNormalizedShapeSize(shpRange(shapeCount), useHeight)
        Case "MEAN"
            valueSum = 0
            For i = 1 To shapeCount
                valueSum = valueSum + GetNormalizedShapeSize(shpRange(i), useHeight)
            Next i
            GetSelectedSizeValue = valueSum / shapeCount
        Case "MIN"
            For i = 2 To shapeCount
                currentValue = GetNormalizedShapeSize(shpRange(i), useHeight)
                If currentValue < GetSelectedSizeValue Then GetSelectedSizeValue = currentValue
            Next i
        Case Else
            For i = 2 To shapeCount
                currentValue = GetNormalizedShapeSize(shpRange(i), useHeight)
                If currentValue > GetSelectedSizeValue Then GetSelectedSizeValue = currentValue
            Next i
    End Select
End Function

Private Function GetNormalizedShapeSize(ByRef shp As Shape, ByVal useHeight As Boolean) As Single
    If ShapeUsesSwappedSize(shp) Then
        If useHeight Then
            GetNormalizedShapeSize = shp.Width
        Else
            GetNormalizedShapeSize = shp.Height
        End If
    ElseIf useHeight Then
        GetNormalizedShapeSize = shp.Height
    Else
        GetNormalizedShapeSize = shp.Width
    End If
End Function

Private Function ShapeUsesSwappedSize(ByRef shp As Shape) As Boolean
    ShapeUsesSwappedSize = (shp.Rotation = 90 Or shp.Rotation = 270)
End Function

Public Sub SwapPosition()
    Dim shpRange As ShapeRange
    Dim i As Long
    Dim n As Long
    Dim lastLeft As Single
    Dim lastTop As Single
    
    Set shpRange = GetActiveShapeRange()
    If shpRange Is Nothing Then Exit Sub
    n = shpRange.Count
    If n < 2 Then Exit Sub
    
    ' Store last shape position
    lastLeft = shpRange(n).Left
    lastTop = shpRange(n).Top
    
    ' Shift positions down: shape(i) <- shape(i-1)
    For i = n To 2 Step -1
        shpRange(i).Left = shpRange(i - 1).Left
        shpRange(i).Top = shpRange(i - 1).Top
    Next i
    
    ' Put last position on first shape
    shpRange(1).Left = lastLeft
    shpRange(1).Top = lastTop
End Sub

Public Sub SwapPositionSize()
    Dim shpRange As ShapeRange
    Dim i As Long
    Dim n As Long
    Dim lastLeft As Single
    Dim lastTop As Single
    Dim lastWidth As Single
    Dim lastHeight As Single
    
    Set shpRange = GetActiveShapeRange()
    If shpRange Is Nothing Then Exit Sub
    n = shpRange.Count
    If n < 2 Then Exit Sub
    
    ' Store last shape position and size
    lastLeft = shpRange(n).Left
    lastTop = shpRange(n).Top
    lastWidth = shpRange(n).Width
    lastHeight = shpRange(n).Height
    
    ' Shift position and size down: shape(i) <- shape(i-1)
    For i = n To 2 Step -1
        shpRange(i).Left = shpRange(i - 1).Left
        shpRange(i).Top = shpRange(i - 1).Top
        shpRange(i).Width = shpRange(i - 1).Width
        shpRange(i).Height = shpRange(i - 1).Height
    Next i
    
    ' Put last values on first shape
    shpRange(1).Left = lastLeft
    shpRange(1).Top = lastTop
    shpRange(1).Width = lastWidth
    shpRange(1).Height = lastHeight
End Sub


' Bei allen markierten Shapes werden die Texte in separate Textboxen kopiert
Public Sub MoveTextOutOfShapes()
    Dim shp As Shape
    Dim shpTxt As Shape
    Dim shpRange As ShapeRange
    
    Set shpRange = GetActiveShapeRange()
    If shpRange Is Nothing Then Exit Sub

    For Each shp In shpRange
        If shp.HasTextFrame And shp.TextFrame.HasText Then   'shp.TextFrame.TextRange.text <> "" Then
            Set shpTxt = ActiveWindow.View.Slide.Shapes.AddTextbox(msoTextOrientationHorizontal, shp.Left, shp.Top, shp.Width, shp.Height)
            ' WordWrap / AutoSize
            shpTxt.TextFrame.WordWrap = msoTrue
            shpTxt.TextFrame.AutoSize = ppAutoSizeNone
            shpTxt.Height = shp.Height
            shpTxt.Rotation = shp.Rotation
            ' Seitenraender
            shpTxt.TextFrame.MarginBottom = shp.TextFrame.MarginBottom
            shpTxt.TextFrame.MarginTop = shp.TextFrame.MarginTop
            shpTxt.TextFrame.MarginLeft = shp.TextFrame.MarginLeft
            shpTxt.TextFrame.MarginRight = shp.TextFrame.MarginRight
            ' Ausrichtung
            shpTxt.TextFrame.Orientation = shp.TextFrame.Orientation
            shpTxt.TextFrame.HorizontalAnchor = shp.TextFrame.HorizontalAnchor
            shpTxt.TextFrame.VerticalAnchor = shp.TextFrame.VerticalAnchor
            ' Text kopieren
            shp.TextFrame.TextRange.Copy
            shpTxt.TextFrame.TextRange.Paste
            'shp.TextFrame.TextRange.text = ""
            shp.TextFrame.DeleteText
            ' Gr??e wiederherstellen
            shp.Top = shpTxt.Top
            shp.Height = shpTxt.Height
            shp.Width = shpTxt.Width
            ' Textfeld selektieren
            shpTxt.Select msoFalse
        End If
    Next
End Sub

' Markierung muss eine Textbox und ein weiteres Shape-Objekt enthalten
' Text aus der Textbox wird in das Shape-Objekt kopiert und Textbox entfernt
Public Sub MoveTextIntoShape()
    Dim shp As Shape
    Dim shpTxt As Shape
    Dim shpRange As ShapeRange
    
    Set shpRange = GetActiveShapeRange()
    If shpRange Is Nothing Then Exit Sub

    If shpRange.Count <> 2 Then
        MsgBox "Bitte eine Textbox und ein Shape-Objekt ausw?hlen.", vbInformation
        Exit Sub
    End If
    
    If shpRange(1).Type = msoTextBox Then
        Set shpTxt = shpRange(1)
        Set shp = shpRange(2)
    Else
        Set shpTxt = shpRange(2)
        Set shp = shpRange(1)
    End If
    
    ' Text kopieren
    shpTxt.TextFrame.TextRange.Copy
    ' DoEvents umgeht Bug, dass in manchen Situation kein Text eingefuegt wird
    DoEvents
    shp.TextFrame.TextRange.Paste
    ' Textbox loeschen
    shpTxt.Delete
End Sub


' Textabsaetze der markierten Shapes werden auf einzelne Shapes aufgeteilt
Public Sub SplitShapeByParagraphs()
    Dim shp As Shape
    Dim shpCopy As Object
    Dim par As TextRange
    Dim parIndex As Long
    Dim index As Long
    Dim parHeight As Single
    Dim rngSelection As ShapeRange
    
    Set rngSelection = GetActiveShapeRange()
    If rngSelection Is Nothing Then Exit Sub
    For Each shp In rngSelection
        If shp.HasTextFrame And shp.TextFrame.HasText Then   'shp.TextFrame.TextRange.text <> "" Then
            shp.Select msoTrue
            
            For parIndex = 2 To shp.TextFrame.TextRange.Paragraphs.Count
                Set par = shp.TextFrame.TextRange.Paragraphs(parIndex)
                ' Shape dublizieren
                Set shpCopy = shp.Duplicate
                shpCopy.Select msoFalse
                shpCopy.Top = shp.Top
                shpCopy.Left = shp.Left
                ' Absaetze 1..i-1 entfernen und Shape entsprechend verschieben
                For index = 1 To parIndex - 1
                    ' Textbox Position entsprechend Absatzhoehe anpassen
                    shpCopy.Top = shpCopy.Top + ParagraphHeight(shpCopy.TextFrame.TextRange.Paragraphs(1))
                    ' Absatz entfernen
                    shpCopy.TextFrame.TextRange.Paragraphs(1).Delete
                Next
                
                ' Absaetze i+1..n entfernen
                For index = parIndex + 1 To shp.TextFrame.TextRange.Paragraphs.Count
                    shpCopy.TextFrame.TextRange.Paragraphs(2).Delete
                Next
                ' Letztes CR-Zeichen loesen
                TrimNewLineCharacters shpCopy.TextFrame.TextRange
                
                ' Shape Hoehe abhaengig von Absaetzhoehe
                shpCopy.Height = ParagraphHeight(shpCopy.TextFrame.TextRange.Paragraphs(1)) + shpCopy.TextFrame.MarginTop + shpCopy.TextFrame.MarginBottom
                
                ' --> ein Absatz bleibt ?brig
            Next
            
            ' letzten Shape nach unten schieben
            shpCopy.Top = Max(shpCopy.Top, shp.Top + shp.Height - shpCopy.Height)
            
            ' Absaetze 2..n im Original-Shape entfernen
            For index = 2 To shp.TextFrame.TextRange.Paragraphs.Count
                shp.TextFrame.TextRange.Paragraphs(2).Delete
            Next
            ' Letztes CR-Zeichen loesen
            TrimNewLineCharacters shp.TextFrame.TextRange
            ' Textbox Hoehe an Absatzhoehe anpassen
            shp.Height = ParagraphHeight(shp.TextFrame.TextRange.Paragraphs(1)) + shp.TextFrame.MarginTop + shp.TextFrame.MarginBottom
            ' Objekte vertikal verteilen
            rngSelection.Distribute msoDistributeVertically, msoFalse
        End If
    Next
End Sub

' Text der markierte Shapes wird im ersten Shape zusammengefassen,
' andere Shapes werden geloescht.
' Reihenfolge abhaengig von Shape.top
Public Sub JoinShapesWithText()
    Dim Shapes As Variant
    Dim index As Long
    Dim txtRange As TextRange
    Dim shp As Shape
    Dim parCount As Long
    
    ' Shapes nach top sortieren
    Shapes = ActiveWindowSelectionSortedByTop
    ' Anapssung Gr??e des ersten Shapes (Master-Shape)
    Set shp = Shapes(1)
    shp.Height = Max(shp.Height, Shapes(UBound(Shapes)).Top + Shapes(UBound(Shapes)).Height - shp.Top)
    
    For index = 2 To UBound(Shapes)
        ' Text aus Shape kopieren
        Shapes(index).TextFrame.TextRange.Copy
        ' neuen Absatz in Master-Shape erstellen
        parCount = shp.TextFrame.TextRange.Paragraphs.Count
        Do
            Set txtRange = shp.TextFrame.TextRange.Paragraphs.InsertAfter(vbCr) 'vbNewline not supported by mac
        Loop Until parCount < shp.TextFrame.TextRange.Paragraphs.Count
        ' Text in Master-Shape einfuegen
        DoEvents
        txtRange.Paste
        ' Letztes CR-Zeichen loeschen
        If txtRange.Characters(txtRange.Length, 1) = vbCr Then
            txtRange.Characters(txtRange.Length, 1).Delete
        End If
        ' Shape loeschen
        Shapes(index).Delete
    Next
End Sub


Public Sub ReplaceAllText()
    Dim newText As String
    Dim shp As Shape
    Dim rngSelection As ShapeRange
    Set rngSelection = GetActiveShapeRange()
    If rngSelection Is Nothing Then Exit Sub

    newText = InputBox("Neuen Text eingeben", "Text ersetzen", "tbd")
    ' Bei Abbruch ist R?ckgabewert leer
    If newText = "" Then Exit Sub

    On Error Resume Next
    rngSelection.TextFrame.TextRange.text = newText
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0

    For Each shp In rngSelection
        If shp.HasTextFrame Then
            shp.TextFrame.TextRange.text = newText
        End If
    Next
End Sub

Public Sub RemoveAllText()
    Dim shp As Shape
    Dim rngSelection As ShapeRange
    Set rngSelection = GetActiveShapeRange()
    If rngSelection Is Nothing Then Exit Sub

    On Error Resume Next
    rngSelection.TextFrame.TextRange.Delete
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0

    For Each shp In rngSelection
        If shp.HasTextFrame Then
            shp.TextFrame.TextRange.Delete
        End If
    Next
End Sub

Public Sub TextMarginZero()
    Dim shp As Shape
    Dim rngSelection As ShapeRange
    Set rngSelection = GetActiveShapeRange()
    If rngSelection Is Nothing Then Exit Sub

    On Error Resume Next
    rngSelection.TextFrame2.MarginLeft = 0
    rngSelection.TextFrame2.MarginRight = 0
    rngSelection.TextFrame2.MarginTop = 0
    rngSelection.TextFrame2.MarginBottom = 0
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0

    For Each shp In rngSelection
        If shp.HasTextFrame Then
            shp.TextFrame2.MarginLeft = 0
            shp.TextFrame2.MarginRight = 0
            shp.TextFrame2.MarginTop = 0
            shp.TextFrame2.MarginBottom = 0
        End If
    Next
End Sub


Public Sub HideShapes()
    Dim shp As Shape
    Dim rngSelection As ShapeRange
    
    Set rngSelection = GetActiveShapeRange()
    If rngSelection Is Nothing Then Exit Sub

    On Error Resume Next
    rngSelection.Visible = msoFalse
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0

    For Each shp In rngSelection
        shp.Visible = msoFalse
    Next
End Sub


Public Sub ShowShapes()
    Dim shp As Shape
    Dim allShapes As Shapes
    
    Set allShapes = ActiveWindow.View.Slide.Shapes
    ActiveWindow.Selection.Unselect
    For Each shp In allShapes
        If shp.Visible = msoFalse Then
            shp.Visible = msoTrue
            shp.Select msoFalse
        End If
    Next
End Sub

'Issue with Mac: Paste as picture pastes PDF/Vector, that is not properly shown on Windows
'BUT: PasteSpecial not available on Mac, so no solution at the moment
'Public Sub PasteAsPicturePng()
'    Dim sld As Slide
'
'    Set sld = ActiveWindow.View.Slide
'    sld.shapes.PasteSpecial ppPastePNG
'End Sub

Public Sub PasteOnSlides()
    Dim sld As Slide
    Dim rngSelection As SlideRange
    
    Set rngSelection = GetActiveSlideRange()
    If rngSelection Is Nothing Then Exit Sub
    For Each sld In rngSelection
        sld.Shapes.Paste
    Next
End Sub

Public Sub PasteAndReplace()
    Dim shp As Shape
    Dim rngSelection As ShapeRange
    Set rngSelection = GetActiveShapeRange()
    If rngSelection Is Nothing Then Exit Sub
    
    'On Error Resume Next
    For Each shp In rngSelection
        PasteAndReplaceShape shp
    Next
End Sub

Private Sub PasteAndReplaceShape(shp As Shape)
    Dim sld As Slide
    Dim pastedShapes As ShapeRange
    Dim pastedShape As Shape
    Dim targetZ As Long
    
    Set sld = ActiveWindow.View.Slide
    targetZ = shp.ZOrderPosition
    Set pastedShapes = sld.Shapes.Paste
    
    If pastedShapes.Count > 1 Then
        Set pastedShape = pastedShapes.Group
    Else
        Set pastedShape = pastedShapes(1)
    End If
    
    pastedShape.Width = shp.Width
    If pastedShape.LockAspectRatio = 0 Or pastedShape.Height > shp.Height Then
        pastedShape.Height = shp.Height
    End If
    pastedShape.LockAspectRatio = shp.LockAspectRatio
    
    pastedShape.Top = shp.Top
    pastedShape.Left = shp.Left
    pastedShape.Rotation = shp.Rotation
    SetShapeZOrder pastedShape, targetZ
    
    shp.Delete
    pastedShape.Select msoFalse
End Sub

Public Sub ReplaceKeepSize()
    Dim rngSelection As ShapeRange
    Dim masterShape As Shape
    Dim refShape As Shape
    Dim newShape As Shape
    Dim duplicatedRange As ShapeRange
    Dim refs As New Collection
    Dim i As Long
    Dim targetZ As Long
    
    Set rngSelection = GetActiveShapeRange()
    If rngSelection Is Nothing Then Exit Sub
    If rngSelection.Count < 2 Then Exit Sub
    
    Set masterShape = rngSelection(1)

    For i = 2 To rngSelection.Count
        refs.Add rngSelection(i)
    Next i
    
    For i = 1 To refs.Count
        Set refShape = refs(i)
        targetZ = refShape.ZOrderPosition
        
        If i = 1 Then
            Set newShape = masterShape
        Else
            Set duplicatedRange = masterShape.Duplicate()
            Set newShape = duplicatedRange(1)
        End If
        
        newShape.Rotation = refShape.Rotation
        newShape.Width = refShape.Width
        If newShape.LockAspectRatio = 0 Or newShape.Height > refShape.Height Then
            newShape.Height = refShape.Height
        End If
        newShape.Top = refShape.Top
        newShape.Left = refShape.Left
        
        SetShapeZOrder newShape, targetZ
        
        refShape.Delete
        newShape.Select msoFalse
    Next i
End Sub

Private Sub SetShapeZOrder(ByVal shp As Shape, ByVal targetPosition As Long)
    Dim prevPosition As Long
    
    On Error Resume Next
    
    Do While shp.ZOrderPosition > targetPosition
        prevPosition = shp.ZOrderPosition
        shp.ZOrder msoSendBackward
        If shp.ZOrderPosition = prevPosition Then Exit Do
    Loop
    
    Do While shp.ZOrderPosition < targetPosition
        prevPosition = shp.ZOrderPosition
        shp.ZOrder msoBringForward
        If shp.ZOrderPosition = prevPosition Then Exit Do
    Loop
End Sub

Public Sub SetFillTransparency(transp As Single)
    Dim shp As Shape
    Dim shpRange As ShapeRange

    Set shpRange = GetActiveShapeRange()
    If shpRange Is Nothing Then Exit Sub

    On Error Resume Next
    shpRange.Fill.Transparency = transp
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0

    For Each shp In shpRange
        shp.Fill.Transparency = transp
    Next
End Sub

Public Sub SetLineTransparency(transp As Single)
    Dim shp As Shape
    Dim shpRange As ShapeRange

    Set shpRange = GetActiveShapeRange()
    If shpRange Is Nothing Then Exit Sub

    On Error Resume Next
    shpRange.Line.Transparency = transp
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0

    For Each shp In shpRange
        shp.Line.Transparency = transp
    Next
End Sub

Public Sub CleanAuthor()
    activePresentation.BuiltInDocumentProperties.Item("author").value = ""
End Sub

Public Sub CleanSlideMasters()
    Dim i As Long
    Dim j As Long
    Dim deletedLayouts As Integer
    Dim deletedDesigns As Integer
    Dim oPres As Presentation
    Set oPres = activePresentation

    deletedLayouts = 0
    deletedDesigns = 0
    With oPres
        On Error Resume Next
        For i = .Designs.Count To 1 Step -1
            'remove layouts, throws error if in use
            For j = .Designs(i).SlideMaster.CustomLayouts.Count To 1 Step -1
                Err.Clear
                .Designs(i).SlideMaster.CustomLayouts(j).Delete
                If Err.Number = 0 Then
                    deletedLayouts = deletedLayouts + 1
                End If
            Next j
            'remove empty design
            If .Designs(i).SlideMaster.CustomLayouts.Count = 0 Then
                .Designs(i).Delete
                deletedDesigns = deletedDesigns + 1
            End If
        Next i
    End With
    
    MsgBox "Es wurden " & deletedLayouts & " ungenutzte Folienlayouts und " & deletedDesigns & " nicht mehr verwendete Designs gel?scht!", vbInformation
End Sub

Public Sub CleanUnusedDesigns()
    Dim i As Long
    Dim deletedDesigns As Integer
    Dim oPres As Presentation
    Dim usedDesigns() As Boolean

    Set oPres = activePresentation
    deletedDesigns = 0

    If oPres.Designs.Count = 0 Then
        MsgBox "Es wurden 0 nicht verwendete Designs gelšscht!", vbInformation
        Exit Sub
    End If

    ReDim usedDesigns(1 To oPres.Designs.Count)

    On Error Resume Next
    For i = 1 To oPres.Slides.Count
        usedDesigns(oPres.Slides(i).Design.index) = True
    Next i

    For i = oPres.Designs.Count To 1 Step -1
        If Not usedDesigns(i) Then
            Err.Clear
            oPres.Designs(i).Delete
            If Err.Number = 0 Then
                deletedDesigns = deletedDesigns + 1
            End If
        End If
    Next i
    On Error GoTo 0

    MsgBox "Es wurden " & deletedDesigns & " nicht verwendete Designs gelšscht!", vbInformation
End Sub

Sub SendEmailFromSlideSelection()
    Dim sldRange As SlideRange
    Dim newPres As Presentation
    Dim fileName As String, tempFullName As String
    
    Dim OutApp As Object
    Dim OutMail As Object

    ' markierte Folien
    Set sldRange = GetActiveSlideRange()
    If sldRange Is Nothing Then Exit Sub
    
    ' bisheriger Dateiname
    If InStrRev(ActiveWindow.Presentation.Name, ".") = 0 Then
        fileName = ActiveWindow.Presentation.Name
    Else
        fileName = Left(ActiveWindow.Presentation.Name, InStrRev(ActiveWindow.Presentation.Name, ".") - 1)
    End If
    ' Folien (tbd)
    If sldRange.Count = 1 Then
        fileName = fileName & " Folie " & sldRange(1).SlideIndex
    Else
        fileName = fileName & " Folien " & SlideRangeIdentifier(sldRange)
    End If
    ' Zeitstempel
    fileName = fileName & " " & Format(Now, "YYYYMMDDThhnn")
    ' Dateiendung
    fileName = fileName & ".pptx"
    fileName = InputBox("Dateiname eingeben", "Markierte Folien per Mail versenden", fileName)
    ' Bei Abbruch ist Rueckgabewert leer
    If fileName = "" Then Exit Sub
    
    ' Kopie speichern und oeffnen
    tempFullName = Environ("temp") & "\" & fileName
    ActiveWindow.Presentation.SaveCopyAs tempFullName
    Set newPres = Application.Presentations.Open(tempFullName, msoFalse, msoFalse, msoFalse)
    
    ' Folien entfernen, die nicht ausgewaehlt waren
    On Error GoTo ErrorClosePres
    DeleteUnselectedSlides newPres, sldRange
    
    ' Speichern und schliessen
    tempFullName = newPres.FullName
    newPres.Save
    newPres.Saved = msoTrue
    newPres.Close
    
    ' per Mail versenden
'    If Application.Version = "15.0" Then
'        ' PowerPoint 2013
'        SendKeys "%dgoh"
'    Else
'        ' andere Version, auf PowerPoint 2010 getestet
'        SendKeys "%dpt"
'    End If
'    ActivePresentation.sendmail
    
    On Error GoTo ErrorSendEmail
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    With OutMail
        .Subject = fileName
        .Attachments.Add tempFullName, 1 'olByValue
        .Display
    End With
Exit Sub
ErrorClosePres:
    newPres.Saved = msoTrue
    newPres.Close
    MsgBox "Fehler " & Err.Number & ":" & vbNewLine & Err.Description
ErrorSendEmail:
    MsgBox "Fehler " & Err.Number & ":" & vbNewLine & Err.Description
End Sub

Sub CreatePresentationFromSlideSelection()
    Dim sldRange As SlideRange
    Dim newPres As Presentation
    Dim fileName As String
    Dim tempFullName As String
    
    If ActiveWindow.Presentation.Path = "" Then
        MsgBox "Bitte Praesentation erst speichern", vbExclamation
        Exit Sub
    End If
    
    Set sldRange = GetActiveSlideRange()
    If sldRange Is Nothing Then Exit Sub
    fileName = BuildSlideSelectionFileName(sldRange)
    tempFullName = BuildPresentationSelectionCopyPath(ActiveWindow.Presentation, fileName)
    
    ActiveWindow.Presentation.SaveCopyAs tempFullName
    Set newPres = Application.Presentations.Open(tempFullName, msoFalse, msoFalse, msoTrue)
    
    ' Folien entfernen, die nicht ausgewaehlt waren
    DeleteUnselectedSlides newPres, sldRange
    newPres.Save
    newPres.Saved = msoFalse
End Sub

Private Function BuildSlideSelectionFileName(ByRef sldRange As SlideRange) As String
    Dim fileName As String
    
    If InStrRev(ActiveWindow.Presentation.Name, ".") = 0 Then
        fileName = ActiveWindow.Presentation.Name
    Else
        fileName = Left(ActiveWindow.Presentation.Name, InStrRev(ActiveWindow.Presentation.Name, ".") - 1)
    End If
    
    If sldRange.Count = 1 Then
        fileName = fileName & " Folie " & sldRange(1).SlideIndex
    Else
        fileName = fileName & " Folien " & SlideRangeIdentifier(sldRange)
    End If
    
    fileName = fileName & " " & Format(Now, "YYYYMMDDThhnn") & ".pptx"
    BuildSlideSelectionFileName = fileName
End Function

Private Function BuildPresentationSelectionCopyPath(ByRef sourcePresentation As Presentation, ByVal fileName As String) As String
    Dim pathSeparator As String
    
    #If Mac Then
        pathSeparator = "/"
    #Else
        pathSeparator = "\"
    #End If
    
    BuildPresentationSelectionCopyPath = sourcePresentation.Path & pathSeparator & fileName
End Function

Sub ApplyThemeFromFile()
    #If Mac Then
        'not supported by mac
    #Else
    With Application.FileDialog(msoFileDialogOpen)
        '.Title "PowerPoint-Datei auswŠhlen"
        .Show
        If .SelectedItems.Count = 0 Then
            Exit Sub
        Else
            ActiveWindow.Presentation.ApplyTemplate .SelectedItems(1)
        End If
    End With
    #End If
End Sub


' Blendet Foliennummerierungen ein bzw. aus
' Auf jeder Folie wird eine Textbox mit der Foliennr. eingef?gt, welche sich
' bei Umsortierungen der Folien nicht ?ndert und dadurch Diskussionen in Teams vereinfacht
Public Sub ToggleSlideNumbering()
    Dim sld As Slide
    Dim shp As Shape
    Dim hasNumbering As Boolean
    
    hasNumbering = False
    
    ' Alle Shapes in allen Slides durchlaufen
    For Each sld In activePresentation.Slides
        For Each shp In sld.Shapes
            ' Shape mit SlideNumberTag gefunden
            If shp.Tags.Item(SLIDENUMBERING) = SLIDENUMBERING Then
                hasNumbering = True
                Exit For
            End If
        Next
        If hasNumbering Then Exit For
    Next
    
    If hasNumbering Then
        RemoveSlideNumbers
    Else
        AddSlideNumbers
    End If
End Sub

' Auf jeder Folie wird eine Textbox mit der Foliennr. eingef?gt
Private Sub AddSlideNumbers()
    Dim sld As Slide
    Dim shp As Shape
    
    ' Alle Slides durchlaufen
    For Each sld In activePresentation.Slides
        Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 100, 100)
        shp.TextFrame.TextRange.Font.Size = 32
        shp.TextFrame.TextRange.Font.Bold = msoTrue
        shp.TextFrame.TextRange.Font.Color = 192 + 0 * 256 + 0 * 256 * CDbl(256)
        shp.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignRight
        shp.TextFrame.TextRange.text = sld.SlideNumber
        shp.TextFrame.MarginBottom = 0
        shp.TextFrame.MarginTop = 0
        shp.TextFrame.MarginRight = 0
        shp.TextFrame.MarginLeft = 0
        shp.Left = activePresentation.PageSetup.SlideWidth - shp.Width - 15
        shp.Top = 15
        shp.Tags.Add SLIDENUMBERING, SLIDENUMBERING
    Next
End Sub

' Textboxen mit den Foliennummern werden von allen Folien entfernt
Private Sub RemoveSlideNumbers()
    Dim sld As Slide
    Dim shp As Shape
    
    ' Alle Shapes in allen Slides durchlaufen
    For Each sld In activePresentation.Slides
        For Each shp In sld.Shapes
            ' Shape mit SlideNumberTag loeschen
            If shp.Tags.Item(SLIDENUMBERING) = SLIDENUMBERING Then
                shp.Delete
                Exit For
            End If
        Next
    Next
End Sub

' Sprache f?r gesamte Pr?sentation setzen
Public Sub setLanguage(ByVal langCode As Integer)
    On Error Resume Next
    Dim sld As Slide
    Dim shp As Shape


    #If Mac Then
        'not supported by mac
    #Else
        activePresentation.DefaultLanguageID = langCode
    #End If
    
    ' Alle Shapes in allen Slides durchlaufen
    For Each sld In activePresentation.Slides
        For Each shp In sld.Shapes
            setLanguageForShape shp, langCode
'            If shp.HasTextFrame Then
'                shp.TextFrame2.TextRange.LanguageID = langCode
'            End If
        Next
    Next
End Sub

' Sprache je Shape setzen
Private Sub setLanguageForShape(shp As Shape, langCode As Integer)
    Dim i As Long, row As Long, col As Long
    
    On Error GoTo Err_Handler

    If shp.HasTextFrame Then
        shp.TextFrame2.TextRange.LanguageID = langCode
    End If
    
    If shp.HasTable Then
        For row = 1 To shp.Table.Rows.Count
            For col = 1 To shp.Table.Columns.Count
                shp.Table.Cell(row, col).Shape.TextFrame2.TextRange.LanguageID = langCode
            Next
        Next
    End If
    
    Select Case shp.Type
        Case msoGroup
            For i = 1 To shp.GroupItems.Count
                setLanguageForShape shp.GroupItems.Item(i), langCode
            Next i
        Case msoSmartArt
            For i = 1 To shp.SmartArt.AllNodes.Count
                shp.SmartArt.AllNodes.Item(i).TextFrame2.TextRange.LanguageID = langCode
            Next i
        Case msoChart, msoDiagram
            If shp.Chart.HasTitle Then
                shp.Chart.ChartTitle.Format.TextFrame2.TextRange.LanguageID = langCode
            End If
    End Select

Exit Sub
Err_Handler:
End Sub


Public Sub MultiplyShapes(RowsCols As Integer, Sep As Single, Optional vertical As Boolean = False)
    Dim i As Long
    Dim shp As Shape, newShp As ShapeRange
    Dim shpRange As ShapeRange
    
    Set shpRange = GetActiveShapeRange()
    If shpRange Is Nothing Then Exit Sub

    For Each shp In shpRange
        For i = 1 To RowsCols - 1
            Set newShp = shp.Duplicate
            
            If vertical = False Then
                newShp.Left = shp.Left + i * (shp.Width + Sep)
                newShp.Top = shp.Top
            Else
                newShp.Top = shp.Top + i * (shp.Height + Sep)
                newShp.Left = shp.Left
            End If
            newShp.Select False
        Next i
    Next
End Sub


Public Sub SplitShapes(RowsCols As Integer, Sep As Single, Optional vertical As Boolean = False)
    Dim i As Long
    Dim shp As Shape, newShp As ShapeRange
    Dim targetSize As Single
    Dim shpRange As ShapeRange
    
    Set shpRange = GetActiveShapeRange()
    If shpRange Is Nothing Then Exit Sub

    For Each shp In shpRange
        If vertical = False Then
            targetSize = (shp.Width - (RowsCols - 1) * Sep) / RowsCols
            shp.Width = targetSize
        Else
            targetSize = (shp.Height - (RowsCols - 1) * Sep) / RowsCols
            shp.Height = targetSize
        End If
        For i = 1 To RowsCols - 1
            Set newShp = shp.Duplicate
            
            If vertical = False Then
                newShp.Left = shp.Left + i * (targetSize + Sep)
                newShp.Top = shp.Top
                newShp.Width = targetSize
            Else
                newShp.Top = shp.Top + i * (targetSize + Sep)
                newShp.Left = shp.Left
                newShp.Height = targetSize
            End If
            newShp.Select False
        Next i
    Next
End Sub


Public Sub ArrangeByLast(Position As Integer)
    Dim selectedShapes As ShapeRange
    Dim shp As Shape, masterShp As Shape
    Dim masterTop As Single, masterHeight As Single
    Dim masterLeft As Single, masterWidth As Single
    Dim masterBottom As Single, masterRight As Single
    Dim masterCenterX As Single, masterCenterY As Single
    Dim shpTop As Single, shpHeight As Single
    Dim shpLeft As Single, shpWidth As Single
    Dim dx As Single, dy As Single
    
    Set selectedShapes = GetActiveShapeRange()
    If selectedShapes Is Nothing Then Exit Sub
    Set masterShp = selectedShapes(selectedShapes.Count)
    
    GetVisualBounds masterShp, masterLeft, masterTop, masterWidth, masterHeight
    masterBottom = masterTop + masterHeight
    masterRight = masterLeft + masterWidth
    masterCenterX = masterLeft + (masterWidth / 2)
    masterCenterY = masterTop + (masterHeight / 2)
    
    For Each shp In selectedShapes
        GetVisualBounds shp, shpLeft, shpTop, shpWidth, shpHeight
        dx = 0
        dy = 0
        
        Select Case Position
            Case 1
                dy = masterTop - shpTop
            Case 2
                dy = masterBottom - (shpTop + shpHeight)
            Case 3
                dx = masterLeft - shpLeft
            Case 4
                dx = masterRight - (shpLeft + shpWidth)
            Case 5
                dy = masterCenterY - (shpTop + shpHeight / 2)
            Case 6
                dx = masterCenterX - (shpLeft + shpWidth / 2)
        End Select
        
        If dx <> 0 Then shp.Left = shp.Left + dx
        If dy <> 0 Then shp.Top = shp.Top + dy
    Next
End Sub

Public Sub StretchByLast(Position As Integer)
    Dim selectedShapes As ShapeRange
    Dim shp As Shape, masterShp As Shape
    Dim masterTop As Single, masterHeight As Single
    Dim masterLeft As Single, masterWidth As Single
    Dim masterBottom As Single, masterRight As Single
    Dim shpTop As Single, shpHeight As Single
    Dim shpLeft As Single, shpWidth As Single
    
    Set selectedShapes = GetActiveShapeRange()
    If selectedShapes Is Nothing Then Exit Sub
    If selectedShapes.Count < 2 Then Exit Sub
    
    Set masterShp = selectedShapes(selectedShapes.Count)
    GetVisualBounds masterShp, masterLeft, masterTop, masterWidth, masterHeight
    masterBottom = masterTop + masterHeight
    masterRight = masterLeft + masterWidth
    
    For Each shp In selectedShapes
        If shp.Id <> masterShp.Id Then
            GetVisualBounds shp, shpLeft, shpTop, shpWidth, shpHeight
            
            Select Case Position
                Case 1
                    StretchShapeToVisualWidth shp, shpLeft + shpWidth - masterLeft, True
                Case 2
                    StretchShapeToVisualWidth shp, masterRight - shpLeft, False
                Case 3
                    StretchShapeToVisualHeight shp, shpTop + shpHeight - masterTop, True
                Case 4
                    StretchShapeToVisualHeight shp, masterBottom - shpTop, False
            End Select
        End If
    Next
End Sub

Private Sub GetVisualBounds(ByVal shp As Shape, ByRef Left As Single, ByRef Top As Single, ByRef Width As Single, ByRef Height As Single)
    Dim angleRad As Double
    Dim cosA As Double, sinA As Double
    Dim bbWidth As Double, bbHeight As Double
    Dim cx As Double, cy As Double
    
    angleRad = shp.Rotation * (3.14159265358979 / 180#)
    cosA = Abs(Cos(angleRad))
    sinA = Abs(Sin(angleRad))
    
    bbWidth = shp.Width * cosA + shp.Height * sinA
    bbHeight = shp.Width * sinA + shp.Height * cosA
    
    cx = shp.Left + shp.Width / 2
    cy = shp.Top + shp.Height / 2
    
    Left = cx - bbWidth / 2
    Top = cy - bbHeight / 2
    Width = bbWidth
    Height = bbHeight
End Sub

Private Sub StretchShapeToVisualWidth(ByVal shp As Shape, ByVal targetVisualWidth As Single, ByVal keepRight As Boolean)
    Dim originalLeft As Single, originalTop As Single
    Dim originalWidth As Single, originalHeight As Single
    Dim angleRad As Double
    Dim cosA As Double, sinA As Double
    Dim targetSize As Double
    Dim newLeft As Single, newTop As Single
    Dim newWidth As Single, newHeight As Single
    
    If targetVisualWidth < 1 Then targetVisualWidth = 1
    
    GetVisualBounds shp, originalLeft, originalTop, originalWidth, originalHeight
    angleRad = shp.Rotation * (3.14159265358979 / 180#)
    cosA = Abs(Cos(angleRad))
    sinA = Abs(Sin(angleRad))
    
    If cosA >= sinA And cosA > 0.0001 Then
        targetSize = (targetVisualWidth - sinA * shp.Height) / cosA
        shp.Width = Max(1, targetSize)
    ElseIf sinA > 0.0001 Then
        targetSize = (targetVisualWidth - cosA * shp.Width) / sinA
        shp.Height = Max(1, targetSize)
    Else
        Exit Sub
    End If
    
    GetVisualBounds shp, newLeft, newTop, newWidth, newHeight
    If keepRight Then
        shp.Left = shp.Left + (originalLeft + originalWidth) - (newLeft + newWidth)
    Else
        shp.Left = shp.Left + (originalLeft - newLeft)
    End If
End Sub

Private Sub StretchShapeToVisualHeight(ByVal shp As Shape, ByVal targetVisualHeight As Single, ByVal keepBottom As Boolean)
    Dim originalLeft As Single, originalTop As Single
    Dim originalWidth As Single, originalHeight As Single
    Dim angleRad As Double
    Dim cosA As Double, sinA As Double
    Dim targetSize As Double
    Dim newLeft As Single, newTop As Single
    Dim newWidth As Single, newHeight As Single
    
    If targetVisualHeight < 1 Then targetVisualHeight = 1
    
    GetVisualBounds shp, originalLeft, originalTop, originalWidth, originalHeight
    angleRad = shp.Rotation * (3.14159265358979 / 180#)
    cosA = Abs(Cos(angleRad))
    sinA = Abs(Sin(angleRad))
    
    If cosA >= sinA And cosA > 0.0001 Then
        targetSize = (targetVisualHeight - sinA * shp.Width) / cosA
        shp.Height = Max(1, targetSize)
    ElseIf sinA > 0.0001 Then
        targetSize = (targetVisualHeight - cosA * shp.Height) / sinA
        shp.Width = Max(1, targetSize)
    Else
        Exit Sub
    End If
    
    GetVisualBounds shp, newLeft, newTop, newWidth, newHeight
    If keepBottom Then
        shp.Top = shp.Top + (originalTop + originalHeight) - (newTop + newHeight)
    Else
        shp.Top = shp.Top + (originalTop - newTop)
    End If
End Sub


