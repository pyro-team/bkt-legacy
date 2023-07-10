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
    Set shp = sld.shapes.AddTextbox(msoTextOrientationHorizontal, 0, 60, 100, 20)
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
    With sld.shapes.AddConnector(Type:=msoConnectorStraight, BeginX:=0, _
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
    With sld.shapes.AddConnector(Type:=msoConnectorStraight, BeginX:=0, _
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
    
    ' Text auswählen
    shp.Select
    shp.TextFrame.TextRange.Select
End Sub





Public Sub AddConnectorShapeLeftRight()
    Dim shpLeft As Shape
    Dim shpRight As Shape
    Dim shpConnector As Shape
    Dim shapes As Variant
    
    On Error GoTo Err_Handler
    
    shapes = ActiveWindowSelectionSortedByLeft()
    
    Set shpLeft = shapes(1)
    Set shpRight = shapes(2)
    
    Set shpConnector = ActiveWindow.View.Slide.shapes.AddShape(msoShapeRectangle, shpLeft.Left + shpLeft.Width, shpLeft.Top, shpRight.Left - shpLeft.Left - shpLeft.Width, shpLeft.Height)
    
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
    Dim shapes As Variant
    
    On Error GoTo Err_Handler
    
    shapes = ActiveWindowSelectionSortedByTop()
    
    Set shpTop = shapes(1)
    Set shpBottom = shapes(2)
    
    Set shpConnector = ActiveWindow.View.Slide.shapes.AddShape(msoShapeRectangle, shpTop.Left, shpTop.Top + shpTop.Height, shpTop.Width, shpBottom.Top - shpTop.Top - shpTop.Height)
    
    ' node 3: bottom right
    shpConnector.Nodes.SetPosition 3, shpBottom.Left + shpBottom.Width, shpBottom.Top
    ' node 4: bottom left
    shpConnector.Nodes.SetPosition 4, shpBottom.Left, shpBottom.Top
    shpConnector.Fill.ForeColor.RGB = 12566463 '193
    shpConnector.Line.ForeColor.RGB = 8355711 ' 127 127 127
    shpConnector.Line.Weight = 0.75
Err_Handler:
End Sub


Public Sub SetSameHeight(Optional func As String = "Max")
    Dim shp As Shape
    Dim Height As Single
    
    Height = ActiveWindow.selection.ShapeRange(1).Height
    For Each shp In ActiveWindow.selection.ShapeRange
'        height = Max(height, shp.height)
        Height = Application.Run(func, Height, shp.Height)
    Next
    For Each shp In ActiveWindow.selection.ShapeRange
        shp.Height = Height
    Next
End Sub

Public Sub SetSameWidth(Optional func As String = "Max")
    Dim shp As Shape
    Dim Width As Single
    
    Width = ActiveWindow.selection.ShapeRange(1).Width
    For Each shp In ActiveWindow.selection.ShapeRange
'        width = Max(width, shp.width)
        Width = Application.Run(func, Width, shp.Width)
    Next
    For Each shp In ActiveWindow.selection.ShapeRange
        shp.Width = Width
    Next
End Sub

Public Sub SwapPosition()
    Dim shape1 As Shape
    Dim shape2 As Shape
    Dim tmpLeft As Single
    Dim tmpTop As Single
    
    Set shape1 = ActiveWindow.selection.ShapeRange(1)
    Set shape2 = ActiveWindow.selection.ShapeRange(2)
    
    tmpLeft = shape1.Left
    tmpTop = shape1.Top
    shape1.Left = shape2.Left
    shape1.Top = shape2.Top
    shape2.Left = tmpLeft
    shape2.Top = tmpTop
End Sub

Public Sub SwapPositionSize()
    Dim shape1 As Shape
    Dim shape2 As Shape
    Dim tmpWidth As Single
    Dim tmpHeight As Single
    
    Set shape1 = ActiveWindow.selection.ShapeRange(1)
    Set shape2 = ActiveWindow.selection.ShapeRange(2)
    
    tmpWidth = shape1.Width
    tmpHeight = shape1.Height
    shape1.Width = shape2.Width
    shape1.Height = shape2.Height
    shape2.Width = tmpWidth
    shape2.Height = tmpHeight
    
    SwapPosition
End Sub


' Bei allen markierten Shapes werden die Texte in separate Textboxen kopiert
Public Sub MoveTextOutOfShapes()
    Dim shp As Shape
    Dim shpTxt As Shape
    
    For Each shp In ActiveWindow.selection.ShapeRange
        If shp.HasTextFrame And shp.TextFrame.HasText Then   'shp.TextFrame.TextRange.text <> "" Then
            Set shpTxt = ActiveWindow.View.Slide.shapes.AddTextbox(msoTextOrientationHorizontal, shp.Left, shp.Top, shp.Width, shp.Height)
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
            ' Größe wiederherstellen
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
    
    If ActiveWindow.selection.ShapeRange.Count <> 2 Then
        MsgBox "Bitte eine Textbox und ein Shape-Objekt auswählen.", vbInformation
        Exit Sub
    End If
    
    If ActiveWindow.selection.ShapeRange(1).Type = msoTextBox Then
        Set shpTxt = ActiveWindow.selection.ShapeRange(1)
        Set shp = ActiveWindow.selection.ShapeRange(2)
    Else
        Set shpTxt = ActiveWindow.selection.ShapeRange(2)
        Set shp = ActiveWindow.selection.ShapeRange(1)
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
    Dim selection As ShapeRange
    
    Set selection = ActiveWindow.selection.ShapeRange
    For Each shp In selection
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
                
                ' --> ein Absatz bleibt übrig
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
            ActiveWindow.selection.ShapeRange.Distribute msoDistributeVertically, msoFalse
        End If
    Next
End Sub

' Text der markierte Shapes wird im ersten Shape zusammengefassen,
' andere Shapes werden geloescht.
' Reihenfolge abhaengig von Shape.top
Public Sub JoinShapesWithText()
    Dim shapes As Variant
    Dim index As Long
    Dim txtRange As TextRange
    Dim shp As Shape
    Dim parCount As Long
    
    ' Shapes nach top sortieren
    shapes = ActiveWindowSelectionSortedByTop
    ' Anapssung Größe des ersten Shapes (Master-Shape)
    Set shp = shapes(1)
    shp.Height = Max(shp.Height, shapes(UBound(shapes)).Top + shapes(UBound(shapes)).Height - shp.Top)
    
    For index = 2 To UBound(shapes)
        ' Text aus Shape kopieren
        shapes(index).TextFrame.TextRange.Copy
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
        shapes(index).Delete
    Next
End Sub


Public Sub ReplaceAllText()
    Dim newText As String
    Dim shp As Shape
    Dim selection As ShapeRange
    Set selection = ActiveWindow.selection.ShapeRange
    On Error Resume Next
    
    newText = InputBox("Neuen Text eingeben", "Text ersetzen", "tbd")
    ' Bei Abbruch ist Rückgabewert leer
    If newText = "" Then Exit Sub
    
    For Each shp In selection
        If shp.HasTextFrame Then
            shp.TextFrame.TextRange.text = newText
        End If
    Next
End Sub

Public Sub RemoveAllText()
    Dim shp As Shape
    Dim selection As ShapeRange
    Set selection = ActiveWindow.selection.ShapeRange
    On Error Resume Next
    For Each shp In selection
        If shp.HasTextFrame Then
            shp.TextFrame.TextRange.Delete
        End If
    Next
End Sub

Public Sub TextMarginZero()
    Dim shp As Shape
    Dim selection As ShapeRange
    Set selection = ActiveWindow.selection.ShapeRange
    On Error Resume Next
    For Each shp In selection
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
    Dim selection As ShapeRange
    
    Set selection = ActiveWindow.selection.ShapeRange
    For Each shp In selection
        shp.Visible = msoFalse
    Next
End Sub


Public Sub ShowShapes()
    Dim shp As Shape
    Dim allShapes As shapes
    
    Set allShapes = ActiveWindow.View.Slide.shapes
    ActiveWindow.selection.Unselect
    For Each shp In allShapes
        If shp.Visible = msoFalse Then
            shp.Visible = msoTrue
            shp.Select msoFalse
        End If
    Next
End Sub


Public Sub PasteOnSlides()
    Dim sld As Slide
    Dim selection As SlideRange
    
    Set selection = ActiveWindow.selection.SlideRange
    For Each sld In selection
        sld.shapes.Paste
    Next
End Sub

Public Sub PasteAndReplace()
    Dim shp As Shape
    Dim selection As ShapeRange
    Set selection = ActiveWindow.selection.ShapeRange
    
    'On Error Resume Next
    For Each shp In selection
        PasteAndReplaceShape shp
    Next
End Sub

Private Sub PasteAndReplaceShape(shp As Shape)
    Dim sld As Slide
    Dim pastedShape As ShapeRange
    
    Set sld = ActiveWindow.View.Slide
    Set pastedShape = sld.shapes.Paste
    
'    If pastedShape.Count > 1 Then
'        Set pastedShape = pastedShape.Group
'    End If
    
    pastedShape.Width = shp.Width
    If pastedShape.LockAspectRatio = 0 Or pastedShape.Height > shp.Height Then
        pastedShape.Height = shp.Height
    End If
    pastedShape.LockAspectRatio = shp.LockAspectRatio
    
    pastedShape.Top = shp.Top
    pastedShape.Left = shp.Left
    pastedShape.Rotation = shp.Rotation
    
    shp.Delete
    pastedShape.Select msoFalse
End Sub

Public Sub SetFillTransparency(transp As Single)
    Dim shp As Shape
    
    For Each shp In ActiveWindow.selection.ShapeRange
        shp.Fill.Transparency = transp
    Next
End Sub

Public Sub SetLineTransparency(transp As Single)
    Dim shp As Shape
    
    For Each shp In ActiveWindow.selection.ShapeRange
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
    
    MsgBox "Es wurden " & deletedLayouts & " ungenutzte Folienlayouts und " & deletedDesigns & " nicht mehr verwendete Designs gelöscht!", vbInformation
End Sub

Sub SendEmailFromSlideSelection()
    Dim sldRange As SlideRange
    Dim newPres As Presentation
    Dim fileName As String, tempFullName As String
    
    Dim OutApp As Object
    Dim OutMail As Object

    ' markierte Folien
    Set sldRange = ActiveWindow.selection.SlideRange
    
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
    ' Bei Abbruch ist Rückgabewert leer
    If fileName = "" Then Exit Sub
    
    ' Kopie speichern und öffnen
    tempFullName = Environ("temp") & "\" & fileName
    ActiveWindow.Presentation.SaveCopyAs tempFullName
    Set newPres = Application.Presentations.Open(tempFullName, msoFalse, msoFalse, msoFalse)
    
    ' Folien entfernen, die nicht ausgewählt waren
    On Error GoTo ErrorClosePres
    DeleteUnselectedSlides newPres, sldRange
    
    ' Speichern und schließen
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
    
    If ActiveWindow.Presentation.Path = "" Then
        MsgBox "Bitte Präsentation erst speichern", vbExclamation
        Exit Sub
    End If
    
    Set sldRange = ActiveWindow.selection.SlideRange
    fileName = ActiveWindow.Presentation.FullName
    
    ' Kopie öffnen
    Set newPres = Application.Presentations.Open(fileName, msoFalse, msoTrue, msoTrue)
    
    ' Folien entfernen, die nicht ausgewählt waren
    DeleteUnselectedSlides newPres, sldRange
End Sub

Sub ApplyThemeFromFile()
    ' File dialog not supported on mac
    With Application.FileDialog(msoFileDialogOpen)
        '.Title "PowerPoint-Datei auswählen"
        .Show
        If .SelectedItems.Count = 0 Then
            Exit Sub
        Else
            ActiveWindow.Presentation.ApplyTemplate .SelectedItems(1)
        End If
    End With
End Sub


' Blendet Foliennummerierungen ein bzw. aus
' Auf jeder Folie wird eine Textbox mit der Foliennr. eingefügt, welche sich
' bei Umsortierungen der Folien nicht ändert und dadurch Diskussionen in Teams vereinfacht
Public Sub ToggleSlideNumbering()
    Dim sld As Slide
    Dim shp As Shape
    Dim hasNumbering As Boolean
    
    hasNumbering = False
    
    ' Alle Shapes in allen Slides durchlaufen
    For Each sld In activePresentation.Slides
        For Each shp In sld.shapes
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

' Auf jeder Folie wird eine Textbox mit der Foliennr. eingefügt
Private Sub AddSlideNumbers()
    Dim sld As Slide
    Dim shp As Shape
    
    ' Alle Slides durchlaufen
    For Each sld In activePresentation.Slides
        Set shp = sld.shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 100, 100)
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
        For Each shp In sld.shapes
            ' Shape mit SlideNumberTag loeschen
            If shp.Tags.Item(SLIDENUMBERING) = SLIDENUMBERING Then
                shp.Delete
                Exit For
            End If
        Next
    Next
End Sub

' Sprache für gesamte Präsentation setzen
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
        For Each shp In sld.shapes
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
    
    For Each shp In ActiveWindow.selection.ShapeRange
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
    
    For Each shp In ActiveWindow.selection.ShapeRange
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
