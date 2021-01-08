VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ToolboxAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const TOOLBOX_AGENDA = "Toolbox-Agenda"
Private Const TOOLBOX_AGENDA_SlideNo = "Toolbox-Agenda-SlideNo"
Private Const TOOLBOX_AGENDA_Selector = "Toolbox-Agenda-Selector"
Private Const TOOLBOX_AGENDA_TextBox = "Toolbox-Agenda-TextBox"


Private selectorFillColor As Single
Private selectorLineColor As Single


Private Sub Class_Initialize()
    selectorFillColor = RGB(193, 193, 193)
    selectorLineColor = RGB(127, 127, 127)
End Sub


Public Sub CreateAgendaTextBoxOnSlide(ByVal sld As Slide)
    Dim shp As Shape
    
    On Error GoTo Err_Handler
    
    'sld.shapes(1).TextFrame.TextRange.text = "Agenda"
    ' Shape rechts oben auf slide erstellen
    Set shp = sld.shapes.AddShape(msoShapeRectangle, 160, 240, 400, 100)
    ' Shape-Typ ist links-rechts-Pfeil, weil es die passenden Connector-Ecken hat
    shp.AutoShapeType = msoShapeLeftRightArrow
    ' Shape-Anpassung, so dass es wie ein Rechteck aussieht
    shp.Adjustments(1) = 1
    shp.Adjustments(2) = 0
    ' Shape-Stil
    shp.Line.Weight = 0.75
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
    ' Text-Stil
    shp.TextFrame.TextRange.Font.Color.RGB = 0
    shp.TextFrame.TextRange.Font.Size = 14
    shp.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft
    ' Autosize / Textumbruch
    shp.TextFrame.WordWrap = msoTrue
    shp.TextFrame.AutoSize = ppAutoSizeShapeToFitText
    shp.Width = 400
    ' Text
    ' Note: vbNewline does not work properly on Mac! (even though is should work https://docs.microsoft.com/de-de/office/vba/language/reference/user-interface-help/miscellaneous-constants)
    '       New bullets are inserted but in VBA the whole text is recognized as 1 paragraph, so agenda update fails
    shp.TextFrame.TextRange.text = "Abschnitt 1" & vbCr & "Abschnitt 2" & vbCr & "Abschnitt 3"
    shp.TextFrame.VerticalAnchor = msoAnchorTop
    shp.TextFrame.Ruler.Levels(1).FirstMargin = 0
    shp.TextFrame.Ruler.Levels(1).LeftMargin = 14
    ' Innenabstand
    shp.TextFrame.MarginBottom = 12
    shp.TextFrame.MarginTop = 12
    shp.TextFrame.MarginLeft = 6
    shp.TextFrame.MarginRight = 6
    shp.TextFrame.TextRange.ParagraphFormat.Bullet.Type = ppBulletUnnumbered
    shp.TextFrame.TextRange.ParagraphFormat.Bullet.Character = 167
    shp.TextFrame.TextRange.ParagraphFormat.Bullet.Font.Name = "Wingdings"
    shp.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 18
    
    ' Connectoren erstellen und mit Connector-Ecken des Shapes verbinden
    With sld.shapes.AddConnector(Type:=msoConnectorStraight, BeginX:=0, _
            BeginY:=0, EndX:=100, EndY:=100)
        .ConnectorFormat.BeginConnect ConnectedShape:=shp, ConnectionSite:=1
        .ConnectorFormat.EndConnect ConnectedShape:=shp, ConnectionSite:=3
        .Line.ForeColor.RGB = 0
        .Line.Weight = 0.75
    End With
    With sld.shapes.AddConnector(Type:=msoConnectorStraight, BeginX:=0, _
            BeginY:=0, EndX:=100, EndY:=100)
        .ConnectorFormat.BeginConnect ConnectedShape:=shp, ConnectionSite:=5
        .ConnectorFormat.EndConnect ConnectedShape:=shp, ConnectionSite:=7
        .Line.ForeColor.RGB = 0
        .Line.Weight = 0.75
    End With
    
    
    SetTagsForSlide sld, 0
    SetTagsForTextbox shp
    
    Exit Sub
Err_Handler:
    MsgBox "Fehler " & Err.Number & ":" & vbNewLine & Err.Description
End Sub


Public Sub CreateOrUpdateAgenda()
    If CanUpdateAgenda Then
        ' vorhandene Agenda aktualisieren
        UpdateAgenda
    Else
        CreateFromCurrentSlide
    End If
End Sub

Public Sub CreateFromCurrentSlide()
    Dim shp As Shape
    Dim sld As Slide
    Dim answer As Integer
    
    On Error GoTo Err_Handler
    
    
    ' In normaler Ansicht
    If ActiveWindow.View.Type <> ppViewNormal Then
        MsgBox "In Slide-View wechseln!"
        Exit Sub
    End If
    
    ' Aktuelles Slide
    Set sld = ActiveWindow.View.Slide
    
    ' Markiertes Shape oder erstes Shape mit Bullet-Point-Liste auswaehlen
    If ActiveWindow.selection.Type = ppSelectionShapes Or ActiveWindow.selection.Type = ppSelectionText Then
        Set shp = ActiveWindow.selection.ShapeRange(1)
    Else
        For Each shp In sld.shapes
            If shp.TextFrame.HasText Then
                If shp.TextFrame.TextRange.ParagraphFormat.Bullet.Type = ppBulletUnnumbered Then
                    Exit For
                End If
            End If
        Next
        
        If shp Is Nothing Then
            MsgBox "Textbox mit Agenda-Einträgen auswählen oder als Bullet-Liste formatieren."
            Exit Sub
        End If
    End If
    
    
    If Not shp Is Nothing Then
        If sld.Tags.Item(TOOLBOX_AGENDA) <> "" Then
            ' Vorhandene Meta-Informationen einer anderen Agenda werden gelöscht.
            answer = MsgBox("Es gibt bereits andere Agenda-Slides in der Präsentation." & vbNewLine & _
                            "Die vorhandenen Meta-Informationen werden gelöscht bevor die neue Agenda erstellt wird." & vbNewLine & vbNewLine & _
                            "Meta-Informationen löschen und fortfahren?", vbYesNo, "Meta-Informationen der bisherigen Agenda werden gelöscht")
            If answer = vbNo Then
                Exit Sub
            End If
            
            RemoveAgendaB False
        End If
        
        CreateAgenda sld, shp
    End If
    
    Exit Sub
Err_Handler:
    If Err.Number = -2 Then
    Else
        MsgBox "Fehler " & Err.Number & ":" & vbNewLine & Err.Description
    End If
End Sub


' Agenda aus aktuellem Slide auf Basis der angegebenen textBox erstellen
Private Sub CreateAgenda(ByVal sld As Slide, ByVal textbox As Shape)
    SetTagsForSlide sld, 0
    SetTagsForTextbox textbox
    UpdateAgendaB True
End Sub

Public Function CanUpdateAgenda() As Boolean
    ' Aktualisieren oder Löschen nur möglich, wenn es Slides mit Agenda-Meta-Informationen gibt
    CanUpdateAgenda = (GetAgendaSlides().Count > 0)
End Function

Public Sub UpdateAgenda()

    If Not CanUpdateAgenda Then Exit Sub
    
    ' Neue Kopien der ersten Agenda-Folie erstellen
    DetermineSelectorSettings
    UpdateAgendaB True
    
'    Dim answer As Integer
'
'    answer = MsgBox("Neue Kopien vom ersten Agenda-Slide erstennen?" & vbNewLine & vbNewLine & _
'                    "In diesem Fall gehen alle Änderungen an den anderen Agenda-Slides verloren." & vbNewLine & _
'                    "Ansonsten werden nur Agenda-Texte, Agenda-Position und Markierungsposition aktualisiert.", _
'                    vbYesNoCancel, "Neue Kopien vom ersten Agenda-Slide erstennen?")
'
'    Select Case answer
'    Case vbCancel
'        Exit Sub
'    Case vbYes
'        ' Neue Kopien der ersten Agenda-Folie erstellen
'        UpdateAgendaB True
'    Case vbNo
'        ' Agenda aktualisieren ohne neue Kopien vom ersten Agenda-Slide zu machen
'        ' Aktualisiert nur Agenda-Texte, Agenda-Position und Markierungsposition
'        UpdateAgendaB False
'    End Select
'
End Sub

Private Sub DetermineSelectorSettings()
    Dim Slides As Collection
    Dim sld As Slide
    
    Set Slides = GetAgendaSlides
    For Each sld In Slides
        Set shp = GetShapeWithTagItem(sld, TOOLBOX_AGENDA_Selector)
        If Not shp Is Nothing Then
            selectorFillColor = shp.Fill.ForeColor.RGB
            selectorLineColor = shp.Line.ForeColor.RGB
            Exit Sub
        End If
    Next
    
End Sub



Private Sub UpdateAgendaB(ByVal recreateFromMaster As Boolean)
    Dim Slides As Collection
    Dim idx As Integer
    Dim sld As Slide
    Dim sldMaster As Slide
    Dim sldCopy As Slide
    Dim textbox As Shape
    Dim newSldCount As Integer
    
    On Error GoTo Err_Handler
    
    ' Textbox auf erstem Agenda-Slide holen
    newSldCount = 0
    Set Slides = GetAgendaSlides
    If Slides.Count = 0 Then Exit Sub
    
    Set sldMaster = Slides(1)
    Set textbox = GetAgendaTextBoxOnSlide(sldMaster)
    If textbox Is Nothing Then
        MsgBox "Update nicht möglich! Agenda-Textbox fehlt auf erstem Agenda-Slide."
        Exit Sub
    End If
    
    ' Pro Absatz das zugehörige Slide aktualisieren (bzw. neu erstellen)
    For idx = 1 To textbox.TextFrame.TextRange.Paragraphs.Count
        Set sld = Nothing
        On Error Resume Next
        Set sld = Slides("slide-" & idx)
        On Error GoTo Err_Handler
        
        If TrimParagraphText(textbox.TextFrame.TextRange.Paragraphs(idx).text) <> "" Then
            If sld Is Nothing Then
                ' Slide nicht vorhanden, neu erstellen und nach Master-Agenda-Slide einfuegen
                newSldCount = newSldCount + 1
                Set sld = sldMaster.Duplicate(1)
                sld.MoveTo sldMaster.SlideIndex + newSldCount
                SetTagsForSlide sld, idx
            Else
                If recreateFromMaster Then
                    If sld.SlideID <> sldMaster.SlideID Then
                        ' Slide neu erstellen erzwungen, alte Position wird beibehalten
                        Set sld = sldMaster.Duplicate(1)
                        sld.MoveTo Slides("slide-" & idx).SlideIndex
                        Slides("slide-" & idx).Delete
                        SetTagsForSlide sld, idx
                    End If
                Else
                    MsgBox "Not implemented yet!", vbCritical
                    Exit Sub
                End If
                Slides.Remove "slide-" & idx
            End If
            ' Agenda aktualisieren
            UpdateAgendaOnSlide sld, textbox, idx
        End If
    Next
    
    On Error Resume Next
    Slides.Remove "slide-0"
    On Error GoTo Err_Handler
        
    For Each sld In Slides
        sld.Delete
    Next
    
    
    Exit Sub
Err_Handler:
    MsgBox "Fehler " & Err.Number & ":" & vbNewLine & Err.Description
    Debug.Assert False
End Sub

Private Sub UpdateAgendaOnSlide(ByVal sld As Slide, ByVal textbox As Shape, ByVal parIdx As Integer)
    Dim selectorTop As Single
    Dim selectorHeight As Single
    Dim selectorMargin As Single
    Dim idx As Integer
    Dim selector As Shape
    Dim oTextBox As Shape
    
    On Error GoTo Err_Handler
    
    ' Tag auf Slide setzen
    SetTagsForSlide sld, parIdx
    ' Rand ober- und unterhalb / um wieviel ist Markierung größer als Absatz?
    selectorMargin = textbox.TextFrame.TextRange.Font.Size * 0.15
    ' Position der Markierung für ersten Eintrag der Agenda
    selectorTop = textbox.Top + textbox.TextFrame.MarginTop - selectorMargin
    
    ' Position der Markierung bestimmen
    ' Absatz-Höhe und -Abstände pro Absatz addieren
    For idx = 1 To parIdx - 1
        ' Absaetzhoehe addieren
        selectorTop = selectorTop + ParagraphHeight(textbox.TextFrame.TextRange.Paragraphs(idx), False)
        ' Absatzabsatand danach
        selectorTop = selectorTop + textbox.TextFrame.TextRange.Paragraphs(idx).ParagraphFormat.SpaceAfter
        ' Absatzabsatand davor vom naechsten Absatz
        selectorTop = selectorTop + textbox.TextFrame.TextRange.Paragraphs(idx + 1).ParagraphFormat.SpaceBefore
    Next
    selectorHeight = ParagraphHeight(textbox.TextFrame.TextRange.Paragraphs(parIdx), False) + 2 * selectorMargin
    
    ' Selector (Markierung) aktualisieren
    Set selector = GetOrCreateSelectorOnSlide(sld)
    selector.Top = selectorTop
    selector.Left = textbox.Left
    selector.Height = selectorHeight
    selector.Width = textbox.Width
    
    ' Text und Textbox-Position aktualisieren
    Set oTextBox = GetAgendaTextBoxOnSlide(sld)
    
    'oTextBox.TextFrame.TextRange.text = textbox.TextFrame.TextRange.text
    textbox.TextFrame.TextRange.Copy
    oTextBox.TextFrame.TextRange.Paste
    oTextBox.TextFrame.TextRange.Font.Bold = msoFalse
    oTextBox.TextFrame.TextRange.Paragraphs(parIdx).Font.Bold = msoTrue
    oTextBox.Left = textbox.Left
    oTextBox.Top = textbox.Top
    oTextBox.Width = textbox.Width
    oTextBox.Height = textbox.Height
    
    Exit Sub
Err_Handler:
    MsgBox "Fehler " & Err.Number & ":" & vbNewLine & Err.Description
    Debug.Assert False
End Sub


' Collection mit Agenda-Slides liefern
' Slides werden anhand der Meta-Informationen (Tags) wiedererkannt
Private Function GetAgendaSlides() As Collection
    Dim agendaSlides As Collection
    Dim sld As Slide
    
    Set agendaSlides = New Collection
    On Error Resume Next
    
    For Each sld In ActiveWindow.Presentation.Slides
        If sld.Tags.Item(TOOLBOX_AGENDA) <> "" Then
            agendaSlides.Add sld, "slide-" & CInt(sld.Tags.Item(TOOLBOX_AGENDA_SlideNo))
        End If
    Next
    
    Set GetAgendaSlides = agendaSlides
End Function

' Selector auf Slide liefern oder neu erstellen
' Shape wird anhand der Meta-Informationen (Tags) wiedererkannt
Private Function GetOrCreateSelectorOnSlide(ByVal sld As Slide) As Shape
    Dim shp As Shape
    
    
    Set shp = GetShapeWithTagItem(sld, TOOLBOX_AGENDA_Selector)
    If Not shp Is Nothing Then
        Set GetOrCreateSelectorOnSlide = shp
        Exit Function
    End If
    
    ' Neues Selector-Shape erstellen
    Set shp = sld.shapes.AddShape(msoShapeRectangle, 0, 0, 100, 20)
    SetTagsForSelector shp
    shp.ZOrder msoSendToBack
    ' Grauer Hintergrund/Rand
    shp.Fill.ForeColor.RGB = selectorFillColor
    shp.Line.ForeColor.RGB = selectorLineColor
    shp.Line.Weight = 0.75
    shp.Line.Visible = msoFalse
    
    Set GetOrCreateSelectorOnSlide = shp
End Function

' Agenda-TextBox auf Slide liefern oder neu erstellen
' Shape wird anhand der Meta-Informationen (Tags) wiedererkannt
Private Function GetAgendaTextBoxOnSlide(ByVal sld As Slide) As Shape
    Set GetAgendaTextBoxOnSlide = GetShapeWithTagItem(sld, TOOLBOX_AGENDA_TextBox)
End Function


Public Sub RemoveAgenda()
    If Not CanUpdateAgenda Then Exit Sub
    ' Agenda Meta-Informationen inkl. Slides löschen
    RemoveAgendaB True
End Sub

Private Sub RemoveAgendaB(ByVal deleteSlides As Boolean)
    Dim sld As Slide
    Dim agendaSlides As Collection
    Dim shp As Shape
    Dim idx As Integer
    
    On Error GoTo Err_Handler
    
    Set agendaSlides = New Collection
    
    For Each sld In ActiveWindow.Presentation.Slides
        If sld.Tags.Item(TOOLBOX_AGENDA) <> "" Then
            ' Tags von Slide loeschen
            sld.Tags.Add TOOLBOX_AGENDA, ""
            sld.Tags.Add TOOLBOX_AGENDA_SlideNo, ""
            
            ' Tags von Shapes auf Slide loeschen
            Set shp = GetAgendaTextBoxOnSlide(sld)
            If Not shp Is Nothing Then
                shp.Tags.Add TOOLBOX_AGENDA_TextBox, ""
                Set shp = Nothing
            End If
            Set shp = GetAgendaTextBoxOnSlide(sld)
            If Not shp Is Nothing Then
                shp.Tags.Add TOOLBOX_AGENDA_Selector, ""
                Set shp = Nothing
            End If
            
            If deleteSlides Then
                agendaSlides.Add sld
            End If
        End If
    Next

    If deleteSlides Then
        For idx = 2 To agendaSlides.Count
            agendaSlides(idx).Delete
        Next
    End If

    Exit Sub
Err_Handler:
    MsgBox "Fehler " & Err.Number & ":" & vbNewLine & Err.Description
    Debug.Assert False
End Sub


' Meta-Informationen für Slide einstellen
Private Sub SetTagsForSlide(ByVal sld As Slide, ByVal slideNo As Long)
    sld.Tags.Add TOOLBOX_AGENDA, "1"
    sld.Tags.Add TOOLBOX_AGENDA_SlideNo, CStr(slideNo)
End Sub

' Meta-Informationen für Textbox einstellen
Private Sub SetTagsForTextbox(ByVal textbox As Shape)
    textbox.Tags.Add TOOLBOX_AGENDA_TextBox, "1"
End Sub

' Meta-Informationen für Selector-Shape einstellen
Private Sub SetTagsForSelector(ByVal selector As Shape)
    selector.Tags.Add TOOLBOX_AGENDA_Selector, "1"
End Sub

' Shape auf Slide finden, das einen bestimmten TagKey enthaelt
Private Function GetShapeWithTagItem(ByVal sld As Slide, ByVal tagKey As String) As Shape
    Dim shp As Shape
    
    For Each shp In sld.shapes
        If shp.Tags.Item(tagKey) <> "" Then
            Set GetShapeWithTagItem = shp
            Exit Function
        End If
    Next
End Function


