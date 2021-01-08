Attribute VB_Name = "ToolboxRibbon"
Option Explicit
Public myRibbon As IRibbonUI

Private ConvertPointsToCentimeters As Boolean

Private oTrigger As TriggerInvalidate

Private AdjustmentValue As Integer



'Public Sub Test()
'    Dim sld As Slide
'    Dim idx As Integer
'    Dim shp As Shape
'
'    Set sld = ActiveWindow.View.Slide
'
'
'    IntToRGB sld.ColorScheme.Colors(ppBackground) ' Hintergrund 1
'    IntToRGB sld.ColorScheme.Colors(ppForeground) ' Text 1
'    IntToRGB sld.ColorScheme.Colors(ppShadow) ' Hintergrund 2
'    IntToRGB sld.ColorScheme.Colors(ppTitle) ' Text 2
'    IntToRGB sld.ColorScheme.Colors(ppFill) ' Akzent 1
'    IntToRGB sld.ColorScheme.Colors(ppAccent1) ' Akzent 2
'    IntToRGB sld.ColorScheme.Colors(ppAccent2) ' Hyperlink-Farbe
'    IntToRGB sld.ColorScheme.Colors(ppAccent3) ' Besuchte-Hyperlink-Farbe
'
'    'sld.ColorScheme.Colors(9).RGB = RGB(255, 0, 0)
'End Sub
'
'Public Sub IntToRGB(ByVal color As Long)
'    'Debug.Print Int(color / 256 ^ 2) Mod 256 & " / " & Int(color / 256) Mod 256 & " / " & color Mod 256
'    Debug.Print color Mod 256 & " / " & Int(color / 256) Mod 256 & " / " & Int(color / 256 ^ 2) Mod 256
'End Sub



' Initialisierung nach dem Laden des Ribbons
Sub ribbonLoaded(Ribbon As IRibbonUI)
    'On Error Resume Next
    Set myRibbon = Ribbon
    
    On Error GoTo Err_Handler
    
    ConvertPointsToCentimeters = True
    
    ScaleFrom = msoScaleFromTopLeft
    
    AdjustmentValue = 1
    
    Set oTrigger = New TriggerInvalidate
    Set oTrigger.Ribbon = myRibbon
    Set oTrigger.PPTEvent = ActiveWindow.Application
    
Exit Sub
Err_Handler:
End Sub

' Funktion zum Neu-Laden der Ribbon-Werte
Sub ReloadValues(control As IRibbonControl)
    On Error Resume Next
    myRibbon.Invalidate
End Sub

Sub isVisibleWin(control As IRibbonControl, ByRef Visible)
    #If Mac Then
        Visible = False
    #Else
        Visible = True
    #End If
End Sub

Sub isVisibleMac(control As IRibbonControl, ByRef Visible)
    #If Mac Then
        Visible = True
    #Else
        Visible = False
    #End If
End Sub


Sub GetLabelValue(control As IRibbonControl, ByRef label)
    Dim ctlId As String
    
    On Error GoTo Err_Handler
    
    ctlId = control.Id
    
    Select Case ctlId
    Case "menu-RectCorner"
        label = "Wert " & AdjustmentValue & " von 8"
    End Select

Exit Sub
Err_Handler:
    label = "ERROR"
End Sub

Sub isEnabled(control As IRibbonControl, ByRef enabled)
    Dim ctlId As String
    Dim oAgenda As ToolboxAgenda
    
    On Error GoTo Err_Handler
    
    ctlId = control.Id
    If Left(ctlId, 3) = "inc" Or Left(ctlId, 3) = "dec" Or Left(ctlId, 3) = "res" Then
        ctlId = "eb" & Mid(ctlId, 4)
    End If
    
    Select Case ctlId
    Case "actUpdateAgenda", "actRemoveAgenda"
        'Set oAgenda = New ToolboxAgenda
        'enabled = oAgenda.CanUpdateAgenda
        enabled = True
        
    Case Else
        ' Enabled-Status von Selection abhängig
        If ActiveWindow.selection.Type = ppSelectionNone Then
            enabled = False
        Else
            Select Case ctlId
            Case "ebVSep", "ebHSep", "actJoinShapesWithText", "actSplitSameWidth", "actSplitSameHeight"
                ' Enabled bei Selection-Count > 1
                enabled = (ActiveWindow.selection.ShapeRange.Count > 1)
                
            Case "actHConnect", "actVConnect", "actTextIntoShape", "actSplitSwap"
                ' Enabled bei Selection-Count = 2
                enabled = (ActiveWindow.selection.ShapeRange.Count = 2)
                
            Case "ebRectCorner"
                ' Enabled von Shape-Typ abhaengig
                ' enabled = (ActiveWindow.selection.ShapeRange(1).AutoShapeType = 51 Or ActiveWindow.selection.ShapeRange(1).AutoShapeType = 52 Or ActiveWindow.selection.ShapeRange(1).AutoShapeType = msoShapeRoundedRectangle)
                enabled = (ActiveWindow.selection.ShapeRange(1).Adjustments.Count >= AdjustmentValue)
'            Case "ebRectCorner2"
'                enabled = (ActiveWindow.selection.ShapeRange(1).Adjustments.Count >= AdjustmentValue + 1)
    
            Case Else
                If Left(ctlId, 2) = "eb" Then
                    ' Enabled bei Textboxen, wenn Wert nicht-leer
                    enabled = (GetEditBoxValue(ctlId) <> "")
                'ElseIf Left(ctlId, 3) = "act" Then
                '    ' Enabled bei Selection-Count = 1
                '    enabled = True
                Else
                    ' Enabled bei Selection-Count >= 1
                    enabled = (ActiveWindow.selection.Type <> ppSelectionNone)
                End If
            End Select
        End If
    End Select

Exit Sub
Err_Handler:
    enabled = False
End Sub



' Initialisierungsfunktion für CheckBoxen
' Gibt anhand der control.id den Pressed-Status der Ceckboxen zurueck
Function cbValue_init(control As IRibbonControl, ByRef returnedVal) As Boolean
    
    On Error GoTo Err_Handler
    
    Select Case control.Id
    ' RoundCorner
    Case "menu-RectCorner-t1"
        returnedVal = (AdjustmentValue = 1)
        Exit Function
    Case "menu-RectCorner-t2"
        returnedVal = (AdjustmentValue = 2)
        Exit Function
    Case "menu-RectCorner-t3"
        returnedVal = (AdjustmentValue = 3)
        Exit Function
    Case "menu-RectCorner-t4"
        returnedVal = (AdjustmentValue = 4)
        Exit Function
    Case "menu-RectCorner-t5"
        returnedVal = (AdjustmentValue = 5)
        Exit Function
    Case "menu-RectCorner-t6"
        returnedVal = (AdjustmentValue = 6)
        Exit Function
    Case "menu-RectCorner-t7"
        returnedVal = (AdjustmentValue = 7)
        Exit Function
    Case "menu-RectCorner-t8"
        returnedVal = (AdjustmentValue = 8)
        Exit Function
    
    ' Locpin
    Case "toggleLocPin1"
        returnedVal = (ScaleFrom = msoScaleFromTopLeft)
        Exit Function
    Case "toggleLocPin2"
        returnedVal = (ScaleFrom = msoScaleFromMiddle)
        Exit Function
    Case "toggleLocPin3"
        returnedVal = (ScaleFrom = msoScaleFromBottomRight)
        Exit Function
    End Select
    
    If ActiveWindow.selection.Type = ppSelectionNone Then
        returnedVal = False
    Else
        Select Case control.Id
        ' Textboxen
        Case "cbWordWrap", "cbWordWrap2"
            returnedVal = (ActiveWindow.selection.ShapeRange(1).TextFrame.WordWrap = msoTrue)
        Case "cbAutoSize", "cbAutoSize2"
            returnedVal = (ActiveWindow.selection.ShapeRange(1).TextFrame.AutoSize = ppAutoSizeShapeToFitText)
        ' Shapes
        Case "cbLockAspectRatio", "cbLockAspectRatio2"
            returnedVal = (ActiveWindow.selection.ShapeRange(1).LockAspectRatio = msoTrue)
        End Select
    End If
 
Exit Function
Err_Handler:
    cbValue_init = returnedVal
End Function


' Funktion für Änderungen an Checkboxen
' Zu control.id gehörige Eigenschaft wird an dem Pressed-Status der Checkbox angepasst
Sub cbValue_onChange(control As IRibbonControl, pressed As Boolean)
    Dim shp As Shape
    
    On Error GoTo Err_Handler
    
    Select Case control.Id
    ' RoundCorner
    Case "menu-RectCorner-t1"
        AdjustmentValue = 1
    Case "menu-RectCorner-t2"
        AdjustmentValue = 2
    Case "menu-RectCorner-t3"
        AdjustmentValue = 3
    Case "menu-RectCorner-t4"
        AdjustmentValue = 4
    Case "menu-RectCorner-t5"
        AdjustmentValue = 5
    Case "menu-RectCorner-t6"
        AdjustmentValue = 6
    Case "menu-RectCorner-t7"
        AdjustmentValue = 7
    Case "menu-RectCorner-t8"
        AdjustmentValue = 8
    
    ' Locpin
    Case "toggleLocPin1"
        ScaleFrom = msoScaleFromTopLeft
    Case "toggleLocPin2"
        ScaleFrom = msoScaleFromMiddle
    Case "toggleLocPin3"
        ScaleFrom = msoScaleFromBottomRight
    
    Case Else:
        ' Alle Controls die selektierte Shapes benötigen
        If ActiveWindow.selection.Type = ppSelectionNone Then Exit Sub
    
        For Each shp In ActiveWindow.selection.ShapeRange
            Select Case control.Id
            ' Textboxen
            Case "cbWordWrap", "cbWordWrap2"
                shp.TextFrame.WordWrap = CInt(pressed)
            Case "cbAutoSize", "cbAutoSize2"
                shp.TextFrame.AutoSize = Abs(CInt(pressed))
            ' Shapes
            Case "cbLockAspectRatio", "cbLockAspectRatio2"
                shp.LockAspectRatio = CInt(pressed)
            End Select
        Next
    End Select
    myRibbon.Invalidate

Exit Sub
Err_Handler:
End Sub


' Initialisierungsfunktion für TextBoxen
' Liefert Wert zu control.id, der in den Ribbon-Textboxen angezeigt wird
Sub ebPixelValue_init(control As IRibbonControl, ByRef returnedVal)
    
    On Error GoTo Err_Handler
    
    If ActiveWindow.selection.Type = ppSelectionNone Then
        returnedVal = ""
    Else
        returnedVal = GetEditBoxValue(control.Id)
    End If

Exit Sub
Err_Handler:
    returnedVal = ""
End Sub

Private Function GetEditBoxValue(ByVal ctlId As String) As String
    Dim returnedVal As String
    Dim shps As Variant
    
    On Error GoTo Err_Handler
    
    If ActiveWindow.selection.Type = ppSelectionNone Then
        returnedVal = ""
    Else
        Select Case ctlId
        ' Innenabstand
        Case "ebMarginLeft"
            returnedVal = ActiveWindow.selection.ShapeRange(1).TextFrame2.MarginLeft
            If returnedVal <> "" And ConvertPointsToCentimeters Then returnedVal = Round(PointsToCentimeters(returnedVal), 2)
        Case "ebMarginRight"
            returnedVal = ActiveWindow.selection.ShapeRange(1).TextFrame2.MarginRight
            If returnedVal <> "" And ConvertPointsToCentimeters Then returnedVal = Round(PointsToCentimeters(returnedVal), 2)
        Case "ebMarginTop"
            returnedVal = ActiveWindow.selection.ShapeRange(1).TextFrame2.MarginTop
            If returnedVal <> "" And ConvertPointsToCentimeters Then returnedVal = Round(PointsToCentimeters(returnedVal), 2)
        Case "ebMarginBottom"
            returnedVal = ActiveWindow.selection.ShapeRange(1).TextFrame2.MarginBottom
            If returnedVal <> "" And ConvertPointsToCentimeters Then returnedVal = Round(PointsToCentimeters(returnedVal), 2)
            
        ' Position/Groesse
        Case "ebPosLeft"
'            returnedVal = ActiveWindow.selection.ShapeRange(1).left
            returnedVal = GetLeft(ActiveWindow.selection.ShapeRange(1))
            If returnedVal <> "" And ConvertPointsToCentimeters Then returnedVal = Round(PointsToCentimeters(returnedVal), 2)
        Case "ebPosTop"
            returnedVal = GetTop(ActiveWindow.selection.ShapeRange(1))
            If returnedVal <> "" And ConvertPointsToCentimeters Then returnedVal = Round(PointsToCentimeters(returnedVal), 2)
        Case "ebPosRight"
            returnedVal = GetWidth(ActiveWindow.selection.ShapeRange(1))
            If returnedVal <> "" And ConvertPointsToCentimeters Then returnedVal = Round(PointsToCentimeters(returnedVal), 2)
        Case "ebPosBottom"
            returnedVal = GetHeight(ActiveWindow.selection.ShapeRange(1))
            If returnedVal <> "" And ConvertPointsToCentimeters Then returnedVal = Round(PointsToCentimeters(returnedVal), 2)
        
        ' Rotation
        Case "ebRotation"
            returnedVal = ActiveWindow.selection.ShapeRange(1).Rotation
        
        ' Objektabstand
        Case "ebVSep"
            If ActiveWindow.selection.ShapeRange.Count >= 2 Then
                shps = ActiveWindowSelectionSortedByTop
                returnedVal = shps(2).Top - shps(1).Top - shps(1).Height
                If returnedVal <> "" And ConvertPointsToCentimeters Then returnedVal = Round(PointsToCentimeters(returnedVal), 2)
            End If
        Case "ebHSep"
            If ActiveWindow.selection.ShapeRange.Count >= 2 Then
                shps = ActiveWindowSelectionSortedByLeft
                returnedVal = shps(2).Left - shps(1).Left - shps(1).Width
                If returnedVal <> "" And ConvertPointsToCentimeters Then returnedVal = Round(PointsToCentimeters(returnedVal), 2)
            End If
            
        ' Absatzabstand
'        Case "ebParIndent"
'            returnedVal = Round(PointsToCentimeters(ActiveWindow.selection.ShapeRange(1).TextFrame.Ruler.Levels(1).FirstMargin), 2)
        Case "ebParIndentLeft"
            ' returnedVal = Round(PointsToCentimeters(ActiveWindow.selection.ShapeRange(1).TextFrame.Ruler.Levels(1).LeftMargin), 2)
            With ActiveWindow.selection.ShapeRange(1).TextFrame2.TextRange
                If .Paragraphs.Count > 0 Then
                    returnedVal = Round(PointsToCentimeters(.Paragraphs(1).ParagraphFormat.LeftIndent), 2)
                Else
                    returnedVal = Round(PointsToCentimeters(.ParagraphFormat.LeftIndent), 2)
                End If
            End With
        Case "ebParIndentRight"
            With ActiveWindow.selection.ShapeRange(1).TextFrame2.TextRange
                If .Paragraphs.Count > 0 Then
                    returnedVal = Round(PointsToCentimeters(.Paragraphs(1).ParagraphFormat.RightIndent), 2)
                Else
                    returnedVal = Round(PointsToCentimeters(.ParagraphFormat.RightIndent), 2)
                End If
            End With
        Case "ebParIndentFirst"
            With ActiveWindow.selection.ShapeRange(1).TextFrame2.TextRange
                If .Paragraphs.Count > 0 Then
                    returnedVal = Round(PointsToCentimeters(.Paragraphs(1).ParagraphFormat.FirstLineIndent), 2)
                Else
                    returnedVal = Round(PointsToCentimeters(.ParagraphFormat.FirstLineIndent), 2)
                End If
            End With
        Case "ebParPreSep"
            With ActiveWindow.selection.ShapeRange(1).TextFrame2.TextRange
                If .Paragraphs.Count > 0 Then
                    returnedVal = .Paragraphs(1).ParagraphFormat.SpaceBefore
                Else
                    returnedVal = .ParagraphFormat.SpaceBefore
                End If
            End With
        Case "ebParPostSep"
            With ActiveWindow.selection.ShapeRange(1).TextFrame2.TextRange
                If .Paragraphs.Count > 0 Then
                    returnedVal = .Paragraphs(1).ParagraphFormat.SpaceAfter
                Else
                    returnedVal = .ParagraphFormat.SpaceAfter
                End If
            End With
        Case "ebParWithin"
            With ActiveWindow.selection.ShapeRange(1).TextFrame2.TextRange
                If .Paragraphs.Count > 0 Then
                    returnedVal = .Paragraphs(1).ParagraphFormat.SpaceWithin
                Else
                    returnedVal = .ParagraphFormat.SpaceWithin
                End If
            End With
        
        ' Gerundete Ecken
        Case "ebRectCorner"
            If ActiveWindow.selection.ShapeRange(1).Adjustments.Count >= AdjustmentValue Then
                returnedVal = Round(RoundedCornerSize(ActiveWindow.selection.ShapeRange(1), AdjustmentValue), 2)
            Else
                returnedVal = ""
            End If
            'If ConvertPointsToCentimeters Then returnedVal = Round(PointsToCentimeters(returnedVal), 2)
'        Case "ebRectCorner2"
'            If ActiveWindow.selection.ShapeRange(1).Adjustments.Count >= AdjustmentValue + 1 Then
'                returnedVal = Round(RoundedCornerSize(ActiveWindow.selection.ShapeRange(1), AdjustmentValue + 1), 2)
'            Else
'                returnedVal = ""
'            End If
'            'If ConvertPointsToCentimeters Then returnedVal = Round(PointsToCentimeters(returnedVal), 2)
        
        ' Transparenz und Rahmen
        Case "ebTranspFill"
            returnedVal = Max(0, Round(ActiveWindow.selection.ShapeRange(1).Fill.Transparency * 100))
        Case "ebTranspLine"
            returnedVal = Max(0, Round(ActiveWindow.selection.ShapeRange(1).Line.Transparency * 100))
        Case "ebLineWeight"
            returnedVal = Max(0, Round(ActiveWindow.selection.ShapeRange(1).Line.Weight, 2))
        
        Case Else
            Debug.Print ctlId
        End Select
    End If
    GetEditBoxValue = returnedVal

Exit Function
Err_Handler:
    GetEditBoxValue = ""
End Function


' Funktion für Änderungen an Textboxen mit Integer-Werten
' Zu control.id gehörige Eigenschaft wird an dem Wert in der Textbox angepasst
Sub ebIntValue_onChange(control As IRibbonControl, text As String)
    Dim value As Integer
    Dim shp As Shape
    
    On Error GoTo Err_Handler
    
    value = CInt(text)
    If ActiveWindow.selection.Type = ppSelectionNone Then Exit Sub
    
    For Each shp In ActiveWindow.selection.ShapeRange
        Select Case control.Id
        ' Rotation
        Case "ebRotation"
            shp.Rotation = ((value Mod 360) + 360) Mod 360
        ' Transparency
        Case "ebTranspFill"
            shp.Fill.Transparency = Min(100, Max(0, value / 100))
        Case "ebTranspLine"
            shp.Line.Transparency = Min(100, Max(0, value / 100))
        End Select
    Next
    myRibbon.Invalidate

Exit Sub
Err_Handler:
End Sub

' Funktion für Änderungen an Textboxen mit Pixel-Werten (Längen)
' Zu control.id gehörige Eigenschaft wird an dem Wert in der Textbox angepasst
Sub ebPixelValue_onChange(control As IRibbonControl, text As String)
    Dim value As Single
    Dim shp As Shape
    Dim delta As Single
    Dim lastShp As Shape
    Dim shps As Variant
    Dim shpIdx As Integer
    
    On Error GoTo Err_Handler
    
    If ActiveWindow.selection.Type = ppSelectionNone Then Exit Sub
    
    Select Case control.Id
    Case "ebHSep"
        shps = ActiveWindowSelectionSortedByLeft
        value = CSng(text)
    Case "ebVSep"
        shps = ActiveWindowSelectionSortedByTop
        value = CSng(text)
    Case "ebRectCorner", "ebParFirst" ', "ebRectCorner2"
        Set shps = ActiveWindow.selection.ShapeRange
        value = CSng(text)
    Case Else
        Set shps = ActiveWindow.selection.ShapeRange
        value = Max(0, CSng(text))
    End Select
    
    Select Case control.Id
    Case "ebParPreSep", "ebParPostSep", "ebParWithin", "ebRotation", "ebLineWeight", "ebRectCorner" ', "ebRectCorner2"
    Case Else
        If ConvertPointsToCentimeters Then value = CentimetersToPoints(value)
    End Select
    
    For shpIdx = 1 To ActiveWindow.selection.ShapeRange.Count
        'For Each shp In ActiveWindow.Selection.ShapeRange
        Set shp = shps(shpIdx)
        Select Case control.Id
        ' Innenabstand
        Case "ebMarginLeft"
            shp.TextFrame2.MarginLeft = value
        Case "ebMarginRight"
            shp.TextFrame2.MarginRight = value
        Case "ebMarginTop"
            shp.TextFrame2.MarginTop = value
        Case "ebMarginBottom"
            shp.TextFrame2.MarginBottom = value
        
        ' Position/Groesse
        Case "ebPosLeft"
'            shp.left = value
            SetLeft shp, value
        Case "ebPosTop"
            SetTop shp, value
        Case "ebPosRight"
            SetWidth shp, value
        Case "ebPosBottom"
            SetHeight shp, value
        
        ' Objektabstand
        Case "ebVSep"
            If Not lastShp Is Nothing Then
                shp.Top = lastShp.Top + lastShp.Height + value
            End If
        Case "ebHSep"
            If Not lastShp Is Nothing Then
                shp.Left = lastShp.Left + lastShp.Width + value
            End If
        
        ' Absatzabstand
'        Case "ebParIndent"
'            delta = value - shp.TextFrame.Ruler.Levels(1).FirstMargin
'            shp.TextFrame.Ruler.Levels(1).FirstMargin = value
'            shp.TextFrame.Ruler.Levels(1).LeftMargin = shp.TextFrame.Ruler.Levels(1).LeftMargin + delta
            
        Case "ebParIndentFirst"
            shp.TextFrame2.TextRange.ParagraphFormat.FirstLineIndent = value
        Case "ebParIndentLeft"
            shp.TextFrame2.TextRange.ParagraphFormat.LeftIndent = value
        Case "ebParIndentRight"
            shp.TextFrame2.TextRange.ParagraphFormat.RightIndent = value
        Case "ebParPreSep"
            shp.TextFrame2.TextRange.ParagraphFormat.SpaceBefore = value
        Case "ebParPostSep"
            shp.TextFrame2.TextRange.ParagraphFormat.SpaceAfter = value
        Case "ebParWithin"
            shp.TextFrame2.TextRange.ParagraphFormat.SpaceWithin = value
        ' Rechteck Rundungen
        Case "ebRectCorner"
            SetRoundedCornerSize shp, value, AdjustmentValue
'        Case "ebRectCorner2"
'            SetRoundedCornerSize shp, value, AdjustmentValue + 1

        ' Line Weight
        Case "ebLineWeight"
            shp.Line.Weight = value
        
        End Select
        Set lastShp = shp
    Next
    myRibbon.Invalidate

Exit Sub
Err_Handler:
End Sub

' Funktionen zur Anpassung der Werte durch die Buttons
Sub IncreasePixelValue(control As IRibbonControl)
    On Error Resume Next
    ChangeValueBy control, 1
End Sub
Sub DecreasePixelValue(control As IRibbonControl)
    On Error Resume Next
    ChangeValueBy control, -1
End Sub

Sub ResetPixelValue(control As IRibbonControl)
    Dim shp As Shape
    Dim shps As Variant
    Dim shpIdx As Integer
    Dim lastShp As Shape
    
    On Error GoTo Err_Handler
    
    If ActiveWindow.selection.Type = ppSelectionNone Then Exit Sub
    
    Select Case control.Id
    Case "resHSep"
        shps = ActiveWindowSelectionSortedByLeft
    Case "resVSep"
        shps = ActiveWindowSelectionSortedByTop
    Case Else
        Set shps = ActiveWindow.selection.ShapeRange
    End Select
    
    For shpIdx = 1 To ActiveWindow.selection.ShapeRange.Count
        'For Each shp In ActiveWindow.Selection.ShapeRange
        Set shp = shps(shpIdx)
        Select Case control.Id
        ' Objektabstand
        Case "resVSep"
            If Not lastShp Is Nothing Then
                shp.Top = lastShp.Top + lastShp.Height
            End If
        Case "resHSep", "decHSep"
            If Not lastShp Is Nothing Then
                shp.Left = lastShp.Left + lastShp.Width
            End If
        
        Case Else
            SetShapeSettingSingle shp, "eb" & Mid(control.Id, 4), 0
        End Select
        Set lastShp = shp
    Next
    
    myRibbon.Invalidate

Exit Sub
Err_Handler:
End Sub

' Zu control.id gehörige Eigenschaft wird um value erhöht/verringert
' Je nach control.id erfolgt Umrechnung von cm in punkte
' Bei nicht gedrückter control-taste wird ein Vielfaches von value verwendet
Private Sub ChangeValueBy(control As IRibbonControl, ByVal value As Integer)
    Dim shp As Shape
    Dim delta As Single
    Dim ptValue As Single
    Dim cmValue As Single
    Dim intValue As Single
    Dim shps As Variant
    Dim shpIdx As Integer
    Dim lastShp As Shape
    Dim newValue As Single
    Dim oldValue As Single
    Dim firstDelta As Single
    
    On Error GoTo Err_Handler
    
    If ActiveWindow.selection.Type = ppSelectionNone Then Exit Sub
    
    
    ptValue = value
    cmValue = value * 0.1
    intValue = value
    If Not IsControlKeyDown Then
        ptValue = 3 * ptValue
        cmValue = 2 * cmValue
        intValue = 5 * intValue
    End If
    
    
    Select Case control.Id
    Case "incHSep", "decHSep"
        shps = ActiveWindowSelectionSortedByLeft
        oldValue = shps(2).Left - shps(1).Left - shps(1).Width
    Case "incVSep", "decVSep"
        shps = ActiveWindowSelectionSortedByTop
        oldValue = shps(2).Top - shps(1).Top - shps(1).Height
    Case Else
        ' cmValue
        Set shps = ActiveWindow.selection.ShapeRange
        oldValue = GetShapeSettingSingle(shps(1), control.Id)
    End Select
    
    Select Case control.Id
    Case "incParPreSep", "decParPreSep", "incParPostSep", "decParPostSep"
        newValue = Max(0, Round(oldValue) + ptValue)
    Case "incParWithin", "decParWithin"
        newValue = Max(0, Round(oldValue, 1) + cmValue)
    Case "incParIndentFirst", "decParIndentFirst"
        newValue = CentimetersToPoints(Round(PointsToCentimeters(oldValue), 1) + cmValue)
    Case "incRotation", "decRotation"
        newValue = Round(oldValue) + intValue
    Case "incTranspFill", "decTranspFill", "incTranspLine", "decTranspLine"
        newValue = Min(oldValue + (intValue / 100), 1)
    Case "incHSep", "decHSep", "incVSep", "decVSep"
        newValue = CentimetersToPoints(Round(PointsToCentimeters(oldValue), 1) + cmValue)
    Case "decRectCorner", "incRectCorner" ', "decRectCorner2", "incRectCorner2"
        newValue = Round(oldValue, 1) + cmValue
    Case "incLineWeight", "decLineWeight"
        newValue = oldValue + (cmValue * 2.5)
    Case Else
        newValue = CentimetersToPoints(Max(0, Round(PointsToCentimeters(oldValue), 1) + cmValue))
    End Select
    
    firstDelta = newValue - oldValue
    
    For shpIdx = 1 To ActiveWindow.selection.ShapeRange.Count
        'For Each shp In ActiveWindow.Selection.ShapeRange
        Set shp = shps(shpIdx)
        Select Case control.Id
        ' Objektabstand
        Case "incVSep", "decVSep"
            If Not lastShp Is Nothing Then
                shp.Top = lastShp.Top + lastShp.Height + newValue
            End If
        Case "incHSep", "decHSep"
            If Not lastShp Is Nothing Then
                shp.Left = lastShp.Left + lastShp.Width + newValue
            End If
        ' Absatzabstand
'        Case "incParIndent", "decParIndent"
'            If Not IsAltKeyDown Then
'                oldValue = shp.TextFrame.Ruler.Levels(1).FirstMargin
'                delta = newValue - oldValue
'                shp.TextFrame.Ruler.Levels(1).FirstMargin = shp.TextFrame.Ruler.Levels(1).FirstMargin + delta
'                shp.TextFrame.Ruler.Levels(1).LeftMargin = shp.TextFrame.Ruler.Levels(1).LeftMargin + delta
'            Else
'                shp.TextFrame.Ruler.Levels(1).FirstMargin = shp.TextFrame.Ruler.Levels(1).FirstMargin + firstDelta
'                shp.TextFrame.Ruler.Levels(1).LeftMargin = shp.TextFrame.Ruler.Levels(1).LeftMargin + firstDelta
'            End If
        
        Case Else
            If Not IsAltKeyDown Then
                SetShapeSettingSingle shp, control.Id, newValue
            Else
                oldValue = GetShapeSettingSingle(shp, control.Id)
                SetShapeSettingSingle shp, control.Id, oldValue + firstDelta
            End If
        End Select
        Set lastShp = shp
    Next
    
    myRibbon.Invalidate

Exit Sub
Err_Handler:
End Sub



' Liefert die Größe der Rundung eines abgerundeten Rechtecks in pt
' Prozent-Wert in den adjustments des shapes wird anhand der Recheckgröße umgerechnet
' Funktioniert auch für Pfeile (Fünfecke)
Private Function RoundedCornerSize(ByVal shp As Shape, Optional adj As Integer = 1) As Double
    Dim ref As Double
    
    On Error GoTo Err_Handler
    
    'Dim shp As Shape
    'Set shp = ActiveWindow.Selection.ShapeRange(1)
    
    If shp.Adjustments.Count < adj Then
        RoundedCornerSize = 0
        Exit Function
    End If
    
    ' Umrechnung Prozent zu Minimum aus Höhe und Breite bei bekannten Shape-Typen
    If adj = 1 Then
        Select Case shp.AutoShapeType
        Case msoShapePentagon, msoShapeChevron, msoShapeHexagon, msoShapeRoundedRectangle, msoShapeSnip1Rectangle, msoShapeSnip2DiagRectangle, msoShapeSnip2SameRectangle, msoShapeSnipRoundRectangle, msoShapeRound1Rectangle, msoShapeRound2DiagRectangle, msoShapeRound2SameRectangle
            ref = Min(shp.Height, shp.Width)
            ref = shp.Adjustments(1) * ref
            If ConvertPointsToCentimeters Then ref = Round(PointsToCentimeters(ref), 2)
            RoundedCornerSize = ref
        Case Else
            RoundedCornerSize = shp.Adjustments(1) * 100
        End Select
    Else
        RoundedCornerSize = shp.Adjustments(adj) * 100
    End If

Exit Function
Err_Handler:
End Function

' Ändert die Größe der Rundung eines abgerundeten Rechtecks
' Wert (in pt) wird anhand der Reckeckgröße auf Prozent-Wert in den adjustments des shapes umgerechnet
' Funktioniert auch für Pfeile (Fünfecke)
Public Function SetRoundedCornerSize(ByVal shp As Shape, ByVal newValue As Double, Optional adj As Integer = 1)
    Dim ref As Double
    
    On Error GoTo Err_Handler
    
    ' Umrechnung Prozent zu Minimum aus Höhe und Breite bei bekannten Shape-Typen
    If adj = 1 Then
        Select Case shp.AutoShapeType
        Case msoShapePentagon, msoShapeChevron, msoShapeHexagon, msoShapeRoundedRectangle, msoShapeSnip1Rectangle, msoShapeSnip2DiagRectangle, msoShapeSnip2SameRectangle, msoShapeSnipRoundRectangle, msoShapeRound1Rectangle, msoShapeRound2DiagRectangle, msoShapeRound2SameRectangle
            newValue = Max(0, newValue)
            ref = Min(shp.Height, shp.Width)
            If ConvertPointsToCentimeters Then newValue = CentimetersToPoints(newValue)
            shp.Adjustments(1) = newValue / ref
        Case Else
            shp.Adjustments(1) = newValue / 100
        End Select
    Else
        shp.Adjustments(adj) = newValue / 100
    End If

Exit Function
Err_Handler:
End Function






Function GetShapeSettingSingle(ByVal shp As Shape, controlID As String) As Single
    Select Case controlID
    ' Innenabstand
    Case "ebMarginLeft", "decMarginLeft", "incMarginLeft"
        GetShapeSettingSingle = shp.TextFrame2.MarginLeft
    Case "ebMarginRight", "decMarginRight", "incMarginRight"
        GetShapeSettingSingle = shp.TextFrame2.MarginRight
    Case "ebMarginTop", "decMarginTop", "incMarginTop"
        GetShapeSettingSingle = shp.TextFrame2.MarginTop
    Case "ebMarginBottom", "decMarginBottom", "incMarginBottom"
        GetShapeSettingSingle = shp.TextFrame2.MarginBottom
    
    ' Position/Groesse
    Case "ebPosLeft", "decPosLeft", "incPosLeft"
        GetShapeSettingSingle = GetLeft(shp)
    Case "ebPosTop", "decPosTop", "incPosTop"
        GetShapeSettingSingle = GetTop(shp)
    Case "ebPosRight", "decPosRight", "incPosRight"
        GetShapeSettingSingle = GetWidth(shp)
    Case "ebPosBottom", "decPosBottom", "incPosBottom"
        GetShapeSettingSingle = GetHeight(shp)
    
    ' Rotation
    Case "ebRotation", "decRotation", "incRotation"
        GetShapeSettingSingle = shp.Rotation
    
    ' Transparency
    Case "ebTranspFill", "decTranspFill", "incTranspFill"
        GetShapeSettingSingle = Max(0, shp.Fill.Transparency)
    Case "ebTranspLine", "decTranspLine", "incTranspLine"
        GetShapeSettingSingle = Max(0, shp.Line.Transparency)
        
    ' Line Weight
    Case "ebLineWeight", "decLineWeight", "incLineWeight"
        GetShapeSettingSingle = Max(0, shp.Line.Weight)
    
    ' Objektabstand
    Case "ebVSep"
    Case "ebHSep"
        
    ' Absatzabstand
'    Case "ebParIndent", "decParIndent", "incParIndent"
'        GetShapeSettingSingle = shp.TextFrame.Ruler.Levels(1).FirstMargin
        
    Case "ebParIndentFirst", "decParIndentFirst", "incParIndentFirst"
        With shp.TextFrame2.TextRange
            If .Paragraphs.Count > 0 Then
                GetShapeSettingSingle = .Paragraphs(1).ParagraphFormat.FirstLineIndent
            Else
                GetShapeSettingSingle = .ParagraphFormat.FirstLineIndent
            End If
        End With
    Case "ebParIndentLeft", "decParIndentLeft", "incParIndentLeft"
        With shp.TextFrame2.TextRange
            If .Paragraphs.Count > 0 Then
                GetShapeSettingSingle = .Paragraphs(1).ParagraphFormat.LeftIndent
            Else
                GetShapeSettingSingle = .ParagraphFormat.LeftIndent
            End If
        End With
    Case "ebParIndentRight", "decParIndentRight", "incParIndentRight"
        With shp.TextFrame2.TextRange
            If .Paragraphs.Count > 0 Then
                GetShapeSettingSingle = .Paragraphs(1).ParagraphFormat.RightIndent
            Else
                GetShapeSettingSingle = .ParagraphFormat.RightIndent
            End If
        End With
    Case "ebParPreSep", "decParPreSep", "incParPreSep"
        With shp.TextFrame2.TextRange
            If .Paragraphs.Count > 0 Then
                GetShapeSettingSingle = .Paragraphs(1).ParagraphFormat.SpaceBefore
            Else
                GetShapeSettingSingle = .ParagraphFormat.SpaceBefore
            End If
        End With
    Case "ebParPostSep", "decParPostSep", "incParPostSep"
        With shp.TextFrame2.TextRange
            If .Paragraphs.Count > 0 Then
                GetShapeSettingSingle = .Paragraphs(1).ParagraphFormat.SpaceAfter
            Else
                GetShapeSettingSingle = .ParagraphFormat.SpaceAfter
            End If
        End With
    Case "ebParWithin", "decParWithin", "incParWithin"
        With shp.TextFrame2.TextRange
            If .Paragraphs.Count > 0 Then
                GetShapeSettingSingle = .Paragraphs(1).ParagraphFormat.SpaceWithin
            Else
                GetShapeSettingSingle = .ParagraphFormat.SpaceWithin
            End If
        End With
    
    ' Gerundete Ecken
    Case "ebRectCorner", "decRectCorner", "incRectCorner"
        GetShapeSettingSingle = RoundedCornerSize(shp, AdjustmentValue)
'    Case "ebRectCorner2", "decRectCorner2", "incRectCorner2"
'        GetShapeSettingSingle = RoundedCornerSize(shp, AdjustmentValue + 1)
    Case Else
        Debug.Print controlID
    End Select
End Function

Sub SetShapeSettingSingle(ByVal shp As Shape, controlID As String, ByVal newValue As Single)
    Select Case controlID
    ' Innenabstand
    Case "ebMarginLeft", "decMarginLeft", "incMarginLeft"
        shp.TextFrame2.MarginLeft = newValue
    Case "ebMarginRight", "decMarginRight", "incMarginRight"
        shp.TextFrame2.MarginRight = newValue
    Case "ebMarginTop", "decMarginTop", "incMarginTop"
        shp.TextFrame2.MarginTop = newValue
    Case "ebMarginBottom", "decMarginBottom", "incMarginBottom"
        shp.TextFrame2.MarginBottom = newValue
    
    ' Position/Groesse
    Case "ebPosLeft", "decPosLeft", "incPosLeft"
        SetLeft shp, newValue
    Case "ebPosTop", "decPosTop", "incPosTop"
        SetTop shp, newValue
    Case "ebPosRight", "decPosRight", "incPosRight"
        SetWidth shp, newValue
    Case "ebPosBottom", "decPosBottom", "incPosBottom"
        SetHeight shp, newValue
    
    ' Rotation
    Case "ebRotation", "decRotation", "incRotation"
        shp.Rotation = newValue
    
    ' Transparency
    Case "ebTranspFill", "decTranspFill", "incTranspFill"
        shp.Fill.Transparency = Min(1, Max(0, newValue))
    Case "ebTranspLine", "decTranspLine", "incTranspLine"
        shp.Line.Transparency = Min(1, Max(0, newValue))
        
    ' Line Weight
    Case "ebLineWeight", "decLineWeight", "incLineWeight"
        shp.Line.Weight = Max(0, newValue)
    
    
    ' Objektabstand
    Case "ebVSep"
    Case "ebHSep"
        
    ' Absatzabstand
'    Case "ebParIndent", "decParIndent", "incParIndent"
'        shp.TextFrame.Ruler.Levels(1).FirstMargin = newValue
        
    Case "ebParIndentFirst", "decParIndentFirst", "incParIndentFirst"
        shp.TextFrame2.TextRange.ParagraphFormat.FirstLineIndent = newValue
    Case "ebParIndentLeft", "decParIndentLeft", "incParIndentLeft"
        shp.TextFrame2.TextRange.ParagraphFormat.LeftIndent = newValue
    Case "ebParIndentRight", "decParIndentRight", "incParIndentRight"
        shp.TextFrame2.TextRange.ParagraphFormat.RightIndent = newValue
    Case "ebParPreSep", "decParPreSep", "incParPreSep"
        shp.TextFrame2.TextRange.ParagraphFormat.SpaceBefore = newValue
    Case "ebParPostSep", "decParPostSep", "incParPostSep"
        shp.TextFrame2.TextRange.ParagraphFormat.SpaceAfter = newValue
    Case "ebParWithin", "decParWithin", "incParWithin"
        shp.TextFrame2.TextRange.ParagraphFormat.SpaceWithin = newValue
    
    ' Gerundete Ecken
    Case "ebRectCorner", "decRectCorner", "incRectCorner"
        SetRoundedCornerSize shp, newValue, AdjustmentValue
'    Case "ebRectCorner2", "decRectCorner2", "incRectCorner2"
'        SetRoundedCornerSize shp, newValue, AdjustmentValue + 1
    Case Else
        Debug.Print controlID
    End Select
End Sub


' Funktionen fuer Buttons
Sub btnAction(control As IRibbonControl)
    Dim oAgenda As ToolboxAgenda
    
    On Error GoTo Err_Handler
    Select Case control.Id
    ' Objekte auswaehlen
    Case "actSelectByShape"
        SelectShapesByShapeType
    Case "actSelectByLine"
        SelectShapesByLineType
    Case "actSelectByFill"
        SelectShapesByFillType
    
    ' Objekthoehe/-breite angleichen
    Case "actSameHeight", "actSameHeight2"
        If IsShiftKeyDown Then
            SetSameHeight "Min"
        Else
            SetSameHeight
        End If
    Case "actSameHeightMin"
        SetSameHeight "Min"
    
    Case "actSameWidth", "actSameWidth2"
        If IsShiftKeyDown Then
            SetSameWidth "Min"
        Else
            SetSameWidth
        End If
    Case "actSameWidthMin"
        SetSameWidth "Min"
    Case "actSwap", "actSwap2"
        SwapPosition
    Case "actSwapSize"
        SwapPositionSize
    
    ' Horizontaler/Vertikaler Verbinder
    Case "actHConnect"
        AddConnectorShapeLeftRight
    Case "actVConnect"
        AddConnectorShapeTopBottom
    
    ' Texte auf/in Shapes
    Case "actTextIntoShape"
        MoveTextIntoShape
    Case "actTextOutOfShape"
        MoveTextOutOfShapes
    
    ' Texte ersetzen/löschen
    Case "actShapeTextReplace"
        ReplaceAllText
    Case "actShapeTextRemove"
        RemoveAllText
    
    ' Einfügen
    Case "actPasteToSlides"
        PasteOnSlides
    Case "actPasteAndReplace"
        PasteAndReplace
        
    ' Aufräumen
    Case "actCleanAuthor"
        CleanAuthor
    Case "actCleanSlideMasters"
        CleanSlideMasters
    
    ' Shapes ein/ausblende
    Case "actShapesHide"
        HideShapes
    Case "actShapesShow"
        ShowShapes
    
    ' Objekte zerlegen/zusammenführen
    Case "actSplitShapeByParagraphs"
        SplitShapeByParagraphs
    Case "actJoinShapesWithText"
        JoinShapesWithText
        
    'Sticker
    Case "actSticker"
        AddSticker
    
    ' Slides / Templates
    Case "actSlideNumbering"
        ToggleSlideNumbering
    Case "actTemplateSlide"
        ShowTemplateDialog
    Case "actThemeApply"
        ApplyThemeFromFile
    
    ' Agenda
    Case "actCreateAgendaTextBox"
        Set oAgenda = New ToolboxAgenda
        oAgenda.CreateAgendaTextBoxOnSlide ActiveWindow.View.Slide
    Case "actCreateAgendaFromSlide"
        Set oAgenda = New ToolboxAgenda
        oAgenda.CreateFromCurrentSlide
    Case "actUpdateAgenda"
        Set oAgenda = New ToolboxAgenda
        oAgenda.CreateOrUpdateAgenda
    Case "actRemoveAgenda"
        Set oAgenda = New ToolboxAgenda
        oAgenda.RemoveAgenda
    
    ' Sprache
    Case "actLangDe", "actLangDe2"
        setLanguage msoLanguageIDGerman
    Case "actLangEnUs", "actLangEnUs2"
        setLanguage msoLanguageIDEnglishUS
    Case "actLangEnUk", "actLangEnUk2"
        setLanguage msoLanguageIDEnglishUK
    
    ' Slide-Selection
    Case "btnSendSelectionAsAttachment"
        SendEmailFromSlideSelection
    Case "btnSelectedSlidesSaveAs"
        CreatePresentationFromSlideSelection
    
    ' Text
    Case "textMarginZero"
        TextMarginZero
    
    ' Info
    Case "lblxInfo", "lblxWebsite"
        MsgBox "More information on www.bkt-toolbox.de/legacy"
    
    End Select
    
    myRibbon.Invalidate

Exit Sub
Err_Handler:
End Sub

Sub galAction(control As IRibbonControl, selectedID As String, selectedIndex As Integer)

    On Error GoTo Err_Handler
    Select Case control.Id
      Case "galFillTransparency"
        SetFillTransparency selectedIndex / 10
      Case "galLineTransparency"
        SetLineTransparency selectedIndex / 10
    End Select

Exit Sub
Err_Handler:
End Sub


Sub GetThinkCellMenuContent(control As IRibbonControl, ByRef xmlStr)
'    Dim i As Integer
'
'    Debug.Print ActiveWindow.Panes.Count
'
'    For i = 1 To Application.CommandBars.Count
'        Debug.Print Application.CommandBars(i).Name
'    Next
'    Debug.Assert False
    
    
'    xmlStr = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" xmlns:tc=""thinkcell.addin"">" & vbNewLine & _
'             "<button idQ=""tc:btnPastePlainText"" />" & vbNewLine & _
'             "<gallery idQ=""tc:galCharts"" />" & vbNewLine & _
'             "<control idQ=""tc:menuBasic"" />" & vbNewLine & _
'             "<button idQ=""tc:btnSaveSlides"" label=""x"" />" & vbNewLine & _
'             "<control idQ=""tc:btnSendSlides"" />" & vbNewLine & _
'             "<control idQ=""tc:btnSameWidth"" />" & vbNewLine & _
'             "<control idQ=""tc:btnSameHeight"" />" & vbNewLine & _
'             "<button id=""xxbutton1"" label=""Button 1"" />" & vbNewLine & _
'             "</menu>"
    
'    xmlStr = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" & vbNewLine & _
'             "<button id=""xxbutton2"" label=""Button 2"" />" & vbNewLine & _
'             "<button id=""xxbutton3"" label=""Button 3"" />" & vbNewLine & _
'             "</menu>"
End Sub

Sub test()
    Dim oAgenda As ToolboxAgenda
    
    Set oAgenda = New ToolboxAgenda
    
    oAgenda.CreateOrUpdateAgenda
    
End Sub















