Attribute VB_Name = "ToolboxRibbon"
Option Explicit
Public myRibbon As IRibbonUI

Private ConvertPointsToCentimeters As Boolean

Private oTrigger As TriggerInvalidate

Private AdjustmentValue As Integer
Private SplitRowsCols As Integer
Private SplitSep As Single


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
    
    KeysEnabled = IsMacScriptFileAccessible()
    
    AdjustmentValue = 1
    SplitRowsCols = 2
    SplitSep = CentimetersToPoints(0.2)
    
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
        label = "Wert " & AdjustmentValue & "/8"
    End Select

Exit Sub
Err_Handler:
    label = "ERROR"
End Sub

Sub isEnabled(control As IRibbonControl, ByRef enabled)
    Dim ctlId As String
    Dim shpRange As ShapeRange
    Dim firstShp As Shape
    Dim shapeCount As Long
    
    On Error GoTo Err_Handler
    
    ctlId = NormalizeShapePropertyControlId(control.Id)
    Set shpRange = GetActiveShapeRange()
    If Not shpRange Is Nothing Then
        shapeCount = shpRange.Count
        If shapeCount > 0 Then Set firstShp = shpRange(1)
    End If
    
    Select Case ctlId
    Case "actUpdateAgenda", "actRemoveAgenda"
        'Set oAgenda = New ToolboxAgenda
        'enabled = oAgenda.CanUpdateAgenda
        enabled = True
        
    Case Else
        ' Enabled-Status von Selection abhängig
        If shpRange Is Nothing Then
            enabled = False
        Else
            Select Case ctlId
            Case "ebVSep", "ebHSep", "actJoinShapesWithText", "actSplitSameWidth", "actSplitSameHeight", "actSplitSwap"
                ' Enabled bei Selection-Count > 1
                enabled = (shapeCount > 1)
                
            Case "actHConnect", "actVConnect", "actTextIntoShape"
                ' Enabled bei Selection-Count = 2
                enabled = (shapeCount = 2)
                
            Case "ebRectCorner"
                enabled = (Not firstShp Is Nothing And firstShp.Adjustments.Count >= AdjustmentValue)

            Case "ebMarginLeft", "ebMarginRight", "ebMarginTop", "ebMarginBottom", _
                 "ebParIndentFirst", "ebParIndentLeft", "ebParIndentRight", _
                 "ebParPreSep", "ebParPostSep", "ebParWithin", _
                 "cbWordWrap", "cbWordWrap2", "cbAutoSize", "cbAutoSize2"
                enabled = SelectionContainsTextFrame(shpRange)
    
            Case Else
                If Left(ctlId, 2) = "eb" Then
                    ' Enabled bei Textboxen, wenn Wert nicht-leer
                    enabled = (GetEditBoxValueForShapeRange(ctlId, shpRange) <> "")
                'ElseIf Left(ctlId, 3) = "act" Then
                '    ' Enabled bei Selection-Count = 1
                '    enabled = True
                Else
                    ' Enabled bei Selection-Count >= 1
                    enabled = (shapeCount > 0)
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
    Dim shpRange As ShapeRange
    Dim firstShp As Shape
    
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
    Case "toggleMacKeys"
        returnedVal = KeysEnabled
        Exit Function
    End Select
    
    Set shpRange = GetActiveShapeRange()
    If shpRange Is Nothing Then
        returnedVal = False
    Else
        Set firstShp = shpRange(1)
        Select Case control.Id
        ' Textboxen
        Case "cbWordWrap", "cbWordWrap2"
            returnedVal = (firstShp.TextFrame.WordWrap = msoTrue)
        Case "cbAutoSize", "cbAutoSize2"
            returnedVal = (firstShp.TextFrame.AutoSize = ppAutoSizeShapeToFitText)
        ' Shapes
        Case "cbLockAspectRatio", "cbLockAspectRatio2"
            returnedVal = (firstShp.LockAspectRatio = msoTrue)
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
    Dim shpRange As ShapeRange
    
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
    
    Case "toggleMacKeys"
        SetKeysEnabled (pressed)
    
    Case Else:
        ' Alle Controls die selektierte Shapes benötigen
        Set shpRange = GetActiveShapeRange()
        If shpRange Is Nothing Then Exit Sub
    
        For Each shp In shpRange
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
    Dim shpRange As ShapeRange
    
    On Error GoTo Err_Handler
    
    Select Case control.Id
        ' Split Multiply
        Case "ebSplitRows"
            returnedVal = SplitRowsCols
        Case "ebSplitSep"
            returnedVal = Round(PointsToCentimeters(SplitSep), 2)
        Case Else
            Set shpRange = GetActiveShapeRange()
            If shpRange Is Nothing Then
                returnedVal = ""
            Else
                returnedVal = GetEditBoxValueForShapeRange(control.Id, shpRange)
            End If
    End Select

Exit Sub
Err_Handler:
    returnedVal = ""
End Sub

Private Function GetEditBoxValueForShapeRange(ByVal ctlId As String, ByVal shpRange As ShapeRange) As String
    Dim returnedVal As String
    Dim shps As Variant
    Dim firstShp As Shape
    Dim shapeCount As Long
    Dim numericValue As Single
    
    On Error GoTo Err_Handler
    
    If shpRange Is Nothing Then
        GetEditBoxValueForShapeRange = ""
        Exit Function
    End If
    
    shapeCount = shpRange.Count
    If shapeCount = 0 Then
        GetEditBoxValueForShapeRange = ""
        Exit Function
    End If
    Set firstShp = shpRange(1)
    
    Select Case ctlId
        ' Objektabstand
        Case "ebVSep"
            If shapeCount >= 2 Then
                shps = ShapeRangeSortedByTop(shpRange)
                returnedVal = shps(2).Top - shps(1).Top - shps(1).Height
                If returnedVal <> "" And ConvertPointsToCentimeters Then returnedVal = Round(PointsToCentimeters(returnedVal), 2)
            End If
        Case "ebHSep"
            If shapeCount >= 2 Then
                shps = ShapeRangeSortedByLeft(shpRange)
                returnedVal = shps(2).Left - shps(1).Left - shps(1).Width
                If returnedVal <> "" And ConvertPointsToCentimeters Then returnedVal = Round(PointsToCentimeters(returnedVal), 2)
            End If
        
        Case Else
            If TryGetShapePropertyValue(firstShp, ctlId, numericValue) Then
                returnedVal = FormatShapePropertyValue(ctlId, numericValue)
            Else
                Debug.Print ctlId
            End If
    End Select
    GetEditBoxValueForShapeRange = returnedVal

Exit Function
Err_Handler:
    GetEditBoxValueForShapeRange = ""
End Function


' Funktion für Änderungen an Textboxen mit Integer-Werten
' Zu control.id gehörige Eigenschaft wird an dem Wert in der Textbox angepasst
Sub ebIntValue_onChange(control As IRibbonControl, text As String)
    Dim value As Integer
    Dim shp As Shape
    Dim shpRange As ShapeRange
    
    On Error GoTo Err_Handler
    
    value = CInt(text)
    
    ' Split Multiply
    If control.Id = "ebSplitRows" Then
        SplitRowsCols = Max(2, value)
        myRibbon.Invalidate
        Exit Sub
    End If
    
    Set shpRange = GetActiveShapeRange()
    If shpRange Is Nothing Then Exit Sub
    
    For Each shp In shpRange
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
    Dim lastShp As Shape
    Dim shps As Variant
    Dim shpIdx As Integer
    Dim shpRange As ShapeRange
    Dim propertyCtlId As String
    
    On Error GoTo Err_Handler
    
    ' Split Multiply
    If control.Id = "ebSplitSep" Then
        SplitSep = CentimetersToPoints(Max(0, CSng(text)))
        myRibbon.Invalidate
        Exit Sub
    End If
    
    Set shpRange = GetActiveShapeRange()
    If shpRange Is Nothing Then Exit Sub
    propertyCtlId = NormalizeShapePropertyControlId(control.Id)
    
    Select Case propertyCtlId
    Case "ebHSep"
        shps = ShapeRangeSortedByLeft(shpRange)
    Case "ebVSep"
        shps = ShapeRangeSortedByTop(shpRange)
    Case Else
        Set shps = shpRange
    End Select
    
    value = ParseShapePropertyInputValue(propertyCtlId, text)
    If RequiresPointConversion(propertyCtlId) Then
        value = CentimetersToPoints(value)
    End If
    
    For shpIdx = 1 To shpRange.Count
        'For Each shp In ActiveWindow.Selection.ShapeRange
        Set shp = shps(shpIdx)
        Select Case propertyCtlId
        Case "ebVSep"
            If Not lastShp Is Nothing Then
                shp.Top = lastShp.Top + lastShp.Height + value
            End If
        Case "ebHSep"
            If Not lastShp Is Nothing Then
                shp.Left = lastShp.Left + lastShp.Width + value
            End If
        Case Else
            SetShapeSettingSingle shp, propertyCtlId, value
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
    Dim shpRange As ShapeRange
    Dim propertyCtlId As String
    
    On Error GoTo Err_Handler
    
    Set shpRange = GetActiveShapeRange()
    If shpRange Is Nothing Then Exit Sub
    propertyCtlId = NormalizeShapePropertyControlId(control.Id)
    
    Select Case control.Id
    Case "resHSep"
        shps = ShapeRangeSortedByLeft(shpRange)
    Case "resVSep"
        shps = ShapeRangeSortedByTop(shpRange)
    Case Else
        Set shps = shpRange
    End Select
    
    For shpIdx = 1 To shpRange.Count
        'For Each shp In ActiveWindow.Selection.ShapeRange
        Set shp = shps(shpIdx)
        Select Case control.Id
        ' Objektabstand
        Case "resVSep"
            If Not lastShp Is Nothing Then
                shp.Top = lastShp.Top + lastShp.Height
            End If
        Case "resHSep"
            If Not lastShp Is Nothing Then
                shp.Left = lastShp.Left + lastShp.Width
            End If
        
        Case Else
            SetShapeSettingSingle shp, propertyCtlId, 0
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
    Dim ptValue As Single
    Dim cmValue As Single
    Dim intValue As Single
    Dim shps As Variant
    Dim shpIdx As Integer
    Dim lastShp As Shape
    Dim newValue As Single
    Dim oldValue As Single
    Dim firstDelta As Single
    Dim shpRange As ShapeRange
    Dim propertyCtlId As String
    
    On Error GoTo Err_Handler
    
    ptValue = value
    cmValue = value * 0.1
    intValue = value
    If IsShiftKeyDown Then
        ptValue = 10 * ptValue
        cmValue = 10 * cmValue
        intValue = 10 * intValue
    ElseIf Not IsControlKeyDown Then
        ptValue = 3 * ptValue
        cmValue = 2 * cmValue
        intValue = 5 * intValue
    End If
    
    ' Split Multiply
    Select Case control.Id
        Case "incSplitRows", "decSplitRows"
            SplitRowsCols = Max(2, SplitRowsCols + value)
            myRibbon.Invalidate
            Exit Sub
        Case "incSplitSep", "decSplitSep"
            SplitSep = Max(0, CentimetersToPoints(Round(PointsToCentimeters(SplitSep), 1) + cmValue))
            myRibbon.Invalidate
            Exit Sub
    End Select
    
    Set shpRange = GetActiveShapeRange()
    If shpRange Is Nothing Then Exit Sub
    propertyCtlId = NormalizeShapePropertyControlId(control.Id)
    
    Select Case control.Id
    Case "incHSep", "decHSep"
        shps = ShapeRangeSortedByLeft(shpRange)
        oldValue = shps(2).Left - shps(1).Left - shps(1).Width
    Case "incVSep", "decVSep"
        shps = ShapeRangeSortedByTop(shpRange)
        oldValue = shps(2).Top - shps(1).Top - shps(1).Height
    Case Else
        ' cmValue
        Set shps = shpRange
        oldValue = GetShapeSettingSingle(shps(1), propertyCtlId)
    End Select
    
    Select Case control.Id
    Case "incParPreSep", "decParPreSep", "incParPostSep", "decParPostSep"
        newValue = Max(0, Round(oldValue) + ptValue)
    Case "incParWithin", "decParWithin"
        newValue = Max(0, Round(oldValue, 1) + cmValue)
    Case "incParIndentFirst", "decParIndentFirst"
        newValue = CentimetersToPoints(Round(PointsToCentimeters(oldValue), 1) + cmValue)
    Case "incRotation", "decRotation"
        If intValue = value * 10 Then
            'Make rotation big step 45 degree
            intValue = intValue * 4.5
        End If
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
    
    For shpIdx = 1 To shpRange.Count
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
        Case Else
            If Not IsAltKeyDown Then
                SetShapeSettingSingle shp, propertyCtlId, newValue
            Else
                oldValue = GetShapeSettingSingle(shp, propertyCtlId)
                SetShapeSettingSingle shp, propertyCtlId, oldValue + firstDelta
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

Private Function NormalizeShapePropertyControlId(ByVal controlID As String) As String
    If Left$(controlID, 3) = "inc" Or Left$(controlID, 3) = "dec" Or Left$(controlID, 3) = "res" Then
        NormalizeShapePropertyControlId = "eb" & Mid$(controlID, 4)
    Else
        NormalizeShapePropertyControlId = controlID
    End If
End Function

Private Function ParseShapePropertyInputValue(ByVal controlID As String, ByVal text As String) As Single
    controlID = NormalizeShapePropertyControlId(controlID)

    Select Case controlID
    Case "ebHSep", "ebVSep", "ebRectCorner"
        ParseShapePropertyInputValue = CSng(text)
    Case Else
        ParseShapePropertyInputValue = Max(0, CSng(text))
    End Select
End Function

Private Function RequiresPointConversion(ByVal controlID As String) As Boolean
    controlID = NormalizeShapePropertyControlId(controlID)

    Select Case controlID
    Case "ebParPreSep", "ebParPostSep", "ebParWithin", "ebRotation", "ebLineWeight", "ebRectCorner"
        RequiresPointConversion = False
    Case Else
        RequiresPointConversion = ConvertPointsToCentimeters
    End Select
End Function

Private Function FormatShapePropertyValue(ByVal controlID As String, ByVal value As Single) As String
    controlID = NormalizeShapePropertyControlId(controlID)

    Select Case controlID
    Case "ebTranspFill", "ebTranspLine"
        FormatShapePropertyValue = Max(0, Round(value * 100))
    Case "ebLineWeight", "ebRectCorner"
        FormatShapePropertyValue = Max(0, Round(value, 2))
    Case "ebParPreSep", "ebParPostSep", "ebParWithin", "ebRotation"
        FormatShapePropertyValue = value
    Case Else
        If ConvertPointsToCentimeters Then
            FormatShapePropertyValue = Round(PointsToCentimeters(value), 2)
        Else
            FormatShapePropertyValue = value
        End If
    End Select
End Function

Private Function TryGetParagraphPropertyValue(ByVal shp As Shape, ByVal controlID As String, ByRef value As Single) As Boolean
    Dim textRange As TextRange2

    On Error GoTo ErrHandler

    controlID = NormalizeShapePropertyControlId(controlID)
    Set textRange = shp.TextFrame2.TextRange

    If textRange.Paragraphs.Count > 0 Then
        Select Case controlID
        Case "ebParIndentFirst"
            value = textRange.Paragraphs(1).ParagraphFormat.FirstLineIndent
        Case "ebParIndentLeft"
            value = textRange.Paragraphs(1).ParagraphFormat.LeftIndent
        Case "ebParIndentRight"
            value = textRange.Paragraphs(1).ParagraphFormat.RightIndent
        Case "ebParPreSep"
            value = textRange.Paragraphs(1).ParagraphFormat.SpaceBefore
        Case "ebParPostSep"
            value = textRange.Paragraphs(1).ParagraphFormat.SpaceAfter
        Case "ebParWithin"
            value = textRange.Paragraphs(1).ParagraphFormat.SpaceWithin
        Case Else
            Exit Function
        End Select
    Else
        Select Case controlID
        Case "ebParIndentFirst"
            value = textRange.ParagraphFormat.FirstLineIndent
        Case "ebParIndentLeft"
            value = textRange.ParagraphFormat.LeftIndent
        Case "ebParIndentRight"
            value = textRange.ParagraphFormat.RightIndent
        Case "ebParPreSep"
            value = textRange.ParagraphFormat.SpaceBefore
        Case "ebParPostSep"
            value = textRange.ParagraphFormat.SpaceAfter
        Case "ebParWithin"
            value = textRange.ParagraphFormat.SpaceWithin
        Case Else
            Exit Function
        End Select
    End If

    TryGetParagraphPropertyValue = True
    Exit Function

ErrHandler:
End Function

Private Function TryGetShapePropertyValue(ByVal shp As Shape, ByVal controlID As String, ByRef value As Single) As Boolean
    On Error GoTo ErrHandler

    controlID = NormalizeShapePropertyControlId(controlID)

    Select Case controlID
    Case "ebMarginLeft"
        value = shp.TextFrame2.MarginLeft
    Case "ebMarginRight"
        value = shp.TextFrame2.MarginRight
    Case "ebMarginTop"
        value = shp.TextFrame2.MarginTop
    Case "ebMarginBottom"
        value = shp.TextFrame2.MarginBottom
    Case "ebPosLeft"
        value = GetLeft(shp)
    Case "ebPosTop"
        value = GetTop(shp)
    Case "ebPosRight"
        value = GetWidth(shp)
    Case "ebPosBottom"
        value = GetHeight(shp)
    Case "ebRotation"
        value = shp.Rotation
    Case "ebTranspFill"
        value = Max(0, shp.Fill.Transparency)
    Case "ebTranspLine"
        value = Max(0, shp.Line.Transparency)
    Case "ebLineWeight"
        value = Max(0, shp.Line.Weight)
    Case "ebRectCorner"
        If shp.Adjustments.Count < AdjustmentValue Then Exit Function
        value = RoundedCornerSize(shp, AdjustmentValue)
    Case Else
        TryGetShapePropertyValue = TryGetParagraphPropertyValue(shp, controlID, value)
        Exit Function
    End Select

    TryGetShapePropertyValue = True
    Exit Function

ErrHandler:
End Function

Private Function TrySetParagraphPropertyValue(ByVal shp As Shape, ByVal controlID As String, ByVal newValue As Single) As Boolean
    Dim textRange As TextRange2

    On Error GoTo ErrHandler

    controlID = NormalizeShapePropertyControlId(controlID)
    Set textRange = shp.TextFrame2.TextRange

    Select Case controlID
    Case "ebParIndentFirst"
        textRange.ParagraphFormat.FirstLineIndent = newValue
    Case "ebParIndentLeft"
        textRange.ParagraphFormat.LeftIndent = newValue
    Case "ebParIndentRight"
        textRange.ParagraphFormat.RightIndent = newValue
    Case "ebParPreSep"
        textRange.ParagraphFormat.SpaceBefore = newValue
    Case "ebParPostSep"
        textRange.ParagraphFormat.SpaceAfter = newValue
    Case "ebParWithin"
        textRange.ParagraphFormat.SpaceWithin = newValue
    Case Else
        Exit Function
    End Select

    TrySetParagraphPropertyValue = True
    Exit Function

ErrHandler:
End Function

Private Function TrySetShapePropertyValue(ByVal shp As Shape, ByVal controlID As String, ByVal newValue As Single) As Boolean
    On Error GoTo ErrHandler

    controlID = NormalizeShapePropertyControlId(controlID)

    Select Case controlID
    Case "ebMarginLeft"
        shp.TextFrame2.MarginLeft = newValue
    Case "ebMarginRight"
        shp.TextFrame2.MarginRight = newValue
    Case "ebMarginTop"
        shp.TextFrame2.MarginTop = newValue
    Case "ebMarginBottom"
        shp.TextFrame2.MarginBottom = newValue
    Case "ebPosLeft"
        SetLeft shp, newValue
    Case "ebPosTop"
        SetTop shp, newValue
    Case "ebPosRight"
        SetWidth shp, newValue
    Case "ebPosBottom"
        SetHeight shp, newValue
    Case "ebRotation"
        shp.Rotation = newValue
    Case "ebTranspFill"
        shp.Fill.Transparency = Min(1, Max(0, newValue))
    Case "ebTranspLine"
        shp.Line.Transparency = Min(1, Max(0, newValue))
    Case "ebLineWeight"
        shp.Line.Weight = Max(0, newValue)
    Case "ebRectCorner"
        SetRoundedCornerSize shp, newValue, AdjustmentValue
    Case Else
        TrySetShapePropertyValue = TrySetParagraphPropertyValue(shp, controlID, newValue)
        Exit Function
    End Select

    TrySetShapePropertyValue = True
    Exit Function

ErrHandler:
End Function

Function GetShapeSettingSingle(ByVal shp As Shape, controlID As String) As Single
    If Not TryGetShapePropertyValue(shp, controlID, GetShapeSettingSingle) Then
        Debug.Print controlID
    End If
End Function

Sub SetShapeSettingSingle(ByVal shp As Shape, controlID As String, ByVal newValue As Single)
    If Not TrySetShapePropertyValue(shp, controlID, newValue) Then
        Debug.Print controlID
    End If
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
    Case "actSameHeightMean"
        SetSameHeight "Mean"
    Case "actSameHeightLast"
        SetSameHeight "Last"
    
    Case "actSameWidth", "actSameWidth2"
        If IsShiftKeyDown Then
            SetSameWidth "Min"
        Else
            SetSameWidth
        End If
    Case "actSameWidthMin"
        SetSameWidth "Min"
    Case "actSameWidthMean"
        SetSameWidth "Mean"
    Case "actSameWidthLast"
        SetSameWidth "Last"

    Case "actSwap", "actSwap2"
        SwapPosition
    Case "actSwapSize"
        SwapPositionSize
    case "actReplaceKeepSize"
        ReplaceKeepSize
    
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
    Case "actAddProtectedHyphen"
        AddProtectedHyphen
    Case "actAddProtectedSpace"
        AddProtectedSpace
    Case "actAddProtectedNarrowSpace"
        AddProtectedNarrowSpace
    
    ' Texte ersetzen/löschen
    Case "actShapeTextReplace"
        ReplaceAllText
    Case "actShapeTextRemove"
        RemoveAllText
    
    ' Einfügen
'    Case "actPasteAsPicturePng"
'        PasteAsPicturePng
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
        
    'Process
    Case "actProcessAdd"
        AddProcessShapes
    
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
    
    ' Multiply Split
    Case "actSplitHorizontal"
        SplitShapes SplitRowsCols, SplitSep
    Case "actSplitVertical"
        SplitShapes SplitRowsCols, SplitSep, True
    Case "actMultiplyHorizontal"
        MultiplyShapes SplitRowsCols, SplitSep
    Case "actMultiplyVertical"
        MultiplyShapes SplitRowsCols, SplitSep, True
    
    ' Arrange
    Case "actArrangeTop"
        ArrangeByLast 1
    Case "actArrangeBottom"
        ArrangeByLast 2
    Case "actArrangeLeft"
        ArrangeByLast 3
    Case "actArrangeRight"
        ArrangeByLast 4
    Case "actArrangeVMiddle"
        ArrangeByLast 5
    Case "actArrangeMiddle"
        ArrangeByLast 6
    
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
