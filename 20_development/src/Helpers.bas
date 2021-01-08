Attribute VB_Name = "Helpers"
Option Explicit

Public ScaleFrom As MsoScaleFrom


Public Function PointsToCentimeters(ByVal pt As Double) As Double
    On Error Resume Next
    PointsToCentimeters = pt * 2.54 / 72
End Function
 
Public Function CentimetersToPoints(ByVal cm As Double) As Double
    On Error Resume Next
    CentimetersToPoints = cm * 72 / 2.54
End Function


Public Function ParagraphHeight(ByVal par As TextRange, Optional ByVal withParSpaces As Boolean = True) As Single
    Dim parHeight As Single
    
    ' Absatzhoehe bestimmen
    With par
        parHeight = .Lines.Count * LineHeight(par)
        If withParSpaces Then
            parHeight = parHeight + Max(0, .ParagraphFormat.SpaceBefore) + Max(0, .ParagraphFormat.SpaceAfter)
        End If
    End With
    
    ParagraphHeight = parHeight
End Function

Public Function LineHeight(ByVal par As TextRange) As Single
    With par
        If .ParagraphFormat.LineRuleWithin Then
            ' spacing = number of lines
            ' Annahme zur Korrektur der Abstände: Abstand zwischen zwei Zeilen ist 0.2pt
            LineHeight = .Font.Size * (Max(0, .ParagraphFormat.SpaceWithin) + 0.2)
        Else
            ' spacing = number of pt
            ' Annahme zur Korrektur der Abstände: Abstand zwischen zwei Zeilen ist 0.2pt
            LineHeight = .ParagraphFormat.SpaceWithin '+ 0.1 * .Font.Size
        End If
    End With
    
End Function



' CR und LF Zeichen am Ende eines Absatzes entfernen
Public Sub TrimNewLineCharacters(ByVal par As TextRange)
    Dim charCode As Long
    Dim deleteChar As Boolean
    
    Do
        charCode = Asc(par.Characters(par.Length, 1))
        deleteChar = (charCode = 10 Or charCode = 13)
        If deleteChar Then
            par.Characters(par.Length, 1).Delete
        End If
    Loop Until deleteChar = False
End Sub


' CR und LF Zeichen am Ende eines Strings entfernen
Public Function TrimParagraphText(ByVal text As String) As String
    Dim charCode As Long
    Dim deleteChar As Boolean
    
    If Len(text) = 0 Then Exit Function
    
    Do
        charCode = Asc(Right(text, 1))
        deleteChar = (charCode = 10 Or charCode = 13)
        If deleteChar Then
            text = Left(text, Len(text) - 1)
        End If
    Loop Until deleteChar = False Or Len(text) = 0
    
    TrimParagraphText = text
End Function



Public Sub DeleteUnselectedSlides(ByRef newPres As Presentation, ByRef selectedSlideRange As SlideRange)
    Dim sld As Slide
    Dim sldIndex As Long
    Dim deleteSlide As Boolean

    ' Folien entfernen, die nicht ausgewählt waren
    ' Folien rueckwaerts durchgehen und schauen, ob Index in Range dabei ist
    For sldIndex = newPres.Slides.Count To 1 Step -1
        deleteSlide = True
        For Each sld In selectedSlideRange
            If sld.SlideIndex = sldIndex Then
                deleteSlide = False
                Exit For
            End If
        Next
        If deleteSlide Then newPres.Slides(sldIndex).Delete
    Next
End Sub


' Funktionen für top, left, width, height unter Berücksichtigung der ScaleFrom-Einstellung

Public Function GetLeft(ByRef shp As Shape) As Single
    Dim pin_add As Single
    pin_add = 0
    If ScaleFrom = msoScaleFromMiddle Then
        pin_add = 0.5 * shp.Width
    ElseIf ScaleFrom = msoScaleFromBottomRight Then
        pin_add = shp.Width
    End If
    
    GetLeft = shp.Left + pin_add
End Function

Public Sub SetLeft(ByRef shp As Shape, ByVal lft As Single)
    shp.IncrementLeft lft - GetLeft(shp)
End Sub

Public Function GetTop(ByRef shp As Shape) As Single
    Dim pin_add As Single
    pin_add = 0
    If ScaleFrom = msoScaleFromMiddle Then
        pin_add = 0.5 * shp.Height
    ElseIf ScaleFrom = msoScaleFromBottomRight Then
        pin_add = shp.Height
    End If
    GetTop = shp.Top + pin_add
End Function

Public Sub SetTop(ByRef shp As Shape, ByVal tp As Single)
    shp.IncrementTop tp - GetTop(shp)
End Sub

Public Function GetWidth(ByRef shp As Shape) As Single
    GetWidth = shp.Width
End Function

Public Sub SetWidth(ByRef shp As Shape, ByVal wdth As Single)
    Dim factor As Single
    On Error GoTo ErrFallback
    'workaround for zero width
    If shp.Width = 0 Then
        shp.Width = 1
    End If
    factor = wdth / shp.Width
    shp.ScaleWidth factor, msoFalse, ScaleFrom
Exit Sub

ErrFallback:
    shp.Width = wdth
End Sub

Public Function GetHeight(ByRef shp As Shape) As Single
    GetHeight = shp.Height
End Function

Public Sub SetHeight(ByRef shp As Shape, ByVal hght As Single)
    Dim factor As Single
    On Error GoTo ErrFallback
    'workaround for zero height
    If shp.Height = 0 Then
        shp.Height = 1
    End If
    factor = hght / shp.Height
    shp.ScaleHeight factor, msoFalse, ScaleFrom
Exit Sub

ErrFallback:
    shp.Height = hght
End Sub
