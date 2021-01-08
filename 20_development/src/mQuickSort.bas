Attribute VB_Name = "mQuickSort"
Option Explicit


' Zwei-dimensionales Array nach Werten in "erster Dimension" sortieren
' Array besteht aus Tupeln (Wert, Objekt), Sortierung erfolgt nach Wert
' Quelle für Quick-Sort Funktion: http://stackoverflow.com/questions/152319/vba-array-sort-function
Public Sub QuickSortM(vArray As Variant, inLow As Long, inHi As Long)

  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2, 1)

  While (tmpLow <= tmpHi)

     While (vArray(tmpLow, 1) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi, 1) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow, 1)
        vArray(tmpLow, 1) = vArray(tmpHi, 1)
        vArray(tmpHi, 1) = tmpSwap
        Set tmpSwap = vArray(tmpLow, 2)
        Set vArray(tmpLow, 2) = vArray(tmpHi, 2)
        Set vArray(tmpHi, 2) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If

  Wend

  If (inLow < tmpHi) Then QuickSortM vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSortM vArray, tmpLow, inHi

End Sub


Public Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)

  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)

     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If

  Wend

  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi

End Sub


