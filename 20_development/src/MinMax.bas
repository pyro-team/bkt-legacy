Attribute VB_Name = "MinMax"
Option Explicit


Public Function Min(ParamArray P())
' Minimum
Dim V, i As Long
  V = Null
  For i = LBound(P) To UBound(P)
    If Not IsNumeric(V) Or (P(i) < V) Then V = P(i)
  Next i
  Min = V
End Function
 
Public Function Max(ParamArray P())
' Maximum
Dim V, i As Long
  V = Null
  For i = LBound(P) To UBound(P)
    If Not IsNumeric(V) Or (P(i) > V) Then V = P(i)
  Next i
  Max = V
End Function

Public Function Mean(ParamArray P())
' Mean / average
Dim S As Double, n As Long, i As Long
  S = 0
  n = 0
  For i = LBound(P) To UBound(P)
    If IsNumeric(P(i)) Then
      S = S + CDbl(P(i))
      n = n + 1
    End If
  Next i
  If n > 0 Then
    Mean = S / n
  Else
    Mean = Null
  End If
End Function

Public Function Last(ParamArray P())
' Last value
  Last = P(UBound(P))
End Function

