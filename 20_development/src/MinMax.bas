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
