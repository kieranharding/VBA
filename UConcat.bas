Attribute VB_Name = "UConcat"
Option Explicit

Public Function CONCAT(ParamArray Text1() As Variant) As String
'PURPOSE: Replicates The Excel 2016 Function CONCAT
'SOURCE: www.TheSpreadsheetGuru.com

Dim RangeArea As Variant
Dim Cell As Range

'Loop Through Each Cell in Given Input
  For Each RangeArea In Text1
    If TypeName(RangeArea) = "Range" Then
      For Each Cell In RangeArea
        If Len(Cell.Value) <> 0 Then
          CONCAT = CONCAT & Cell.Value
        End If
      Next Cell
    Else
      'Text String was Entered
        CONCAT = CONCAT & RangeArea
    End If
  Next RangeArea

End Function
