Attribute VB_Name = "UTextJoin"
Option Explicit

Public Function TEXTJOIN(Delimiter As String, Ignore_Empty As Boolean, ParamArray Text1() As Variant) As String
'PURPOSE: Replicates The Excel 2016 Function CONCAT
'SOURCE: www.TheSpreadsheetGuru.com

Dim RangeArea As Variant
Dim Cell As Range

'Loop Through Each Cell in Given Input
  For Each RangeArea In Text1
    If TypeName(RangeArea) = "Range" Then
      For Each Cell In RangeArea
        If Len(Cell.Value) <> 0 Or Ignore_Empty = False Then
          TEXTJOIN = TEXTJOIN & Delimiter & Cell.Value
        End If
      Next Cell
    Else
      'Text String was Entered
        If Len(RangeArea) <> 0 Or Ignore_Empty = False Then
          TEXTJOIN = TEXTJOIN & Delimiter & RangeArea
        End If
    End If
  Next RangeArea

TEXTJOIN = Mid(TEXTJOIN, Len(Delimiter) + 1)

End Function
