Attribute VB_Name = "ULinestGap"
Option Explicit
'Run LINEST on an array of visible cells that has gaps (blanks, text or errors) in the data.
Function linestGap(Ycells As Range, XA As Variant, Optional Const0 As Boolean = True, Optional Stats As Boolean = False) As Variant
Attribute linestGap.VB_Description = "Run LINEST on an array of visible cells that has gaps (blanks, text or errors) in the data."
Attribute linestGap.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim YA As Variant, Cell As Range
    Dim NumRows As Long, NumXcols As Long, YA2() As Double, XA2() As Double
    Dim i As Long, j As Long, k As Long
    Dim BlankRow As Boolean, RowA() As Long

    YA = Ycells.Value2 'Array to hold the Y Column/Row
    NumRows = UBound(YA) 'Determine the size of the series
    
    If TypeName(XA) = "Range" Then XA = XA.Value2 'Array to hold X Columns/Rows
    If UBound(YA, 2) > NumRows Then 'Transpose the data if it's stored in rows rather than columns
        YA = WorksheetFunction.Transpose(YA)
        NumRows = UBound(YA)
        XA = WorksheetFunction.Transpose(XA)
    End If
    
    If UBound(XA) < NumRows Then NumRows = UBound(XA)
    NumXcols = UBound(XA, 2)
    
    ReDim RowA(1 To NumRows)
    
       
    ' Create array of numbers of visible rows with no blanks
    k = 0
    i = 0
    For Each Cell In Ycells
        i = i + 1
        If Cell.Rows.Hidden = False Then
            BlankRow = False
            If IsNumeric(YA(i, 1)) = False Then
                BlankRow = True
            ElseIf YA(i, 1) = "" Then
                BlankRow = True
            Else
                For j = 1 To NumXcols
                    If IsNumeric(XA(i, j)) = False Then
                        BlankRow = True
                        Exit For
                    ElseIf XA(i, j) = "" Then
                        BlankRow = True
                        Exit For
                    End If
                Next j
            End If
            If BlankRow = False Then
                k = k + 1
                RowA(k) = i
            End If
        End If
    Next Cell


    ' Transfer non-blank visible rows to new array
    NumRows = k
    ReDim YA2(1 To NumRows, 1 To 1)
    ReDim XA2(1 To NumRows, 1 To NumXcols)

    For i = 1 To NumRows
        YA2(i, 1) = YA(RowA(i), 1)
        For j = 1 To NumXcols
            XA2(i, j) = XA(RowA(i), j)
        Next j
    Next i

    linestGap = WorksheetFunction.LinEst(YA2, XA2, Const0, Stats)

End Function


Function LogestGap(Ycells As Range, XA As Variant, Optional Const0 As Boolean = True, Optional Stats As Boolean = False) As Variant
    Dim YA As Variant, Cell As Range
    Dim NumRows As Long, NumXcols As Long, YA2() As Double, XA2() As Double
    Dim i As Long, j As Long, k As Long
    Dim BlankRow As Boolean, RowA() As Long

    YA = Ycells.Value2 'Array to hold the Y Column/Row
    NumRows = UBound(YA) 'Determine the size of the series
    
    If TypeName(XA) = "Range" Then XA = XA.Value2 'Array to hold X Columns/Rows
    If UBound(YA, 2) > NumRows Then 'Transpose the data if it's stored in rows rather than columns
        YA = WorksheetFunction.Transpose(YA)
        NumRows = UBound(YA)
        XA = WorksheetFunction.Transpose(XA)
    End If
    
    If UBound(XA) < NumRows Then NumRows = UBound(XA)
    NumXcols = UBound(XA, 2)
    
    ReDim RowA(1 To NumRows)
    
       
    ' Create array of numbers of visible rows with no blanks
    k = 0
    i = 0
    For Each Cell In Ycells
        i = i + 1
        If Cell.Rows.Hidden = False Then
            BlankRow = False
            If IsNumeric(YA(i, 1)) = False Then
                BlankRow = True
            ElseIf YA(i, 1) = "" Then
                BlankRow = True
            Else
                For j = 1 To NumXcols
                    If IsNumeric(XA(i, j)) = False Then
                        BlankRow = True
                        Exit For
                    ElseIf XA(i, j) = "" Then
                        BlankRow = True
                        Exit For
                    End If
                Next j
            End If
            If BlankRow = False Then
                k = k + 1
                RowA(k) = i
            End If
        End If
    Next Cell


    ' Transfer non-blank visible rows to new array
    NumRows = k
    ReDim YA2(1 To NumRows, 1 To 1)
    ReDim XA2(1 To NumRows, 1 To NumXcols)

    For i = 1 To NumRows
        YA2(i, 1) = YA(RowA(i), 1)
        For j = 1 To NumXcols
            XA2(i, j) = XA(RowA(i), j)
        Next j
    Next i

    LogestGap = WorksheetFunction.LogEst(YA2, XA2, Const0, Stats)

End Function

