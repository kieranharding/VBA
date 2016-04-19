Attribute VB_Name = "UHistogram"
Public Function Histogram(rngData As Range, Optional intBinprecision = 2, Optional boolChart As Boolean = True)
'Create a histogram with charts, using Freedman-Diaconis' Rule for determining bin size.
'Depracated. Use the histogram button instead, once it's built.

Dim dblQ1, dblQ3, dblBinSize, dblTotal As Double
Dim intCount, intMin, intMax, intBinCount As Integer
Dim arrHistogram() As Double
Dim i, j As Integer 'Scratch

'todo:  Check that the data contains only numbers and is non-blank
intCount = rngData.Count
dblQ1 = WorksheetFunction.Quartile(rngData, 1)
dblQ3 = WorksheetFunction.Quartile(rngData, 3)

dblBinSize = 2 * (dblQ3 - dblQ1) / (intCount ^ (1 / 3)) 'That's Freedman-Diaconis' Rule
dblBinSize = WorksheetFunction.Round(dblBinSize, intBinprecision)

intMin = WorksheetFunction.Min(rngData)
intMax = WorksheetFunction.Max(rngData) + 1
intBinCount = WorksheetFunction.RoundUp((intMax - intMin) / dblBinSize, intBinprecision)

'Make sure I end up with the bin sizes I expect
'   Debug.Print "Q1 " & dblQ1
'   Debug.Print "BinSize " & dblBinSize
'   Debug.Print "Q3 " & dblQ3
'   Debug.Print "BinCount " & intBinCount

'Set rngOutput = Range(rngOutput.Range("A1"), rngOutput.Range("A1").Offset(intBinCount, 2))
ReDim arrHistogram(1 To intBinCount + 1, 1 To 3)

'todo:  Add Title data above the histogram
'       Before printing anything, check that the cells to be overwritten are non-blank
'       Create the charts; examples are in my personal workbook
'       Can't do any of these in a function, only if I make it into a macro.

For i = 1 To intBinCount
    arrHistogram(i, 1) = intMin + dblBinSize * (i - 1)
'    Debug.Print i & ":" & arrHistogram(i, 1)
Next i

arrHistogram(intBinCount + 1, 1) = arrHistogram(intBinCount, 1) + dblBinSize

dblTotal = 0

For i = 1 To intBinCount
    arrHistogram(i, 2) = 0
    For Each c In rngData
        If c.Value > arrHistogram(i, 1) And c.Value < arrHistogram(i + 1, 1) Then
            arrHistogram(i, 2) = arrHistogram(i, 2) + 1
        End If
    Next c
    dblTotal = dblTotal + arrHistogram(i, 2) / intCount
    arrHistogram(i, 3) = dblTotal
Next i

arrHistogram(intBinCount + 1, 3) = dblTotal

Histogram = arrHistogram

End Function
