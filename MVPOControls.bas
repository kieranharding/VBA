Attribute VB_Name = "MVPOControls"
Option Explicit


Sub SummarizeControls()
'
' SummarizeControls Macro
' Create a Pivottable from SAP QMS export of Controls.
'

'
    Dim lRecords As Long
    Dim sDataRange As String
    Dim sDataSheet As String
    Dim wsPT1 As Worksheet 'Pivottable to arrange by sample
    Dim wsPT2 As Worksheet 'Pivottable to arrange by shift
    Dim iTimestampColumn As Integer 'Column # containing Task List Description
    
    sDataSheet = "'" & ActiveSheet.Name & "'"
    lRecords = ActiveSheet.UsedRange.Rows.Count
    
    iTimestampColumn = [=MATCH("Task list description",$A$1:$V$1,0)]
    
    'Create Useful Timestamp
    Columns(iTimestampColumn).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, iTimestampColumn).FormulaR1C1 = "Timestamp"
    With Cells(2, iTimestampColumn)
        .FormulaR1C1 = "=RC1+RC[-1]"
        .NumberFormat = "[$-409]m/d/yy h:mm AM/PM;@"
        .AutoFill Destination:=Range(Cells(2, iTimestampColumn), Cells(lRecords, iTimestampColumn))
    End With
    With Range(Cells(2, iTimestampColumn), Cells(lRecords, iTimestampColumn))
        .Copy
        .PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End With
        
    'Remove useless timestamp data
    Range(Columns(1), Columns(iTimestampColumn - 1)).Delete Shift:=xlToLeft
    
    'Rename anaysis columns
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "K2O"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Insol"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "NaCl"
    
    'Remove pH and comments
    Columns("F").Delete Shift:=xlToLeft
    Columns("G").Delete Shift:=xlToLeft
    
    'Replace 0s with blanks
    
    Columns("D:G").Replace What:="0", Replacement:="", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    'Get Shift
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Shift"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = _
        "=((RC[-1]-INT(RC[-1]))<0.25)*(INT(RC[-1])-0.25)+((RC[-1]-INT(RC[-1]))>=0.75)*(INT(RC[-1])+0.75)+AND((RC[-1]-INT(RC[-1])>=0.25),(RC[-1]-INT(RC[-1])<0.75))*(INT(RC[-1])+0.25)"
    Range("B2").Select
    Selection.AutoFill Destination:=Range(Cells(2, 2), Cells(lRecords, 2))
    Range(Cells(2, 2), Cells(lRecords, 2)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    sDataRange = ActiveSheet.UsedRange.Address
    'ActiveSheet.UsedRange.Select
    
    'Sample Summary
    Set wsPT1 = Sheets.Add
    Call RenameWorksheet(wsPT1, "By Sample")
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        sDataSheet & "!" & sDataRange, Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:="'" & wsPT1.Name & "'!R3C1", TableName:="tblBySample", DefaultVersion _
        :=xlPivotTableVersion14
    wsPT1.Select
    Cells(3, 1).Select
    With wsPT1.PivotTables("tblBySample")
        With .PivotFields("Task list description")
            .Orientation = xlRowField
            .Position = 1
        End With
        .AddDataField .PivotFields("K2O"), "Average of K2O", xlAverage
        .AddDataField .PivotFields("NaCl"), "Average of NaCl", xlAverage
        .AddDataField .PivotFields("Insol"), "Average of Insol", xlAverage
        .AddDataField .PivotFields("K2O"), "Samples collected (by K2O)", xlCount
        .PivotFields("Average of Insol").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        .PivotFields("Average of K2O").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        .PivotFields("Average of NaCl").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        .ColumnGrand = False
        .RowGrand = False
    End With
   
    'Shift Summary
    Set wsPT2 = Sheets.Add
    Call RenameWorksheet(wsPT2, "By Shift")
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        sDataSheet & "!" & sDataRange, Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:="'" & wsPT2.Name & "'!R3C1", TableName:="tblByShift", DefaultVersion _
        :=xlPivotTableVersion14
    With wsPT2.PivotTables("tblByShift")
        With .PivotFields("Shift")
            .Orientation = xlRowField
            .Position = 1
        End With
        With .PivotFields("Task list description")
            .Orientation = xlColumnField
        End With
        .AddDataField .PivotFields("K2O"), "Average of K2O", xlAverage
        .AddDataField .PivotFields("NaCl"), "Average of NaCl", xlAverage
        .AddDataField .PivotFields("Insol"), "Average of Insol", xlAverage
        .AddDataField .PivotFields("K2O"), "Samples collected (by K2O)", xlCount
        .PivotFields("Average of Insol").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        .PivotFields("Average of K2O").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        .PivotFields("Average of NaCl").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        .ColumnGrand = False
        .RowGrand = False
    End With
    
    'These fail and I don't know why
'    ThisWorkbook.SlicerCaches.Add(wsPT2.PivotTables(1), "Shift").Slicers.Add wsPT2, , "Shift", "Shift"
'    ThisWorkbook.SlicerCaches.Add(wsPT2.PivotTables(1), "Task List Description").Slicers.Add wsPT2
    
        
End Sub





