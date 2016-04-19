Attribute VB_Name = "MStats"
Public Sub RegressionData()
'Perform a regression using Linest, along with charts and a couple of statistical indicators.'
'Data should already be in the sheet, arranged in columns starting with the XValues. Titles should be in
' row 1 and the first data point should be in column A. Y Values should be the last column, with no spaces in between.

'todo:  Add charting

Const EXTRA_COLUMNS As Integer = 8 'How many columns do I want between the data and the regression?
Const DEFAULT_ALPHA As Double = 0.05

Dim rngYValues, rngXValues, rngXValuesObs As Range 'Ranges that should aleady be populated
Dim rngRegression, rngAlpha, rngTCrit, rngConstant, rngCoeff, _
    rngTStats, rngModel, rngCI, rngD, rngResid As Range 'Ranges we will populate
Dim Cell As Range 'Scratch
Dim intVariables, intObservations As Integer 'Count the X-Variables and the Observations
Dim sh As Worksheet
Dim strLinest, strD, strModel As String
'Dim bool_Constant As Boolean

Set sh = ActiveSheet

'Count the observations and variables.
intVariables = sh.[=Counta($1:$1)] - 1
intObservations = sh.[=Count($A:$A)]

'Set up my ranges
Set rngXValues = sh.Range(Cells(2, 1), Cells(intObservations + 1, intVariables))
Set rngYValues = sh.Range(Cells(2, intVariables + 1), Cells(intObservations + 1, intVariables + 1))
Set rngRegression = sh.Range(Cells(2, intVariables + EXTRA_COLUMNS), _
    Cells(2 + intVariables, intVariables + EXTRA_COLUMNS + 4))
Set rngAlpha = rngRegression.Range("A1").Offset(intVariables + 1, 0)
Set rngConstant = rngAlpha.Offset(1, 5)
Set rngTCrit = rngAlpha.Offset(0, 5)
Set rngTStats = Range(rngTCrit.Offset(-intVariables - 1, 0), rngTCrit.Offset(-1, 0))
Set rngCoeff = Range(rngTStats.Range("A1").Offset(0, 2), rngTStats.Range("A1").Offset(intVariables - 1, 2))

Set rngModel = Range(Cells(2, intVariables + 2), Cells(intObservations + 1, intVariables + 2))
Set rngCI = rngModel.Offset(0, 1)
Set rngResid = rngCI.Offset(0, 1)
Set rngD = rngResid.Offset(0, 1)
Set rngXValuesObs = Range(rngXValues.Range("A1").Offset(-1, 0), rngXValues.Range("A1").Offset(-1, intVariables - 1))

'Create the necessary named ranges in the sheet
With sh.Names
    .Add Name:="XValues", RefersTo:="=" & rngXValues.Address
    .Add Name:="YValues", RefersTo:="=" & rngYValues.Address
    .Add Name:="YPredValues", RefersTo:="=" & rngModel.Address
    .Add Name:="Residuals", RefersTo:="=" & rngResid.Address
End With

'Enter my formulas
rngConstant.Offset(0, -1).Value = "Constant"
rngConstant.Value = True 'Constant first as Linest will need it
rngConstant.Validation.Delete
rngConstant.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="TRUE,FALSE"

rngAlpha.Offset(0, -1).Value = "Alpha"
rngAlpha.Value = DEFAULT_ALPHA 'T-Crit will need to know alpha
rngAlpha.NumberFormat = "0%"

'Run the regression - LINESTGAP doesn't work in here, though it does work if this formula is entered in the sheet.
strLinest = "=TRANSPOSE(IFERROR(LINEST(" & rngYValues.Address(ReferenceStyle:=xlR1C1) & _
    "," & rngXValues.Address(ReferenceStyle:=xlR1C1) & "," & rngConstant.Address(ReferenceStyle:=xlR1C1) & _
    ",TRUE),""--""))" 'We need to define the Linest formula for the appropriate range sizes
rngRegression.FormulaArray = strLinest
rngRegression.NumberFormat = "#,##0.00_);(#,##0.00)"
rngRegression.Range("A1").Offset(1, 3).NumberFormat = "#,##0_)" 'Degrees of Freedom is an integer
rngRegression.Range("A1").Offset(0, 2).NumberFormat = "00%_)" 'R2 in percent

'Apply labels
With rngRegression.Range("A1")
    .Offset(-1, 0).Value = "Coeff"
    .Offset(-1, 1).Value = "SE"
    .Offset(-1, 2).Value = "R2 / s"
    .Offset(-1, 3).Value = "F / dfe"
    .Offset(-1, 4).Value = "SSR / SSE"
    .Offset(-1, 5).Value = "T Stat"
    '.Offset(-1, 6).Value = "p"
End With

'Coefficients will show up in the reverse order to how they were entered in the data.
With sh.Range(rngRegression.Range("A1").Offset(0, -1), rngAlpha.Offset(-2, -1))
    .FormulaArray = "=REVERSEVECTOR(" & rngXValuesObs.Address(ReferenceStyle:=xlR1C1) & ")"
    .EntireColumn.AutoFit
End With
    
rngAlpha.Offset(-1, -1).Value = "Intercept"
    
'Calculate T-Critical to compare the T-Stats of the coefficients
rngTCrit.Offset(0, -1).Value = "Critical"
rngTCrit.Formula = "=TINV(" & rngAlpha.Address & "," & rngRegression.Range("A1").Offset(1, 3).Address & ")"
rngTCrit.NumberFormat = "#,##0.00_);(#,##0.00)"

'...and calculate the T-Stats
For Each Cell In rngTStats
    Cell.FormulaR1C1 = "=abs(R[0]C[-5]/R[0]C[-4])"
    Cell.NumberFormat = "#,##0.00_);(#,##0.00)"
    Next Cell
    
'Highlight cells with a T-Stat less than the critical value
rngTStats.FormatConditions.Delete
rngTStats.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
    Formula1:="=" & rngTCrit.Address
With rngTStats.FormatConditions(1).Font
    .Color = -16383844
    .TintAndShade = 0
End With
    
With rngTStats.FormatConditions(1).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13551615
    .TintAndShade = 0
End With
    
'Reverse the coefficients to use when creating the model
rngCoeff.FormulaArray = "=reversevector(" & rngCoeff.Offset(0, -7).Address(ReferenceStyle:=xlR1C1) & ")"
rngCoeff.Style = "Explanatory Text"

'Create the model values
rngModel.Range("A1").Offset(-1, 0).Value = "Model"
For Each Cell In rngModel
    Cell.FormulaArray = "=MMULT(" & rngXValuesObs.Address(RowAbsolute:=False, columnabsolute:=True, ReferenceStyle:=xlR1C1) _
        & "," & rngCoeff.Address(ReferenceStyle:=xlR1C1) & ")+" & _
        rngRegression.Range("A1").Offset(intVariables, 0).Address(ReferenceStyle:=xlR1C1)
    Cell.NumberFormat = Cell.Offset(0, -1).NumberFormat
    
Next Cell

'Distance from the Model to the Actual for each row
rngResid.Range("A1").Offset(-1, 0).Value = "Residual"
For Each Cell In rngResid
    Cell.FormulaR1C1 = "=R[0]C[-2]-R[0]C[-3]"
    Cell.NumberFormat = Cell.Offset(0, -2).NumberFormat
Next Cell
    
'D value for each row: D = 1+x'(X'X)^(-1)x'
' where x is the data set for the current observation and X is the entire dataset
rngModel.Range("A1").Offset(-1, 3).Value = "D"
For Each Cell In rngD
    Cell.FormulaArray = "=1+MMULT(" & rngXValuesObs.Address(RowAbsolute:=False, columnabsolute:=True, ReferenceStyle:=xlR1C1) _
        & ",MMULT(MINVERSE(MMULT(TRANSPOSE(" & rngXValues.Address(ReferenceStyle:=xlR1C1) & ")," & _
        rngXValues.Address(ReferenceStyle:=xlR1C1) & ")),TRANSPOSE(" & _
        rngXValuesObs.Address(RowAbsolute:=False, columnabsolute:=True, ReferenceStyle:=xlR1C1) & ")))"
    Cell.NumberFormat = "#,##0.00_);(#,##0.00)"
    Next Cell
    
'Confidence interval for each row
rngCI.Range("A1").Offset(-1, 0).Value = "+/-"
For Each Cell In rngCI
    Cell.Formula = "=" & rngTCrit.Address & "*" & rngRegression.Range("A1").Offset(1, 2).Address & _
        "*SQRT(" & Cell.Offset(0, 2).Address & ")"
    Cell.NumberFormat = Cell.Offset(0, -1).NumberFormat
Next Cell

rngAlpha.Offset(1, -1).Value = "Average +/-"
rngAlpha.Offset(1, 0).Formula = "=average(" & rngCI.Address & ")"

MsgBox ("Done")

End Sub

