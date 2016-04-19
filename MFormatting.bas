Attribute VB_Name = "MFormatting"
'   modFormatting: Defines formatting options for things I tend to want to do often.
Option Explicit

Public Sub fmtCurrency()
    'Apply the "Currency" format, alternating between 2 and 0 decimal places.

    Const strCURRENCY As String = "_($* #,##0_);_($* (#,##0);_($* ""-""_);_(@_)"
    Const strCURRENCYZERO As String = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""_);_(@_)"
    
    If TypeName(Selection) = "Range" Then
        If Selection.NumberFormat = strCURRENCY Then
            Selection.NumberFormat = strCURRENCYZERO
        Else
            Selection.NumberFormat = strCURRENCY
        End If
    End If
End Sub

Public Sub fmtComma()
    'Apply the "Comma" format, alternating between 2 and 0 decimal places.
    
    Const strCOMMA As String = "_(* #,##0.00_);_(* (#,##0.00);_(* "" - ""??_);_(@_)"
    Const strCOMMAZERO As String = "_(* #,##0_);_(* (#,##0);_(* "" - ""??_);_(@_)"
    
    If TypeName(Selection) = "Range" Then
        If Selection.NumberFormat = strCOMMA Then
            Selection.NumberFormat = strCOMMAZERO
        Else
            Selection.NumberFormat = strCOMMA
        End If
    End If
    
    
End Sub

Public Sub ClearListObjectFilter(ByRef loTable As ListObject)
'   ClearListObjectFilter: Clear the filter from a ListObject

    If loTable.ShowAutoFilter Then
        loTable.AutoFilter.ShowAllData
    End If
End Sub


Public Sub RebuildDefaultStyles()

'The purpose of this macro is to remove all styles in the active
'workbook and rebuild the default styles.
'It rebuilds the default styles by merging them from a new workbook.
'Copied from http://support.microsoft.com/kb/291321

'Dimension variables.
   Dim MyBook As Workbook
   Dim tempBook As Workbook
   Dim CurStyle As Style

   'Set MyBook to the active workbook.
   Set MyBook = ActiveWorkbook
   On Error Resume Next
   'Delete all the styles in the workbook.
   For Each CurStyle In MyBook.Styles
      'If CurStyle.Name <> "Normal" Then CurStyle.Delete
      Select Case CurStyle.Name
         Case "20% - Accent1", "20% - Accent2", _
               "20% - Accent3", "20% - Accent4", "20% - Accent5", "20% - Accent6", _
               "40% - Accent1", "40% - Accent2", "40% - Accent3", "40% - Accent4", _
               "40% - Accent5", "40% - Accent6", "60% - Accent1", "60% - Accent2", _
               "60% - Accent3", "60% - Accent4", "60% - Accent5", "60% - Accent6", _
               "Accent1", "Accent2", "Accent3", "Accent4", "Accent5", "Accent6", _
               "Bad", "Calculation", "Check Cell", "Comma", "Comma [0]", "Currency", _
               "Currency [0]", "Explanatory Text", "Good", "Heading 1", "Heading 2", _
               "Heading 3", "Heading 4", "Input", "Linked Cell", "Neutral", "Normal", _
               "Note", "Output", "Percent", "Title", "Total", "Warning Text"
            'Do nothing, these are the default styles
         Case Else
            CurStyle.Delete
      End Select

   Next CurStyle

   'Open a new workbook.
   Set tempBook = Workbooks.Add

   'Disable alerts so you may merge changes to the Normal style
   'from the new workbook.
   Application.DisplayAlerts = False

   'Merge styles from the new workbook into the existing workbook.
   MyBook.Styles.Merge Workbook:=tempBook

   'Enable alerts.
   Application.DisplayAlerts = True

   'Close the new workbook.
   tempBook.Close

End Sub

