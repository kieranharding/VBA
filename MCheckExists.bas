Attribute VB_Name = "MCheckExists"
'   modCheckExists: Determine if various items of a specified name exists in the active workbook.
'   This entire module, along with the idea, was initially copied from
'   http://itknowledgeexchange.techtarget.com/beyond-excel/page/8/ on 2013-11-21

Option Explicit

'    'If part of kieranaddins.xlam, these constants are defined in the Admin module. Otherwise, uncomment.
'Global Const Success = False
'Global Const Failure = True



Function NameExists(strName As String) As Boolean
Attribute NameExists.VB_Description = "If part of kieranaddins.xlam, these constants are defined in the Admin module. Otherwise, uncomment.\r\nGlobal Const Success = False\r\nGlobal Const Failure = True"
Attribute NameExists.VB_ProcData.VB_Invoke_Func = " \n14"

'   NameExists:     Determine if a name exists in a spreadsheet

'   Parameters:     strName - Name to be checked
'   Example:        If Not NameExists("Data") then SetupName("Data")


'   Date        Modification
'   2013-11-21  First Programming


    On Error GoTo ErrHandler
    NameExists = False      'Assume not found
    
    Dim objName As Object
    
    For Each objName In Names
        If objName.Name = strName Then
            NameExists = Right(Names(strName).Value, 5) <> "#REF!"
            Exit For
        End If
    Next
      
ErrHandler:

    If Err.Number <> 0 Then MsgBox _
        "NameExists - Error#" & Err.Number & vbCrLf & Err.Description, _
        vbCritical, "Error", Err.HelpFile, Err.HelpContext
        
    On Error GoTo 0
    
End Function

Function ShapeExists(strName As String) As Boolean
'   ShapeExists:    See if a Shape Exists
'   Parameters:     strName - Shape Name to be checked
'   Example:        If not ShapeExists("EasyButton") then _
'           Create_Easy_Button "easy", "Show_Prompt", 10, 8


'   Date        Modification
'   2013-11-21  First Programming
    
    On Error GoTo ErrHandler
    ShapeExists = False     'Assume not found
   
    Dim objName As Object
   
    For Each objName In ActiveSheet.Shapes
        If objName.Name = strName Then
            ShapeExists = True
            Exit For
        End If
    Next

ErrHandler:
   
    If Err.Number <> 0 Then MsgBox _
        "ShapeExists - Error#" & Err.Number & vbCrLf & Err.Description, _
        vbCritical, "Error", Err.HelpFile, Err.HelpContext
    On Error GoTo 0
End Function

Function WorkSheetExists(strName As String) As Boolean
'   WorkSheetExists:See if a Worksheet Exists
'   Parameters:     strName - Worksheet Name to be checked
'   Example:        If not WorkSheetExists("Data") then Setup_Data("Data")


'   Date        Modification
'   2013-11-21  First Programming

    On Error GoTo ErrHandler
    WorkSheetExists = False     'Assume not found
   
    Dim objName As Object
   
    For Each objName In Worksheets
        If objName.Name = strName Then
            WorkSheetExists = True
            Exit For
        End If
    Next
   
ErrHandler:
   
    If Err.Number <> 0 Then MsgBox _
        "WorkSheetExists - Error#" & Err.Number & vbCrLf & Err.Description, _
        vbCritical, "Error", Err.HelpFile, Err.HelpContext
    On Error GoTo 0
End Function
 
Function PivotTableExists(strWorksheet As String, strName As String) As Boolean
'   PivotTableExists:See if a PivotTable Exists
'   Parameters:     strName - PivotTable Name to be checked
'   Example:        If not PivotTableExists("pvtHrs") then Setup_pvtHrs


'   Date        Modification
'   2013-11-21  First Programming

   On Error GoTo ErrHandler
    PivotTableExists = False     'Assume not found
   
    Dim objName As Object
   
    For Each objName In Worksheets(strWorksheet).PivotTables
        If objName.Name = strName Then
            PivotTableExists = True
            Exit For
        End If
    Next
   
ErrHandler:
   
    If Err.Number <> 0 Then MsgBox _
        "PivotTableExists - Error#" & Err.Number & vbCrLf & Err.Description, _
        vbCritical, "Error", Err.HelpFile, Err.HelpContext
    On Error GoTo 0
End Function
 
Function ChartExists(strName As String) As Boolean

'   ChartExists:    See if a Chart Exists
'   Parameters:     strName - Chart Name to be checked
'   Example:        If not ChartExists("chtHrs") then Setup_chtHrs


'   Date        Modification
'   2013-11-21  First Programming

    On Error GoTo ErrHandler
    ChartExists = False     'Assume not found
    Dim objName As Object
   
    For Each objName In Charts
        If objName.Name = strName Then
            ChartExists = True
            Exit For
        End If
    Next
ErrHandler:
   
    If Err.Number <> 0 Then MsgBox _
        "ChartExists - Error#" & Err.Number & vbCrLf & Err.Description, _
        vbCritical, "Error", Err.HelpFile, Err.HelpContext
    On Error GoTo 0

End Function

