Attribute VB_Name = "MHelper"

'   MHelper: Module to hold generic helper subs and functions for use in other macros.

Function wbSelectOrOpenWorkbook(sWorkbookName As String, sWorkbookPath As String) As Workbook
Attribute wbSelectOrOpenWorkbook.VB_Description = "MHelper: Module to hold generic helper subs and functions for use in other macros."
Attribute wbSelectOrOpenWorkbook.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim wb As Workbook
    On Error Resume Next
    
    If wb Is Nothing Then
        Workbooks.Open (sWorkbookPath & Application.PathSeparator & sWorkbookName)
        Set wb = Workbooks(sWorkbookName)
    End If
    On Error GoTo 0
    
    Set wbSelectOrOpenWorkbook = wb
End Function

Public Sub RenameWorksheet(ws As Worksheet, sName As String)
    On Error Resume Next
    If Not IsObject(ws.Parent.Worksheets(sName)) Then ws.Name = sName
    On Error GoTo 0
End Sub

Function sGetCellComment(Cell As Range) As String
    On Error Resume Next

    sGetCellComment = Cell.Comment.Text

    If Err <> 0 Then sGetCellComment = ""

End Function

Function vGetOutputFolder(Optional startFolder As Variant = -1) As Variant
    
    Dim fldr As FileDialog
    Dim vItem As Variant
    
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        If startFolder = -1 Then
            .InitialFileName = Application.DefaultFilePath
        Else
            If Right(startFolder, 1) <> "\" Then
                .InitialFileName = startFolder & "\"
            Else
                .InitialFileName = startFolder
            End If
        End If
        If .Show <> -1 Then GoTo NextCode
        vItem = .SelectedItems(1)
    End With
NextCode:
    vGetOutputFolder = vItem
    Set fldr = Nothing
End Function
