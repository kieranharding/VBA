Attribute VB_Name = "UReverseVector"
Option Explicit
'*********************************************
' User-defined function
' that takes an array and returns it in
' the reverse order.
'*********************************************
Function ReverseVector(rRange As Range) As Variant
Attribute ReverseVector.VB_Description = "*********************************************\r\nUser-defined function\r\nthat takes an array and returns it in\r\nthe reverse order.\r\n*********************************************"
Attribute ReverseVector.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim vRevArr() As Variant 'Storage vectors
    Dim N, i As Long 'Length of Vector and scratch loop
    Dim retCol As Boolean: retCol = False
    Dim c As Range 'Scratch loop
    
    'Make sure we're dealing with just a vector (one-dimensional array)
    With rRange
        'If it's one cell, return it
        If .Rows.Count = 1 And .Columns.Count = 1 Then
            ReverseVector = rRange.Value
            Exit Function
        'If it's two dimensions, return #VALUE
        ElseIf .Rows.Count > 1 And .Columns.Count > 1 Then
            ReverseVector = xlErrValue
            Exit Function
        'Otherwise, get the length of the vector
        ElseIf .Rows.Count > .Columns.Count Then N = .Rows.Count
            Else: N = .Columns.Count
        End If
    End With
    
    'Determine whether to return the results in a row or a column
    'Default is to return results in a row
    If Application.Caller.Rows.Count > 1 Then retCol = True
        
    ReDim vRevArr(1 To N)
    i = 0
    
    'The actual reversal
    For Each c In rRange
        vRevArr(N - i) = c.Value
        i = i + 1
    Next c
    
    'Return the array
    If retCol Then
        ReverseVector = Application.Transpose(vRevArr)
    Else
        ReverseVector = vRevArr
    End If
End Function

