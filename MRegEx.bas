Attribute VB_Name = "MRegEx"
' Based on code found at http://bytecomb.com/regular-expressions-in-vba/ 2014-03-18

Public Function RegExTest( _
    ByVal SourceString As String, _
    ByVal Pattern As String, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = True) As Boolean
  
    With New RegExp
        .MultiLine = MultiLine
        .IgnoreCase = IgnoreCase
        .Global = False
        .Pattern = Pattern
        RegExTest = .test(SourceString)
    End With
      
End Function

Public Function RegExMatch( _
    ByVal SourceString As String, _
    ByVal Pattern As String, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = True) As Variant
  
    Dim oMatches As MatchCollection
    With New RegExp
        .MultiLine = MultiLine
        .IgnoreCase = IgnoreCase
        .Global = False
        .Pattern = Pattern
        Set oMatches = .Execute(SourceString)
        If oMatches.Count > 0 Then
            RegExMatch = oMatches(0).Value
        Else
            RegExMatch = Null
        End If
    End With
  
End Function

Public Function RegExMatches( _
    ByVal SourceString As String, _
    ByVal Pattern As String, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = True, _
    Optional ByVal MatchGlobal As Boolean = True) As Variant
  
    Dim oMatch As Match
    Dim arrMatches
    Dim lngCount As Long
      
    ' Initialize to an empty array
    arrMatches = Array()
    With New RegExp
        .MultiLine = MultiLine
        .IgnoreCase = IgnoreCase
        .Global = MatchGlobal
        .Pattern = Pattern
        For Each oMatch In .Execute(SourceString)
            ReDim Preserve arrMatches(lngCount)
            arrMatches(lngCount) = oMatch.Value
            lngCount = lngCount + 1
        Next
    End With
  
    RegExMatches = arrMatches
End Function

Public Function RegExReplace( _
    ByVal SourceString As String, _
    ByVal Pattern As String, _
    ByVal ReplacePattern As String, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = True, _
    Optional ByVal MatchGlobal As Boolean = True) As String
  
    With New RegExp
        .MultiLine = MultiLine
        .IgnoreCase = IgnoreCase
        .Global = MatchGlobal
        .Pattern = Pattern
        RegExReplace = .Replace(SourceString, ReplacePattern)
    End With
  
End Function

