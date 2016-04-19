Attribute VB_Name = "MReference"
Option Explicit

Sub NewListObject()
    'Example ListObject with connection to POTAHIST
    Dim oCN   As Object    'Connection
    Dim oRS   As Object    'Recordset
    Dim sql   As String
    
    sql = "SET NOCOUNT ON " & _
        "DECLARE @StartDate DateTime " & _
        "DECLARE @EndDate DateTime " & _
        "SET @StartDate = DateAdd(hh,-24,GetDate()) " & _
        "SET @EndDate = GetDate() " & _
        "SET NOCOUNT OFF " & _
        "SELECT DateTime," & _
         "AVG(CASE WHEN TagName = 'WIC3803.MEAS' THEN Value ELSE NULL END) [14705], " & _
         "AVG(CASE WHEN TagName = 'WI8079.MEAS' THEN Value ELSE NULL END) [14701] " & _
         "FROM History " & _
         "WHERE History.TagName IN ('WIC3803.MEAS','WI8079.MEAS') " & _
         "AND Value >= 100 " & _
         "AND wwRetrievalMode = 'Cyclic' " & _
         "AND wwResolution = 3600000 " & _
         "AND wwVersion = 'Latest' " & _
         "AND DateTime >= @StartDate " & _
         "AND DateTime <= @EndDate " & _
        "GROUP BY DateTime "

     
    Set oCN = CreateObject("ADODB.Connection")
    Set oRS = CreateObject("ADODB.Recordset")
         
    oCN.ConnectionString = "Provider=SQLOLEDB;" & _
          "Data Source=SKVANMNPAPP08;" & _
          "Initial Catalog=Runtime;" & _
          "UID=wwUser;" & _
          "PWD=wwUser"
    oCN.Open
    

    oRS.Open sql, oCN
      
    With Selection.Parent.ListObjects.Add(xlSrcQuery, oRS, _
        Destination:=Selection)
        .QueryTable.Refresh
        .Name = "tbl"
    End With
         
    oRS.Close
    oCN.Close
    Set oRS = Nothing
    Set oCN = Nothing
 
End Sub

