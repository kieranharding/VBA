Attribute VB_Name = "MPotahist"
Option Explicit

Function oGetPOTAHISTConnection() As Object
'****************************************************
'   oGetPOTAHISTConnection: Open and return a
'   connection to POTAHIST
'****************************************************
    Dim oCN                     As Object    'Connection
    Dim connStr                 As String
     
    Set oCN = CreateObject("ADODB.Connection")
    
    On Error GoTo CleanUp
         
    connStr = "Provider=SQLOLEDB;" & _
      "Data Source=SKVANMNPAPP08;" & _
      "Initial Catalog=Runtime;" & _
      "UID=wwUser;" & _
      "PWD=wwUser"

    With oCN
        .ConnectionString = connStr
        .Open
    End With
   
    Set oGetPOTAHISTConnection = oCN
    
CleanUp:
    If Err.Number <> 0 Then
        Debug.Print Err.Number
        Debug.Print Err.Description
        Debug.Print Err.HelpContext
        Debug.Print Err.Source
    End If
        
End Function

