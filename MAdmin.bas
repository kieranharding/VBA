Attribute VB_Name = "MAdmin"
'   Admin Module: Contains supporting code for kieranaddins.xlam and other macros.


Option Explicit

Global Const Success As Boolean = False
Global Const Failure As Boolean = True


Sub KeyboardShortcuts()
    'KeyboardShortcuts:     Configure keyboard shortcuts the way I'd like them using Application.OnKey
    '                       This is called in the Workbook_Open method
    '                       + SHIFT
    '                       ^ CTRL
    '                       % ALT
    Application.OnKey "^+4", "fmtCurrency"
    Application.OnKey "^+1", "fmtComma"
End Sub

Sub ReAddin()
    ThisWorkbook.IsAddin = True
End Sub

Sub DeAddin()
    ThisWorkbook.IsAddin = False
End Sub

Sub FindThisFile()
'   FindFile:   Display a message box showing where this addin is saved
    Dim strPath As String
    If ThisWorkbook.FullName = "" Then strPath = "Unsaved" Else strPath = ThisWorkbook.FullName
    MsgBox strPath
End Sub


Function Settings(strMode As String) As Boolean

'   Settings:       Saves, sets, and restores current application settings
'   Parameters:     strMode - "Save", "Restore", "Clear", "Disable", "Debug"
'   Example:        bResult = Settings("Disable")

'   Date        Modification
'   2013-11-21  Initial Programming

'   Initially copied on 2013-11-21 from
'   http://itknowledgeexchange.techtarget.com/beyond-excel/building-a-library-of-routines-settings/
    
    On Error GoTo ErrHandler
    Settings = Failure                  'Assume the worst
   
    Static Setting(999, 4) As Variant  'Limit to 1,000 settings, prevent loops
    Static intLevel As Integer
    Select Case UCase(Trim(strMode))
        Case Is = "SAVE"
            Setting(intLevel, 0) = ActiveSheet.Type
            Setting(intLevel, 1) = ActiveSheet.Name
            Setting(intLevel, 2) = Application.EnableEvents
            Setting(intLevel, 3) = Application.ScreenUpdating
            Setting(intLevel, 4) = Application.Calculation
            intLevel = intLevel + 1
       
        Case Is = "RESTORE"
            If intLevel > 0 Then
                intLevel = intLevel - 1
                If Setting(intLevel, 0) = -4167 Then
                    Worksheets(Setting(intLevel, 1)).Activate
                Else
                    Charts(Setting(intLevel, 1)).Activate
                End If
                Application.EnableEvents = Setting(intLevel, 2)
                Application.ScreenUpdating = Setting(intLevel, 3)
                Application.Calculation = Setting(intLevel, 4)
            End If
      
        Case Is = "CLEAR"       'Remove saved settings
            intLevel = 0
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
           
        Case Is = "DISABLE"
            Application.EnableEvents = False
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
           
        Case Is = "DEBUG"
            Debug.Print intLevel, _
            Setting(intLevel, 0), _
            Setting(intLevel, 1), _
            Setting(intLevel, 2), _
            Setting(intLevel, 3), _
            Setting(intLevel, 4), _
   
    End Select
    
    Settings = Success           'Normal end - no errors
    
ErrHandler:
   
    If Err.Number <> 0 Then MsgBox _
        "Settings - Error#" & Err.Number & vbCrLf & Err.Description, _
        vbCritical, "Error", Err.HelpFile, Err.HelpContext
    On Error GoTo 0
End Function
    
