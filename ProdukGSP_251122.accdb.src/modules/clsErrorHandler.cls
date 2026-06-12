VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'NOTE:
'   This is a static class, meaning the developer does not need to instantiate it, it is always there.
'   To use it in your own projects you have to use this procedure: Export File out of this database, and Import File into your database.
'   A simple copy/paste of the code is NOT enough.

'REFERENCE:
'   https://learn.microsoft.com/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/189fb41b-cc3a-4999-a6d2-ba89f72d2870

Private Const LOGFILENAME As String = "Northwind-Dev.log"
Private m_strLogFilePath As String

Public Sub ErrorHandlerExample()
10        On Error GoTo Err_Handler

          'Code goes here.

Exit_Handler:
20        Exit Sub

Err_Handler:
30        clsErrorHandler.HandleError "YourModuleName", "YourProcedureName"
40        Resume Exit_Handler
End Sub

Private Sub Class_Initialize()
10        On Error GoTo Err_Handler

          'Create full path to error log file in same folder as current database, by concatenating the folder, a backslash, and a filename.
20        m_strLogFilePath = CurrentProject.Path & "\" & LOGFILENAME

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError "clsErrorHandler", "Class_Initialize"
50        Resume Exit_Handler
End Sub

'NOTE:
'   This main error handler should not itself have an error handler.
Public Sub HandleError(ByVal strModuleName As String, ByVal strProcedureName As String)
          Dim lngError        As Long
          Dim lngErrorLine    As Long
          Dim strError        As String
          Dim strMsg          As String

          'Get the current error information
10        strError = Err.Description
20        lngError = Err.Number
30        lngErrorLine = Erl                  'If no line number, this will be 0.

40        If lngError = 2501 Then Exit Sub    '2501 = The OpenForm action was canceled. Not really an error.

          'Log the error, even before we show it to the user.
50        LogErrorToFile strError, lngError, lngErrorLine, strModuleName, strProcedureName

          'Compose the error message by concatenating substrings.

60        strMsg = "The following error has occurred: " & vbCrLf
70        strMsg = strMsg & strError & vbCrLf
80        If lngErrorLine <> 0 Then strMsg = strMsg & "On line " & lngErrorLine & vbCrLf
90        strMsg = strMsg & strModuleName & "." & strProcedureName & vbCrLf & vbCrLf
100       strMsg = strMsg & "Error number: " & lngError & vbCrLf & vbCrLf
110       strMsg = strMsg & "Please notify the Help Desk.  Thank you!"

          'Show the message to the user. A developer running ACCDB can hit Ctrl+Break and Step in the debugger to get to the calling procedure.
120       MsgBox strMsg, vbCritical
End Sub

Private Sub LogErrorToFile(ByVal strError As String, ByVal lngError As Long, ByVal lngErrorLine As Long, ByVal strModuleName As String, ByVal strProcedureName As String)
10        On Error GoTo Err_Handler

          Dim intFile         As Integer

20        intFile = FreeFile
30        Open m_strLogFilePath For Append As intFile

40        Print #intFile, "Current Time   : " & Now
50        Print #intFile, "Module         : " & strModuleName
60        Print #intFile, "Procedure      : " & strProcedureName
70        Print #intFile, "Error String   : " & strError
80        Print #intFile, "Error Number   : " & lngError
90        Print #intFile, "Error Line     : " & IIf(lngErrorLine = 0, "<none>", lngErrorLine)
100       Print #intFile, "User           : " & Environ$("USERNAME")
110       Print #intFile, "Machine        : " & Environ$("COMPUTERNAME")
120       Print #intFile, "---------------"

130       Close #intFile

Exit_Handler:
140       Exit Sub

Err_Handler:
150       clsErrorHandler.HandleError "clsErrorHandler", "LogErrorToFile"
160       Resume Exit_Handler
End Sub
