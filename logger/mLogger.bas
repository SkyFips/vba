Attribute VB_Name = "mLogger"
Option Explicit
' ##############################################################################\
'   Author(s):   Philipp Gorkiewicz                                             |
'   Description: log level identifier                                           |
'                                                                               |
' ##############################################################################|
'   Version | Description                                                       |
'     1.0.0 | initial Version                                                   |
' ##############################################################################/
Public Enum logLevels
  UNKNOWN
  FATAL
  WARN
  FAILURE
  INFO
  DEBUGGER
End Enum

Public Function logLevelAsName(level As logLevels) As String
  Select Case level
  Case 0: logLevelAsName = "UNKNOWN"
  Case 1: logLevelAsName = "FATAL"
  Case 2: logLevelAsName = "WARN"
  Case 3: logLevelAsName = "FAILURE"
  Case 4: logLevelAsName = "INFO"
  Case 5: logLevelAsName = "DEBUG"
  End Select
End Function
