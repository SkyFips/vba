Attribute VB_Name = "mLogger"
Option Explicit
' ##############################################################################\
'   Author(s):   Philipp Gorkiewicz                                             |
'   License:     MIT (https://opensource.org/license/mit/)                      |
'   Versioning:  Semantic versioning (https://semver.org)                       |
'   Source:      https://github.com/SkyFips/vba/tree/main/logger                |
'   Description: log level identifier                                           |
'                                                                               |
' ##############################################################################|
'   Version | Description                                                       |
'     1.0.0 | initial Version                                                   |
'     1.1.0 | add method to convert from name to log level id                   |
'     1.1.1 | respond "UNKNOWN" if not known, use different name and keep old   |
'           | for backward compatibility                                        |
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
  logLevelAsName = asString(level)
End Function

Public Function logNameAsLevel(name As String) As logLevels
  logNameAsLevel = asLogLevel(name)
End Function

Public Function asString(level As logLevels) As String
  Select Case level
  Case 1: asString = "FATAL"
  Case 2: asString = "WARN"
  Case 3: asString = "FAILURE"
  Case 4: asString = "INFO"
  Case 5: asString = "DEBUG"
  Case Else: asString = "UNKNOWN"
  End Select
End Function

Public Function asLogLevel(name As String) As logLevels
  Select Case name
  Case "FATAL": asLogLevel = 1
  Case "WARN": asLogLevel = 2
  Case "FAILURE": asLogLevel = 3
  Case "INFO": asLogLevel = 4
  Case "DEBUG": asLogLevel = 5
  Case Else: asLogLevel = 0
  End Select
End Function
