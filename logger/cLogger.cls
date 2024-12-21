VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ##############################################################################\
'   Author(s):   Philipp Gorkiewicz                                             |
'   License:     MIT (https://opensource.org/license/mit/)                      |
'   Versioning:  https://semver.org                                             |
'   Description: log level identifier                                           |
'                                                                               |
' ##############################################################################|
'   Version | Description                                                       |
'     1.0.0 | initial Version                                                   |
' ##############################################################################/
Public Event printLog(logLevel As logLevels, Text As String)
Dim oPrinters As Collection
Dim eLevel    As logLevels

Public Enum logLevels
  UNKNOWN
  FATAL
  WARN
  FAILURE
  INFO
  DEBUGGER
End Enum
Public Property Get level() As logLevels
  If eLevel = UNKNOWN Then eLevel = INFO
  level = eLevel
End Property
Public Property Let level(l As logLevels)
  eLevel = l
End Property
Public Function asName(level As logLevels) As String
  Select Case level
  Case 0: asName = "UNKNOWN"
  Case 1: asName = "FATAL"
  Case 2: asName = "WARN"
  Case 3: asName = "FAILURE"
  Case 4: asName = "INFO"
  Case 5: asName = "DEBUG"
  End Select
End Function
Public Sub log(logLevel As logLevels, logText As String)
  RaiseEvent printLog(logLevel, logText)
End Sub
Public Sub LogWARN(logText As String)
  RaiseEvent printLog(WARN, logText)
End Sub
Public Sub LogFATAL(logText As String)
  RaiseEvent printLog(FATAL, logText)
End Sub
Public Sub LogINFO(logText As String)
  RaiseEvent printLog(INFO, logText)
End Sub
Public Sub LogDEBUG(logText As String)
  RaiseEvent printLog(DEBUGGER, logText)
End Sub
Public Sub AddPrinter(logPrinter As iLogPrinter)
  Set logPrinter.logger = Me
  If oPrinters Is Nothing Then Set oPrinters = New Collection
  oPrinters.Add logPrinter
End Sub

