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
'   Versioning:  Semantic versioning (https://semver.org)                       |
'   Source:      https://github.com/SkyFips/vba/tree/main/logger                |
'   Description: cLogger is the main class for logging with event printLog      |
'                for each Log                                                   |
'                Level printer available as LogWARN, LogDEBUG, ...              |
'                Default level is INFO                                          |
'                                                                               |
' ##############################################################################|
'   Version | Description                                                       |
'     1.0.0 | initial Version                                                   |
'     1.0.1 | remove "asName/enum" as not instance specific, improve description|
' ##############################################################################/
Implements iLogger
Public Event printLog(logLevel As logLevels, Text As String)
Dim oPrinters As Collection
Dim eLevel    As logLevels

Public Property Get level() As logLevels
  If eLevel = UNKNOWN Then eLevel = INFO
  level = eLevel
End Property
Public Property Let level(l As logLevels)
  eLevel = l
End Property
Public Sub log(logLevel As logLevels, logText As String)
  RaiseEvent printLog(logLevel, logText)
End Sub
Public Sub LogFATAL(logText As String)
  RaiseEvent printLog(FATAL, logText)
End Sub
Public Sub LogWARN(logText As String)
  RaiseEvent printLog(WARN, logText)
End Sub
Public Sub LogFAILURE(logText As String)
  RaiseEvent printLog(FAILURE, logText)
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
Public Property Get iLogger_level() As logLevels
  iLogger_level = level
End Property
Public Property Let iLogger_level(l As logLevels)
  level = l
End Property
Public Sub iLogger_Log(logLevel As logLevels, logMessage As String)
  Call log(logLevel, logMessage)
End Sub
