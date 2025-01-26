VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLogPrinter_watchWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ##############################################################################\
' Author(s):   Philipp Gorkiewicz                                               |
' License:     MIT (https://opensource.org/license/mit/)                        |
' Versioning:  https://semver.org                                               |
' Source:      https://github.com/SkyFips/vba/tree/main/logger                  |
' Description: iLogPrinter as watchWindow                                       |
' ------------------------------------------------------------------------------|
' Version | Description                                                         |
'   1.0.0 | initial Version                                                     |
'   1.0.1 | use "asString" instead of "asName"                                  |
'   1.0.2 | add time to log message                                             |
' ##############################################################################/
Implements iLogPrinter
Public WithEvents logger As cLogger
Attribute logger.VB_VarHelpID = -1
Private level As logLevels

Private Sub iLogPrinter_Log(logLevel As logLevels, logText As String)
  If logLevel <= level Then
    Debug.Print "[" & mLogger.asString(logLevel) & " " & Format(Now, "hh:mm:ss") & Right(Format(Timer, "0\.000"), 4) & "] " & logText
  End If
End Sub
Private Sub logger_printLog(logLevel As logLevels, logText As String)
  Call iLogPrinter_Log(logLevel, logText)
End Sub
Private Property Let iLogPrinter_level(l As logLevels)
  level = l
End Property
Private Property Get iLogPrinter_level() As logLevels
  If level = UNKNOWN Then level = logger.level
  iLogPrinter_level = level
End Property
Public Property Set iLogPrinter_logger(l As cLogger)
  Set logger = l
End Property
