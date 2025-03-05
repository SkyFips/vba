VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLogPrinter_file"
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
' Description: iLogPrinter implementation for log output to txt file            |
' ------------------------------------------------------------------------------|
' Version | Description                                                         |
'   1.0.0 | initial version                                                     |
'   1.0.1 | use "asString" instead of "asName"                                  |
'   1.0.2 | add error handling in case something is not printable               |
' ##############################################################################/
Implements iLogPrinter
Private WithEvents logger As cLogger
Attribute logger.VB_VarHelpID = -1

Dim oFileSys As Object
Dim oTxt     As Object
Dim level    As logLevels
Dim overwritten As Boolean

Private Sub Class_Initialize()
  level = INFO
End Sub
Private Sub Class_Terminate()
  txt.Close
End Sub
Private Property Get fileSys() As Object
  If oFileSys Is Nothing Then Set oFileSys = CreateObject("Scripting.FileSystemObject")
  Set fileSys = oFileSys
End Property
Private Property Get txt() As Object
  If oTxt Is Nothing Then
    Set oTxt = fileSys.CreateTextFile(fileSys.BuildPath(ThisWorkbook.Path, _
                                                        "log_" & Format(Now, "yyyy-mm-ddThh-mm-ss") & ".txt"))
  End If
  Set txt = oTxt
End Property
Private Sub logger_printLog(logLevel As logLevels, logText As String)
  On Error GoTo PrintProblem
  If logLevel <= level Then txt.WriteLine "[" & mLogger.asString(logLevel) & "][" & Now() & "] " & logText
  Exit Sub
PrintProblem:
  txt.WriteLine "[" & mLogger.asString(logLevel) & "][" & Now() & "] not able to print into txt file -> check the watch window"
  Debug.Print logText
  Debug.Print vbLf
  Err.Clear
End Sub

' ------------------------------------
' iLogPrinter methods
' ------------------------------------
Private Sub iLogPrinter_Log(logLevel As logLevels, Text As String)
  If logLevel <= level Then txt.WriteLine "[" & mLogger.asString(logLevel) & "][" & Now() & "] " & Text
End Sub
Private Property Get iLogPrinter_level() As logLevels
  If level = UNKNOWN Then level = logger.level
  iLogPrinter_level = level
End Property
Private Property Let iLogPrinter_level(l As logLevels)
  level = l
End Property
Private Property Set iLogPrinter_logger(l As cLogger)
  Set logger = l
End Property
