VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLogger_file"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ##############################################################################\
' Author(s):   Philipp Gorkiewicz                                               |
' Description: iLogger implementation for log output to txt file                |
'                                                                               |
' ##############################################################################|
'   version | description                                                       |
'     1.0.0 | initial version                                                   |
' ##############################################################################/
Implements iLogger

Dim oFileSys  As Object
Dim oTxt      As Object
Dim level     As logLevels

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
' ------------------------------------
' iLogger methods
' ------------------------------------
Private Sub iLogger_Log(logLevel As logLevels, Text As String)
  If logLevel <= level Then txt.WriteLine "[" & mLogger.logLevelAsName(logLevel) & "][" & Now() & "] " & Text
End Sub
Private Property Get iLogger_level() As logLevels
  iLogger_level = level
End Property
Private Property Let iLogger_level(l As logLevels)
  level = l
End Property
