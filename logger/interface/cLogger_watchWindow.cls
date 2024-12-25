VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLogger_watchWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ##############################################################################\
'   Author(s):   Philipp Gorkiewicz                                             |
'   Description: iLogger as watchWindow                                         |
'                                                                               |
' ##############################################################################|
'   Version | Description                                                       |
'     1.0.0 | initial Version                                                   |
' ##############################################################################/
Implements iLogger

Private level As logLevels

Private Sub Class_Initialize()
  level = INFO
End Sub
Private Sub iLogger_Log(logLevel As logLevels, logText As String)
  If logLevel <= level Then
    Debug.Print "[" & mLogger.logLevelAsName(logLevel) & "] " & logText
  End If
End Sub
Private Property Get iLogger_level() As logLevels
  iLogger_level = level
End Property
Private Property Let iLogger_level(l As logLevels)
  level = l
End Property
