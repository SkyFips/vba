VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ##############################################################################\
' (c)2022 Scientific Games International GmbH                                   |
'                                                                               |
' Author(s):      Philipp Gorkiewicz                                            |
' Version:        1.0.0                                                         |
' Creation date:  2022-04-02                                                    |
' Last change:    2022-04-02                                                    |
'                                                                               |
' Description:    logger interface                                              |
'                                                                               |
' ##############################################################################|
'     Version | Description                                                     |
'       1.0.0 | Initial Version                                                 |
'###############################################################################/
Public Enum enumLogLevel
  UNKNOWN
  FATAL
  WARN
  FAILURE
  INFO
  DEBUGGER
End Enum

Public Property Get level() As enumLogLevel
End Property
Public Property Let level(l As enumLogLevel)
End Property
Public Sub log(logLevel As enumLogLevel, _
               logMessage As String)
' log implementation, how it is logged0
End Sub
