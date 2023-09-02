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
'   Author(s):   Philipp Gorkiewicz                                             |
'   Description: logger interface                                               |
'                                                                               |
' ##############################################################################|
'   Version | Description                                                       |
'     1.0.0 | initial version                                                   |
'###############################################################################/
Public Property Get level() As logLevels
End Property

Public Property Let level(l As logLevels)
End Property

Public Sub log(logLevel As logLevels, logMessage As String)
End Sub
