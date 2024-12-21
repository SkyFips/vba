VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iLogPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ##############################################################################\
'   Author(s):   Philipp Gorkiewicz                                             |
'   License:     MIT (https://opensource.org/license/mit/)                      |
'   Versioning:  https://semver.org                                             |
'   Description: log printer interface                                          |
'                                                                               |
' ##############################################################################|
'   Version | Description                                                       |
'     1.0.0 | initial version                                                   |
'###############################################################################/
Public Property Get level() As logLevels
End Property
Public Property Let level(l As logLevels)
End Property
Public Sub log(logLevel As logLevels, logText As String)
End Sub
Public Property Set logger(l As cLogger)
End Property
