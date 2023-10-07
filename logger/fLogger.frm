VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogger 
   Caption         =   "LOGGER"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12330
   OleObjectBlob   =   "frmLogger.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
' ##############################################################################\
' Author(s):   Philipp Gorkiewicz                                               |
' Version:     1.0.0                                                            |
' Description: implementation for iLogPrinter                                   |
'              used to print to userform                                        |
'                                                                               |
' ##############################################################################|
'   Version | Description                                                       |
'     1.0.0 | Initial Version                                                   |
' ##############################################################################/
Implements iLogger

Dim collMessages  As Collection
Private level     As enumLogLevel

Const ctrlMessage         As String = "lblMessage_"
Const countLabelMessages  As Byte = 20

Private Sub UserForm_Initialize()
  Dim counter As Byte
  Set collMessages = New Collection
  For counter = 1 To countLabelMessages
    With Me.Controls(ctrlMessage & counter)
      .Visible = False
    End With
  Next counter
  Call Me.Show(vbModeless)
End Sub
Private Sub UserForm_Terminate()
  Me.Hide
  Unload Me
End Sub

Private Sub refreshMessages()
  Dim Text    As Variant
  Dim counter As Byte
  For Each Text In collMessages
    counter = counter + 1
    With Me.Controls(ctrlMessage & counter)
      .Visible = True
      .Caption = Text
    End With
  Next Text
End Sub
Private Function convertIdToName(logLevel As enumLogLevel) As String
  Select Case logLevel
    Case 0: convertIdToName = "UNKNOWN"
    Case 1: convertIdToName = "FATAL"
    Case 2: convertIdToName = "WARN"
    Case 3: convertIdToName = "FAILURE"
    Case 4: convertIdToName = "INFO"
    Case 5: convertIdToName = "DEBUG"
  End Select
End Function

' ######################################################
' implementation iLogger
' ######################################################

Private Sub iLogger_Log(logLevel As enumLogLevel, Text As String)
  If logLevel <= level Then
    If collMessages.count >= countLabelMessages Then
      collMessages.Remove 1
    End If
    collMessages.Add "[" & convertLogLevelToName(logLevel) & "][" & Format(Now, "hh:mm:ss") & Right(Format(Timer, "0\.000"), 4) & "]" & Text
    refreshMessages
  End If
End Sub

Private Property Get iLogger_level() As enumLogLevel
  iLogger_level = level
End Property
Private Property Let iLogger_level(l As enumLogLevel)
  level = l
End Property
