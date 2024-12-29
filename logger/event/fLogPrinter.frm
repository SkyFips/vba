VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fLogPrinter 
   Caption         =   "LOGGER"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12330
   OleObjectBlob   =   "fLogPrinter.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "fLogPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
' ##############################################################################\
' Author(s):   Philipp Gorkiewicz                                               |
' License:     MIT (https://opensource.org/license/mit/)                        |
' Versioning:  https://semver.org                                               |
' Source:      https://github.com/SkyFips/vba/tree/main/logger                  |
' Description: a logger form to print log messages into an excel userform       |
' ------------------------------------------------------------------------------|
' Version | Description                                                         |
'   1.0.0 | Initial Version                                                     |
'   1.0.1 | correct enum name                                                   |
'   1.1.0 | redraw the messages every time                                      |
'   1.1.1 | use "asString" instead of "asName"                                  |
' ##############################################################################/
Implements iLogPrinter
Private WithEvents logger As cLogger
Attribute logger.VB_VarHelpID = -1

Dim collMessages As Collection
Dim level        As logLevels
Dim mCancel      As Boolean
Dim ctrl         As Control
Dim topPos       As Integer
Dim oLogger      As cLogger

Private Sub UserForm_Initialize()
  Dim counter As Byte
  Set collMessages = New Collection
  level = INFO
  Call Me.Show(vbModeless)
End Sub
Private Sub UserForm_Terminate()
  Me.Hide
  Unload Me
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = vbFormControlMenu Then Cancel = True
  Hide
  mCancel = True
End Sub
Private Sub refreshMessages()
  Dim v As Variant
  Dim c As Byte

  If mCancel Then
    mCancel = False
    Call Me.Show(vbModeless)
  End If

  For c = 1 To collMessages.count
    With Me.Controls("lblMessage_" & c)
      .Caption = collMessages(c)
    End With
  Next c
End Sub
Private Sub logger_printLog(logLevel As logLevels, logText As String)
  Call iLogPrinter_Log(logLevel, logText)
End Sub

' ######################################################
' implementation iLogger
' ######################################################
Private Sub iLogPrinter_Log(logLevel As logLevels, Text As String)
  If logLevel <= level Then
    collMessages.Add "[" & mLogger.asString(logLevel) & " " & Format(Now, "hh:mm:ss") & Right(Format(Timer, "0\.000"), 4) & "] " & Text
    If Me.Controls.count <= 25 Then
      Set ctrl = Me.Controls.Add("Forms.Label.1", "lblMessage_" & collMessages.count)
      With ctrl
        .Top = topPos
        .Width = 589
        .font.name = "Courier New"
      End With
      topPos = topPos + 10
    Else
      collMessages.Remove 1
    End If
    refreshMessages
    Me.Repaint
  End If
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
