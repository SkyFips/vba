VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLogger_wks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ##############################################################################\
' Author(s):   Philipp Gorkiewicz                                               |
' Description: iLogger as worksheet                                             |
'                                                                               |
' ##############################################################################|
'   Version | Description                                                       |
'     1.0.0 | Initial Version                                                   |
' ##############################################################################/
Implements iLogger

Private level     As enumLogLevel
Private objSheet  As Excel.Worksheet

Private Sub Class_Initialize()
  level = INFO
End Sub
Private Property Get row() As LongPtr
  row = sheet.Cells(sheet.Rows.count, Columns("A").Column).End(xlUp).row + 1
End Property
Private Property Get sheet() As Excel.Worksheet
  If objSheet Is Nothing Then
    Dim a As Boolean
    a = Excel.Application.DisplayAlerts
    Excel.Application.DisplayAlerts = False
    If mHelper.sheetExist("log") Then Excel.ThisWorkbook.Sheets("log").Delete
    Excel.Application.DisplayAlerts = a
    Set objSheet = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    objSheet.name = "log"
    objSheet.Select
    objSheet.Cells(1, Columns("A").Column).Value = "level"
    objSheet.Cells(1, Columns("B").Column).Value = "time"
    objSheet.Cells(1, Columns("C").Column).Value = "text"
  End If
  Set sheet = objSheet
End Property
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
Private Sub iLogger_Log(logLevel As enumLogLevel, Text As String)
  If logLevel <= level Then
    Dim c As LongPtr
    Dim s As Boolean
    c = row
    s = Excel.Application.ScreenUpdating
    sheet.Select
    Excel.Application.ScreenUpdating = True
    sheet.Cells(c, Columns("A").Column).Value = "[" & convertLogLevelToName(logLevel) & "] "
    sheet.Cells(c, Columns("B").Column).Value = Now()
    sheet.Cells(c, Columns("C").Column).Value = Text
    sheet.Columns("A:C").EntireColumn.AutoFit
    Excel.Application.ScreenUpdating = s
  End If
End Sub

Private Property Get iLogger_level() As enumLogLevel
  iLogger_level = level
End Property
Private Property Let iLogger_level(l As enumLogLevel)
  level = l
End Property

