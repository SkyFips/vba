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
'     1.0.1 | use enum from module                                              |
' ##############################################################################/
Implements iLogger
Dim oLevel  As logLevels

Private objSheet  As Excel.Worksheet

Private Sub Class_Initialize()
  oLevel = INFO
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
Private Sub iLogger_Log(logLevel As logLevels, Text As String)
  If logLevel <= oLevel Then
    Dim c As LongPtr
    Dim s As Boolean
    c = row
    s = Excel.Application.ScreenUpdating
    sheet.Select
    Excel.Application.ScreenUpdating = True
    sheet.Cells(c, Columns("A").Column).Value = "[" & mLogger.logLevelAsName(logLevel) & "] "
    sheet.Cells(c, Columns("B").Column).Value = Now()
    sheet.Cells(c, Columns("C").Column).Value = Text
    sheet.Columns("A:C").EntireColumn.AutoFit
    Excel.Application.ScreenUpdating = s
  End If
End Sub

Private Property Get iLogger_level() As logLevels
  iLogger_level = oLevel
End Property
Private Property Let iLogger_level(l As logLevels)
  oLevel = l
End Property

