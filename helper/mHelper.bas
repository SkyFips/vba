Attribute VB_Name = "mHelper"
Option Explicit
' ##############################################################################\
' Author(s):   Philipp Gorkiewicz                                               |
' Description: singleton modules                                                |
'                                                                               |
' ##############################################################################|
'   Version | Description                                                       |
'     1.0.0 | Initial Version                                                   |
'     2.0.0 | remove concatenate function                                       |
'     2.1.0 | add concatenate function                                          |
' ##############################################################################/
Dim oLogger As iLogger

Private Property Get logger() As iLogger
  If oLogger Is Nothing Then Set oLogger = config.logger
  Set logger = oLogger
End Property
Public Property Set logger(l As iLogger)
  Set oLogger = l
End Property
Public Function sheetExist(strSheetName As String, Optional wkb As Workbook) As Boolean
  On Error Resume Next
  If wkb Is Nothing Then Set wkb = ThisWorkbook
  sheetExist = wkb.Sheets(strSheetName).Index > 0
End Function
Public Function strftime(d As Date, f As String) As String
' uses the same format(s) as ruby strftime
  logger.log DEBUGGER, "d => " & CStr(d)
  logger.log DEBUGGER, "f => " & f
  Dim o As String
  o = f
  o = Replace(o, "%m", Format(Month(d), "00"))
  o = Replace(o, "%Y", Format(Year(d), "0000"))
  o = Replace(o, "%d", Format(day(d), "00"))
  o = Replace(o, "%b", Format(Month(d), "mmm"))
  o = Replace(o, "%B", Format(Month(d), "mmmm"))
  o = Replace(o, "%H", Format(Hour(d), "00"))
  o = Replace(o, "%M", Format(Minute(d), "00"))
  logger.log DEBUGGER, "o => " & o
  strftime = o
End Function
Public Function nextWeekday(d As Date, _
                            wd As VBA.VbDayOfWeek, _
                            Optional inclDay As Boolean = False) As Date
  logger.log DEBUGGER, "d => " & d
  logger.log DEBUGGER, "wd=> " & wd
  logger.log DEBUGGER, "incleDay => " & inclDay
  If inclDay Then
    If Weekday(d) > wd Then
      nextWeekday = d + 7 - Weekday(d) + wd
    Else
      nextWeekday = d - Weekday(d) + wd
    End If
  Else
    If Weekday(d) >= wd Then
      nextWeekday = d + 7 - Weekday(d) + wd
    Else
      nextWeekday = d - Weekday(d) + wd
    End If
  End If
End Function
Public Function hasKey(c As Collection, k As String) As Boolean
  On Error Resume Next
  Dim d As LongPtr
  d = VarType(c.Item(k))
  hasKey = CBool(Err.Number = 0)
  Err.Clear
End Function
Public Function hasValidation(c As Range) As Boolean
  On Error Resume Next
  Dim t: t = Null
  t = c.Validation.Type
  On Error GoTo 0
  hasValidation = Not IsNull(t)
End Function
Public Function colToChr(c As Long) As String
  If c < 27 Then
  'first letter "A" has a column number of 65, hence the use of 64 (65 minus 1)
    colToChr = Chr(c + 64)

  'for two letter column
  ElseIf c < 703 Then
  'first letter "A" has a column number of 65, and one round of alphabets "A" to "Z" equals 26
    colToChr = Chr(Int((c - 1) / 26) + 64) & _
               Chr(((c - 1) Mod 26) + 65)

  'for three letter column
  ElseIf c < 18279 Then
  '676 equates to 26*26; 702 equates to 26*26+26;
    colToChr = Chr((Int((c - 703) / 676) Mod 26) + 65) & _
               Chr((Int((c - 27) / 26) Mod 26) + 65) & Chr(Int((c - 1) Mod 26) + 65)

  'for four letter column
  ElseIf c < 475255 Then
  '17576 equates to 26*26*26; 18278 equates to 26*26*26+702;
    colToChr = Chr((Int((c - 18279) / 17576) Mod 26) + 65) & _
               Chr((Int((c - 703) / 676) Mod 26) + 65) & _
               Chr((Int((c - 27) / 26) Mod 26) + 65) & Chr(Int((c - 1) Mod 26) + 65)
  End If
End Function
Public Function performanceTuning(b As Boolean)
  If b Then
    logger.log INFO, "ScreenUpdating/DisplayAlerts/Calculation => OFF"
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
  Else
    logger.log INFO, "ScreenUpdating/DisplayAlerts/Calculation => ON"
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
  End If
End Function
Public Function splitString(i As String, _
                            DELEMITER As String) As Collection
  Set splitString = New Collection
  Dim v As Variant

  For Each v In Split(i, DELEMITER)
    splitString.add v
  Next v
End Function
Public Function concatenate(DELEMITER As String, _
                            collInput As Collection) As String

  Dim c As Integer
  If (IsNull(collInput)) Then Exit Function
  For c = 1 To collInput.count
    concatenate = concatenate & collInput(c)
    If c < collInput.count Then
      concatenate = concatenate & DELEMITER
    End If
  Next c
End Function
