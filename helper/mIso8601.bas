Attribute VB_Name = "mIso8601"
Option Explicit
' ##############################################################################\
' Author(s):   Philipp Gorkiewicz                                               |
' License:     MIT (https://opensource.org/license/mit/)                        |
' Description: Iso8601 parser/generater                                         |
'              https://en.wikipedia.org/wiki/ISO_8601                           |
' ##############################################################################|
'   Version | Description                                                       |
'     1.0.0 | Initial Version                                                   |
' ##############################################################################/
Dim oRegEx  As Object
Dim oLogger As iLogger

' https://learn.microsoft.com/en-us/windows/win32/api/timezoneapi/nf-timezoneapi-gettimezoneinformation
Private Declare PtrSafe Function GetTimeZoneInformation Lib "kernel32" _
  (timeZoneInformation As TIME_ZONE_INFORMATION) As LongPtr
' https://learn.microsoft.com/en-us/windows/win32/api/timezoneapi/nf-timezoneapi-systemtimetotzspecificlocaltime
Private Declare PtrSafe Function SystemTimeToTzSpecificLocalTime Lib "kernel32" _
  (timeZoneInformation As TIME_ZONE_INFORMATION, universalTime As SYSTEMTIME, localTime As SYSTEMTIME) As LongPtr
' https://learn.microsoft.com/en-us/windows/win32/api/timezoneapi/nf-timezoneapi-tzspecificlocaltimetosystemtime
Private Declare PtrSafe Function TzSpecificLocalTimeToSystemTime Lib "kernel32" _
  (timeZoneInformation As TIME_ZONE_INFORMATION, localTime As SYSTEMTIME, universalTime As SYSTEMTIME) As LongPtr

'https://learn.microsoft.com/en-us/windows/win32/api/minwinbase/ns-minwinbase-systemtime
Private Type SYSTEMTIME
  wYear         As Integer
  wMonth        As Integer
  wDayOfWeek    As Integer
  wDay          As Integer
  wHour         As Integer
  wMinute       As Integer
  wSecond       As Integer
  wMilliseconds As Integer
End Type

' https://learn.microsoft.com/en-us/windows/win32/api/timezoneapi/ns-timezoneapi-time_zone_information
Private Type TIME_ZONE_INFORMATION
  bias                  As Long
  standardName(0 To 31) As Integer
  standardDate          As SYSTEMTIME
  standardBias          As Long
  daylightName(0 To 31) As Integer
  daylightDate          As SYSTEMTIME
  daylightBias          As Long
End Type

Public Property Get logger() As iLogger
  If oLogger Is Nothing Then Set oLogger = config.logger
  Set logger = oLogger
End Property

Public Property Set logger(l As iLogger)
  Set oLogger = l
End Property

Private Property Get regEx() As Object
  If oRegEx Is Nothing Then Set oRegEx = CreateObject("VBScript.RegExp")
  Set regEx = oRegEx
End Property

Private Function convertToLocalDate(utcDate As Date) As Date
  Dim timeZoneInfo  As TIME_ZONE_INFORMATION
  Dim localDate     As SYSTEMTIME

  GetTimeZoneInformation timeZoneInfo
  SystemTimeToTzSpecificLocalTime timeZoneInformation:=timeZoneInfo, _
                                  universalTime:=asSYSTEMTIME(utcDate), _
                                  localTime:=localDate
  convertToLocalDate = asDate(localDate)
End Function

Private Function convertToUtcDate(localDate As Date) As Date
  Dim timeZoneInfo  As TIME_ZONE_INFORMATION
  Dim utcDate       As SYSTEMTIME

  GetTimeZoneInformation timeZoneInfo
  SystemTimeToTzSpecificLocalTime timeZoneInformation:=timeZoneInfo, _
                                  universalTime:=asSYSTEMTIME(localDate), _
                                  localTime:=utcDate
  convertToUtcDate = asDate(utcDate)
End Function

Public Function parse(isoString As String, Optional asLocalDate As Boolean = False) As Date
  Dim matches As Object
  Dim match   As Object
  regEx.Pattern = "^(\d{4})-(\d{2})-(\d{2})[tT](\d{2}):(\d{2}):(\d{2})(?:\.\d{3})?([zZ]|(?:[+-])(\d{2}):?(\d{2}))$"
  regEx.Global = True
  
  If Not regEx.test(isoString) Then
    logger.log FAILURE, "the provided string (" & isoString & ") does not match expectation (e.g. 2020-03-04T05:31:11+00:00)"
    Err.Raise 5, _
              "mConverter_date.parseIso8601", _
              "the provided string (" & isoString & ") does not match expectation (e.g. 2020-03-04T05:31:11+00:00)"
  End If
  
  Set matches = regEx.execute(isoString)
  With matches(0)
    parse = VBA.DateSerial(VBA.CInt(.submatches(0)), VBA.CInt(.submatches(1)), VBA.CInt(.submatches(2))) + _
            VBA.TimeSerial(.submatches(3), .submatches(4), .submatches(5))
    'logger.log INFO, CStr(parseIso8601)
    Select Case Left(.submatches(6), 1)
    Case "+"
      parse = parse - TimeSerial(.submatches(7), .submatches(8), 0)
    Case "-"
      parse = parse + TimeSerial(.submatches(7), .submatches(8), 0)
    End Select
  End With
  If asLocalDate Then parse = convertToLocalDate(parse)
End Function

Public Function generate(localDate As Date) As String
  Dim i As TIME_ZONE_INFORMATION
  Dim h As Integer
  Dim m As Integer
  
  GetTimeZoneInformation i
  h = i.bias / 60
  m = i.bias Mod 60
  If h < 0 Or m < 0 Then
    generate = VBA.Format$(localDate, "yyyy-mm-ddTHH:mm:ss.000") & "+" & VBA.Format$(Abs(h), "00") & ":" & VBA.Format$(Abs(m), "00")
  ElseIf h > 0 Or m > 0 Then
    generate = VBA.Format$(localDate, "yyyy-mm-ddTHH:mm:ss.000") & "-" & VBA.Format$(h, "00") & ":" & VBA.Format$(m, "00")
  Else
    generate = VBA.Format$(localDate, "yyyy-mm-ddTHH:mm:ss.000Z")
  End If
End Function

Private Function asSYSTEMTIME(d As Date) As SYSTEMTIME
  asSYSTEMTIME.wYear = VBA.year(d)
  asSYSTEMTIME.wMonth = VBA.month(d)
  asSYSTEMTIME.wDay = VBA.day(d)
  asSYSTEMTIME.wHour = VBA.Hour(d)
  asSYSTEMTIME.wMinute = VBA.Minute(d)
  asSYSTEMTIME.wSecond = VBA.Second(d)
  asSYSTEMTIME.wMilliseconds = 0
End Function

Private Function asDate(s As SYSTEMTIME) As Date
  asDate = DateSerial(s.wYear, s.wMonth, s.wDay) + TimeSerial(s.wHour, s.wMinute, s.wSecond)
End Function
