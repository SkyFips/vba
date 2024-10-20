Attribute VB_Name = "mWebHelper"
Option Explicit
' ##############################################################################\
' Author(s):   Philipp Gorkiewicz                                               |
' License:     MIT (https://opensource.org/license/mit/)                        |
' Versioning:  https://semver.org                                               |
' Description: useful web helper functions                                      |
' ------------------------------------------------------------------------------|
' Version | Description                                                         |
'   1.0.0 | Initial Version                                                     |
' ##############################################################################/

Function rfc7636(length As Byte) As String
  ' https://datatracker.ietf.org/doc/html/rfc7636#section-4.1
  If length < 43 Or length > 128 Then
    Err.Raise 5, _
              "mWebHelper.rfc7636", _
              "as per IETF standard, length must be >=43 and <=128"
  End If
  Const characters As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~"
  Dim r As String
  Dim i As Byte
  Randomize
  For i = 1 To length
    r = r & Mid(characters, Int((Len(characters) * Rnd) + 1), 1)
  Next i
  rfc7636 = r
End Function
