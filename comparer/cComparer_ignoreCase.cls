VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cComparer_ignoreCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ##############################################################################\
' Author(s):   Philipp Gorkiewicz                                               |
' License:     MIT (https://opensource.org/license/mit/)                        |
' Versioning:  https://semver.org                                               |
' Description: implementation for IComparer for string comparison               |
'              non case sensitive                                               |
' ------------------------------------------------------------------------------|
' Version | Description                                                         |
'   1.0.0 | Initial Version                                                     |
' ##############################################################################/
Implements iComparer

Private Function iComparer_compare(input_A As Variant, _
                                   input_B As Variant) As compareResult
  Dim a As String
  Dim b As String

  a = input_A
  b = input_B
  iComparer_compare = StrComp(a, b, vbTextCompare)
End Function