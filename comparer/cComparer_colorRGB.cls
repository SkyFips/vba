VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cComparer_colorRGB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "comparison of RGB colors (A compared to B)"
Option Explicit
' ##############################################################################\
' Author(s):   Philipp Gorkiewicz                                               |
' License:     MIT (https://opensource.org/license/mit/)                        |
' Versioning:  https://semver.org                                               |
' Description: iComparer implementationfor RGB color comparison                 |
' ------------------------------------------------------------------------------|
' Version | Description                                                         |
'   1.0.0 | Initial Version                                                     |
' ##############################################################################/
Implements iComparer

Private Function iComparer_compare(input_A As Variant, _
                                   input_B As Variant) As compareResult
  Dim a As LongPtr
  Dim b As LongPtr

  a = input_A
  b = input_B

  If a > b Then
    iComparer_compare = greater
  ElseIf a < b Then
    iComparer_compare = less
  Else
    iComparer_compare = equal
  End If
End Function
