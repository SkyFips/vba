VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cComparer_colorRGB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ##############################################################################\
' (c)2020 Scientific Games International GmbH                                   |
'                                                                               |
' Author(s):      Philipp Gorkiewicz                                            |
' Version:        1.0.0                                                         |
' Creation date:  2020-04-08                                                    |
' Last change:    2020-04-08                                                    |
'                                                                               |
' Description:    implementation for IComparer for RGB color comparison         |
'                                                                               |
' ##############################################################################|
'     Version | Description                                                     |
'       1.0.0 | Initial Version                                                 |
' ##############################################################################/
Implements iComparer

Const MODULE  As String = "cComparer_colorRGB"

Private Function iComparer_compare(input_A As Variant, _
                                   input_B As Variant) As compareResult
  Dim colorA  As LongPtr
  Dim colorB  As LongPtr
  
  colorA = input_A
  colorB = input_B
  
  iComparer_compare = compare(colorA, colorB)
End Function
Private Function compare(colorA As LongPtr, colorB As LongPtr) As compareResult
  If colorA > colorB Then
    compare = greater
  ElseIf colorA < colorB Then
    compare = less
  Else
    compare = equal
  End If
End Function
