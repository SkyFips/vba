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
' (c)2020 Scientific Games International GmbH                                   |
'                                                                               |
' Author(s):      Philipp Gorkiewicz                                            |
' Version:        1.1.0                                                         |
' Creation date:  2020-01-17                                                    |
' Last change:    2020-09-22                                                    |
'                                                                               |
' Description:    implementation for IComparer for string comparison            |
'                 non case sensitive                                            |
' ##############################################################################|
'     Version | Description                                                     |
'       1.0.0 | Initial Version                                                 |
'       1.0.1 | small variables adaptions                                       |
'       1.1.0 | remove globLogger                                               |
' ##############################################################################/
Implements iComparer

Const MODULE  As String = "cComparer_caseSensitive"

Private Function iComparer_compare(input_A As Variant, _
                                   input_B As Variant) As compareResult
  Dim stringA As String
  Dim stringB As String
  
  stringA = input_A
  stringB = input_B
  iComparer_compare = StrComp(stringA, stringB, vbTextCompare)
End Function
