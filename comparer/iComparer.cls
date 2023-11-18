VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "interface description for iComparer"
' ##############################################################################\
' (c)2018 Scientific Games International GmbH
'                                                                               |
' Author(s):      Philipp Gorkiewicz                                            |
' Version:        1.0.1                                                         |
' Creation date:  2018-09-23                                                    |
' Last change:    2020-01-17                                                    |
'                                                                               |
' Description:    Interface Description for string Comparisons                  |
'                                                                               |
' ##############################################################################|
'     Version | Description                                                     |
'       1.0.0 | Initial Version                                                 |
'       1.0.1 | make it flexibel to compare Objects as well                     |
' ##############################################################################/
Option Explicit

Public Enum compareResult
  equal = 0
  greater = 1
  less = -1
End Enum

Public Function compare(a As Variant, _
                        b As Variant) As compareResult
Attribute compare.VB_Description = "compares two objects and return enum (0=equal;1=greater;-1=less)"
Attribute compare.VB_UserMemId = 0
  ' A compared to B
  ' declaration only
  ' return values
  '  0 = A equal B
  '  1 = A greater than B
  ' -1 = A less than B
End Function
