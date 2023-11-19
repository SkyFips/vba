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
' Author(s):   Philipp Gorkiewicz                                               |
' License:     MIT (https://opensource.org/license/mit/)                        |
' Versioning:  https://semver.org                                               |
' Description: interface description of comparer                                |
' ------------------------------------------------------------------------------|
' Version | Description                                                         |
'   1.0.0 | Initial Version                                                     |
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
  ' return values
  '  0 = A equal B
  '  1 = A greater than B
  ' -1 = A less than B
End Function
