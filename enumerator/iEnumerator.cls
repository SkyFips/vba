VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iEnumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ##############################################################################\
' Author(s):   Philipp Gorkiewicz                                               |
' License:     MIT (https://opensource.org/license/mit/)                        |
' Versioning:  https://semver.org                                               |
' Source:      https://github.com/SkyFips/vba/tree/main/enumerator              |
' Description: enumerator interface                                             |
' ------------------------------------------------------------------------------|
'   version | description                                                       |
'     1.0.0 | Initial Version                                                   |
' ##############################################################################/
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
  ' used to make the class iteratable with a for each loop
End Property
Public Property Get count() As Long
End Property
Public Property Get Item(id As Variant) As Variant
Attribute Item.VB_UserMemId = 0
End Property
Public Sub Swap(a As LongPtr, b As LongPtr)
End Sub
