Attribute VB_Name = "mBubbleSorter"
Option Explicit
' ##############################################################################\
' Author(s):   Philipp Gorkiewicz                                               |
' License:     MIT (https://opensource.org/license/mit/)                        |
' Versioning:  https://semver.org/                                              |
' Description: bubble sorter                                                    |
'              comparison on iComparer object                                   |
' ##############################################################################|
' Version | Description                                                         |
'   1.0.0 | Initial Version                                                     |
' ##############################################################################/
Dim compare As iComparer
Dim order   As sortOrder
Public Enum sortOrder
  ascending = 0
  descending = 1
End Enum
Public Sub Sort(inputObject As Variant, _
                comparer As iComparer, _
                Optional order As sortOrder = ascending)
  Set compare = c
  order = order
  If TypeOf inputObject Is Collection Then
    Dim c As Collection
    Set c = inputObject
    Call sort_collection(c)
  ElseIf TypeOf inputObject Is iEnumerator Then
    Dim e As iEnumerator
    Set e = inputObject
    Call sort_enumerator(e)
  Else
    Err.Raise 17, _
              "mBubbleSorter.Sort", _
              """" & TypeName(inputObject) & """ not supported to sort"
  End If
End Sub
Private Sub sort_collection(coll As Collection)
  Dim result  As compareResult
  Select Case sortOrder
  Case descending
    result = greater
  Case Else
    result = less
  End Select

  Dim i   As LongPtr
  Dim j   As LongPtr
  Dim tmp As Variant
  For i = 1 To (coll.count - 1)
    For j = (i + 1) To coll.count
      If compare(coll(j), coll(i)) = result Then
        If IsObject(coll.item(i)) Then
          Set tmp = coll.item(i)
        Else
          tmp = coll.item(i)
        End If
        coll.Add item:=coll.item(j), After:=i
        Call coll.Remove(i)
        coll.Add item:=tmp, After:=j
        Call coll.Remove(j)
      End If
    Next j
  Next i
End Sub

Private Sub sort_enumerator(e As iEnumerator)
  Dim result As compareResult
  Select Case sortOrder
  Case descending
    r = greater
  Case Else
    r = less
  End Select

  Dim i As LongPtr
  Dim j As LongPtr
  For i = 1 To (e.count - 1)
    For j = (i + 1) To e.count
      If compare(e(j), e(i)) = result Then Call e.Swap(e(j), e(i))
    Next j
  Next i
End Sub
