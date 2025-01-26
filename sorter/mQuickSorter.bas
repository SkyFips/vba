Attribute VB_Name = "mQuickSorter"
Option Explicit
' ##############################################################################\
' Author(s):   Philipp Gorkiewicz                                               |
' License:     MIT (https://opensource.org/license/mit/)                        |
' Versioning:  https://semver.org/                                              |
' Source:      https://github.com/SkyFips/vba/tree/main/sorter                  |
' Description: quicksort for collection/iEnumerator                             |
'              comparison done on iComparer object                              |
' ------------------------------------------------------------------------------|
' Version | Description                                                         |
'   1.0.0 | Initial Version                                                     |
'   1.0.1 | correct call var                                                    |
'   1.0.2 | only call sort, if input count greater 0                            |
' ##############################################################################/
Dim compare As iComparer
Dim order   As sortOrder
Public Enum sortOrder
  ascending = 0
  descending = 1
End Enum
Public Sub Sort(inputObject As Object, _
                comparer As iComparer, _
                Optional order As sortOrder = ascending)
  Set compare = comparer
  order = order
  If TypeOf inputObject Is Collection Then
    Dim c As Collection
    Set c = inputObject
    If c.count > 0 Then Call sort_collection(c, 1, c.count)
  ElseIf TypeOf inputObject Is iEnumerator Then
    Dim e As iEnumerator
    Set e = inputObject
    If e.count > 0 Then Call sort_enumerator(e, 1, e.count)
  Else
    Err.Raise 17, _
              "mQuickSorter.Sort", _
              """" & TypeName(inputObject) & """ not supported to sort"

  End If
End Sub

Private Sub sort_collection(coll As Collection, _
                            lowerBound As Long, _
                            upperBound As Long)
  Dim center  As Variant
  Dim lower   As Long
  Dim upper   As Long

  lower = lowerBound
  upper = upperBound

  If IsObject(coll.Item((lower + upper) / 2)) Then
    Set center = coll.Item((lower + upper) / 2)
  Else
    center = coll.Item((lower + upper) / 2)
  End If
  While (lower <= upper)
    If order = ascending Then
      While compare(coll.Item(lower), center) = less And lower < upperBound
        lower = lower + 1
      Wend
      While compare(center, coll.Item(upper)) = less And upper > lowerBound
        upper = upper - 1
      Wend
    Else
      While compare(coll.Item(lower), center) = greater And lower < upperBound
        lower = lower + 1
      Wend
      While (compare(center, coll.Item(upper)) = greater And upper > lowerBound)
        upper = upper - 1
      Wend
    End If
    If (lower <= upper) Then
      Dim varTmp  As Variant

      If IsObject(coll.Item(lower)) Then
        Set varTmp = coll.Item(lower)
      Else
        varTmp = coll.Item(lower)
      End If
      coll.Add Item:=coll.Item(upper), After:=lower
      Call coll.Remove(lower)
      coll.Add Item:=varTmp, After:=upper
      Call coll.Remove(upper)
      lower = lower + 1
      upper = upper - 1
    End If
  Wend
  If (lowerBound < upper) Then Call sort_collection(coll, lowerBound, upper)
  If (lower < upperBound) Then Call sort_collection(coll, lower, upperBound)
End Sub

Private Sub sort_enumerator(e As iEnumerator, _
                            lowerBound As LongPtr, _
                            upperBound As LongPtr)
  Dim center  As Variant
  Dim lower   As LongPtr
  Dim upper   As LongPtr

  lower = lowerBound
  upper = upperBound

  If IsObject(e(CLng((lower + upper) / 2))) Then
    Set center = e(CLng((lower + upper) / 2))
  Else
    center = e(CLng((lower + upper) / 2))
  End If
  While (lower <= upper)
    If order = ascending Then
      While compare(e(lower), center) = less And lower < upperBound
        lower = lower + 1
      Wend
      While compare(center, e.Item(upper)) = less And upper > lowerBound
        upper = upper - 1
      Wend
    Else
      While compare(e(lower), center) = greater And lower < upperBound
        lower = lower + 1
      Wend
      While (compare(center, e(upper)) = greater And upper > lowerBound)
        upper = upper - 1
      Wend
    End If
    If (lower <= upper) Then
      Call e.Swap(lower, upper)
      lower = lower + 1
      upper = upper - 1
    End If
  Wend
  If (lowerBound < upper) Then Call sort_enumerator(e, lowerBound, upper)
  If (lower < upperBound) Then Call sort_enumerator(e, lower, upperBound)
End Sub
