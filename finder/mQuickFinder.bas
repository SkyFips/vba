Attribute VB_Name = "mQuickFinder"
Option Explicit
' ##############################################################################\
' Author(s):   Philipp Gorkiewicz                                               |
' License:     MIT (https://opensource.org/license/mit/)                        |
' Versioning:  https://semver.org/                                              |
' Source:      https://github.com/SkyFips/vba/tree/main/finder                  |
' Description: quick finder for an object implements iEnumerator                |
'              comparison done on iComparer object                              |
' ------------------------------------------------------------------------------|
' Version | Description                                                         |
'   1.0.0 | Initial Version                                                     |
'   1.0.1 | correct call var                                                    |
'   1.0.2 | declare inputObject as object instead of Variant                    |
' ##############################################################################/
Dim compare As iComparer
Dim toFind As Variant
Public Function Find(inputObject As Object, _
                     comparer As iComparer, _
                     objToFind As Variant) As Variant
  Set compare = comparer
  Set toFind = objToFind
  Call mQuickSorter.Sort(inputObject, comparer, ascending)
  If TypeOf inputObject Is Collection Then
    Dim c As Collection
    Set c = inputObject
    Call find_collection(c, 1, c.count)
  ElseIf TypeOf inputObject Is iEnumerator Then
    Dim e As iEnumerator
    Set e = inputObject
    Call find_enumerator(e, 1, e.count)
  Else
    Err.Raise 17, _
              "mQuickFinder.Find", _
              """" & TypeName(inputObject) & """ not supported for search"

  End If
End Function

Private Function find_collection(coll As Collection, _
                                 lowerBound As Long, _
                                 upperBound As Long) As Variant
  Dim center As Long
  Dim lower  As Long
  Dim upper  As Long

  lower = lowerBound
  upper = upperBound
  center = (lowerBound + upperBound) / 2
  
  If compare(toFind, coll(center)) = greater Then
    Set find_collection = find_collection(coll, center, upperBound)
  ElseIf compare(toFind, coll(center)) = less Then
    Set find_collection = find_collection(coll, lowerBound, center)
  Else
    
  End If

End Function

Private Sub find_enumerator(e As iEnumerator, _
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
    While compare(e(lower), center) = less And lower < upperBound
      lower = lower + 1
    Wend
    While compare(center, e.Item(upper)) = less And upper > lowerBound
      upper = upper - 1
    Wend
    If (lower <= upper) Then
      Call e.Swap(lower, upper)
      lower = lower + 1
      upper = upper - 1
    End If
  Wend
  If (lowerBound < upper) Then Call find_enumerator(e, lowerBound, upper)
  If (lower < upperBound) Then Call find_enumerator(e, lower, upperBound)
End Sub


