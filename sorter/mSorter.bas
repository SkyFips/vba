Attribute VB_Name = "mSorter"
Option Explicit
' ##############################################################################\
' Author(s):   Philipp Gorkiewicz                                               |
' Description: collection sorter, using an iComparer object                     |
'                                                                               |
' ##############################################################################|
'     Version | Description                                                     |
'       1.0.0 | Initial Version                                                 |
' ##############################################################################/
Dim comparer As iComparer

Public Enum sortOrder
  ascending = 0
  descending = 1
End Enum

Public Sub Sort(inputObject As Object, _
                comparer As iComparer, _
                Optional order As sortOrder = ascending)
  If TypeOf inputObject Is Collection Then

  ElseIf TypeOf inputObject Is iEnumerator Then

  Else
    Err.Raise 17, _
              "mSorter.Sort", _
              TypeName(inputObject) & " is not supported"
  End If

  Set ccomparer = c
  quickSort collObject, _
            1, _
            collObject.count, _
            order
'  bubbleSort collObject, _
'             1, _
'             collObject.count, _
'             order
End Sub

Private Sub qs_collection(coll As Collection, _
                          lowerBound As Long, _
                          upperBound As Long, _
                          Optional order As sortOrder)
  Dim center  As Variant
  Dim lower   As Long
  Dim upper   As Long

  lower = lowerBound
  upper = upperBound

  If IsObject(coll.item((lower + upper) / 2)) Then
    Set center = coll.item((lower + upper) / 2)
  Else
    center = coll.item((lower + upper) / 2)
  End If
  While (lower <= upper)
    If order = ascending Then
      While c.compare(coll.item(lower), center) = less And lower < upperBound
        lower = lower + 1
      Wend
      While c.compare(center, coll.item(upper)) = less And upper > lowerBound
        upper = upper - 1
      Wend
    Else
      While c.compare(coll.item(lower), center) = greater And lower < upperBound
        lower = lower + 1
      Wend
      While (c.compare(center, coll.item(upper)) = greater And upper > lowerBound)
        upper = upper - 1
      Wend
    End If
    If (lower <= upper) Then
      Dim varTmp  As Variant

      If IsObject(coll.item(lower)) Then
        Set varTmp = coll.item(lower)
      Else
        varTmp = coll.item(lower)
      End If
      coll.Add item:=coll.item(upper), After:=lower
      Call coll.Remove(lower)
      coll.Add item:=varTmp, After:=upper
      Call coll.Remove(upper)
      lower = lower + 1
      upper = upper - 1
    End If
  Wend
  If (lowerBound < upper) Then Call qs_collection(coll, lowerBound, upper, order)
  If (lower < upperBound) Then Call qs_collection(coll, lower, upperBound, order)
End Sub

Private Sub bs_collection(coll As Collection, _
                          lowerBound As Long, _
                          upperBound As Long, _
                          order As sortOrder)
  Dim i As Long
  Dim j As Long
  Dim vTmp  As Variant
  Dim comp  As compareResult
  Select Case order
  Case descending
    comp = greater
  Case Else
    comp = less
  End Select

  For i = lowerBound To (upperBound - 1)
    For j = (i + 1) To upperBound
      If c.compare(coll(j), coll(i)) = comp Then
        If IsObject(coll.item(i)) Then
          Set vTmp = coll.item(i)
        Else
          vTmp = coll.item(i)
        End If
        coll.Add item:=coll.item(j), After:=i
        Call coll.Remove(i)
        coll.Add item:=vTmp, After:=j
        Call coll.Remove(j)
      End If
    Next j
  Next i
End Sub

Private Sub qs_enumerator(e As iEnumerator, _
                          lowerBound As Long, _
                          upperBound As Long, _
                          Optional order As sortOrder)
  Dim center  As Variant
  Dim lower   As Long
  Dim upper   As Long

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
      While compare(center, e.item(upper)) = less And upper > lowerBound
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
  If (lowerBound < upper) Then Call qs_enumerator(e, lowerBound, upper, order)
  If (lower < upperBound) Then Call qs_enumerator(e, lower, upperBound, order)
End Sub

Private Sub bs_enumerator(e As iEnumerator, _
                          lowerBound As Long, _
                          upperBound As Long, _
                          order As sortOrder)
  Dim i As Long
  Dim j As Long
  Dim vTmp  As Variant
  Dim comp  As compareResult
  Select Case order
  Case descending
    comp = greater
  Case Else
    comp = less
  End Select

  For i = lowerBound To (upperBound - 1)
    For j = (i + 1) To upperBound
      If compare(e(j), e(i)) = comp Then Call e.Swap(e(j), e(i))
    Next j
  Next i
End Sub
