Attribute VB_Name = "mSorter_collection"
Option Explicit
' ##############################################################################\
' Author(s):   Philipp Gorkiewicz                                               |
' Description: collection sorter, using an iComparer object                     |
'                                                                               |
' ##############################################################################|
'     Version | Description                                                     |
'       1.0.0 | Initial Version                                                 |
' ##############################################################################/
Dim c As iComparer

Public Enum sortOrder
  ascending = 0
  descending = 1
End Enum

Public Sub Sort(collObject As Collection, _
                comparer As iComparer, _
                Optional order As sortOrder = ascending)
  Set c = comparer
  quickSort collObject, _
            1, _
            collObject.count, _
            order
'  bubbleSort collObject, _
'             1, _
'             collObject.count, _
'             order
End Sub

Private Sub quickSort(coll As Collection, _
                      lowerBound As Long, _
                      upperBound As Long, _
                      Optional order As sortOrder)
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
      While c.compare(coll.Item(lower), center) = less And lower < upperBound
        lower = lower + 1
      Wend
      While c.compare(center, coll.Item(upper)) = less And upper > lowerBound
        upper = upper - 1
      Wend
    Else
      While c.compare(coll.Item(lower), center) = greater And lower < upperBound
        lower = lower + 1
      Wend
      While (c.compare(center, coll.Item(upper)) = greater And upper > lowerBound)
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
  If (lowerBound < upper) Then Call quickSort(coll, lowerBound, upper, order)
  If (lower < upperBound) Then Call quickSort(coll, lower, upperBound, order)
End Sub

Private Sub bubbleSort(coll As Collection, _
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
        If IsObject(coll.Item(i)) Then
          Set vTmp = coll.Item(i)
        Else
          vTmp = coll.Item(i)
        End If
        coll.Add Item:=coll.Item(j), After:=i
        Call coll.Remove(i)
        coll.Add Item:=vTmp, After:=j
        Call coll.Remove(j)
      End If
    Next j
  Next i
End Sub
