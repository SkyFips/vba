Attribute VB_Name = "mSorter_enumerator"
Option Explicit
' ##############################################################################\
' Author(s):   Philipp Gorkiewicz                                               |
' Description: collection sorter, using an iComparer object                     |
'                                                                               |
' ##############################################################################|
'     Version | Description                                                     |
'       1.0.0 | Initial Version                                                 |
' ##############################################################################/
Dim compare As iComparer

Public Sub Sort(e As iEnumerator, _
                comparer As iComparer, _
                Optional order As sortOrder = ascending)
  Set compare = comparer
  quickSort e, _
            1, _
            e.count, _
            order
'  bubbleSort collObject, _
'             1, _
'             collObject.count, _
'             order
End Sub

Private Sub quickSort(e As iEnumerator, _
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
  If (lowerBound < upper) Then Call quickSort(e, lowerBound, upper, order)
  If (lower < upperBound) Then Call quickSort(e, lower, upperBound, order)
End Sub

Private Sub bubbleSort(e As iEnumerator, _
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


