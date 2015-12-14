Option Strict Off
Option Explicit On

Friend Class Segments

    Implements System.Collections.IEnumerable

    Private mcolSegments As New Collection
    Private intSegNum As Short


    Public Function Add() As Segment

        Dim segNew As New Segment

        intSegNum = intSegNum + 1

        If intSegNum > MaxSegments Then
            Err.Raise(Number:=vbObjectError + 32101, Description:="Maximum number of segments is " & MaxSegments)
            Add = Nothing
        Else
            segNew.SegNum = intSegNum
            mcolSegments.Add(segNew, CStr(segNew.SegNum))
            Add = segNew
        End If

    End Function


    Public Function Count() As Integer

        Count = mcolSegments.Count()

    End Function


    Public Sub Delete(ByVal Index As Object)

        mcolSegments.Remove(Index)
        If intSegNum > 0 Then
            intSegNum = intSegNum - 1
        Else
            intSegNum = 0
        End If

    End Sub


    Public Function Item(ByVal Index As Object) As Segment

        Item = mcolSegments.Item(Index)

    End Function


    'UPGRADE_NOTE: NewEnum property was commented out. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B3FC1610-34F3-43F5-86B7-16C984F0E88E"'
    'Public Function NewEnum() As stdole.IUnknown
    '
    'NewEnum = mcolSegments.GetEnumerator
    '
    'End Function

    Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator

        GetEnumerator = mcolSegments.GetEnumerator

    End Function


    Public Sub New()

        MyBase.New()
        intSegNum = 0

    End Sub

End Class