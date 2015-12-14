Option Strict Off
Option Explicit On

Friend Class StaticSolution

    Implements System.Collections.IEnumerable

    Private mcolStaticSolution As New Collection


    Public Function Add() As StaticSolutionPoint

        Dim sspNew As New StaticSolutionPoint
        Dim intSSPNum As Short

        intSSPNum = mcolStaticSolution.Count()

        If intSSPNum > MaxSSPoints Then
            Err.Raise(Number:=vbObjectError + 32101, Description:="Maximum number of Static Solution Points is " & MaxSSPoints)
            Add = Nothing
        Else
            sspNew.PointNum = intSSPNum
            mcolStaticSolution.Add(sspNew, CStr(sspNew.PointNum))
            Add = sspNew
        End If

    End Function


    Public Function Count() As Integer

        Count = mcolStaticSolution.Count()

    End Function


    Public Sub Delete(ByVal Index As Object)

        mcolStaticSolution.Remove(Index)

    End Sub


    Public Function Item(ByVal Index As Object) As StaticSolutionPoint

        Item = mcolStaticSolution.Item(Index)

    End Function


    'UPGRADE_NOTE: NewEnum property was commented out. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B3FC1610-34F3-43F5-86B7-16C984F0E88E"'
    'Public Function NewEnum() As stdole.IUnknown
    '
    'NewEnum = mcolStaticSolution.GetEnumerator
    '
    'End Function

    Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator

        GetEnumerator = mcolStaticSolution.GetEnumerator

    End Function

End Class