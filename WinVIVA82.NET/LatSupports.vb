Option Strict Off
Option Explicit On

Friend Class LatSupports

    Implements System.Collections.IEnumerable

    Private mcolLatSupports As New Collection


    Public Sub Add(ByRef Dist As Single, ByRef Stiffness As Single, ByRef Damping As Single)

        Dim supNew As New LatSupport

        If mcolLatSupports.Count() < MaxNumLatSupports Then
            With supNew
                .Dist = Dist
                .Stiffness = Stiffness
                .Damping = Damping
            End With
            mcolLatSupports.Add(supNew)
        End If

    End Sub


    Public Sub Clear()

        Dim NumLatSupports, i As Short

        NumLatSupports = mcolLatSupports.Count()
        For i = NumLatSupports To 1 Step -1
            mcolLatSupports.Remove(i)
        Next i

    End Sub


    Public Function Count() As Integer

        Count = mcolLatSupports.Count()

    End Function


    Public Sub Delete(ByVal index As Object)

        mcolLatSupports.Remove(index)

    End Sub


    Public Function Item(ByVal index As Object) As LatSupport

        Item = mcolLatSupports.Item(index)

    End Function


    'UPGRADE_NOTE: NewEnum property was commented out. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B3FC1610-34F3-43F5-86B7-16C984F0E88E"'
    'Public Function NewEnum() As stdole.IUnknown
    '
    'NewEnum = mcolLatSupports.GetEnumerator
    '
    'End Function

    Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator

        GetEnumerator = mcolLatSupports.GetEnumerator

    End Function

End Class