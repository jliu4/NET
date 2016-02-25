Option Strict Off
Option Explicit On
Friend Class threeDObjects
    Implements System.Collections.IEnumerable

    Private mcolthreeDObjects As New Collection
	Private mintIndex As Short
	
	Public Function Add(ByVal New3DObj As threeDObject, Optional ByVal OName As String = "") As threeDObject
		If OName = "" Then
			mintIndex = mintIndex + 1
			OName = CStr(mintIndex)
		End If
		mcolthreeDObjects.Add(New3DObj, OName)
		Add = New3DObj
	End Function
	
	Public Function Count() As Integer
		Count = mcolthreeDObjects.Count()
	End Function
	
	Public Function Delete(ByVal Index As Object) As Object
		mcolthreeDObjects.Remove(Index)
	End Function
	
	Public Function Item(ByVal Index As Object) As threeDObject
		Item = mcolthreeDObjects.Item(Index)
	End Function


    Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        GetEnumerator = mcolthreeDObjects.GetEnumerator
    End Function
	
	Public Function Copy(ByRef Obj3D As threeDObject, Optional ByVal NewName As String = "") As threeDObject
		Dim New3DObj As New threeDObject
		Dim Line3D As threeDLine
        Dim Lines(,) As Single
        If NewName = "" Then
			mintIndex = mintIndex + 1
			NewName = CStr(mintIndex)
		End If
		Dim Num3DLines, LN As Short
		mcolthreeDObjects.Add(New3DObj, NewName)
		Num3DLines = Obj3D.Count
		ReDim Lines(Num3DLines, 6)
		LN = 0
		For	Each Line3D In Obj3D
			LN = LN + 1
			Line3D.getCoords(Lines(LN, 1), Lines(LN, 2), Lines(LN, 3), Lines(LN, 4), Lines(LN, 5), Lines(LN, 6))
		Next Line3D
        'New3DObj.AddLines(Num3DLines, Lines)
        Copy = New3DObj
	End Function
End Class