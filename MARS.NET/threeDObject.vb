Option Strict Off
Option Explicit On
Friend Class threeDObject
	Implements System.Collections.IEnumerable
	Private mcolthreeDLines As New Collection
	
	'Public Function Add(x1 As Single, y1 As Single, z1 As Single, _
	''                    x2 As Single, y2 As Single, z2 As Single) As threeDLine
	'Dim l3dNew As New threeDLine
	'Static int3DLNum As Long
	'    With l3dNew
	'        int3DLNum = int3DLNum + 1
	'        .setCoords x1, y1, z1, x2, y2, z2
	'        mcolthreeDLines.Add l3dNew, int3DLNum
	'    End With
	'    Set Add = l3dNew
	'End Function
	
	Public Function AddLines(ByRef Num3DLines As Short, ByRef Lines() As Single) As Object
		Dim LN As Short
		Dim New3DLine As threeDLine
		Static mintIndex As Integer
		For LN = 1 To Num3DLines
			mintIndex = mintIndex + 1
			New3DLine = New threeDLine
            'New3DLine.setCoords(Lines(LN, 1), Lines(LN, 2), Lines(LN, 3), Lines(LN, 4), Lines(LN, 5), Lines(LN, 6))
            mcolthreeDLines.Add(New3DLine, CStr(mintIndex))
		Next LN
	End Function
	
	Public Function Count() As Integer
		Count = mcolthreeDLines.Count()
	End Function
	
	Public Function DeleteLine(ByVal Index As Object) As Object
		mcolthreeDLines.Remove(Index)
	End Function
	
	Public Function Item(ByVal Index As Object) As threeDLine
		Item = mcolthreeDLines.Item(Index)
	End Function
	
	'UPGRADE_NOTE: NewEnum property was commented out. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B3FC1610-34F3-43F5-86B7-16C984F0E88E"'
	'Public Function NewEnum() As stdole.IUnknown
		'NewEnum = mcolthreeDLines.GetEnumerator
	'End Function
	
	Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
		'UPGRADE_TODO: Uncomment and change the following line to return the collection enumerator. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="95F9AAD0-1319-4921-95F0-B9D3C4FF7F1C"'
		'GetEnumerator = mcolthreeDLines.GetEnumerator
	End Function
	
	Public Function Translate(ByRef dx As Single, ByRef dy As Single, ByRef Dz As Single) As Object
		Dim Line3D As threeDLine
		'UPGRADE_WARNING: Lower bound of array X was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim X(6) As Single
		For	Each Line3D In mcolthreeDLines
			Line3D.getCoords(X(1), X(2), X(3), X(4), X(5), X(6))
			X(1) = X(1) + dx
			X(2) = X(2) + dy
			X(3) = X(3) + Dz
			X(4) = X(4) + dx
			X(5) = X(5) + dy
			X(6) = X(6) + Dz
			Line3D.setCoords(X(1), X(2), X(3), X(4), X(5), X(6))
		Next Line3D
	End Function
	
	Public Function Rotate(ByRef ThetaX As Single, ByRef ThetaY As Single, ByRef ThetaZ As Single) As Object
		Dim Line3D As threeDLine
		'UPGRADE_WARNING: Lower bound of array X was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim X(6) As Single
		'UPGRADE_WARNING: Lower bound of array X0 was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim X0(3) As Single
		'UPGRADE_WARNING: Lower bound of array x1 was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim x1(3) As Single
		'UPGRADE_WARNING: Lower bound of array C was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim C(3) As Single
		'UPGRADE_WARNING: Lower bound of array S was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim S(3) As Single
		'UPGRADE_WARNING: Lower bound of array A was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim A(3, 3) As Single
		Dim i, j As Short
		' Build the rotation matrix; first, determine sines and cosines for the
		' rotation angles provided
		C(mX) = System.Math.Cos(ThetaX)
		C(MY_Renamed) = System.Math.Cos(ThetaY)
		C(mZ) = System.Math.Cos(ThetaZ)
		S(mX) = System.Math.Sin(ThetaX)
		S(MY_Renamed) = System.Math.Sin(ThetaY)
		S(mZ) = System.Math.Sin(ThetaZ)
		' The rotation matrix is derived from the multiplication of three independent
		' matrices, one for each rotation.  The matrix product is:
		A(1, 1) = C(mX) * C(MY_Renamed)
		A(1, 2) = S(mX) * S(MY_Renamed) * C(mZ) - S(mX) * S(mZ)
		A(1, 3) = C(mX) * S(MY_Renamed) * C(mZ) + S(mX) * S(mZ)
		A(2, 1) = C(MY_Renamed) * S(mZ)
		A(2, 2) = C(mX) * C(mZ) + S(mX) * S(MY_Renamed) * S(mZ)
		A(2, 3) = C(mX) * S(MY_Renamed) * S(mZ) - S(mX) * C(mZ)
		A(3, 1) = -S(MY_Renamed)
		A(3, 2) = S(mX) * C(MY_Renamed)
		A(3, 3) = C(mX) * C(MY_Renamed)
		' Now loop through each line segment in the object
		For	Each Line3D In mcolthreeDLines
			Line3D.getCoords(X(1), X(2), X(3), X(4), X(5), X(6))
			' Copy first point vector
			X0(1) = X(1)
			X0(2) = X(2)
			X0(3) = X(3)
			' Rotate
			MatMul(A, X0, x1)
			' Replace original with result
			X(1) = x1(1)
			X(2) = x1(2)
			X(3) = x1(3)
			'Copy second point vector
			X0(1) = X(4)
			X0(2) = X(5)
			X0(3) = X(6)
			' Rotate
			MatMul(A, X0, x1)
			' Replace original with result
			X(4) = x1(1)
			X(5) = x1(2)
			X(6) = x1(3)
			'Reset the line coordinates
			Line3D.setCoords(X(1), X(2), X(3), X(4), X(5), X(6))
		Next Line3D
	End Function
End Class