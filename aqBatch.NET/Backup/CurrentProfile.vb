Option Strict Off
Option Explicit On
Friend Class CurrentProfile
	Implements System.Collections.IEnumerable
	Private mcolCurrentProfile As New Collection
	Private mstrCurrentFile As String
	Private mstrProfileName As String
	Private WithEvents crpNew As CurrentPair
	Event Modified()
	
	
	Public Property CurrentFile() As String
		Get
			CurrentFile = mstrCurrentFile
		End Get
		Set(ByVal Value As String)
			mstrCurrentFile = Value
			RaiseEvent Modified()
		End Set
	End Property
	
	
	Public Property ProfileName() As String
		Get
			ProfileName = mstrProfileName
		End Get
		Set(ByVal Value As String)
			mstrProfileName = Value
			RaiseEvent Modified()
		End Set
	End Property
	
	Public Function Add(ByVal CDepth As Single, ByVal CVelocity As Single) As CurrentPair
		
		crpNew = New CurrentPair
		Static intPairNum As Integer
		With crpNew
			intPairNum = intPairNum + 1
			.PairNum = intPairNum
			.Depth.Value = CDepth
			.Velocity.Value = CVelocity
			mcolCurrentProfile.Add(crpNew, CStr(.PairNum))
		End With
		Add = crpNew
		
		RaiseEvent Modified()
	End Function
	
	Public Function Count() As Integer
		Count = mcolCurrentProfile.Count()
	End Function
	
	Public Function Delete(ByVal Index As Short) As Object
		mcolCurrentProfile.Remove(Index)
	End Function
	
	Public Function Item(ByVal Index As Object) As CurrentPair
		Item = mcolCurrentProfile.Item(Index)
	End Function
	
	'UPGRADE_NOTE: NewEnum property was commented out. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B3FC1610-34F3-43F5-86B7-16C984F0E88E"'
	'Public Function NewEnum() As stdole.IUnknown
		'NewEnum = mcolCurrentProfile.GetEnumerator
	'End Function
	
	Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
		'UPGRADE_TODO: Uncomment and change the following line to return the collection enumerator. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="95F9AAD0-1319-4921-95F0-B9D3C4FF7F1C"'
		'GetEnumerator = mcolCurrentProfile.GetEnumerator
	End Function
	
	Public Sub Clear()
		Dim i As Short
		For i = Count To 1 Step -1
			Delete(i)
		Next i
		
		RaiseEvent Modified()
	End Sub
	
	Public Function GetVelocity(ByVal Depth As Single) As Single
		' This function interpolates the current velocity array
		' to provide the current at a particular depth.  Simple
		' linear interpolation is used.  Depth values outside the
		' range of the array will be assigned the velocity of the
		' appropriate endpoint (first or last)
		Dim Pair As Object
		Dim Matched As Boolean
		Dim LastDepth, LastVel As Single
		Matched = False
		'UPGRADE_WARNING: Couldn't resolve default property of object mcolCurrentProfile.Item().Depth. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		LastDepth = mcolCurrentProfile.Item(1).Depth
		'UPGRADE_WARNING: Couldn't resolve default property of object mcolCurrentProfile.Item().Velocity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		LastVel = mcolCurrentProfile.Item(1).Velocity
		If Depth <= LastDepth Then
			GetVelocity = LastVel
		Else
			For	Each Pair In mcolCurrentProfile
				With Pair
					'UPGRADE_WARNING: Couldn't resolve default property of object Pair.Depth. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If Depth >= LastDepth And Depth < .Depth Then
						Matched = True
						'UPGRADE_WARNING: Couldn't resolve default property of object Pair.Depth. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If .Depth = LastDepth Then
							GetVelocity = LastVel
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object Pair.Depth. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Pair.Velocity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							GetVelocity = LastVel + (.Velocity - LastVel) * (Depth - LastDepth) / (.Depth - LastDepth)
						End If
						Exit For
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object Pair.Depth. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						LastDepth = .Depth
						'UPGRADE_WARNING: Couldn't resolve default property of object Pair.Velocity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						LastVel = .Velocity
					End If
				End With
			Next Pair
		End If
		If Not Matched Then
			GetVelocity = LastVel
		End If
	End Function
	
	Private Sub crpNew_Modified() Handles crpNew.Modified
		RaiseEvent Modified()
	End Sub
End Class