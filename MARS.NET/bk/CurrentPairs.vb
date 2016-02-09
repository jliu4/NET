Option Strict Off
Option Explicit On
Friend Class CurrentPairs
	Implements System.Collections.IEnumerable
	Private mcolCurrentPairs As New Collection
	Private mstrCurrentFile As String
	
	
	Public Property CurrentFile() As String
		Get
			CurrentFile = mstrCurrentFile
		End Get
		Set(ByVal Value As String)
			mstrCurrentFile = Value
		End Set
	End Property
	
	Public Function Add(ByVal CDepth As Single, ByVal CVelocity As Single) As CurrentPair
		Dim crpNew As New CurrentPair
		Static intPairNum As Integer
		With crpNew
			intPairNum = intPairNum + 1
			.PairNum = intPairNum
			.Depth = CDepth
			.Velocity = CVelocity
			mcolCurrentPairs.Add(crpNew, CStr(.PairNum))
		End With
		Add = crpNew
	End Function
	
	Public Function Count() As Integer
		Count = mcolCurrentPairs.Count()
	End Function
	
	Public Function Delete(ByVal Index As Object) As Object
		mcolCurrentPairs.Remove(Index)
	End Function
	
	Public Function Item(ByVal Index As Object) As CurrentPair
		Item = mcolCurrentPairs.Item(Index)
	End Function
	
	'UPGRADE_NOTE: NewEnum property was commented out. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B3FC1610-34F3-43F5-86B7-16C984F0E88E"'
	'Public Function NewEnum() As stdole.IUnknown
		'NewEnum = mcolCurrentPairs.GetEnumerator
	'End Function
	
	Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
		'UPGRADE_TODO: Uncomment and change the following line to return the collection enumerator. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="95F9AAD0-1319-4921-95F0-B9D3C4FF7F1C"'
		'GetEnumerator = mcolCurrentPairs.GetEnumerator
	End Function
End Class