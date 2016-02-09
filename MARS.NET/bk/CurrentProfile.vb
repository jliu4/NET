Option Strict Off
Option Explicit On
Friend Class CurrentProfile
	Implements System.Collections.IEnumerable
	Private mcolCurrentProfile As New Collection
	Private mstrCurrentFile As String
	Private mstrProfileName As String
	
	
	Public Property CurrentFile() As String
		Get
			CurrentFile = mstrCurrentFile
		End Get
		Set(ByVal Value As String)
			mstrCurrentFile = Value
		End Set
	End Property
	
	
	Public Property ProfileName() As String
		Get
			ProfileName = mstrProfileName
		End Get
		Set(ByVal Value As String)
			mstrProfileName = Value
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
			mcolCurrentProfile.Add(crpNew, CStr(.PairNum))
		End With
		Add = crpNew
	End Function
	
	Public Function Count() As Integer
		Count = mcolCurrentProfile.Count()
	End Function
	
	Public Function Delete(ByVal Index As Object) As Object
		mcolCurrentProfile.Remove(Index)
	End Function
	
	Public Function Item(ByVal Index As Object) As CurrentPair
		Item = mcolCurrentProfile.Item(Index)
	End Function


    Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        GetEnumerator = mcolCurrentProfile.GetEnumerator
    End Function
	
	Public Function VofD(ByVal Depth As Single) As Single
		' This function interpolates the current velocity array
		' to provide the current at a particular depth.  Simple
		' linear interpolation is used.  Depth values outside the
		' range of the array will be assigned the velocity of the
		' appropriate endpoint (first or last)
		Dim Pair As Object
		Dim Matched As Boolean
		Dim LastDepth, LastVel As Single
		Matched = False
        LastDepth = mcolCurrentProfile.Item(1).Depth
        LastVel = mcolCurrentProfile.Item(1).Velocity
        If Depth <= LastDepth Then
			VofD = LastVel
		Else
			For	Each Pair In mcolCurrentProfile
				With Pair
                    If Depth >= LastDepth And Depth < .Depth Then
                        Matched = True
                        If .Depth = LastDepth Then
                            VofD = LastVel
                        Else
                            VofD = LastVel + (.Velocity - LastVel) * (Depth - LastDepth) / (.Depth - LastDepth)
                        End If
                        Exit For
                    Else
                        LastDepth = .Depth
                        LastVel = .Velocity
                    End If
				End With
			Next Pair
		End If
		If Not Matched Then
			VofD = LastVel
		End If
	End Function
End Class