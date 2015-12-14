Option Strict Off
Option Explicit On
Friend Class CurrentPair
	' Properties of the class
	Private msngDepth As Single
	Private msngVelocity As Single
	Private mintPairNum As Short
	
	
	Public Property PairNum() As Short
		Get
			PairNum = mintPairNum
		End Get
		Set(ByVal Value As Short)
			Static blnAlreadySet As Boolean
			If Not blnAlreadySet Then
				blnAlreadySet = True
				mintPairNum = Value
			Else
				Err.Raise(Number:=vbObjectError + 32107, Description:="Depth/Velocity" & CStr(Value) & " already defined.")
			End If
		End Set
	End Property
	
	
	Public Property Depth() As Single
		Get
			Depth = msngDepth
		End Get
		Set(ByVal Value As Single)
			msngDepth = Value
		End Set
	End Property
	
	
	Public Property Velocity() As Single
		Get
			Velocity = msngVelocity
		End Get
		Set(ByVal Value As Single)
			msngVelocity = Value
		End Set
	End Property
End Class