Option Strict Off
Option Explicit On
Friend Class CurrentData
	
	' current profile data
	
	' properties
	' Depth:        water depth (ft)
	' Velocity:     current velocity (ft/s)
	
	Private msngDepth As Single
	Private msngVelocity As Single
	
	
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