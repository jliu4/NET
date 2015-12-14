Option Strict Off
Option Explicit On
Friend Class Motion
	
	' motion components
	
	' properties
	' Move:         total movement (ft)
	' Surge:        move in surge direction (ft)
	' Sway:         move in sway direction (ft)
	' Yaw:          yaw movement (rad)
	
	Private msngSurge As Single
	Private msngSway As Single
	Private msngYaw As Single
	
	Public ReadOnly Property Move() As Single
		Get
			
			Move = System.Math.Sqrt(msngSurge ^ 2 + msngSway ^ 2)
			
		End Get
	End Property
	
	
	Public Property Surge() As Single
		Get
			
			Surge = msngSurge
			
		End Get
		Set(ByVal Value As Single)
			
			msngSurge = Value
			
		End Set
	End Property
	
	
	Public Property Sway() As Single
		Get
			
			Sway = msngSway
			
		End Get
		Set(ByVal Value As Single)
			
			msngSway = Value
			
		End Set
	End Property
	
	
	Public Property Yaw() As Single
		Get
			
			Yaw = msngYaw
			
		End Get
		Set(ByVal Value As Single)
			
			msngYaw = Value
			
		End Set
	End Property
End Class