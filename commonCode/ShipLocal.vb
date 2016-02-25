Option Strict Off
Option Explicit On
Friend Class ShipLocal
	
	' ship in local coordinates
	
	' properties
	' Heading:      ship heading (N - clockwise, 0 - 2 pi) (rad)
	' Surge:        surge from well head (fwd) (ft)
	' Sway:         sway from well head (port) (ft)
	' Yaw:          yaw from design heading (count-clockwise) (rad)
	
	Private msngHead As Single
	Private msngSurge As Single
	Private msngSway As Single
	Private msngYaw As Single

    Public Property Heading() As Single
		Get
			
			Heading = msngHead
			
		End Get
		Set(ByVal Value As Single)
			
			msngHead = Value
			
		End Set
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