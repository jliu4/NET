Option Strict Off
Option Explicit On
Friend Class ShipYawRate
	
	' ship yaw rate drag
	
	' properties
	' Draft:        ship draft (ft)
	
	' YawRateDrag:  yaw rate drag (lb-ft/(rad/sec)^2)
	
	Private msngDraft As Single
	
	Private msngYawRateDrag As Single
	
	
	Public Property Draft() As Single
		Get
			
			Draft = msngDraft
			
		End Get
		Set(ByVal Value As Single)
			
			msngDraft = Value
			
		End Set
	End Property
	
	
	Public Property YawRateDrag() As Single
		Get
			
			YawRateDrag = msngYawRateDrag
			
		End Get
		Set(ByVal Value As Single)
			
			msngYawRateDrag = Value
			
		End Set
	End Property
End Class