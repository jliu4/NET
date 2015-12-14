Option Strict Off
Option Explicit On
Friend Class ShipDamp
	
	' ship linear damping of LFM
	
	' properties
	' Draft:        ship draft (ft)
	' DampSurge:    damping in surge direction (lbs/(ft/s))
	' DampSway:     damping in sway direction (lbs/(ft/s))
	' DampYaw:      damping in yaw direction (lbs/(ft/s))
	
	Private msngDraft As Single
	
	Private msngDampSurge As Single
	Private msngDampSway As Single
	Private msngDampYaw As Single
	
	
	Public Property Draft() As Single
		Get
			
			Draft = msngDraft
			
		End Get
		Set(ByVal Value As Single)
			
			msngDraft = Value
			
		End Set
	End Property
	
	
	Public Property DampSurge() As Single
		Get
			
			DampSurge = msngDampSurge
			
		End Get
		Set(ByVal Value As Single)
			
			msngDampSurge = Value
			
		End Set
	End Property
	
	
	Public Property DampSway() As Single
		Get
			
			DampSway = msngDampSway
			
		End Get
		Set(ByVal Value As Single)
			
			msngDampSway = Value
			
		End Set
	End Property
	
	
	Public Property DampYaw() As Single
		Get
			
			DampYaw = msngDampYaw
			
		End Get
		Set(ByVal Value As Single)
			
			msngDampYaw = Value
			
		End Set
	End Property
End Class