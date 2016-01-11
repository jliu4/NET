Option Strict Off
Option Explicit On
Friend Class ShipMass
	
	' ship mass and added mass
	
	' properties
	' Draft:        ship draft (ft)
	
	' MassSurge:    mass in surge direction (slug)
	' MassSway:     mass in sway direction (slug)
	'Invalid_string_refer_to_original_code
	
	' AddMassSurge: added mass in surge direction (slug)
	' AddMassSway:  added mass in sway direction (slug)
	'Invalid_string_refer_to_original_code
	
	' VirMassSurge: virtual mass in surge direction (slug)
	' VirMassSway:  virtual mass in sway direction (slug)
	'Invalid_string_refer_to_original_code
	
	Private msngDraft As Single
	
	Private msngMassSurge As Single
	Private msngMassSway As Single
	Private msngMassYaw As Single
	
	Private msngAddMassSurge As Single
	Private msngAddMassSway As Single
	Private msngAddMassYaw As Single
	
	
	Public Property Draft() As Single
		Get
			
			Draft = msngDraft
			
		End Get
		Set(ByVal Value As Single)
			
			msngDraft = Value
			
		End Set
	End Property
	
	
	Public Property MassSurge() As Single
		Get
			
			MassSurge = msngMassSurge
			
		End Get
		Set(ByVal Value As Single)
			
			msngMassSurge = Value
			
		End Set
	End Property
	
	
	Public Property MassSway() As Single
		Get
			
			MassSway = msngMassSway
			
		End Get
		Set(ByVal Value As Single)
			
			msngMassSway = Value
			
		End Set
	End Property
	
	
	Public Property MassYaw() As Single
		Get
			
			MassYaw = msngMassYaw
			
		End Get
		Set(ByVal Value As Single)
			
			msngMassYaw = Value
			
		End Set
	End Property
	
	
	
	Public Property AddMassSurge() As Single
		Get
			
			AddMassSurge = msngAddMassSurge
			
		End Get
		Set(ByVal Value As Single)
			
			msngAddMassSurge = Value
			
		End Set
	End Property
	
	
	Public Property AddMassSway() As Single
		Get
			
			AddMassSway = msngAddMassSway
			
		End Get
		Set(ByVal Value As Single)
			
			msngAddMassSway = Value
			
		End Set
	End Property
	
	
	Public Property AddMassYaw() As Single
		Get
			
			AddMassYaw = msngAddMassYaw
			
		End Get
		Set(ByVal Value As Single)
			
			msngAddMassYaw = Value
			
		End Set
	End Property
	
	Public ReadOnly Property VirMassSurge() As Single
		Get
			
			VirMassSurge = msngMassSurge + msngAddMassSurge
			
		End Get
	End Property
	
	Public ReadOnly Property VirMassSway() As Single
		Get
			
			VirMassSway = msngMassSway + msngAddMassSway
			
		End Get
	End Property
	
	Public ReadOnly Property VirMassYaw() As Single
		Get
			
			VirMassYaw = msngMassYaw + msngAddMassYaw
			
		End Get
	End Property
End Class