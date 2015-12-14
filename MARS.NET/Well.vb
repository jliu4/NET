Option Strict Off
Option Explicit On
Friend Class Well
	
	' well properties
	
	' properties
	' NameID:       well name and/or id
	' Xg:           x coordinate (ft)
	' Yg:           y coordinate (ft)
	' Depth:        water depth (ft)
	
	Private mstrNameID As String
	Private msngXg As Single
	Private msngYg As Single
	Private msngDepth As Single
	
	' properties
	
	
	Public Property NameID() As String
		Get
			
			NameID = mstrNameID
			
		End Get
		Set(ByVal Value As String)
			
			mstrNameID = Value
			
		End Set
	End Property
	
	
	Public Property Xg() As Single
		Get
			
			Xg = msngXg
			
		End Get
		Set(ByVal Value As Single)
			
			msngXg = Value
			
		End Set
	End Property
	
	
	Public Property Yg() As Single
		Get
			
			Yg = msngYg
			
		End Get
		Set(ByVal Value As Single)
			
			msngYg = Value
			
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
End Class