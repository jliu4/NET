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
	Private msngXg As Double
	Private msngYg As Double
	Private msngDepth As Double
	
	' properties
	
	
	Public Property NameID() As String
		Get
			
			NameID = mstrNameID
			
		End Get
		Set(ByVal Value As String)
			
			mstrNameID = Value
			
		End Set
	End Property
	
	
	Public Property Xg() As Double
		Get
			
			Xg = msngXg
			
		End Get
		Set(ByVal Value As Double)
			
			msngXg = Value
			
		End Set
	End Property
	
	
	Public Property Yg() As Double
		Get
			
			Yg = msngYg
			
		End Get
		Set(ByVal Value As Double)
			
			msngYg = Value
			
		End Set
	End Property
	
	
	Public Property Depth() As Double
		Get
			
			Depth = msngDepth
			
		End Get
		Set(ByVal Value As Double)
			
			msngDepth = Value
			
		End Set
	End Property
End Class