Option Strict Off
Option Explicit On
Friend Class Force
	
	' force components
	
	' properties
	' Ft:           total horizontal force (lbs)
	' Fx:           force in x direction (lbs)
	' Fy:           force in y direction (lbs)
	' MYaw:         yaw moment (count-clockwise) (lbs-ft)
	
	Private msngFx As Single
	Private msngFy As Single
	Private msngMYaw As Single
	
	Public ReadOnly Property Ft() As Single
		Get
			
			Ft = System.Math.Sqrt(msngFx ^ 2 + msngFy ^ 2)
			
		End Get
	End Property
	
	
	Public Property Fx() As Single
		Get
			
			Fx = msngFx
			
		End Get
		Set(ByVal Value As Single)
			
			msngFx = Value
			
		End Set
	End Property
	
	
	Public Property Fy() As Single
		Get
			
			Fy = msngFy
			
		End Get
		Set(ByVal Value As Single)
			
			msngFy = Value
			
		End Set
	End Property
	
	
	Public Property MYaw() As Single
		Get
			
			MYaw = msngMYaw
			
		End Get
		Set(ByVal Value As Single)
			
			msngMYaw = Value
			
		End Set
	End Property
End Class