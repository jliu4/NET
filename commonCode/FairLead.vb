Option Strict Off
Option Explicit On
Friend Class FairLead
	
	' fairlead locations
	
	' properties
	' NeedUpdate:   coordinates changed
	' SprdAngle:    designed spreading angle (0 - 2 pi) (rad)
	' Xg:           x coordinate in global system (ft)
	' Xs:           x coordinate in ship system (ft)
	' Yg:           y coordinate in global system (ft)
	' Ys:           y coordinate in ship system (ft)
	' Z:            z coordinate (above bl of ship) (ft)
	
	Private mblnNeedUpdate As Boolean
	Private msngSprdAngle As Single
	Private msngXs As Single
	Private msngYs As Single
    Private msngZ As Single
    Private msngNode As Integer 'node in the aqwa model, default as 950i


    Friend Property NeedUpdate() As Boolean
		Get
			
			NeedUpdate = mblnNeedUpdate
			
		End Get
		Set(ByVal Value As Boolean)
			
			mblnNeedUpdate = Value
			
		End Set
	End Property
	
	
	Public Property SprdAngle() As Single
		Get
			
			SprdAngle = msngSprdAngle
			
		End Get
		Set(ByVal Value As Single)
			
			msngSprdAngle = Value
			
		End Set
	End Property
	
	Public ReadOnly Property Xg(ByVal ShipGlob As ShipGlobal) As Single
		Get
            If ShipGlob Is Nothing Then Exit Property

            ' Input
            '   ShipGlob:   ship coordinates and heading in global system (N-E)

            Dim Alpha As Single

            Alpha = PI / 2# - ShipGlob.Heading
			
			Xg = System.Math.Cos(Alpha) * msngXs - System.Math.Sin(Alpha) * msngYs + ShipGlob.Xg
			
		End Get
	End Property
	
	
	Public Property Xs() As Single
		Get
			
			Xs = msngXs
			
		End Get
		Set(ByVal Value As Single)
			
			msngXs = Value
			mblnNeedUpdate = True
			
		End Set
	End Property
	
	Public ReadOnly Property Yg(ByVal ShipGlob As ShipGlobal) As Single
		Get
            If ShipGlob Is Nothing Then Exit Property

            ' Input
            '   ShipGlob:   ship coordinates and heading in global system (N-E)

            Dim Alpha As Single
			
			Alpha = PI / 2# - ShipGlob.Heading
			
			Yg = System.Math.Sin(Alpha) * msngXs + System.Math.Cos(Alpha) * msngYs + ShipGlob.Yg
			
		End Get
	End Property
	
	
	Public Property Ys() As Single
		Get
			
			Ys = msngYs
			
		End Get
		Set(ByVal Value As Single)
			
			msngYs = Value
			mblnNeedUpdate = True
			
		End Set
	End Property


    Public Property z() As Single
        Get

            z = msngZ

        End Get
        Set(ByVal Value As Single)

            msngZ = Value

        End Set
    End Property

    Public Property Node() As Integer
        Get

            Node = msngNode

        End Get
        Set(ByVal Value As Integer)

            msngNode = Value

        End Set
    End Property
End Class