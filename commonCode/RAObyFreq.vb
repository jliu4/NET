Option Strict Off
Option Explicit On
Friend Class RAObyFreq
	
	' r.a.o.s (0 - 180 deg even spacing) in ship local system
	' by frequency
	
	' properties
	' Freq:         wave frequency
	' RAOx:         r.a.o.s in x direction (surge) (ft/ft)
	' RAOy:         r.a.o.s in y direction (sway) (ft/ft)
	' RAOr:         r.a.o.s in yaw direction (rad/ft)
	
	Private msngFreq As Single
	Private mclsRAOx As ForceCoef
	Private mclsRAOy As ForceCoef
	Private mclsRAOr As ForceCoef


    Public Sub New()
        MyBase.New()
        mclsRAOx = New ForceCoef
        mclsRAOy = New ForceCoef
        mclsRAOr = New ForceCoef

        mclsRAOx.ClampedEnd = True
        mclsRAOy.ClampedEnd = False
        mclsRAOr.ClampedEnd = False
    End Sub
    ' properties

    Public Property Freq() As Single
		Get
			
			Freq = msngFreq
			
		End Get
		Set(ByVal Value As Single)
			
			msngFreq = Value
			
		End Set
	End Property
	
	Public ReadOnly Property RAOx() As ForceCoef
		Get
			
			RAOx = mclsRAOx
			
		End Get
	End Property
	
	Public ReadOnly Property RAOy() As ForceCoef
		Get
			
			RAOy = mclsRAOy
			
		End Get
	End Property
	
	Public ReadOnly Property RAOr() As ForceCoef
		Get
			
			RAOr = mclsRAOr
			
		End Get
	End Property
End Class