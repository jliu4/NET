Option Strict Off
Option Explicit On
Friend Class DrftFCbyFreq
	
	' drift force coefficient (0 - 180 deg even spacing) in ship local system
	' by frequency
	
	' properties
	' Freq:         wave frequency
	
	Private msngFreq As Single
	Private mclsFCx As ForceCoef
	Private mclsFCy As ForceCoef
	Private mclsFCm As ForceCoef
	
	
	Public Sub New()
		MyBase.New()
        mclsFCx = New ForceCoef
        mclsFCy = New ForceCoef
        mclsFCm = New ForceCoef

        mclsFCx.ClampedEnd = True
        mclsFCy.ClampedEnd = False
        mclsFCm.ClampedEnd = False

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
	
	Public ReadOnly Property FCx() As ForceCoef
		Get
			
			FCx = mclsFCx
			
		End Get
	End Property
	
	Public ReadOnly Property FCy() As ForceCoef
		Get
			
			FCy = mclsFCy
			
		End Get
	End Property
	
	Public ReadOnly Property FCm() As ForceCoef
		Get
			
			FCm = mclsFCm
			
		End Get
	End Property
End Class