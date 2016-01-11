Option Strict Off
Option Explicit On
Friend Class Metocean
	
	' metocean criteria
	
	' properties
	' Name:         name
	' Heading:      set uniform enviroment heading to wind wave and current
	' Current:      current
	' Wave:         wave
	' Wind:         wind
	
	Private mstrName As String
	Private mclsCurrent As Current
	Private mclsWave As Wave
	Private mclsWind As Wind
	
	Public Sub New()
		MyBase.New()
        mclsWind = New Wind
        mclsWave = New Wave
        mclsCurrent = New Current
	End Sub
	

	Public Property Name() As String
		Get
			
			Name = mstrName
			
		End Get
		Set(ByVal Value As String)
			
			mstrName = Value
			
		End Set
	End Property
	
	Public WriteOnly Property Heading() As Single
		Set(ByVal Value As Single)
			
			mclsWind.Heading = Value
			mclsCurrent.Heading = Value
			mclsWave.Heading = Value
			
		End Set
	End Property
	
	Public ReadOnly Property Current() As Current
		Get
			
			Current = mclsCurrent
			
		End Get
	End Property
	
	Public ReadOnly Property Wave() As Wave
		Get
			
			Wave = mclsWave
			
		End Get
	End Property
	
	Public ReadOnly Property Wind() As Wind
		Get
			
			Wind = mclsWind
			
		End Get
	End Property
End Class