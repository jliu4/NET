Option Strict Off
Option Explicit On
Friend Class Current
	
	' current properties
	
	' properties
	' Heading:      current heading (N clock-wise)
	' Profile:      Current profile (water depth, Velocity)
	
	' Physical constants
	Private WithEvents mclsHeading As CData
	Private mclsSurfVel As Double ' surface current velocity
	Private WithEvents moProfile As CurrentProfile
	Event Modified()
	
	' initializing and terminating
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		moProfile = New CurrentProfile
		mclsHeading = New CData
		mclsSurfVel = 0
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		
		'UPGRADE_NOTE: Object moProfile may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		moProfile = Nothing
		'UPGRADE_NOTE: Object mclsHeading may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mclsHeading = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	
	Public Property SurfaceVel() As Double
		Get
			
			SurfaceVel = mclsSurfVel
			
		End Get
		Set(ByVal Value As Double)
			
			mclsSurfVel = Value
			'  moProfile.Clear
			' moProfile.Add 0, vData
			RaiseEvent Modified()
			
		End Set
	End Property
	
	
	' properties
	
	Public Property Heading() As CData
		Get
			
			Heading = mclsHeading
			
		End Get
		Set(ByVal Value As CData)
			
			mclsHeading = Value
			RaiseEvent Modified()
			
		End Set
	End Property
	
	
	Public Property Profile() As CurrentProfile
		Get
			Profile = moProfile
		End Get
		Set(ByVal Value As CurrentProfile)
			moProfile = Value
			RaiseEvent Modified()
		End Set
	End Property
	
	Private Sub moProfile_Modified() Handles moProfile.Modified
		RaiseEvent Modified()
	End Sub
	
	Private Sub mclsHeading_IsModified() Handles mclsHeading.IsModified
		RaiseEvent Modified()
	End Sub
End Class