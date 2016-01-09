Option Strict Off
Option Explicit On
Friend Class Wind
	
	' wind properties
	
	' properties
	' Duration:     time period over which the wind velocity is measured (sec)
	' Elevation:    where the wind velocity is measured (ft)
	' Heading:      wind heading (N clock-wise) (rad)
	' Velocity:     wind velocity (ft/s)
	' VelCorr:      wind velocity corrected for elevation (ft/s)
	
	' functions
	' DurCorr:      duration correction factor
	
	Private mintDuration As Short ' e.g. 3-min gust, 1-min sustained
	Private msngElevation As Single ' reference height
	Private mstrSpecName As String
	Private mstrSpecDataString As String
	Private WithEvents msngHeading As CData
	Private WithEvents msngVelocity As CData
	Private msngVelCorr As Single
	
	Private mblnUpdated As Boolean
	Event Modified()
	
	Private Const RefElev As Single = 32.80839295
	
	' initializing and terminating
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		msngElevation = RefElev
		mintDuration = 60
		msngHeading = New CData
		msngVelocity = New CData
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub

    ' properties	

    Public Property SpecName() As String
		Get
			SpecName = mstrSpecName
		End Get
		Set(ByVal Value As String)
			mstrSpecName = Value
		End Set
	End Property
	
	
	Public Property SpecDataString() As String
		Get
			SpecDataString = mstrSpecDataString
		End Get
		Set(ByVal Value As String)
			mstrSpecDataString = Value
		End Set
	End Property
	
	
	Public Property Duration() As Single
		Get
			
			Duration = mintDuration
			
		End Get
		Set(ByVal Value As Single)
			
			mintDuration = Value
			
		End Set
	End Property
	
	
	Public Property Elevation() As Single
		Get
			
			Elevation = msngElevation
			
		End Get
		Set(ByVal Value As Single)
			
			msngElevation = Value
			mblnUpdated = False
			RaiseEvent Modified()
		End Set
	End Property
	
	
	Public Property Heading() As CData
		Get
			
			Heading = msngHeading
			
		End Get
		Set(ByVal Value As CData)
			
			msngHeading = Value
			RaiseEvent Modified()
			
		End Set
	End Property
	
	
	Public Property Velocity() As CData
		Get
			
			Velocity = msngVelocity
			
		End Get
		Set(ByVal Value As CData)
			
			msngVelocity = Value
			mblnUpdated = False
			RaiseEvent Modified()
			
		End Set
	End Property
	
	Public ReadOnly Property VelCorr() As Single
		Get
			
			If Not mblnUpdated Then
                msngVelCorr = Convert.ToDouble(msngVelocity) * (msngElevation / RefElev) ^ 0.2 * DurCorr(mintDuration)
                mblnUpdated = True
			End If
			VelCorr = msngVelCorr
			
		End Get
	End Property
	
	' functions
	
	Private Function DurCorr(ByRef Duration As Short) As Single
		
		' Input
		'   Duration:   wind measurement duration
		
		Dim Ns, N, i As Short
		Dim Dx As Single
		Dim RefDur(6) As Single
		Dim Fact(6) As Single
		Dim B(6) As Single
		Dim c(6) As Single
		Dim D(6) As Single
		
		N = 6
		RefDur(1) = 3600#
		RefDur(2) = 600#
		RefDur(3) = 60#
		RefDur(4) = 15#
		RefDur(5) = 5#
		RefDur(6) = 3#
		Fact(1) = 1#
		Fact(2) = 1.06
		Fact(3) = 1.18
		Fact(4) = 1.26
		Fact(5) = 1.31
		Fact(6) = 1.33
		
		'   initiate array
		Ns = 0
		For i = 1 To N
			If Duration >= RefDur(i) And Ns = 0 Then Ns = i
		Next 
		
		'   interpolate by spline
		Dx = Duration - RefDur(Ns)
		If Dx = 0# Then
			DurCorr = Fact(Ns)
		Else
			Call Spline3(N - 1, RefDur, Fact, B, c, D, False)
			DurCorr = Fact(Ns) + B(Ns) * Dx + c(Ns) * Dx ^ 2 + D(Ns) * Dx ^ 3
		End If
		
		If DurCorr <= 0# Then
			DurCorr = 10000000000#
		Else
			DurCorr = Fact(3) / DurCorr
		End If
		
	End Function

    Private Sub Class_Terminate_Renamed()
        msngHeading = Nothing
        msngVelocity = Nothing
    End Sub
    Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Private Sub msngHeading_IsModified() Handles msngHeading.IsModified
		RaiseEvent Modified()
	End Sub
	
	Private Sub msngVelocity_IsModified() Handles msngVelocity.IsModified
		RaiseEvent Modified()
	End Sub
End Class