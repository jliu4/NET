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
	
	Private mintDuration As Short
	Private msngElevation As Single
	Private msngHeading As Single
	Private msngVelocity As Single
	Private msngVelCorr As Single
	
	Private mblnUpdated As Boolean
	
	Private Const RefElev As Single = 32.80839295
	

	Public Sub New()
		MyBase.New()
        msngElevation = RefElev
        mintDuration = 60
	End Sub
	
	' properties
	
	
	Public Property Duration() As Short
		Get
			
			Duration = mintDuration
			
		End Get
		Set(ByVal Value As Short)
			
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
			
		End Set
	End Property
	
	
	Public Property Heading() As Single
		Get
			
			Heading = msngHeading
			
		End Get
		Set(ByVal Value As Single)
			
			msngHeading = Value
			
		End Set
	End Property
	
	
	Public Property Velocity() As Single
		Get
			
			Velocity = msngVelocity
			
		End Get
		Set(ByVal Value As Single)
			
			msngVelocity = Value
			mblnUpdated = False
			
		End Set
	End Property
	
	Public ReadOnly Property VelCorr() As Single
		Get
			
			If Not mblnUpdated Then
				msngVelCorr = msngVelocity * (msngElevation / RefElev) ^ 0.2 * DurCorr(mintDuration)
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
		Dim dx As Single
		Dim RefDur(6) As Single
		Dim Fact(6) As Single
		Dim B(6) As Single
		Dim C(6) As Single
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
		dx = Duration - RefDur(Ns)
		If dx = 0# Then
			DurCorr = Fact(Ns)
		Else
			Call Spline3(N - 1, RefDur, Fact, B, C, D, False)
			DurCorr = Fact(Ns) + B(Ns) * dx + C(Ns) * dx ^ 2 + D(Ns) * dx ^ 3
		End If
		
		If DurCorr <= 0# Then
			DurCorr = 10000000000#
		Else
			DurCorr = Fact(3) / DurCorr
		End If
		
	End Function
End Class