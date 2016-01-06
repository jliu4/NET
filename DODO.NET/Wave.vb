Option Strict Off
Option Explicit On
Friend Class Wave
	
	' wave properties
	
	' properties
	' Heading:      wave heading (N clockwise) (rad)
	' Height:       wave height (significant) (ft)
	' Period:       wave period (peak) (sec)
	' Frequency:    wave frequency (peak)
	' Spectrum:     wave spectrum value
	
	' functions
	' ISSC:         calculations of spectrum
	
	Private msngHeading As Single
	Private msngHeight As Single
	Private msngPeriod As Single
	Private msngFrequency As Single
	
	
	Public Property Heading() As Single
		Get
			
			Heading = msngHeading
			
		End Get
		Set(ByVal Value As Single)
			
			msngHeading = Value
			
		End Set
	End Property
	
	
	Public Property Height() As Single
		Get
			
			Height = msngHeight
			
		End Get
		Set(ByVal Value As Single)
			
			msngHeight = Value
			
		End Set
	End Property
	
	
	Public Property Period() As Single
		Get
			
			Period = msngPeriod
			
		End Get
		Set(ByVal Value As Single)
			
			msngPeriod = Value
			If msngPeriod = 0# Then
				msngFrequency = 10000000000#
			Else
				msngFrequency = 2# * PI / msngPeriod
			End If
			
		End Set
	End Property
	
	Public ReadOnly Property Frequency() As Single
		Get
			
			Frequency = msngFrequency
			
		End Get
	End Property
	
	Public ReadOnly Property Spectrum(ByVal Frequency As Single) As Single
		Get
			
			' Input
			'   Frequency:  look-up wave frequency
			
			Spectrum = ISSC(Frequency)
			
		End Get
	End Property
	
	' functions
	
	'wave spectrum ISSC
	Private Function ISSC(ByRef w As Single) As Single
		
		' Input
		'   w:          wave frequency
		
		Dim B1, A1, wp As Single
		
		If w = 0# Then
			ISSC = 0#
		Else
			wp = msngFrequency
			If wp = 0# Then
				A1 = 0#
			Else
				A1 = ((0.3125 * msngHeight ^ 2) / wp)
			End If
			If w = 0# Then
				B1 = 0#
			Else
				B1 = wp / w
			End If
			
			ISSC = A1 * B1 ^ 5 * System.Math.Exp(-1.25 * B1 ^ 4)
		End If
		
	End Function
End Class