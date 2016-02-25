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
	Private msngGamma As Single
    Private msngSpectrumName As String
    Private msngSwellHeading As Single
	Private msngSwellHeight As Single
	Private msngSwellPeriod As Single
	Private msngSwellFrequency As Single
	Private msngSwellGamma As Single
    Private msngSwellSpectrumName As String
	
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
                msngFrequency = 2.0# * Math.PI / msngPeriod
            End If
			
		End Set
	End Property
	
	Public ReadOnly Property Frequency() As Single
		Get
			
			Frequency = msngFrequency
			
		End Get
	End Property

    Public Property SpectrumName() As String
        Get

            SpectrumName = msngSpectrumName

        End Get
        Set(ByVal Value As String)

            msngSpectrumName = Value

        End Set
    End Property

    Public Property gamma() As Single
        Get

            gamma = msngGamma

        End Get
        Set(ByVal Value As Single)

            msngGamma = Value

        End Set
    End Property

    Public Property SwellHeading() As Single
        Get

            SwellHeading = msngSwellHeading

        End Get
        Set(ByVal Value As Single)

            msngSwellHeading = Value

        End Set
    End Property


    Public Property SwellHeight() As Single
        Get

            SwellHeight = msngSwellHeight

        End Get
        Set(ByVal Value As Single)

            msngSwellHeight = Value

        End Set
    End Property


    Public Property SwellPeriod() As Single
        Get

            SwellPeriod = msngSwellPeriod

        End Get
        Set(ByVal Value As Single)

            msngSwellPeriod = Value
            If msngSwellPeriod = 0# Then
                msngSwellFrequency = 10000000000.0#
            Else
                msngSwellFrequency = 2.0# * Math.PI / msngSwellPeriod
            End If

        End Set
    End Property

    Public ReadOnly Property SwellFrequency() As Single
        Get

            SwellFrequency = msngSwellFrequency

        End Get
    End Property

    Public Property SwellSpectrumName() As String
        Get

            SwellSpectrumName = msngSwellSpectrumName

        End Get
        Set(ByVal Value As String)

            msngSwellSpectrumName = Value

        End Set
    End Property

    Public Property Swellgamma() As Single
        Get

            Swellgamma = msngSwellGamma

        End Get
        Set(ByVal Value As Single)

            msngSwellGamma = Value

        End Set
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

    Private Function Gaus(ByVal w As Double) As Double

        Dim M, f, fp, pai As Double
        pai = 4 * System.Math.Atan(1)


        If w <= 0# Or msngHeight <= 0# Or msngPeriod <= 0# Or msngGamma <= 0# Then
            Gaus = 0#
        Else
            f = w / 2.0# / pai
            fp = 1.0# / msngPeriod
            M = msngHeight ^ 2 / 16.0#

            Gaus = M / msngGamma / System.Math.Sqrt(2.0# * pai) * System.Math.Exp(-(f - fp) ^ 2 / 2.0# / msngGamma ^ 2) / 2.0# / pai
        End If

    End Function

    'wave spectrum JonSWAP
    Private Function Jons(ByVal w As Double) As Double

        Dim beta, pai, f, fp, alpha, Sigma As Double

        pai = 4 * System.Math.Atan(1)

        If w <= 0# Or msngHeight <= 0# Or msngPeriod <= 0# Or msngGamma <= 0# Then
            Jons = 0#
        Else
            f = w / 2.0# / pai
            fp = 1 / msngPeriod

            If f <= fp Then
                Sigma = 0.07
            Else
                Sigma = 0.09
            End If

            alpha = 0.0624 / (0.23 + 0.0336 * msngGamma - 0.185 / (1.9 + msngGamma))
            beta = System.Math.Exp(-(f / fp - 1.0#) ^ 2 / 2.0# / Sigma ^ 2)

            Jons = alpha * msngHeight ^ 2 * (fp / f) ^ 4 / w * System.Math.Exp(-1.25 * (fp / f) ^ 4) * msngGamma ^ beta
        End If

    End Function

    Public Function GetSpectrumValue(ByVal Omega As Double) As Double

        ' Input
        '   Frequency:  look-up wave frequency

        If CStr(msngSpectrumName) = "PSMZ" Then
            GetSpectrumValue = ISSC(Omega) + Gaus(Omega)
        Else
            GetSpectrumValue = Jons(Omega) + Gaus(Omega)
        End If

    End Function

End Class