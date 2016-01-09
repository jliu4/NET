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
	
	Private WithEvents msngHeading As CData
	Private WithEvents msngHeight As CData
	Private WithEvents msngPeriod As CData
	Private WithEvents msngFrequency As CData
	Private WithEvents msngGamma As CData
	Private WithEvents msSpectrumName As CData
	Private WithEvents msngSwellHeight As CData
	Private WithEvents msngSwellPeriod As CData
	Private WithEvents msngSwellGamma As CData
	Private WithEvents msSwellSpectrumName As CData
	Private WithEvents msngSwellHeading As CData
	Event Modified()

    Public Property Heading() As CData
		Get
			
			Heading = msngHeading
			
		End Get
		Set(ByVal Value As CData)
			
			msngHeading = Value
			
			RaiseEvent Modified()
		End Set
	End Property
	
	
	Public Property Height() As CData
		Get
			
			Height = msngHeight
			
		End Get
		Set(ByVal Value As CData)
			
			msngHeight = Value
			
			RaiseEvent Modified()
		End Set
	End Property
	
	
	Public Property Period() As CData
		Get
			
			Period = msngPeriod
			
		End Get
		Set(ByVal Value As CData)
			
			msngPeriod = Value
			
			RaiseEvent Modified()
		End Set
	End Property


    Public Property Frequency() As Object
        Get

            Frequency = msngFrequency

        End Get
        Set(ByVal Value As Object)

            msngFrequency = Value
            RaiseEvent Modified()
        End Set
    End Property


    Public Property SpectrumName() As CData
		Get
			
			SpectrumName = msSpectrumName
			
		End Get
		Set(ByVal Value As CData)
			
			msSpectrumName = Value
			
			RaiseEvent Modified()
		End Set
	End Property
	
	
	Public Property SwellSpectrumName() As CData
		Get
			
			SwellSpectrumName = msSwellSpectrumName
			
		End Get
		Set(ByVal Value As CData)
			
			msSwellSpectrumName = Value
			
			RaiseEvent Modified()
		End Set
	End Property
	
	Public Property gamma() As CData
		Get
			
			gamma = msngGamma
			
		End Get
		Set(ByVal Value As CData)
			
			msngGamma = Value
			
			RaiseEvent Modified()
		End Set
	End Property
	
	
	Public Property SwellHeight() As CData
		Get
			
			SwellHeight = msngSwellHeight
			
		End Get
		Set(ByVal Value As CData)
			
			msngSwellHeight = Value
			
			RaiseEvent Modified()
		End Set
	End Property
	
	
	Public Property SwellPeriod() As CData
		Get
			
			SwellPeriod = msngSwellPeriod
			
		End Get
		Set(ByVal Value As CData)
			
			msngSwellPeriod = Value
			
			RaiseEvent Modified()
		End Set
	End Property


    Public Property SwellGamma() As CData
		Get
			
			SwellGamma = msngSwellGamma
			
		End Get
		Set(ByVal Value As CData)
			
			msngSwellGamma = Value
			
			RaiseEvent Modified()
		End Set
	End Property
	
	
	Public Property SwellHeading() As CData
		Get
			
			SwellHeading = msngSwellHeading
			
		End Get
		Set(ByVal Value As CData)
			
			msngSwellHeading = Value
			
			RaiseEvent Modified()
		End Set
	End Property
	
	' functions
	
	'wave spectrum ISSC
	Private Function ISSC(ByVal w As Double) As Double
		
		' Input
		'   w:          wave frequency
		
		Dim B1, A1, wp As Double
		
		If w = 0# Then
			ISSC = 0#
		Else
            wp = Convert.ToDouble(msngFrequency)
            If wp = 0# Then
				A1 = 0#
			Else
				A1 = ((0.3125 * msngHeight.Value ^ 2) / wp)
			End If
			If w = 0# Then
				B1 = 0#
			Else
				B1 = wp / w
			End If
			
			ISSC = A1 * B1 ^ 5 * System.Math.Exp(-1.25 * B1 ^ 4)
		End If
		
	End Function
	
	'wave spectrum Gausian
	Private Function Gaus(ByVal w As Double) As Double
		
		Dim M, f, fp, pai As Double
		Dim Hs As Single
		
		pai = 4 * System.Math.Atan(1)
		
		
		If w <= 0# Or msngHeight.Value <= 0# Or msngPeriod.Value <= 0# Or msngGamma.Value <= 0# Then
			Gaus = 0#
		Else
			f = w / 2# / pai
			fp = 1# / msngPeriod.Value
			M = msngHeight.Value ^ 2 / 16#
			
			Gaus = M / msngGamma.Value / System.Math.Sqrt(2# * pai) * System.Math.Exp(-(f - fp) ^ 2 / 2# / msngGamma.Value ^ 2) / 2# / pai
		End If
		
	End Function
	
	'wave spectrum JonSWAP
	Private Function Jons(ByVal w As Double) As Double
		
		Dim beta, pai, f, fp, alpha, Sigma As Double
		
		pai = 4 * System.Math.Atan(1)
		
		If w <= 0# Or msngHeight.Value <= 0# Or msngPeriod.Value <= 0# Or msngGamma.Value <= 0# Then
			Jons = 0#
		Else
			f = w / 2# / pai
			fp = 1 / msngPeriod.Value
			
			If f <= fp Then
				Sigma = 0.07
			Else
				Sigma = 0.09
			End If
			
			alpha = 0.0624 / (0.23 + 0.0336 * msngGamma.Value - 0.185 / (1.9 + msngGamma.Value))
			beta = System.Math.Exp(-(f / fp - 1#) ^ 2 / 2# / Sigma ^ 2)
			
			Jons = alpha * msngHeight.Value ^ 2 * (fp / f) ^ 4 / w * System.Math.Exp(-1.25 * (fp / f) ^ 4) * msngGamma.Value ^ beta
		End If
		
	End Function
	
	Public Function GetSpectrumValue(ByVal Omega As Double) As Double
		
		' Input
		'   Frequency:  look-up wave frequency
		
		If CStr(msSpectrumName.Value) = "PSMZ" Then
			GetSpectrumValue = ISSC(Omega) + Gaus(Omega)
		Else
			GetSpectrumValue = Jons(Omega) + Gaus(Omega)
		End If
		
	End Function
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		msngHeading = New CData
		msngHeight = New CData
		msngPeriod = New CData
		msngFrequency = New CData
		msngGamma = New CData
		msSpectrumName = New CData
		msngSwellHeight = New CData
		msngSwellPeriod = New CData
		msngSwellGamma = New CData
		msngSwellHeading = New CData
		msSwellSpectrumName = New CData
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object msngHeading may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		msngHeading = Nothing
		'UPGRADE_NOTE: Object msngHeight may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		msngHeight = Nothing
		'UPGRADE_NOTE: Object msngPeriod may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		msngPeriod = Nothing
		'UPGRADE_NOTE: Object msngFrequency may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		msngFrequency = Nothing
		'UPGRADE_NOTE: Object msngGamma may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		msngGamma = Nothing
		'UPGRADE_NOTE: Object msSpectrumName may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		msSpectrumName = Nothing
		'UPGRADE_NOTE: Object msSwellSpectrumName may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		msSwellSpectrumName = Nothing
		'UPGRADE_NOTE: Object msngSwellHeight may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		msngSwellHeight = Nothing
		'UPGRADE_NOTE: Object msngSwellPeriod may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		msngSwellPeriod = Nothing
		'UPGRADE_NOTE: Object msngSwellGamma may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		msngSwellGamma = Nothing
		'UPGRADE_NOTE: Object msngSwellHeading may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		msngSwellHeading = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Private Sub msngFrequency_IsModified() Handles msngFrequency.IsModified
		RaiseEvent Modified()
	End Sub
	
	Private Sub msngGamma_IsModified() Handles msngGamma.IsModified
		RaiseEvent Modified()
	End Sub
	
	Private Sub msngHeading_IsModified() Handles msngHeading.IsModified
		RaiseEvent Modified()
	End Sub
	
	Private Sub msngHeight_IsModified() Handles msngHeight.IsModified
		RaiseEvent Modified()
	End Sub
	
	Private Sub msngPeriod_IsModified() Handles msngPeriod.IsModified
		RaiseEvent Modified()
	End Sub
	
	Private Sub msSpectrumName_IsModified() Handles msSpectrumName.IsModified
		RaiseEvent Modified()
	End Sub
	Private Sub msSwellSpectrumName_IsModified() Handles msSwellSpectrumName.IsModified
		RaiseEvent Modified()
	End Sub
	
	Private Sub msngSwellHeight_IsModified() Handles msngSwellHeight.IsModified
		RaiseEvent Modified()
	End Sub
	
	Private Sub msngSwellPeriod_IsModified() Handles msngSwellPeriod.IsModified
		RaiseEvent Modified()
	End Sub
	
	
	Private Sub msngSwellGamma_IsModified() Handles msngSwellGamma.IsModified
		RaiseEvent Modified()
	End Sub
	
	Private Sub msngSwellHeading_IsModified() Handles msngSwellHeading.IsModified
		RaiseEvent Modified()
	End Sub
End Class