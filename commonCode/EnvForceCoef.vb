Option Strict Off
Option Explicit On
Friend Class EnvForceCoef
	
	' environemnt force coefficient (0 - 180 deg even spacing) in ship local system
	
	' properties
	' Draft:        vessel draft (ft)
	'Invalid_string_refer_to_original_code
	'Invalid_string_refer_to_original_code
	'Invalid_string_refer_to_original_code
	
	Private msngDraft As Single
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
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		
		'UPGRADE_NOTE: Object mclsFCx may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mclsFCx = Nothing
		'UPGRADE_NOTE: Object mclsFCy may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mclsFCy = Nothing
		'UPGRADE_NOTE: Object mclsFCm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mclsFCm = Nothing
		
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	' properties
	
	
	Public Property Draft() As Single
		Get
			
			Draft = msngDraft
			
		End Get
		Set(ByVal Value As Single)
			
			msngDraft = Value
			
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