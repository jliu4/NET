Option Strict Off
Option Explicit On
Friend Class CurrentPair
	' Properties of the class
	Private WithEvents msngDepth As CData
	Private WithEvents msngVelocity As CData
	Private mintPairNum As Short
	
	Event Modified()
	
	
	
	Public Property PairNum() As Short
		Get
			PairNum = mintPairNum
		End Get
		Set(ByVal Value As Short)
			Static blnAlreadySet As Boolean
			If Not blnAlreadySet Then
				blnAlreadySet = True
				mintPairNum = Value
			Else
				Err.Raise(Number:=vbObjectError + 32107, Description:="Depth/Velocity" & CStr(Value) & " already defined.")
			End If
		End Set
	End Property
	
	
	Public Property Depth() As CData
		Get
			Depth = msngDepth
		End Get
		Set(ByVal Value As CData)
			msngDepth = Value
			RaiseEvent Modified()
		End Set
	End Property
	
	
	Public Property Velocity() As CData
		Get
			Velocity = msngVelocity
		End Get
		Set(ByVal Value As CData)
			msngVelocity = Value
			RaiseEvent Modified()
		End Set
	End Property
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		msngDepth = New CData
		msngVelocity = New CData
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object msngVelocity may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		msngVelocity = Nothing
		'UPGRADE_NOTE: Object msngDepth may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		msngDepth = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Private Sub msngDepth_IsModified() Handles msngDepth.IsModified
		RaiseEvent Modified()
	End Sub
	
	Private Sub msngVelocity_IsModified() Handles msngVelocity.IsModified
		RaiseEvent Modified()
	End Sub
End Class