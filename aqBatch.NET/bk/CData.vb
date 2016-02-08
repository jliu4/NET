Option Strict Off
Option Explicit On
Friend Class CData
	' self-aware
	
	Private mbIsEmpty As Boolean
	Private mValue As Object
	
	Event IsModified()
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		mbIsEmpty = True
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	Public ReadOnly Property IsEmpty() As Boolean
		Get
			IsEmpty = mbIsEmpty
		End Get
	End Property
	
	
	Public Property Value() As Object
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object mValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Value = mValue
		End Get
		Set(ByVal Value As Object)
			'UPGRADE_WARNING: Couldn't resolve default property of object NewValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mValue = Value
			mbIsEmpty = False
			RaiseEvent IsModified()
		End Set
	End Property
End Class