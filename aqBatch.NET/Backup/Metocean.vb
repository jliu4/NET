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
	
	Private mbIsEmpty As Boolean
	Private mbIsModified As Boolean
	
	Private WithEvents mclsCurrent As Current
	Private WithEvents mclsWave As Wave
	Private WithEvents mclsWind As Wind
	
	' initializing and terminating
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		mclsWind = New Wind
		mclsWave = New Wave
		mclsCurrent = New Current
		mbIsModified = False
		mbIsEmpty = True
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		
		'UPGRADE_NOTE: Object mclsWind may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mclsWind = Nothing
		'UPGRADE_NOTE: Object mclsWave may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mclsWave = Nothing
		'UPGRADE_NOTE: Object mclsCurrent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mclsCurrent = Nothing
		
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	' properties
	
	
	Public Property Name() As String
		Get
			
			Name = mstrName
			
		End Get
		Set(ByVal Value As String)
			
			mstrName = Value
			mbIsModified = True
			mbIsEmpty = False
		End Set
	End Property
	
	Public WriteOnly Property Heading() As Single
		Set(ByVal Value As Single)
			
			mclsWind.Heading = Value
			mclsCurrent.Heading = Value
			mclsWave.Heading = Value
			mclsWave.SwellHeading = Value
			
			mbIsModified = True
			mbIsEmpty = False
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
	
	Public ReadOnly Property IsModified() As Boolean
		Get
			IsModified = mbIsModified
		End Get
	End Property
	
	Public ReadOnly Property IsEmpty() As Boolean
		Get
			IsEmpty = mbIsEmpty
		End Get
	End Property
	
	Public Function WriteData(ByVal fnum As Integer, ByVal UDEF As Boolean) As Object
		Dim i As Short
		
		If InStr(mclsWave.SpectrumName.Value, "PSMZ") > 0 Then
			mclsWave.gamma.Value = ""
		End If
		If InStr(mclsWave.SwellSpectrumName.Value, "PSMZ") > 0 Then
			mclsWave.SwellGamma.Value = ""
		End If
		If UDEF Then
			PrintLine(fnum, VB6.Format(mclsWind.Velocity.Value, "#0.0000"), mclsWind.Heading.Value, mclsWave.SpectrumName.Value, mclsWave.Height.Value, mclsWave.Period.Value, mclsWave.gamma.Value, mclsWave.Heading.Value, mclsCurrent.Heading.Value, VB6.Format(mclsCurrent.SurfaceVel, "#0.0000"), mclsCurrent.Profile.Count, mclsWave.SwellHeight.Value, mclsWave.SwellPeriod.Value, mclsWave.SwellSpectrumName.Value, mclsWave.SwellGamma.Value, mclsWave.SwellHeading.Value)
		Else
			PrintLine(fnum, VB6.Format(mclsWind.Velocity.Value, "#0.0000"), mclsWind.Heading.Value, mclsWave.SpectrumName.Value, mclsWave.Height.Value, mclsWave.Period.Value, mclsWave.gamma.Value, mclsWave.Heading.Value, mclsCurrent.Heading.Value, VB6.Format(mclsCurrent.SurfaceVel, "#0.0000"), mclsCurrent.Profile.Count)
		End If
		If mclsCurrent.Profile.Count > 1 Then
			PrintLine(fnum, "WD" & Space(10) & "Current Vel")
			For i = 1 To mclsCurrent.Profile.Count
				PrintLine(fnum, mclsCurrent.Profile.Item(i).Depth.Value, VB6.Format(mclsCurrent.Profile.Item(i).Velocity.Value, "#0.0000"))
			Next i
		End If
		mbIsModified = False
		
	End Function
	
	Public Function ReadData(ByVal fnum As Integer, ByVal UDEF As Boolean) As Object
		Dim i, TmpCount As Short
		Dim aline As String
		Dim Fields() As String
		Dim NumFields As Short
		
		'  file already opened
		aline = LineInput(fnum)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Split_Renamed(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Fields = Split_Renamed(aline, " ")
		
		mclsWind.Velocity.Value = Fields(1) '  skip case no
		mclsWind.Heading.Value = Fields(2)
		
		mclsWave.SpectrumName.Value = Fields(3)
		mclsWave.Height.Value = Fields(4)
		mclsWave.Period.Value = Fields(5)
		If InStr(mclsWave.SpectrumName.Value, "JONH") > 0 Then
			mclsWave.gamma.Value = Fields(6)
			mclsWave.Heading.Value = Fields(7)
			mclsCurrent.Heading.Value = Fields(8)
			mclsCurrent.SurfaceVel = CDbl(Fields(9))
			TmpCount = CShort(Fields(10))
			If UDEF Then
				mclsWave.SwellHeight.Value = Fields(11)
				mclsWave.SwellPeriod.Value = Fields(12)
				mclsWave.SwellSpectrumName.Value = Fields(13)
				mclsWave.SwellGamma.Value = Fields(14)
				mclsWave.SwellHeading.Value = Fields(15)
			End If
		Else
			mclsWave.gamma.Value = ""
			mclsWave.Heading.Value = Fields(6)
			mclsCurrent.Heading.Value = Fields(7)
			mclsCurrent.SurfaceVel = CDbl(Fields(8))
			TmpCount = CShort(Fields(9))
			
			If UDEF Then
				mclsWave.SwellHeight.Value = Fields(10)
				mclsWave.SwellPeriod.Value = Fields(11)
				mclsWave.SwellSpectrumName.Value = Fields(12)
				mclsWave.SwellGamma.Value = Fields(13)
				mclsWave.SwellHeading = Fields(14)
				
			End If
		End If
		
		If TmpCount > 1 Then
			aline = LineInput(fnum) ' discard header WD Vel
			mclsCurrent.Profile.Clear()
			For i = 1 To TmpCount
				aline = LineInput(fnum)
				'UPGRADE_WARNING: Couldn't resolve default property of object Split_Renamed(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Fields = Split_Renamed(aline, " ")
				mclsCurrent.Profile.Add(CSng(Fields(0)), CSng(Fields(1)))
			Next i
		Else
			mclsCurrent.Profile.Clear()
			mclsCurrent.Profile.Add(0, CSng(Fields(TmpCount)))
		End If
		mbIsModified = False
	End Function
	
	Private Sub mclsCurrent_Modified() Handles mclsCurrent.Modified
		mbIsModified = True
		mbIsEmpty = False
	End Sub
	
	Private Sub mclsWave_Modified() Handles mclsWave.Modified
		mbIsModified = True
		mbIsEmpty = False
	End Sub
	
	Private Sub mclsWind_Modified() Handles mclsWind.Modified
		mbIsModified = True
		mbIsEmpty = False
	End Sub
End Class