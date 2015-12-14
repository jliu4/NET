Option Strict Off
Option Explicit On
Friend Class Wells
	Implements System.Collections.IEnumerable
	' Wells     well collection
	' Version 1.0.1
	' 2001, Copyright DTCEL, All Rights Reserved
	
	
	' properties
	' CurWellNo current well no
	
	' methods
	' Add       add wave component
	' Clear     clear all components
	' Count     count component number
	' Delete    delete wave component
	' Item      find wave component item
	
	' InputWS   input well site data from file
	' OutputWS  output well site data to file
	
	Private mintCurWellNo As Short
	Private mcolWells As Collection
	
	' initiation and termination
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		mcolWells = New Collection
		DefaultWell()
		mintCurWellNo = 1
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		
		'UPGRADE_NOTE: Object mcolWells may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mcolWells = Nothing
		
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	' properties
	
	
	Public Property CurWellNo() As Short
		Get
			
			CurWellNo = mintCurWellNo
			
		End Get
		Set(ByVal Value As Short)
			
			mintCurWellNo = Value
			
		End Set
	End Property
	
	' methods
	
	Public Sub Add(ByVal NameID As String, ByVal Xg As Single, ByVal Yg As Single, ByVal Depth As Single)
		
		Dim WellNew As New Well
		
		With WellNew
			.NameID = NameID
			.Xg = Xg
			.Yg = Yg
			.Depth = Depth
		End With
		mcolWells.Add(WellNew)
		
		mintCurWellNo = mcolWells.Count()
		
	End Sub
	
	Public Sub Clear()
		
		Dim NumWells, i As Short
		
		NumWells = mcolWells.Count()
		For i = NumWells To 1 Step -1
			mcolWells.Remove(i)
		Next i
		
	End Sub
	
	Public Function Count() As Integer
		
		Count = mcolWells.Count()
		
	End Function
	
	Private Sub Delete(ByVal Index As Object)
		
		mcolWells.Remove(Index)
		
	End Sub
	
	Public Function Item(ByVal Index As Object) As Well
		
		Item = mcolWells.Item(Index)
		
	End Function
	
	'UPGRADE_NOTE: NewEnum property was commented out. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B3FC1610-34F3-43F5-86B7-16C984F0E88E"'
	'Public Function NewEnum() As stdole.IUnknown
		'
		'NewEnum = mcolWells.GetEnumerator
		'
	'End Function
	
	Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
		'UPGRADE_TODO: Uncomment and change the following line to return the collection enumerator. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="95F9AAD0-1319-4921-95F0-B9D3C4FF7F1C"'
		'GetEnumerator = mcolWells.GetEnumerator
	End Function
	
	Public Function InputWS(ByVal FileNum As Short) As Boolean
		
		Dim i As Short
		Dim NumWells As Short
		Dim NameID As String
		Dim Y, X, Depth As Single
		
		Dim NewMoorLine As MoorLine
		
		InputWS = False
		
		On Error GoTo ErrorHandler
		
		Input(FileNum, NumWells)
		
		Clear()
		If NumWells <= 0 Then
			DefaultWell()
		Else
			For i = 1 To NumWells
				Input(FileNum, X)
				Input(FileNum, Y)
				Input(FileNum, Depth)
				Input(FileNum, NameID)
				Add(NameID, X, Y, Depth)
			Next i
		End If
		mintCurWellNo = 1
		
		InputWS = True
		
ErrorHandler: 
		
	End Function
	
	Public Function OutputWS(ByVal FileNum As Short) As Boolean
		
		Dim i As Short
		Dim NumWells As Short
		
		OutputWS = False
		
		NumWells = mcolWells.Count()
		WriteLine(FileNum, NumWells)
		
		For i = 1 To NumWells
			With mcolWells.Item(i)
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolWells(i).NameID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolWells(i).Depth. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolWells(i).Yg. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolWells(i).Xg. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				WriteLine(FileNum, .Xg, .Yg, .Depth, .NameID)
			End With
		Next i
		
		OutputWS = True
		
	End Function
	
	Private Sub DefaultWell()
		
		Dim wellDefault As New Well
		
		With wellDefault
			.NameID = "DEFAULT"
			.Xg = 0#
			.Yg = 0#
			.Depth = 5000
		End With
		
		mcolWells.Add(wellDefault)
		
	End Sub
End Class