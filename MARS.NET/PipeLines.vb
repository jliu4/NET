Option Strict Off
Option Explicit On
Friend Class PltPipeLines
	
	' this class creats pinelines by read data file.
	
	Public pGrapher As threeDGrapher
	'UPGRADE_WARNING: Lower bound of array thePipeLines was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Private thePipeLines(MaxNumPipes) As threeDLine
	Private NumPipes As Short
	
	'UPGRADE_WARNING: Lower bound of array NumPipeNodes was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Private NumPipeNodes(MaxNumPipes) As Short '...total pipe knot points
	'UPGRADE_WARNING: Lower bound of array xyz was changed from 1,0,1 to 0,0,0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Private xyz(MaxNumPipes, MaxNumPipeNodes - 1, 3) As Single '...xyz(pipenumber,pipepoints,GPS sys. coords)
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		Dim i As Short
		
		NumPipes = 0
		
		For i = 1 To MaxNumPipes
			thePipeLines(i) = New threeDLine
			thePipeLines(i).setColor(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Cyan))
		Next i
		
		pGrapher = New threeDGrapher
		ReadFile()
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	Private Sub ReadFile()
		
		Dim i, j As Short
		'UPGRADE_NOTE: FileOpen was upgraded to FileOpen_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim FileOpen_Renamed As Boolean
		
		On Error GoTo ExitErr '...if not find data file then don't draw pipelines
		FileOpen_Renamed = False
		
		FileOpen(FileNumRes, MarsDir & PipeFile, OpenMode.Input) ' Open file for input.
		FileOpen_Renamed = True
		
		Input(FileNumRes, NumPipes) ' Read data into two variables.
		For i = 1 To NumPipes
			Input(FileNumRes, NumPipeNodes(i))
			For j = 0 To NumPipeNodes(i) - 1
				Input(FileNumRes, xyz(i, j, 1))
				Input(FileNumRes, xyz(i, j, 2))
				Input(FileNumRes, xyz(i, j, 3)) '...read xyz coords data
			Next j
		Next i
		
		FileClose(FileNumRes) ' Close file.
		FileOpen_Renamed = False
		Exit Sub
		
ExitErr: 
		If FileOpen_Renamed Then FileClose(FileNumRes)
		NumPipes = 0
		Exit Sub
		
	End Sub
	
	Public Sub drawPipeLines(ByRef ShipGlob As ShipGlobal, ByVal ShipDraft As Single)
		
		Dim i, j As Short
		'UPGRADE_WARNING: Lower bound of array Rxyz was changed from 0,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim Rxyz(MaxNumPipeNodes - 1, 3) As Single
		
		'...reference to ship sys. coords
		With ShipGlob
			For i = 1 To NumPipes Step 1
				For j = 0 To NumPipeNodes(i) - 1 Step 1
					Rxyz(j, 1) = xyz(i, j, 1) - .Xg
					Rxyz(j, 2) = xyz(i, j, 2) - .Yg
					Rxyz(j, 3) = xyz(i, j, 3) + ShipDraft
					If j > 0 Then
						thePipeLines(i).setCoords(Rxyz(j - 1, 1), Rxyz(j - 1, 2), Rxyz(j - 1, 3), Rxyz(j, 1), Rxyz(j, 2), Rxyz(j, 3))
						pGrapher.drawTheLine(thePipeLines(i))
					End If
				Next j
			Next i
		End With
		
	End Sub
	
	Public Sub angleChange(ByVal hChange As Short, ByRef vChange As Short)
		pGrapher.newAngles(hChange, vChange)
	End Sub
	
	Public Sub setGraph(ByRef aGrapher As System.Windows.Forms.PictureBox, ByRef X As Single, ByRef Y As Single, ByRef Incremental As Boolean)
		pGrapher.setDrawSurface(aGrapher, X, Y, Incremental)
	End Sub
	
	Public Function ZoomGraph(ByRef aChange As Single, ByRef ShipGlob As ShipGlobal, ByRef ShipDraft As Single) As Boolean
		
		Dim zoomVar As Boolean
		
		zoomVar = pGrapher.changeScale(aChange)
		If zoomVar = True Then
			drawPipeLines(ShipGlob, ShipDraft)
		End If
		ZoomGraph = zoomVar
		
	End Function
End Class