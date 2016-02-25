Option Strict Off
Option Explicit On
Friend Class PltSeabed
	
	Public WaterDepth As Single
	
	Public seaGrapher As threeDGrapher
	Private Grid As threeDLine
	
	Private X As Single '...the min. x coords data (GPS data)
	Private Y As Single '...the min. y coords data (GPS data)
	Private z(MaxNumGrids - 1, MaxNumGrids - 1) As Single '...the seabed z(x grid points, y grid points) coords data
	Private xlength As Single '...x direct total length
	Private ylength As Single '...y direct total length
	Private xgridpt As Short '...x direct total grid points
	Private ygridpt As Short '...y direct total grid points

    Public Sub New()
        MyBase.New()
        Grid = New threeDLine
        Grid.setColor(&H40C0) 'RGB(0, 60, 60)
        seaGrapher = New threeDGrapher
        ReadFile()
    End Sub

    Public Sub ReadFile()
		
		Dim i, j As Short
        Dim FileOpen_Renamed As Boolean

        On Error GoTo ExitErr '...if not find data file then don't draw pipelines
		FileOpen_Renamed = False
		
		FileOpen(FileNumRes, MarsDir & SeaBedFile, OpenMode.Input) ' Open file for input.
		FileOpen_Renamed = True
		
		Input(FileNumRes, xlength)
		Input(FileNumRes, ylength)
		Input(FileNumRes, xgridpt)
		Input(FileNumRes, ygridpt)
		Input(FileNumRes, X)
		Input(FileNumRes, Y) ' Read data into two variables.
		
		If xgridpt > MaxNumGrids Then xgridpt = MaxNumGrids
		If ygridpt > MaxNumGrids Then ygridpt = MaxNumGrids
		
		For i = 0 To xgridpt - 1
			For j = 0 To ygridpt - 1
				Input(FileNumRes, z(i, j)) '...read z data
			Next j
		Next i
		FileClose(FileNumRes) ' Close file.
		FileOpen_Renamed = False
		
		Exit Sub
		
ExitErr: 
		If FileOpen_Renamed Then FileClose(FileNumRes)
		
		For i = 0 To xgridpt - 1
			For j = 0 To ygridpt - 1
				If z(i, j) >= 0# Then z(i, j) = -WaterDepth
			Next j
		Next i
		Exit Sub
		
	End Sub
	
	Public Sub drawSeabed(ByRef ShipGlob As ShipGlobal, ByVal ShipDraft As Single)
		
		Dim i, j As Short
		Dim X0 As Single '...the min. x coords data (reference ship location data)
		Dim Y0 As Single '...the min. y coords data (reference ship location data)
		Dim x1 As Single '...the start point x coords data (reference ship location data)
		Dim y1 As Single '...the start point y coords data (reference ship location data)
		Dim x2 As Single '...the end point x coords data (reference ship location data)
		Dim y2 As Single '...the end point x coords data (reference ship location data)
		Dim dx As Single '...x direct spacing
		Dim dy As Single '...y direct spacing
		
		'   ...defined dx,dy,x1,y1,x2,y2
		dx = xlength / (xgridpt - 1)
		dy = ylength / (ygridpt - 1)
		With ShipGlob
			X0 = X - .Xg '...reference to ship sys. coords
			Y0 = Y - .Yg
		End With
		
		'   ...draw y direction grid
		x1 = X0
		For i = 0 To xgridpt - 1
			y1 = Y0
			For j = 1 To ygridpt - 1
				y2 = y1 + dy
				If z(i, j - 1) < 0# And z(i, j) < 0# Then
					Grid.setCoords(x1, y1, z(i, j - 1) + ShipDraft, x1, y2, z(i, j) + ShipDraft)
					seaGrapher.drawTheLine(Grid)
				End If
				y1 = y2
			Next j
			x1 = x1 + dx
		Next i
		
		'   ...draw x direction grid
		y1 = Y0
		For j = 0 To ygridpt - 1
			x1 = X0
			For i = 1 To xgridpt - 1
				x2 = x1 + dx
				If z(i - 1, j) < 0# And z(i, j) < 0# Then
					Grid.setCoords(x1, y1, z(i - 1, j) + ShipDraft, x2, y1, z(i, j) + ShipDraft)
					seaGrapher.drawTheLine(Grid)
				End If
				x1 = x2
			Next i
			y1 = y1 + dy
		Next j
		
	End Sub
	
	Public Sub angleChange(ByVal hChange As Short, ByRef vChange As Short)
		
		seaGrapher.newAngles(hChange, vChange)
		
	End Sub
	
	Public Sub setGraph(ByRef aGrapher As System.Windows.Forms.PictureBox, ByRef X As Single, ByRef Y As Single, ByRef Incremental As Boolean)
		
		seaGrapher.setDrawSurface(aGrapher, X, Y, Incremental)
		
	End Sub
	
	Public Function ZoomGraph(ByVal aChange As Single, ByRef ShipGlob As ShipGlobal, ByVal ShipDraft As Single) As Object
		
		Dim zoomVar As Boolean
		
		zoomVar = seaGrapher.changeScale(aChange)
		If zoomVar = True Then
			drawSeabed(ShipGlob, ShipDraft)
		End If
        ZoomGraph = zoomVar

    End Function
End Class