Option Strict Off
Option Explicit On
Friend Class PltWaterSurface
	
	' ...Modify by neil
	' this class creats water surface.
	
	'...defined water surface coords (x,y,z)
	'Private Const z As Single = 3.28 * -10
	Private z As Single
	Private Const X As Single = 0 '...ship sys. coords(E direction)
	Private Const Y As Single = 0 '...ship sys. coords(N direction)
	'...defined water surface range
	Private Const xlength As Single = 3000
	Private Const ylength As Single = 3000
	
	Private i As Short
	'...defined how many lines to draw
	Private Const NumLines As Short = 11
	
	Public sGrapher As threeDGrapher
	
	Private Grid As threeDLine
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		Grid = New threeDLine
		sGrapher = New threeDGrapher
		
		Grid.setColor(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue))
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	Public Sub drawSurface(ByVal ShipDraft As Single)
		
		Dim x1 As Single
		Dim y1 As Single
		Dim x2 As Single
		Dim y2 As Single
		Dim x3 As Single
		Dim y3 As Single
		Dim dx As Single
		Dim dy As Single
		
		'X = XVessel(mX) '...read vessel x position form mainform input data
		'Y = XVessel(mY) '...read vessel y position form mainform input data
		z = ShipDraft '...read vessel draft form mainform input data
		
		'...defined points to draw lines
		x1 = X - xlength / 2
		y1 = -Y - ylength / 2
		x2 = X + xlength / 2
		y2 = y1
		x3 = x1
		y3 = -Y + ylength / 2
		dx = xlength / (NumLines - 1) '...x direct spacing
		dy = ylength / (NumLines - 1) '...y direct spacing
		
		'...draw x direction grid
		For i = 1 To NumLines
			Grid.setCoords(x1, y1, z, x2, y2, z)
			sGrapher.drawTheLine(Grid)
			y1 = y1 + dy
			y2 = y2 + dy
		Next i
		
		'...draw y direction grid
		x1 = X - xlength / 2 '...set to initial value
		y1 = -Y - ylength / 2 '...set to initial value
		For i = 1 To NumLines
			Grid.setCoords(x1, y1, z, x3, y3, z)
			sGrapher.drawTheLine(Grid)
			x1 = x1 + dx
			x3 = x3 + dx
		Next i
		
	End Sub
	Public Sub angleChange(ByVal hChange As Short, ByRef vChange As Short)
		sGrapher.newAngles(hChange, vChange)
	End Sub
	Public Sub setGraph(ByRef aGrapher As System.Windows.Forms.PictureBox, ByRef X As Single, ByRef Y As Single, ByRef Incremental As Boolean)
		sGrapher.setDrawSurface(aGrapher, X, Y, Incremental)
	End Sub
	Public Function ZoomGraph(ByVal aChange As Single, ByVal ShipDraft As Single) As Object
		
		Dim zoomVar As Boolean
		zoomVar = sGrapher.changeScale(aChange)
		If zoomVar = True Then
			drawSurface(ShipDraft)
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object ZoomGraph. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ZoomGraph = zoomVar
	End Function
End Class