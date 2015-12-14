Option Strict Off
Option Explicit On
Friend Class PltAxes
	
	
	
	' this class tells us what makes us the x,y,z coords.
	' specifically it is made up of 3 lines.
	
	'Private grapher As threeDGrapher
	Public xAxis As threeDLine
	Public yAxis As threeDLine
	Public zAxis As threeDLine
	
	' initialize the lines
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		Dim grapher As Object
		xAxis = New threeDLine
		yAxis = New threeDLine
		zAxis = New threeDLine
		grapher = New threeDGrapher
		xAxis.setCoords(0, 0, 0, 50, 0, 0)
		yAxis.setCoords(0, 0, 0, 0, 50, 0)
		zAxis.setCoords(0, 0, 0, 0, 0, 50)
		xAxis.setColor(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		yAxis.setColor(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Lime))
		zAxis.setColor(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red))
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'Public Sub setGraph(ByRef aGraph As PictureBox)
	'Set Grapher = New threeDGrapher
	'grapher.setDrawSurface aGraph
	
	'End Sub
	
	Public Sub drawAxis(ByRef grapher As threeDGrapher)
		' need to call the grapher's line drawing method to
		' draw the three lines
		'Grapher.clearGraph
		Call grapher.drawTheLine(xAxis)
		Call grapher.drawTheLine(yAxis)
		Call grapher.drawTheLine(zAxis)
	End Sub
	
	Public Sub angleChange(ByVal hChange As Short, ByRef vChange As Short)
		Dim grapher As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object grapher.newAngles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		grapher.newAngles(hChange, vChange)
	End Sub
End Class