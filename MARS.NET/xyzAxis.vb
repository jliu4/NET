Option Strict Off
Option Explicit On
Friend Class PltAxes

    ' this class tells us what makes us the x,y,z coords.
    ' specifically it is made up of 3 lines.

    'Private grapher As threeDGrapher
    Public xAxis As threeDLine
	Public yAxis As threeDLine
	Public zAxis As threeDLine

    Public Sub New()
		MyBase.New()
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
        grapher.newAngles(hChange, vChange)
    End Sub
End Class