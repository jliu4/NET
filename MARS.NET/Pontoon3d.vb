Option Strict Off
Option Explicit On
Friend Class PltPontoon3d
	' manages data associated with a pontoon.  Also contains
	' a method to draw the pontoons.  Must create both the
	' left and right pontoons
	
	' note all dimensions are in feet and radians
	' 1meter = 3.28 feet
	
	' dimensions is a structure used to hold the pontoons
	' different dimensions
	
	
	Private Structure Dimensions
		Dim PLength As Single
		Dim PointLength As Single
		Dim PWidth As Single
		Dim PHeight As Single
		Dim PDistance As Single
	End Structure
	
	' panels is used to hold the drawing components of the
	' pontoon
	
	Private Structure Panels
		Dim FrontP1 As threeDRect '...front-triangle left side rect
		Dim FrontP2 As threeDRect '...front-triangle right side rect
		Dim RearP1 As threeDRect '...rear-triangle left side rect
		Dim RearP2 As threeDRect '...rear-triangle right side rect
		Dim InsideP1 As threeDRect '...inside left side rect
		Dim InsideP2 As threeDRect '...inside right side rect
		Dim Top As threeDRect '...top rect
		Dim Bottom As threeDRect '...bottom side
		'...front-triangle
		Dim bottomFL As threeDRect '...bottom left triangle
		Dim topFL As threeDRect '...top left triangle
		Dim bottomFR As threeDRect '...bottom right triangle
		Dim topFR As threeDRect '...top right triangle
		'...rear-triangle
		Dim bottomRL As threeDRect
		Dim bottomRR As threeDRect
		Dim topRL As threeDRect
		Dim topRR As threeDRect
	End Structure
	
	Private PDim As Dimensions
	'...modify by neil
	Private newPDim As Dimensions
	
	Private LeftPontoon As Panels '...left pontoon (inside)
	Private RightPontoon As Panels '...left pontoon (inside)
	'...modify by neil
	Private newLeftPontoon As Panels '...new left pontoon (outside)
	Private newRightPontoon As Panels '...new right pontoon (outside)
	
	Private theBoards As PltOutboard3d
	Private theColumns As PltColumns
	
	
	' initializes the class.  Used to set the dimensions
	' of a particular pontoon
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		PDim.PHeight = PontHeight
		PDim.PLength = PontLength
		PDim.PWidth = PontWidth
		PDim.PointLength = PLength
		PDim.PDistance = PDistance
		'...modify by neil
		newPDim.PHeight = PontHeight
		newPDim.PLength = PontLength / 3
		newPDim.PWidth = PontWidth
		newPDim.PointLength = 0.1
		newPDim.PDistance = PDistance - 15
		
		initializePanels(LeftPontoon)
		initializePanels(RightPontoon)
		'...modify by neil
		initializePanels(newLeftPontoon)
		initializePanels(newRightPontoon)
		
		createPanels(PDim, LeftPontoon, True, False, True)
		createPanels(PDim, RightPontoon, False, False, False)
		'...modify by neil
		createPanels(newPDim, newLeftPontoon, True, True, True)
		createPanels(newPDim, newRightPontoon, False, True, False)
		
		theColumns = New PltColumns
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	Private Sub initializePanels(ByRef aPanel As Panels)
		aPanel.FrontP1 = New threeDRect
		aPanel.FrontP2 = New threeDRect
		aPanel.InsideP1 = New threeDRect
		aPanel.InsideP2 = New threeDRect
		aPanel.RearP1 = New threeDRect
		aPanel.RearP2 = New threeDRect
		aPanel.bottomFL = New threeDRect
		aPanel.topFL = New threeDRect
		aPanel.bottomFR = New threeDRect
		aPanel.topFR = New threeDRect
		aPanel.bottomRL = New threeDRect
		aPanel.bottomRR = New threeDRect
		aPanel.topRL = New threeDRect
		aPanel.topRR = New threeDRect
		aPanel.Top = New threeDRect
		aPanel.Bottom = New threeDRect
		
		theBoards = New PltOutboard3d
	End Sub
	
	' we assume that we are looking head on to the vessel
	
	'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub createPanels(ByRef aDim As Dimensions, ByRef aPanel As Panels, ByRef Left_Renamed As Boolean, ByRef secondpontoon As Boolean, ByRef newLeft As Boolean)
		
		' note all rectangles will be centered at same z coordinate
		Dim xCenter As Single
		Dim ycenter As Single
		Dim zCenter As Single
		Dim angle As Single
		Dim xOffset As Single
		Dim Length As Single
		
		Dim X As Single
		Dim Y As Single
		X = 0
		Y = 0
		
		If Left_Renamed = True Then
			xOffset = -aDim.PDistance / 2
			xOffset = xOffset - aDim.PWidth / 2
		Else
			xOffset = aDim.PDistance / 2
			xOffset = xOffset + aDim.PWidth / 2
		End If
		
		'...modify by neil
		'...if have second pontoon beside
		If secondpontoon = True Then
			If newLeft = True Then
				xOffset = xOffset - aDim.PWidth - aDim.PDistance '...left pontoon
			Else
				xOffset = xOffset + aDim.PWidth + aDim.PDistance '...right pontoon
			End If
		End If
		
		zCenter = aDim.PHeight / 2
		
		
		Dim radAngle As Single
		Dim width As Single
		Dim zTemp As Single
		
		' rectangle top of pontoon
		xCenter = xOffset
		ycenter = 0
		zTemp = aDim.PHeight
		Length = aDim.PWidth
		width = aDim.PLength - 2 * aDim.PointLength
		angle = 0
		aPanel.Top.createHorizontalRect(Length, width, xCenter, ycenter, zTemp, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		'rectangle bottom of pontoon
		xCenter = xOffset
		ycenter = 0
		zTemp = 0
		Length = aDim.PWidth
		width = aDim.PLength - 2 * aDim.PointLength
		aPanel.Bottom.createHorizontalRect(Length, width, xCenter, ycenter, zTemp, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' bottom rear left
		xCenter = xOffset
		ycenter = -aDim.PLength / 2
		ycenter = ycenter + aDim.PointLength
		zTemp = 0
		angle = 90
		Length = aDim.PointLength
		width = aDim.PWidth / 2
		aPanel.bottomRL.createTriangle(Length, width, xCenter, ycenter, zTemp, angle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' bottom rear right
		xCenter = xOffset
		ycenter = -aDim.PLength / 2
		ycenter = ycenter + aDim.PointLength
		zTemp = 0
		angle = 0
		Length = aDim.PWidth / 2
		width = aDim.PointLength
		aPanel.bottomRR.createTriangle(Length, width, xCenter, ycenter, zTemp, angle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' top rear right
		xCenter = xOffset
		ycenter = -aDim.PLength / 2
		ycenter = ycenter + aDim.PointLength
		zTemp = aDim.PHeight
		angle = 0
		Length = aDim.PWidth / 2
		width = aDim.PointLength
		aPanel.topRR.createTriangle(Length, width, xCenter, ycenter, zTemp, angle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		
		' top rear left
		xCenter = xOffset
		ycenter = -aDim.PLength / 2
		ycenter = ycenter + aDim.PointLength
		zTemp = aDim.PHeight
		angle = 90
		Length = aDim.PointLength
		width = aDim.PWidth / 2
		aPanel.topRL.createTriangle(Length, width, xCenter, ycenter, zTemp, angle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		
		' bottomFL
		xCenter = xOffset
		zTemp = 0
		ycenter = aDim.PLength / 2
		ycenter = ycenter - aDim.PointLength
		angle = -90
		Length = aDim.PointLength
		width = aDim.PWidth / 2
		aPanel.bottomFL.createTriangle(Length, width, xCenter, ycenter, zTemp, angle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' topFL
		xCenter = xOffset
		ycenter = aDim.PLength / 2
		ycenter = ycenter - aDim.PointLength
		angle = -90
		Length = aDim.PointLength
		width = aDim.PWidth / 2
		zTemp = aDim.PHeight
		aPanel.topFL.createTriangle(Length, width, xCenter, ycenter, zTemp, angle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' bottom front right triangle
		xCenter = xOffset
		ycenter = aDim.PLength / 2
		ycenter = ycenter - aDim.PointLength
		angle = 180
		zTemp = 0
		width = aDim.PointLength
		Length = aDim.PWidth / 2
		zTemp = 0
		aPanel.bottomFR.createTriangle(Length, width, xCenter, ycenter, zTemp, angle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' top front right triangle
		xCenter = xOffset
		ycenter = aDim.PLength / 2
		ycenter = ycenter - aDim.PointLength
		angle = 180
		zTemp = aDim.PHeight
		width = aDim.PointLength
		Length = aDim.PWidth / 2
		aPanel.topFR.createTriangle(Length, width, xCenter, ycenter, zTemp, angle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		
		
		' FrontP1
		ycenter = aDim.PLength / 2 - aDim.PointLength / 2
		xCenter = -aDim.PWidth / 4
		xCenter = xCenter + xOffset
		radAngle = System.Math.Atan(aDim.PointLength / (aDim.PWidth / 2))
		angle = radAngle * 180 / PI
		Length = aDim.PointLength / System.Math.Sin(radAngle)
		aPanel.FrontP1.createVerticalRect(Length, aDim.PHeight, xCenter, ycenter, zCenter, angle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' FrontP2
		ycenter = aDim.PLength / 2 - aDim.PointLength / 2
		xCenter = aDim.PWidth / 4
		xCenter = xCenter + xOffset
		radAngle = System.Math.Atan((aDim.PWidth / 2) / aDim.PointLength)
		angle = radAngle * 180 / PI
		Length = aDim.PointLength / System.Math.Cos(radAngle)
		angle = 180 - (90 + angle)
		angle = 180 - angle
		aPanel.FrontP2.createVerticalRect(Length, aDim.PHeight, xCenter, ycenter, zCenter, angle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		'insideP2
		xCenter = xOffset - aDim.PWidth / 2
		ycenter = 0
		Length = aDim.PLength - (aDim.PointLength * 2)
		aPanel.InsideP2.createVerticalRect(Length, aDim.PHeight, xCenter, ycenter, zCenter, 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' insideP1
		xCenter = xOffset + aDim.PWidth / 2
		ycenter = 0
		Length = aDim.PLength - (aDim.PointLength * 2)
		aPanel.InsideP1.createVerticalRect(Length, aDim.PHeight, xCenter, ycenter, zCenter, 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		'rearp1
		xCenter = -aDim.PWidth / 4
		xCenter = xOffset + xCenter
		ycenter = -(aDim.PLength / 2 - aDim.PointLength / 2)
		radAngle = System.Math.Atan(aDim.PointLength / (aDim.PWidth / 2))
		Length = aDim.PointLength / System.Math.Sin(radAngle)
		angle = radAngle * 180 / PI
		angle = -angle
		aPanel.RearP1.createVerticalRect(Length, aDim.PHeight, xCenter, ycenter, zCenter, angle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		'rearP2
		xCenter = aDim.PWidth / 4
		xCenter = xOffset + xCenter
		ycenter = -(aDim.PLength / 2 - aDim.PointLength / 2)
		radAngle = System.Math.Atan(aDim.PointLength / (aDim.PWidth / 2))
		Length = aDim.PointLength / System.Math.Sin(radAngle)
		angle = radAngle * 180 / PI
		aPanel.RearP2.createVerticalRect(Length, aDim.PHeight, xCenter, ycenter, zCenter, angle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		
	End Sub
	
	
	Public Sub drawPontoon(ByRef a3dGrapher As threeDGrapher)
		'...left pontoon(inside)
		a3dGrapher.drawRect(LeftPontoon.FrontP1)
		a3dGrapher.drawRect(LeftPontoon.FrontP2)
		a3dGrapher.drawRect(LeftPontoon.InsideP1)
		a3dGrapher.drawRect(LeftPontoon.InsideP2)
		'...modify by neil
		'a3dGrapher.drawRect LeftPontoon.RearP1
		'a3dGrapher.drawRect LeftPontoon.RearP2
		a3dGrapher.drawRect(LeftPontoon.bottomFL)
		a3dGrapher.drawRect(LeftPontoon.bottomFR)
		a3dGrapher.drawRect(LeftPontoon.topFL)
		a3dGrapher.drawRect(LeftPontoon.topFR)
		'...modify by neil
		'a3dGrapher.drawRect LeftPontoon.bottomRL
		'a3dGrapher.drawRect LeftPontoon.bottomRR
		'a3dGrapher.drawRect LeftPontoon.topRL
		'a3dGrapher.drawRect LeftPontoon.topRR
		a3dGrapher.drawRect(LeftPontoon.Top)
		a3dGrapher.drawRect(LeftPontoon.Bottom)
		
		'...right pontoon(inside)
		a3dGrapher.drawRect(RightPontoon.FrontP1)
		a3dGrapher.drawRect(RightPontoon.FrontP2)
		a3dGrapher.drawRect(RightPontoon.InsideP1)
		a3dGrapher.drawRect(RightPontoon.InsideP2)
		'...modify by neil
		'a3dGrapher.drawRect RightPontoon.RearP1
		'a3dGrapher.drawRect RightPontoon.RearP2
		a3dGrapher.drawRect(RightPontoon.bottomFL)
		a3dGrapher.drawRect(RightPontoon.topFL)
		a3dGrapher.drawRect(RightPontoon.bottomFR)
		a3dGrapher.drawRect(RightPontoon.topFR)
		'...modify by neil
		'a3dGrapher.drawRect RightPontoon.bottomRL
		'a3dGrapher.drawRect RightPontoon.bottomRR
		'a3dGrapher.drawRect RightPontoon.topRL
		'a3dGrapher.drawRect RightPontoon.topRR
		a3dGrapher.drawRect(RightPontoon.Top)
		a3dGrapher.drawRect(RightPontoon.Bottom)
		
		'...modify by neil
		'...new left pontoon (outside)
		a3dGrapher.drawRect(newLeftPontoon.InsideP1)
		a3dGrapher.drawRect(newLeftPontoon.InsideP2)
		a3dGrapher.drawRect(newLeftPontoon.Top)
		a3dGrapher.drawRect(newLeftPontoon.Bottom)
		'...modify by neil
		'...new right pontoon (outside)
		a3dGrapher.drawRect(newRightPontoon.InsideP1)
		a3dGrapher.drawRect(newRightPontoon.InsideP2)
		a3dGrapher.drawRect(newRightPontoon.Top)
		a3dGrapher.drawRect(newRightPontoon.Bottom)
		
		
		'...modify by neil
		'theBoards.DrawOutboards a3dGrapher
		'...modify by neil
		theColumns.drawColumns(a3dGrapher)
		
	End Sub
End Class