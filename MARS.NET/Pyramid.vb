Option Strict Off
Option Explicit On
Friend Class PltPyramid
	' this class holds data that is used to draw a
	' pyramid
	
	' member variables.  Has a threeDGrapher as well as
	' threeDLines which are needed to construct a pyramid
	Private Graphing As System.Windows.Forms.PictureBox
	Public theGraph As threeDGrapher
	Private bLength1 As threeDLine
	Private bLength2 As threeDLine
	Private bWidth1 As threeDLine
	Private bWidth2 As threeDLine
	Private L1W1 As threeDLine
	Private L1W2 As threeDLine
	Private L2W2 As threeDLine
	Private L2W1 As threeDLine
	Private testRect As threeDRect
	Private testRect2 As threeDRect
	Private testRect3 As threeDRect
	Private xyz As PltAxes
	
	
	Public Sub drawPyramid()
		theGraph.clearGraph()
		'xyz.setGraph Graphing
		'xyz.drawAxis
		theGraph.drawRect(testRect)
		theGraph.drawRect(testRect2)
		theGraph.drawRect(testRect3)
		theGraph.drawTheLine(bLength1)
		theGraph.drawTheLine(bLength2)
		theGraph.drawTheLine(bWidth1)
		theGraph.drawTheLine(bWidth2)
		theGraph.drawTheLine(L1W1)
		theGraph.drawTheLine(L1W2)
		theGraph.drawTheLine(L2W2)
		theGraph.drawTheLine(L2W1)
	End Sub
	
	Public Sub angleChange(ByVal H As Short, ByVal V As Short)
		theGraph.newAngles(H, V)
		
	End Sub
	
	Public Sub setGraph(ByRef gIn As System.Windows.Forms.PictureBox)
		theGraph = New threeDGrapher
		theGraph.setDrawSurface(gIn, 0, 0, False)
		' xyz.setGraph gIn
		Graphing = gIn
	End Sub
	
	' this method tells us the dimensions of the pyramid
	' pyramid is centered at the center of the base
	' of the pyramid
	
	Public Sub PyramidDim(ByVal bLength As Short, ByVal bWidth As Short, ByVal Height As Short)
		Dim x1 As Single
		Dim y1 As Single
		Dim z1 As Single
		Dim x2 As Single
		Dim y2 As Single
		Dim z2 As Single
		Dim XTemp As Single
		Dim YTemp As Single

        ' let's get the base lines created.  The height is
        ' not important at this point
        ' do bLength1
        XTemp = bLength / 2
		YTemp = bWidth / 2
		x1 = XTemp
		y1 = YTemp
		x2 = -XTemp
		y2 = YTemp
		z1 = 0
		z2 = 0
		bLength1.setCoords(x1, y1, z1, x2, y2, z2)

        'do blength2
        y1 = -YTemp
		y2 = -YTemp
		bLength2.setCoords(x1, y1, z1, x2, y2, z2)
		
		' do bWidth1
		x1 = XTemp
		x2 = XTemp
		y1 = YTemp
		y2 = -YTemp
		bWidth1.setCoords(x1, y1, z1, x2, y2, z2)
		
		' do bWidth2
		
		x1 = -XTemp
		x2 = -XTemp
		bWidth2.setCoords(x1, y1, z1, x2, y2, z2)
		
		' now we need to do the height thing
		' note that z1,x2,y2,z2 will not change for any of the lines
		x2 = 0
		y2 = 0
		z2 = Height
		z1 = 0
		
		' do L1W1
		x1 = XTemp
		y1 = YTemp
		L1W1.setCoords(x1, y1, z1, x2, y2, z2)
		
		' do l1w2
		x1 = -XTemp
		L1W2.setCoords(x1, y1, z1, x2, y2, z2)
		
		' do l2w2
		y1 = -YTemp
		L2W2.setCoords(x1, y1, z1, x2, y2, z2)
		
		'do l2w1
		x1 = XTemp
		L2W1.setCoords(x1, y1, z1, x2, y2, z2)
		
	End Sub

    Public Sub New()
		MyBase.New()
        bLength1 = New threeDLine
        bLength2 = New threeDLine
        bWidth1 = New threeDLine
        bWidth2 = New threeDLine
        L1W1 = New threeDLine
        L2W1 = New threeDLine
        L1W2 = New threeDLine
        L2W2 = New threeDLine
        xyz = New PltAxes
        testRect = New threeDRect
        testRect2 = New threeDRect
        testRect3 = New threeDRect

        testRect.createVerticalRect(50, 50, 200, 0, 100, 45, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red))
        testRect2.createVerticalRect(50, 50, 200, 0, 100, -45, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Lime))
        testRect3.createHorizontalRect(50, 50, 200, 0, 100, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))

    End Sub
	
	Public Function ZoomGraph(ByVal aChange As Short) As Boolean
		Dim zoomVar As Boolean
		zoomVar = theGraph.changeScale(aChange)
		If zoomVar = True Then
			drawPyramid()
		End If
		ZoomGraph = zoomVar
	End Function
End Class