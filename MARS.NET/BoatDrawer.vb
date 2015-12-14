Option Strict Off
Option Explicit On
Friend Class PltVessel
	
	Public aGrapher As threeDGrapher
	Private thePontoon As PltPontoon3d
	Private hBrace1 As PltBraces '...4th brace from front(rear)
	Private hBrace2 As PltBraces '...1st brace
	Private hBrace3 As PltBraces '...2nd brace
	Private hbrace4 As PltBraces '...3rd brace
	Private theDeck As PltDeck3d
	Private anAxis As PltAxes
	Private theAnchor As PltAnchor
	Private theLines As PltAnchorLines
	
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		theLines = New PltAnchorLines
		thePontoon = New PltPontoon3d
		aGrapher = New threeDGrapher
		anAxis = New PltAxes
		hBrace1 = New PltBraces
		hBrace2 = New PltBraces
		hBrace3 = New PltBraces
		hbrace4 = New PltBraces
		theDeck = New PltDeck3d
		theAnchor = New PltAnchor
		initHBraces()
		'initAnchor
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	Private Sub initAnchor()
		theAnchor.createAnchor(200, 200, -500)
		
	End Sub
	Private Sub initHBraces()
		Dim X0 As Single
		Dim Y0 As Single
		Dim z0 As Single
		Dim x1 As Single
		Dim y1 As Single
		Dim z1 As Single
		
		' do the 4th brace
		X0 = -PDistance / 2
		x1 = PDistance / 2
		Y0 = -(PontLength / 2 - PLength - PontWidth / 2)
		y1 = Y0
		z0 = PontHeight + hBraceRadius
		z1 = z0
		hBrace1.createBrace(X0, Y0, z0, x1, y1, z1)
		
		' do the 1st brace
		X0 = -PDistance / 2
		x1 = PDistance / 2
		Y0 = PontLength / 2 - PLength - PontWidth / 2
		y1 = Y0
		z0 = PontHeight + hBraceRadius
		z1 = z0
		
		hBrace2.createBrace(X0, Y0, z0, x1, y1, z1)
		
		' do the 2nd brace
		X0 = -PDistance / 2 - PontWidth - (PDistance - 22)
		x1 = -X0
		Y0 = PontLength / 2 - 141
		y1 = Y0
		
		hBrace3.createBrace(X0, Y0, z0, x1, y1, z1)
		
		' do the 3st brace
		Y0 = -Y0
		y1 = Y0

        hbrace4.createBrace(X0, Y0, z0, x1, y1, z1)
    End Sub
	
	Public Sub setGraph(ByRef gIn As System.Windows.Forms.PictureBox, ByRef X As Single, ByRef Y As Single, ByRef Incremental As Boolean)
		aGrapher.setDrawSurface(gIn, X, Y, Incremental)
	End Sub
	
	Public Sub drawBoat()
		Dim aline As threeDLine

        aGrapher.clearGraph()
        Dim Component As threeDObject
        aline = New threeDLine
		aline.setCoords(0, 0, 0, 0, 60, 0)
		aGrapher.drawTheLine(aline)
		theLines.drawAnchorLines(aGrapher)
        anAxis.drawAxis(aGrapher)

        If VesselPointsFromFile Then

            'For Each Component In CurProj.Vessel
            'For Each aline In Component
            'aGrapher.drawTheLine(aline)
            'Next aline
            'Next Component

        Else
            thePontoon.drawPontoon(aGrapher)
			hBrace1.drawBraces(aGrapher)
			hBrace2.drawBraces(aGrapher)
			hBrace3.drawBraces(aGrapher)
			hbrace4.drawBraces(aGrapher)
			theDeck.drawDeck(aGrapher)
		End If
		
	End Sub
	
	Public Sub angleChange(ByRef H As Short, ByRef V As Short)
		aGrapher.newAngles(H, V)
	End Sub
	
	Public Function ZoomGraph(ByVal aChange As Single) As Object
		Dim zoomVar As Boolean
		zoomVar = aGrapher.changeScale(aChange)
		If zoomVar = True Then
			drawBoat()
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object ZoomGraph. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ZoomGraph = zoomVar
	End Function
	
	Public Sub UpdateCatenary(ByVal FileName As String)
		theLines.ReadFile(FileName)
	End Sub
End Class