Option Strict Off
Option Explicit On
Friend Class PltAnchor
	
	' this class creates an Stevpris Anchor
	
	' length is as seen from a side view
	Private Const anchorLength As Single = 40
	
	' width is as seen from a overhead view, the back
	' part
	Private Const anchorWidth As Single = 30
	
	' estimated as a ratio of Height to Length
	Private Const anchorHeight As Single = 3 / 6 * 40
	
	' estimated length of the base of the anchor
	Private Const BaseLength As Single = 3.28 * 10 '6.294
	
	' estimation of the angle between the fins and the base
	Private Const BaseFinAngle As Single = 31
	
	' these tell us the coordinates of the hook, i.e. where
	' the line connects to the anchor
	
	Private rFinBottom As threeDLine
	Private rFinTop As threeDLine
	Private lFinBottom As threeDLine
	Private lFinTop As threeDLine
	Private lInsideBase As threeDLine
	Private rInsideBase As threeDLine
	Private backBase As threeDLine
	Private rSideBase As threeDLine
	Private lSideBase As threeDLine
	Private notch As threeDRect
	Private xCenter As Single
	Private ycenter As Single
	Private zCenter As Single
	
	Public Sub createAnchor(ByRef X As Single, ByRef Y As Single, ByRef z As Single) ', angle As Single)
		
		' let's calculate the angle from the x,y,z coords.
		Dim angle As Single
		Dim radAngle As Single
		Dim calc As Boolean
		calc = False
		If (X = 0) Or (Y = 0) Then
			calc = True
			If (X = 0) And (Y < 0) Then
				angle = 360 - 90
			ElseIf (X = 0) And (Y > 0) Then 
				angle = 90
				
			ElseIf (Y = 0) And (X > 0) Then 
				angle = 0
			ElseIf (Y = 0) And (X < 0) Then 
				angle = 180
			End If
		End If

        If Not calc = True Then
            angle = System.Math.Atan(Y / X) * 180 / PI
            If (X < 0) Then
                angle = angle + 180
            End If
        End If

        ' first let's calculate the top of the base

        radAngle = PI / 180 * angle
		
		Dim xBase As Single
		Dim yBase As Single
		Dim zBase As Single
		
		xBase = X + anchorLength * System.Math.Cos(radAngle)
		yBase = Y + anchorLength * System.Math.Sin(radAngle)
		zBase = z

        ' draw the back of the base.
        ' z will be the same for the two points
        Dim X0 As Single
		Dim Y0 As Single
		Dim z0 As Single
		Dim x1 As Single
		Dim y1 As Single
		Dim z1 As Single
		
		z0 = zBase
		z1 = zBase
		
		X0 = xBase - anchorWidth / 2 * System.Math.Sin(radAngle)
		Y0 = yBase + anchorWidth / 2 * System.Math.Cos(radAngle)
		
		x1 = xBase + anchorWidth / 2 * System.Math.Sin(radAngle)
		y1 = yBase - anchorWidth / 2 * System.Math.Cos(radAngle)
		
		backBase.setCoords(X0, Y0, z0, x1, y1, z1)
		
		' calculate rSideBase
		X0 = xBase - anchorWidth / 2 * System.Math.Sin(radAngle)
		Y0 = yBase + anchorWidth / 2 * System.Math.Cos(radAngle)
		z0 = zBase
		
		x1 = xBase - BaseLength * System.Math.Cos(BaseFinAngle * PI / 180) * System.Math.Cos(radAngle)
		y1 = yBase - BaseLength * System.Math.Cos(BaseFinAngle * PI / 180) * System.Math.Sin(radAngle)
		z1 = zBase - BaseLength * System.Math.Sin(BaseFinAngle * PI / 180)
		
		x1 = x1 - BaseLength / 4 * System.Math.Sin(radAngle)
		y1 = y1 + BaseLength / 4 * System.Math.Cos(radAngle)
		
		rSideBase.setCoords(X0, Y0, z0, x1, y1, z1)
		
		' let's do the right inside base
		' in this case x1,y1, and z1 will be same since they
		' are connecting points
		
		X0 = xBase - BaseLength / 2 * System.Math.Cos(BaseFinAngle * PI / 180) * System.Math.Cos(radAngle)
		Y0 = yBase - BaseLength / 2 * System.Math.Cos(BaseFinAngle * PI / 180) * System.Math.Sin(radAngle)
		
		z0 = zBase - BaseLength / 2 * System.Math.Sin(BaseFinAngle * PI / 180)
		
		rInsideBase.setCoords(X0, Y0, z0, x1, y1, z1)
		
		' so the left inside base.
		' x0,y0,z0 will be the same!
		
		x1 = xBase - BaseLength * System.Math.Cos(BaseFinAngle * PI / 180) * System.Math.Cos(radAngle)
		y1 = yBase - BaseLength * System.Math.Cos(BaseFinAngle * PI / 180) * System.Math.Sin(radAngle)
		z1 = zBase - BaseLength * System.Math.Sin(BaseFinAngle * PI / 180)
		
		x1 = x1 + BaseLength / 4 * System.Math.Sin(radAngle)
		y1 = y1 - BaseLength / 4 * System.Math.Cos(radAngle)
		
		lInsideBase.setCoords(X0, Y0, z0, x1, y1, z1)
		
		' now let's do the left side base
		' x1,y1,z1 will be the same!
		X0 = xBase + anchorWidth / 2 * System.Math.Sin(radAngle)
		Y0 = yBase - anchorWidth / 2 * System.Math.Cos(radAngle)
		z0 = zBase
		
		lSideBase.setCoords(X0, Y0, z0, x1, y1, z1)
		
		' now lets do the top right fin
		Dim xFin As Single
		Dim yFin As Single
		Dim zFin As Single
		
		zFin = z
		xFin = X '+ 2 * Cos(radAngle)
		yFin = Y '+ 2 * Sin(radAngle)
		
		X0 = xFin
		Y0 = yFin
		z0 = zFin
		
		z1 = zFin
		x1 = X + anchorLength * System.Math.Cos(radAngle)
		y1 = Y + anchorLength * System.Math.Sin(radAngle)
		x1 = x1 - 1 / 4 * anchorWidth * System.Math.Sin(radAngle)
		y1 = y1 + 1 / 4 * anchorWidth * System.Math.Cos(radAngle)
		
		rFinTop.setCoords(X0, Y0, z0, x1, y1, z1)

        ' do the bottom right fin
        x1 = X + System.Math.Cos(radAngle) * anchorLength
		y1 = Y + System.Math.Sin(radAngle) * anchorLength
		x1 = x1 - (BaseLength / 2 * System.Math.Cos(BaseFinAngle) * System.Math.Cos(radAngle))
		y1 = y1 - (BaseLength / 2 * System.Math.Cos(BaseFinAngle) * System.Math.Sin(radAngle))
		x1 = x1 - anchorWidth / 4 * System.Math.Sin(radAngle)
		y1 = y1 + anchorWidth / 4 * System.Math.Cos(radAngle)
		z1 = z - BaseLength / 2 * System.Math.Sin(BaseFinAngle * PI / 180)
		
		rFinBottom.setCoords(X0, Y0, z0, x1, y1, z1)
		
		'do the top left fin
		z1 = zFin
		x1 = X + anchorLength * System.Math.Cos(radAngle)
		y1 = Y + anchorLength * System.Math.Sin(radAngle)
		x1 = x1 + 1 / 4 * anchorWidth * System.Math.Sin(radAngle)
		y1 = y1 - 1 / 4 * anchorWidth * System.Math.Cos(radAngle)
		lFinTop.setCoords(X0, Y0, z0, x1, y1, z1)
		
		' do the bottom left fin
		x1 = X + System.Math.Cos(radAngle) * anchorLength
		y1 = Y + System.Math.Sin(radAngle) * anchorLength
		x1 = x1 - (BaseLength / 2 * System.Math.Cos(BaseFinAngle) * System.Math.Cos(radAngle))
		y1 = y1 - (BaseLength / 2 * System.Math.Cos(BaseFinAngle) * System.Math.Sin(radAngle))
		x1 = x1 + anchorWidth / 4 * System.Math.Sin(radAngle)
		y1 = y1 - anchorWidth / 4 * System.Math.Cos(radAngle)
		z1 = z - BaseLength / 2 * System.Math.Sin(BaseFinAngle * PI / 180)
		
		lFinBottom.setCoords(X0, Y0, z0, x1, y1, z1)
		
		' do the notch!
		
		
	End Sub

    Public Sub New()
        MyBase.New()
        backBase = New threeDLine
        rSideBase = New threeDLine
        rInsideBase = New threeDLine
        lInsideBase = New threeDLine
        lSideBase = New threeDLine
        lFinTop = New threeDLine
        lFinBottom = New threeDLine
        rFinTop = New threeDLine
        rFinBottom = New threeDLine
        notch = New threeDRect


    End Sub
	
	Public Sub drawAnchor(ByRef gIn As threeDGrapher)
		gIn.drawTheLine(backBase)
		gIn.drawTheLine(rSideBase)
		gIn.drawTheLine(rInsideBase)
		gIn.drawTheLine(lInsideBase)
		gIn.drawTheLine(lSideBase)
		gIn.drawTheLine(rFinTop)
		gIn.drawTheLine(rFinBottom)
		gIn.drawTheLine(lFinTop)
		gIn.drawTheLine(lFinBottom)
		
	End Sub
End Class