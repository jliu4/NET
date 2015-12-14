Option Strict Off
Option Explicit On
Friend Class threeDRect
	
	' this is a class used to hold data for a rectangle
	Public Length1 As threeDLine
	Public Length2 As threeDLine
	Public Width1 As threeDLine
	Public Width2 As threeDLine
	Public xCoord As Single
	Public yCoord As Single
	Public zCoord As Single
	Public fillColor As Integer
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		Length1 = New threeDLine
		Length2 = New threeDLine
		Width1 = New threeDLine
		Width2 = New threeDLine
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	' creates lines for a rectangle parallel to the xy plane
	Public Sub createHorizontalRect(ByVal L As Single, ByVal w As Single, ByVal xCenter As Single, ByVal ycenter As Single, ByVal zCenter As Single, ByVal rcolor As Integer, Optional ByRef fillMe As Integer = 0)
		
		' note that z coordinate will remain constant
		Dim X0 As Single
		Dim Y0 As Single
		Dim z0 As Single
		Dim x1 As Single
		Dim y1 As Single
		Dim z1 As Single
		z0 = zCenter
		z1 = zCenter
		xCoord = xCenter
		yCoord = ycenter
		zCoord = zCenter
		ycenter = -ycenter
		
		
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(fillMe) Then
			fillColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
		Else
			fillColor = fillMe
		End If
		
		' length here is the line parallel to the xaxis, w is
		' parallel to the y axis
		
		' L1
		X0 = xCenter - L / 2
		Y0 = ycenter - w / 2
		x1 = xCenter + L / 2
		y1 = ycenter - w / 2
		
		Length1.setCoords(X0, Y0, z0, x1, y1, z1)
		
		'W1
		x1 = X0
		y1 = ycenter + w / 2
		
		Width1.setCoords(X0, Y0, z0, x1, y1, z1)
		
		'l2
		X0 = x1
		Y0 = y1
		x1 = xCenter + L / 2
		
		Length2.setCoords(X0, Y0, z0, x1, y1, z1)
		
		'W2
		X0 = x1
		Y0 = ycenter - w / 2
		
		Width2.setCoords(X0, Y0, z0, x1, y1, z1)
		Length1.setColor(rcolor)
		Length2.setColor(rcolor)
		Width1.setColor(rcolor)
		Width2.setColor(rcolor)
		
		
	End Sub
	
	
	
	' creates lines for a rectangle given an (x,y,z) center.
	' L is length which should be line parallel to xy plane.
	' height is parallel along z-plane
	' xDeg is the angle, in degrees, that the rectangle makes
	' with the x axis.
	' xzDeg is the angle, in degrees, that the rectangle makes
	' with the xz plane.
	
	Public Sub createVerticalRect(ByVal L As Single, ByVal w As Single, ByVal xCenter As Single, ByVal ycenter As Single, ByVal zCenter As Single, ByVal xDeg As Object, ByVal rcolor As Integer, Optional ByRef fillMe As Integer = 0)
		'UPGRADE_WARNING: Couldn't resolve default property of object xDeg. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xDeg = -xDeg
		Dim xRads As Single
		
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(fillMe) Then
			fillColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
		Else
			fillColor = fillMe
		End If
		
		' convert degrees to radians
		'UPGRADE_WARNING: Couldn't resolve default property of object xDeg. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xRads = xDeg * PI / 180
		ycenter = -ycenter
		rotateRect(xCenter, ycenter, zCenter, xRads, L, w, Length1, Length2, Width1, Width2)
		
		Length1.setColor(rcolor)
		Length2.setColor(rcolor)
		Width1.setColor(rcolor)
		Width2.setColor(rcolor)
		
	End Sub
	
	' method rotates angle about the xy plane.  Assumes the
	' rectangle is vertical
	
	Private Sub rotateRect(ByRef xCenter As Single, ByRef ycenter As Single, ByRef zCenter As Single, ByRef xRads As Single, ByRef L As Single, ByRef w As Single, ByRef L1 As threeDLine, ByRef L2 As threeDLine, ByRef W1 As threeDLine, ByRef W2 As threeDLine)
		Dim X0 As Single
		Dim Y0 As Single
		Dim z0 As Single
		Dim x1 As Single
		Dim y1 As Single
		Dim z1 As Single
		
		' calculate L1
		X0 = xCenter - System.Math.Cos(xRads) * L / 2
		Y0 = ycenter - System.Math.Sin(xRads) * L / 2
		z0 = zCenter - w / 2
		
		x1 = xCenter + System.Math.Cos(xRads) * L / 2
		y1 = ycenter + System.Math.Sin(xRads) * L / 2
		z1 = z0
		
		L1.setCoords(X0, Y0, z0, x1, y1, z1)
		
		' calculate w1.
		x1 = X0
		y1 = Y0
		z1 = zCenter + w / 2
		
		W1.setCoords(X0, Y0, z0, x1, y1, z1)
		
		' calc. L2
		X0 = x1
		Y0 = y1
		z0 = z1
		
		x1 = xCenter + System.Math.Cos(xRads) * L / 2
		y1 = ycenter + System.Math.Sin(xRads) * L / 2
		
		L2.setCoords(X0, Y0, z0, x1, y1, z1)
		
		' calc w2.  Note x1,y1,z1 are the same
		z0 = zCenter - w / 2
		X0 = x1
		Y0 = y1
		
		W2.setCoords(X0, Y0, z0, x1, y1, z1)
		
	End Sub
	
	' creates a triangle parallel to the xy plane with the
	' coordinates of the 90 degree angle at (x,y,z)
	
	Public Sub createTriangle(ByVal L As Single, ByVal w As Single, ByVal X As Single, ByVal Y As Single, ByVal Z As Single, ByVal angleOffset As Single, ByVal lColor As Integer, Optional ByRef fColor As Integer = 0)
		Y = -Y
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(fColor) Then
			fillColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
		Else
			fillColor = fColor
		End If
		
		' calculate the center of the triangle. Consider the center
		' to be the center of the hypotenus
		Dim radAngle As Single
		radAngle = -angleOffset * PI / 180
		zCoord = Z
		xCoord = X
		yCoord = Y
		
		Dim X0 As Single
		Dim x1 As Single
		Dim x2 As Single
		Dim Y0 As Single
		Dim y1 As Single
		Dim y2 As Single
		
		' calculate L1 which is parallel to the x axis before
		' angle
		X0 = X
		Y0 = Y
		x1 = X0 + System.Math.Cos(radAngle) * L
		y1 = Y0 - System.Math.Sin(radAngle) * L
		
		
		Length1.setCoords(X0, Y0, Z, x1, y1, Z)
		
		x2 = X0 + System.Math.Sin(radAngle) * w
		y2 = Y0 + System.Math.Cos(radAngle) * w
		
		
		Width1.setCoords(X0, Y0, Z, x2, y2, Z)
		
		Length2.setCoords(x1, y1, Z, x2, y2, Z)
		
		Length1.setColor(lColor)
		Length2.setColor(lColor)
		Width1.setColor(lColor)
	End Sub
End Class