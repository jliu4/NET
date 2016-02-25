Option Strict Off
Option Explicit On
Friend Class PltOctagon3d
	' this class creates a box with octagon top and bottom
	
	
	Private Tlines(8) As threeDLine
	Private BLines(8) As threeDLine
	Private VLines(8) As threeDLine
	Public theLength As Single
	Public theHeight As Single
	Public X As Single
	Public Y As Single
	Public z As Single

    Public Sub New()
		MyBase.New()
        Dim X As Short
        X = 1
        Do While X <= 8
            Tlines(X) = New threeDLine
            BLines(X) = New threeDLine
            VLines(X) = New threeDLine
            X = X + 1
        Loop
    End Sub
	
	Public Sub createOctagon(ByRef Length As Single, ByRef Height As Single, ByRef xCenter As Single, ByRef ycenter As Single, ByRef zBase As Single)
		Dim x1 As Single
		Dim y1 As Single
		Dim zUpper As Single
		Dim x2 As Single
		Dim y2 As Single
		Dim zLower As Single
		theLength = Length
		theHeight = Height
		X = xCenter
		Y = ycenter
		z = zBase
		
		' zupper and zLower will not change so let's solve
		' for them now
		zLower = zBase
		zUpper = zBase + Height
		
		' let's do TLine1 and Bline1
		x1 = xCenter - Length / 2
		x2 = x1
		y1 = ycenter - (System.Math.Tan(45 / 2) * Length / 2)
		y2 = ycenter + (System.Math.Tan(45 / 2) * Length / 2)
		Tlines(1).setCoords(x1, y1, zUpper, x2, y2, zUpper)
		BLines(1).setCoords(x1, y1, zLower, x2, y2, zLower)
		VLines(1).setCoords(x1, y1, zLower, x1, y1, zUpper)
		VLines(2).setCoords(x2, y2, zLower, x2, y2, zUpper)
		
		' let's draw the opposite side
		x1 = xCenter + Length / 2
		x2 = x1
		Tlines(5).setCoords(x1, y1, zUpper, x2, y2, zUpper)
		BLines(5).setCoords(x1, y1, zLower, x2, y2, zLower)
		VLines(5).setCoords(x1, y1, zLower, x1, y1, zUpper)
		VLines(6).setCoords(x2, y2, zLower, x2, y2, zUpper)
		
		' let's draw line 2
		x1 = xCenter - Length / 2
		y1 = ycenter - (System.Math.Tan(45 / 2) * Length / 2)
		
		x2 = xCenter - (System.Math.Tan(45 / 2) * Length / 2)
		y2 = ycenter - Length / 2
		
		Tlines(2).setCoords(x1, y1, zUpper, x2, y2, zUpper)
		BLines(2).setCoords(x1, y1, zLower, x2, y2, zLower)
		VLines(3).setCoords(x2, y2, zLower, x2, y2, zUpper)
		
		'now draw line4
		
		x1 = x1 - xCenter
		x2 = x2 - xCenter
		x1 = -x1
		x2 = -x2
		x1 = x1 + xCenter
		x2 = x2 + xCenter
		Tlines(4).setCoords(x1, y1, zUpper, x2, y2, zUpper)
		BLines(4).setCoords(x1, y1, zLower, x2, y2, zLower)
		VLines(4).setCoords(x2, y2, zLower, x2, y2, zUpper)
		VLines(5).setCoords(x1, y1, zLower, x1, y1, zUpper)
		
		'let's do line 6
		
		y1 = y1 - ycenter
		y2 = y2 - ycenter
		y1 = -y1
		y2 = -y2
		y1 = y1 + ycenter
		y2 = y2 + ycenter
		
		Tlines(6).setCoords(x1, y1, zUpper, x2, y2, zUpper)
		BLines(6).setCoords(x1, y1, zLower, x2, y2, zLower)
		VLines(7).setCoords(x2, y2, zLower, x2, y2, zUpper)
		
		' now let's do line 7
		y1 = ycenter + Length / 2
		y2 = ycenter + Length / 2
		x1 = xCenter - (System.Math.Tan(45 / 2) * Length / 2)
		x2 = xCenter + (System.Math.Tan(45 / 2) * Length / 2)
		
		Tlines(7).setCoords(x1, y1, zUpper, x2, y2, zUpper)
		BLines(7).setCoords(x1, y1, zLower, x2, y2, zLower)
		VLines(8).setCoords(x1, y1, zLower, x1, y1, zUpper)
		
		' let's do line 3
		y1 = y1 - ycenter
		y2 = y2 - ycenter
		y1 = -y1
		y2 = -y2
		y1 = y1 + ycenter
		y2 = y2 + ycenter
		Tlines(3).setCoords(x1, y1, zUpper, x2, y2, zUpper)
		BLines(3).setCoords(x1, y1, zLower, x2, y2, zLower)
		
		' let's do line 8
		y1 = y1 - ycenter
		y1 = -y1
		y1 = y1 + ycenter
		
		x2 = xCenter - Length / 2
		y2 = ycenter + (System.Math.Tan(45 / 2) * Length / 2)
		Tlines(8).setCoords(x1, y1, zUpper, x2, y2, zUpper)
		BLines(8).setCoords(x1, y1, zLower, x2, y2, zLower)

    End Sub
	
	Public Sub getLines(ByRef TLine As threeDLine, ByRef BLine As threeDLine, ByRef VLine As threeDLine, ByRef Index As Short)
		TLine = Tlines(Index)
		BLine = BLines(Index)
		VLine = VLines(Index)
	End Sub
End Class