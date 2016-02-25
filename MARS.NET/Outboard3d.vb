Option Strict Off
Option Explicit On
Friend Class PltOutboard3d
	' class to definte the things that stick out on the
	' side of the pontoon
	
	Private Structure anOutBoard
		Dim side1 As threeDRect
		Dim side2 As threeDRect
		Dim front As threeDRect
		Dim bSide1 As threeDRect
		Dim tSide1 As threeDRect
		Dim bside2 As threeDRect
		Dim tside2 As threeDRect
		Dim Bottom As threeDRect
		Dim Top As threeDRect
	End Structure
	
	Private Structure outboardData
		Dim Length As Single
		Dim Height As Single
		Dim Depth As Single
		Dim sideProtusion As Single
	End Structure
	
	Private Structure aPontoon
		Dim OBack As anOutBoard
		Dim OMiddle As anOutBoard
		Dim OFront As anOutBoard
		Dim OLong As anOutBoard
	End Structure
	
	' note that the outboards will differ by a negative sign
	' in the angle and xcenter coord for the right and left
	' pontoon
	
	Private outBack As outboardData
	Private outFront As outboardData
	Private middle As outboardData
	Private longOne As outboardData
	Private lPontoon As aPontoon
	Private rPontoon As aPontoon

    Public Sub New()
		MyBase.New()
        outBack.Depth = 2.16 * 3.28
        outBack.Height = 5 * 3.28
        outBack.Length = 9 * 3.28
        outBack.sideProtusion = 0.4 * 16.9

        outFront.Depth = 2.16 * 3.28
        outFront.Height = 5 * 3.28
        outFront.Length = 9 * 3.28
        outFront.sideProtusion = 0.4 * 16.9

        middle.Depth = 2.16 * 3.28
        middle.Height = 5 * 3.28
        middle.Length = 52 * 3.28
        middle.sideProtusion = 0.4 * 16.9

        longOne.Depth = 2.16 * 3.28
        longOne.Height = 5 * 3.28
        longOne.Length = 96 * 3.28
        longOne.sideProtusion = 5 * 3.25
        initpontoons()
        createOutboards()

    End Sub
	
	Private Sub initpontoons()
		lPontoon.OBack.side1 = New threeDRect
		lPontoon.OBack.side2 = New threeDRect
		lPontoon.OBack.front = New threeDRect
		lPontoon.OBack.bSide1 = New threeDRect
		lPontoon.OBack.tSide1 = New threeDRect
		lPontoon.OBack.Bottom = New threeDRect
		lPontoon.OBack.Top = New threeDRect
		lPontoon.OBack.bside2 = New threeDRect
		lPontoon.OBack.tside2 = New threeDRect
		
		lPontoon.OFront.side1 = New threeDRect
		lPontoon.OFront.side2 = New threeDRect
		lPontoon.OFront.front = New threeDRect
		lPontoon.OFront.bSide1 = New threeDRect
		lPontoon.OFront.tSide1 = New threeDRect
		lPontoon.OFront.Top = New threeDRect
		lPontoon.OFront.Bottom = New threeDRect
		lPontoon.OFront.bside2 = New threeDRect
		lPontoon.OFront.tside2 = New threeDRect
		
		lPontoon.OLong.side1 = New threeDRect
		lPontoon.OLong.side2 = New threeDRect
		lPontoon.OLong.front = New threeDRect
		lPontoon.OLong.bSide1 = New threeDRect
		lPontoon.OLong.tSide1 = New threeDRect
		lPontoon.OLong.Top = New threeDRect
		lPontoon.OLong.Bottom = New threeDRect
		lPontoon.OLong.bside2 = New threeDRect
		lPontoon.OLong.tside2 = New threeDRect
		
		lPontoon.OMiddle.front = New threeDRect
		lPontoon.OMiddle.side1 = New threeDRect
		lPontoon.OMiddle.side2 = New threeDRect
		lPontoon.OMiddle.Bottom = New threeDRect
		lPontoon.OMiddle.Top = New threeDRect
		lPontoon.OMiddle.bSide1 = New threeDRect
		lPontoon.OMiddle.tSide1 = New threeDRect
		lPontoon.OMiddle.tside2 = New threeDRect
		lPontoon.OMiddle.bside2 = New threeDRect
		
		rPontoon.OBack.front = New threeDRect
		rPontoon.OBack.side1 = New threeDRect
		rPontoon.OBack.side2 = New threeDRect
		rPontoon.OBack.Bottom = New threeDRect
		rPontoon.OBack.Top = New threeDRect
		rPontoon.OBack.bSide1 = New threeDRect
		rPontoon.OBack.tSide1 = New threeDRect
		rPontoon.OBack.tside2 = New threeDRect
		rPontoon.OBack.bside2 = New threeDRect
		
		rPontoon.OFront.front = New threeDRect
		rPontoon.OFront.side1 = New threeDRect
		rPontoon.OFront.side2 = New threeDRect
		rPontoon.OFront.Bottom = New threeDRect
		rPontoon.OFront.Top = New threeDRect
		rPontoon.OFront.bSide1 = New threeDRect
		rPontoon.OFront.tSide1 = New threeDRect
		rPontoon.OFront.bside2 = New threeDRect
		rPontoon.OFront.tside2 = New threeDRect
		
		rPontoon.OLong.front = New threeDRect
		rPontoon.OLong.side1 = New threeDRect
		rPontoon.OLong.side2 = New threeDRect
		rPontoon.OLong.Bottom = New threeDRect
		rPontoon.OLong.Top = New threeDRect
		rPontoon.OLong.bSide1 = New threeDRect
		rPontoon.OLong.tSide1 = New threeDRect
		rPontoon.OLong.bside2 = New threeDRect
		rPontoon.OLong.tside2 = New threeDRect
		
		rPontoon.OMiddle.front = New threeDRect
		rPontoon.OMiddle.side1 = New threeDRect
		rPontoon.OMiddle.side2 = New threeDRect
		rPontoon.OMiddle.Bottom = New threeDRect
		rPontoon.OMiddle.Top = New threeDRect
		rPontoon.OMiddle.bSide1 = New threeDRect
		rPontoon.OMiddle.tSide1 = New threeDRect
		rPontoon.OMiddle.bside2 = New threeDRect
		rPontoon.OMiddle.tside2 = New threeDRect
	End Sub
	
	Private Sub createOutboards()

        Dim xCenter1 As Single
		Dim xCenter2 As Single
		Dim ycenter As Single
		Dim zCenter As Single
		Dim angle1 As Single
		Dim angle2 As Single
		Dim radAngle As Single
        Dim tangle As Single
        Dim tX As Single
		Dim tY As Single
		Dim tZ As Single
		Dim tL As Single
		Dim tW As Single
		
		zCenter = PontHeight / 2
		' create the rectangles that make up the back pontoon
		
		' create P1 for the back outboard
		radAngle = System.Math.Atan(outBack.Depth / outBack.sideProtusion)
		angle1 = radAngle * 180 / PI
		angle2 = radAngle * 180 / PI
		angle1 = 90 + angle1
		angle2 = 90 - angle2
		
		' calc. the centers for both sides.  the ycenters
		' will be the same
		
		ycenter = -PontLength / 2
		ycenter = ycenter + PLength
		ycenter = ycenter + outBack.sideProtusion / 2
		
		xCenter1 = -PDistance / 2
		xCenter1 = xCenter1 - PontWidth
		xCenter1 = xCenter1 - outBack.Depth / 2
		xCenter2 = -xCenter1
		
		Dim Length As Single
		Dim width As Single
		Length = outBack.Depth / System.Math.Sin(radAngle)
		width = outBack.Height
		lPontoon.OBack.side1.createVerticalRect(Length, width, xCenter1, ycenter, zCenter, angle1, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OBack.side1.createVerticalRect(Length, width, xCenter2, ycenter, zCenter, angle2, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' create top and bottom triangles of the back outboard.
		tX = -PDistance / 2
		tX = tX - PontWidth
		tY = -PontLength / 2
		tY = tY + PLength
		tY = tY + outBack.sideProtusion
		'tY = tY + outBack.sideProtusion
		tZ = PontHeight / 2
		tZ = tZ - outBack.Height / 2
		
		tangle = 90
		tL = outBack.Depth
		tW = outBack.sideProtusion
		lPontoon.OBack.bSide1.createTriangle(tL, tW, tX, tY, tZ, tangle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OBack.bSide1.createTriangle(tW, tL, -tX, tY, tZ, tangle - 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' the top triangles
		
		tZ = tZ + outBack.Height
		
		lPontoon.OBack.tSide1.createTriangle(tL, tW, tX, tY, tZ, tangle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OBack.tSide1.createTriangle(tL, tW, -tX, tY, tZ, tangle - 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' now let's create the rectanges at top and bottom
		tX = -PDistance / 2
		tX = tX - PontWidth
		tX = tX - outBack.Depth / 2
		tY = -PontLength / 2
		tY = tY + PLength
		tY = tY + outBack.sideProtusion
		tY = tY + outBack.Length / 2
		tZ = PontHeight / 2
		tZ = tZ - outBack.Height / 2
		tL = outBack.Depth
		tW = outBack.Length
		
		lPontoon.OBack.Bottom.createHorizontalRect(tL, tW, tX, tY, tZ, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OBack.Bottom.createHorizontalRect(tL, tW, -tX, tY, tZ, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' draw the top
		tZ = tZ + outBack.Height
		
		lPontoon.OBack.Top.createHorizontalRect(tL, tW, tX, tY, tZ, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OBack.Top.createHorizontalRect(tL, tW, -tX, tY, tZ, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		
		' draw front panel of the back outboard
		xCenter1 = -PDistance / 2
		xCenter1 = xCenter1 - PontWidth
		xCenter1 = xCenter1 - outBack.Depth
		xCenter2 = -xCenter1
		
		ycenter = -PontLength / 2
		ycenter = ycenter + PLength
		ycenter = ycenter + outBack.sideProtusion
		ycenter = ycenter + outBack.Length / 2
		
		Length = outBack.Length
		width = outBack.Height
		
		lPontoon.OBack.front.createVerticalRect(Length, width, xCenter1, ycenter, zCenter, 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OBack.front.createVerticalRect(Length, width, xCenter2, ycenter, zCenter, 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' draw side2 of the back outboard
		radAngle = System.Math.Atan(outBack.Depth / outBack.sideProtusion)
		angle1 = radAngle * 180 / PI
		angle2 = radAngle * 180 / PI
		angle1 = -angle1 - 90
		angle2 = angle2 - 90
		
		Length = outBack.Depth / System.Math.Sin(radAngle)
		width = outBack.Height
		
		xCenter1 = -PDistance / 2
		xCenter1 = xCenter1 - PontWidth
		xCenter1 = xCenter1 - outBack.Depth / 2
		xCenter2 = -xCenter1
		
		ycenter = -PontLength / 2
		ycenter = ycenter + PLength
		ycenter = ycenter + outBack.sideProtusion
		ycenter = ycenter + outBack.Length
		ycenter = ycenter + outBack.sideProtusion / 2
		
		lPontoon.OBack.side2.createVerticalRect(Length, width, xCenter1, ycenter, zCenter, angle1, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OBack.side2.createVerticalRect(Length, width, xCenter2, ycenter, zCenter, angle2, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' we need to draw the little top and bottom triangles
		tX = -PDistance / 2
		tX = tX - PontWidth
		
		tY = -PontLength / 2
		tY = tY + PLength
		tY = tY + outBack.sideProtusion
		tY = tY + outBack.Length
		
		tZ = PontHeight / 2
		tZ = tZ - outBack.Height / 2
		
		tangle = 180
		
		tL = outBack.Depth
		tW = outBack.sideProtusion
		
		lPontoon.OBack.bside2.createTriangle(tL, tW, tX, tY, tZ, tangle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OBack.bside2.createTriangle(tW, tL, -tX, tY, tZ, tangle + 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' draw the top one
		tZ = tZ + outBack.Height
		
		lPontoon.OBack.tside2.createTriangle(tL, tW, tX, tY, tZ, tangle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OBack.tside2.createTriangle(tW, tL, -tX, tY, tZ, tangle + 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' we have now finished drawing the back outboards.
		
		' we now need to draw the middle outboard
		
		' middle outboard, rear panel
		radAngle = System.Math.Atan(middle.Depth / middle.sideProtusion)
		angle1 = radAngle * 180 / PI
		angle2 = radAngle * 180 / PI
		
		angle1 = 90 + angle1
		angle2 = 90 - angle2
		
		xCenter1 = -PDistance / 2
		xCenter1 = xCenter1 - PontWidth
		xCenter1 = xCenter1 - middle.Depth / 2
		xCenter2 = -xCenter1
		
		ycenter = -middle.Length / 2
		ycenter = ycenter - middle.sideProtusion / 2
		
		Length = middle.Depth / System.Math.Sin(radAngle)
		width = middle.Height
		
		lPontoon.OMiddle.side1.createVerticalRect(Length, width, xCenter1, ycenter, zCenter, angle1, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OMiddle.side1.createVerticalRect(Length, width, xCenter2, ycenter, zCenter, angle2, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' draw the top and bottom triangles
		tZ = PontHeight / 2
		tZ = tZ - middle.Height / 2
		
		tX = -PDistance / 2
		tX = tX - PontWidth
		
		tY = -middle.Length / 2
		tangle = 90
		
		tW = middle.sideProtusion
		tL = middle.Depth
		
		lPontoon.OMiddle.bSide1.createTriangle(tL, tW, tX, tY, tZ, tangle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OMiddle.bSide1.createTriangle(tW, tL, -tX, tY, tZ, tangle - 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' draw the top ones
		
		tZ = tZ + middle.Height
		
		lPontoon.OMiddle.tSide1.createTriangle(tL, tW, tX, tY, tZ, tangle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OMiddle.tSide1.createTriangle(tW, tL, -tX, tY, tZ, tangle - 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		
		' draw the font panel of the middle outboard
		xCenter1 = -PDistance / 2
		xCenter1 = xCenter1 - PontWidth
		xCenter1 = xCenter1 - middle.Depth
		xCenter2 = -xCenter1
		
		ycenter = 0
		Length = middle.Length
		width = middle.Height
		
		lPontoon.OMiddle.front.createVerticalRect(Length, width, xCenter1, ycenter, zCenter, 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OMiddle.front.createVerticalRect(Length, width, xCenter2, ycenter, zCenter, 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' draw the rectangles that make up the upper and lower
		' part of the outboard
		
		tX = -PDistance / 2
		tX = tX - PontWidth
		tX = tX - middle.Depth / 2
		
		tY = 0
		
		tZ = PontHeight / 2
		tZ = tZ - middle.Height / 2
		
		tL = middle.Depth
		tW = middle.Length
		
		lPontoon.OMiddle.Bottom.createHorizontalRect(tL, tW, tX, tY, tZ, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OMiddle.Bottom.createHorizontalRect(tL, tW, -tX, tY, tZ, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' top rectangle
		tZ = tZ + middle.Height
		lPontoon.OMiddle.Top.createHorizontalRect(tL, tW, tX, tY, tZ, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OMiddle.Top.createHorizontalRect(tL, tW, -tX, tY, tZ, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		
		
		' now let's draw side2 of the middle outboard
		xCenter1 = -PDistance / 2
		xCenter1 = xCenter1 - PontWidth
		xCenter1 = xCenter1 - middle.Depth / 2
		xCenter2 = -xCenter1
		
		ycenter = middle.Length / 2
		ycenter = ycenter + middle.sideProtusion / 2
		
		radAngle = System.Math.Atan(middle.Depth / middle.sideProtusion)
		angle1 = radAngle * 180 / PI
		angle2 = radAngle * 180 / PI
		angle1 = -angle1 - 90
		angle2 = angle2 - 90
		
		Length = middle.Depth / System.Math.Sin(radAngle)
		width = middle.Height
		
		lPontoon.OMiddle.side2.createVerticalRect(Length, width, xCenter1, ycenter, zCenter, angle1, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OMiddle.side2.createVerticalRect(Length, width, xCenter2, ycenter, zCenter, angle2, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' let's draw the top and bottom triangles (this part
		' bites)
		
		tX = -PDistance / 2
		tX = tX - PontWidth
		tY = middle.Length / 2
		tZ = PontHeight / 2
		tZ = tZ - middle.Height / 2
		
		tL = middle.sideProtusion
		tW = middle.Depth
		
		tangle = 180
		
		lPontoon.OMiddle.bside2.createTriangle(tL, tW, tX, tY, tZ, tangle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OMiddle.bside2.createTriangle(tW, tL, -tX, tY, tZ, tangle + 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' now let's draw the tops
		tZ = tZ + middle.Height
		
		lPontoon.OMiddle.tside2.createTriangle(tL, tW, tX, tY, tZ, tangle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OMiddle.tside2.createTriangle(tW, tL, -tX, tY, tZ, tangle + 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		
		' we are now finished with the middle outboard
		
		' let's do the front outboard
		
		' the rear panel of the front outboard
		xCenter1 = -PDistance / 2
		xCenter1 = xCenter1 - PontWidth
		xCenter1 = xCenter1 - outFront.Depth / 2
		xCenter2 = -xCenter1
		
		ycenter = PontLength / 2
		ycenter = ycenter - PLength
		
		ycenter = ycenter - outFront.sideProtusion
		ycenter = ycenter - outFront.Length
		ycenter = ycenter - outFront.sideProtusion / 2
		
		radAngle = System.Math.Atan(outFront.Depth / outFront.sideProtusion)
		angle1 = radAngle * 180 / PI
		angle2 = radAngle * 180 / PI
		angle1 = 90 + angle1
		angle2 = 90 - angle2
		
		Length = outFront.Depth / System.Math.Sin(radAngle)
		width = outFront.Height
		
		lPontoon.OFront.side1.createVerticalRect(Length, width, xCenter1, ycenter, zCenter, angle1, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OFront.side1.createVerticalRect(Length, width, xCenter2, ycenter, zCenter, angle2, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' let's create the triangles for the upper and lower shit
		
		tX = -PDistance / 2
		tX = tX - PontWidth
		tY = PontLength / 2
		tY = tY - PLength
		tY = tY - outFront.sideProtusion
		tY = tY - outFront.Length
		tZ = PontHeight / 2
		tZ = tZ - outFront.Height / 2
		
		tangle = 90
		tL = outFront.Depth
		tW = outFront.sideProtusion
		
		lPontoon.OFront.bSide1.createTriangle(tL, tW, tX, tY, tZ, tangle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OFront.bSide1.createTriangle(tW, tL, -tX, tY, tZ, tangle - 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' draw the top ones
		tZ = tZ + outFront.Height
		
		lPontoon.OFront.tSide1.createTriangle(tL, tW, tX, tY, tZ, tangle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OFront.tSide1.createTriangle(tW, tL, -tX, tY, tZ, tangle - 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
        ' now let's do the middle panel of the front outboard
        xCenter1 = -PDistance / 2
		xCenter1 = xCenter1 - PontWidth
		xCenter1 = xCenter1 - outFront.Depth
		xCenter2 = -xCenter1
		
		ycenter = PontLength / 2
		ycenter = ycenter - PLength
		ycenter = ycenter - outFront.sideProtusion
		ycenter = ycenter - outFront.Length / 2
		
		Length = outFront.Length
		width = outFront.Height
		
		lPontoon.OFront.front.createVerticalRect(Length, width, xCenter1, ycenter, zCenter, 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OFront.front.createVerticalRect(Length, width, xCenter2, ycenter, zCenter, 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' lets draw the upper and lower rectangles
		tX = -PDistance / 2
		tX = tX - PontWidth
		tX = tX - outFront.Depth / 2
		
		tY = PontLength / 2
		tY = tY - PLength
		tY = tY - outFront.sideProtusion
		tY = tY - outFront.Length / 2
		
		tZ = PontHeight / 2
		tZ = tZ - outFront.Height / 2
		
		tL = outFront.Depth
		tW = outFront.Length
		
		lPontoon.OFront.Bottom.createHorizontalRect(tL, tW, tX, tY, tZ, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OFront.Bottom.createHorizontalRect(tL, tW, -tX, tY, tZ, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' now let's draw the upper rects
		tZ = tZ + outFront.Height
		
		lPontoon.OFront.Top.createHorizontalRect(tL, tW, tX, tY, tZ, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OFront.Top.createHorizontalRect(tL, tW, -tX, tY, tZ, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
        ' now let's do the front panel of the front outboard
        xCenter1 = -PDistance / 2
		xCenter1 = xCenter1 - PontWidth
		xCenter1 = xCenter1 - outFront.Depth / 2
		xCenter2 = -xCenter1
		
		ycenter = PontLength / 2
		ycenter = ycenter - PLength
		ycenter = ycenter - outFront.sideProtusion / 2
		
		radAngle = System.Math.Atan(outFront.Depth / outFront.sideProtusion)
		angle1 = radAngle * 180 / PI
		angle2 = radAngle * 180 / PI
		angle1 = -angle1 - 90
		angle2 = angle2 - 90
		
		Length = outFront.Depth / System.Math.Sin(radAngle)
		width = outFront.Height
		
		lPontoon.OFront.side2.createVerticalRect(Length, width, xCenter1, ycenter, zCenter, angle1, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OFront.side2.createVerticalRect(Length, width, xCenter2, ycenter, zCenter, angle2, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' let's draw the upper and lower triangles
		tX = -PDistance / 2
		tX = tX - PontWidth
		
		tY = PontLength / 2
		tY = tY - PLength
		tY = tY - outFront.sideProtusion
		
		tZ = PontHeight / 2
		tZ = tZ - outFront.Height / 2
		
		tangle = 180
		
		tL = outFront.sideProtusion
		tW = outFront.Depth
		
		lPontoon.OFront.bside2.createTriangle(tL, tW, tX, tY, tZ, tangle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OFront.bside2.createTriangle(tW, tL, -tX, tY, tZ, tangle + 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' now do the top
		tZ = tZ + outFront.Height
		
		lPontoon.OFront.tside2.createTriangle(tL, tW, tX, tY, tZ, tangle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OFront.tside2.createTriangle(tW, tL, -tX, tY, tZ, tangle + 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' let's draw the inner outboard, the single long one.
		
		' draw rear panel of the long outboard
		
		radAngle = System.Math.Atan(longOne.Depth / longOne.sideProtusion)
		angle1 = radAngle * 180 / PI
		angle2 = radAngle * 180 / PI
		angle1 = 90 - angle1
		angle2 = 90 + angle2
		
		xCenter1 = -PDistance / 2
		xCenter1 = xCenter1 + longOne.Depth / 2
		xCenter2 = -xCenter1
		
		ycenter = -longOne.Length / 2
		ycenter = ycenter - longOne.sideProtusion / 2
		
		Length = longOne.Depth / System.Math.Sin(radAngle)
		width = longOne.Height
		
		lPontoon.OLong.side1.createVerticalRect(Length, width, xCenter1, ycenter, zCenter, angle1, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OLong.side1.createVerticalRect(Length, width, xCenter2, ycenter, zCenter, angle2, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' let's draw the triangles
		tX = -PDistance / 2
		'tY = -PontLength / 2
		'tY = tY + PLength
		'tY = tY + longOne.sideProtusion
		tY = -longOne.Length / 2
		tZ = PontHeight / 2
		tZ = tZ - longOne.Height / 2
		
		tW = longOne.sideProtusion
		tL = longOne.Depth
		
		tangle = 0
		
		lPontoon.OLong.bSide1.createTriangle(tL, tW, tX, tY, tZ, tangle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OLong.bSide1.createTriangle(tW, tL, -tX, tY, tZ, tangle + 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' now do the upper ones
		tZ = tZ + longOne.Height
		
		lPontoon.OLong.tSide1.createTriangle(tL, tW, tX, tY, tZ, tangle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OLong.tSide1.createTriangle(tW, tL, -tX, tY, tZ, tangle + 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		
		' let's do the middle panel of the long inner panel
		xCenter1 = -PDistance / 2
		xCenter1 = xCenter1 + longOne.Depth
		xCenter2 = -xCenter1
		
		ycenter = 0
		
		Length = longOne.Length
		width = longOne.Height
		
		lPontoon.OLong.front.createVerticalRect(Length, width, xCenter1, ycenter, zCenter, 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OLong.front.createVerticalRect(Length, width, xCenter2, ycenter, zCenter, 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' let's do the top and bottom rects
		tX = -PDistance / 2
		tX = tX + longOne.Depth / 2
		
		tY = 0
		
		tZ = PontHeight / 2
		tZ = tZ - longOne.Height / 2
		
		tL = longOne.Depth
		tW = longOne.Length
		
		lPontoon.OLong.Bottom.createHorizontalRect(tL, tW, tX, tY, tZ, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OLong.Bottom.createHorizontalRect(tL, tW, -tX, tY, tZ, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		'now do the top
		
		tZ = tZ + longOne.Height
		
		lPontoon.OLong.Top.createHorizontalRect(tL, tW, tX, tY, tZ, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OLong.Top.createHorizontalRect(tL, tW, -tX, tY, tZ, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		
		' now do the front panel of the long inner outboard
		xCenter1 = -PDistance / 2
		xCenter1 = xCenter1 + longOne.Depth / 2
		xCenter2 = -xCenter1
		
		ycenter = longOne.Length / 2
		ycenter = longOne.sideProtusion / 2
		
		radAngle = System.Math.Atan(longOne.Depth / longOne.sideProtusion)
		angle1 = radAngle * 180 / PI
		angle2 = radAngle * 180 / PI
		angle1 = 90 + angle1
		angle2 = 90 - angle2
		
		xCenter1 = -PDistance / 2
		xCenter1 = xCenter1 + longOne.Depth / 2
		xCenter2 = -xCenter1
		
		ycenter = PontLength / 2
		ycenter = ycenter - PLength
		ycenter = ycenter - longOne.sideProtusion / 2
		
		Length = longOne.Depth / System.Math.Sin(radAngle)
		width = longOne.Height
		
		lPontoon.OLong.side2.createVerticalRect(Length, width, xCenter1, ycenter, zCenter, angle1, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OLong.side2.createVerticalRect(Length, width, xCenter2, ycenter, zCenter, angle2, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' draw the little triangles
		tX = -PDistance / 2
		tY = longOne.Length / 2
		tZ = PontHeight / 2
		tZ = tZ - longOne.Height / 2
		
		tW = longOne.Depth
		tL = longOne.sideProtusion
		
		tangle = -90
		
		lPontoon.OLong.bside2.createTriangle(tL, tW, tX, tY, tZ, tangle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OLong.bside2.createTriangle(tW, tL, -tX, tY, tZ, tangle - 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' now do the top
		tZ = tZ + longOne.Height
		
		lPontoon.OLong.tside2.createTriangle(tL, tW, tX, tY, tZ, tangle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		rPontoon.OLong.tside2.createTriangle(tW, tL, -tX, tY, tZ, tangle - 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))

    End Sub
	
	Public Sub DrawOutboards(ByRef gIn As threeDGrapher)
		gIn.drawRect(lPontoon.OBack.side1)
		gIn.drawRect(rPontoon.OBack.side1)
		gIn.drawRect(lPontoon.OBack.front)
		gIn.drawRect(rPontoon.OBack.front)
		gIn.drawRect(lPontoon.OBack.side2)
		gIn.drawRect(rPontoon.OBack.side2)
		
		gIn.drawRect(lPontoon.OMiddle.side1)
		gIn.drawRect(rPontoon.OMiddle.side1)
		gIn.drawRect(lPontoon.OMiddle.front)
		gIn.drawRect(rPontoon.OMiddle.front)
		gIn.drawRect(lPontoon.OMiddle.side2)
		gIn.drawRect(rPontoon.OMiddle.side2)
		
		gIn.drawRect(lPontoon.OFront.side1)
		gIn.drawRect(rPontoon.OFront.side1)
		gIn.drawRect(lPontoon.OFront.front)
		gIn.drawRect(rPontoon.OFront.front)
		gIn.drawRect(lPontoon.OFront.side2)
		gIn.drawRect(rPontoon.OFront.side2)
		
		gIn.drawRect(lPontoon.OLong.side1)
		gIn.drawRect(rPontoon.OLong.side1)
		gIn.drawRect(lPontoon.OLong.front)
		gIn.drawRect(rPontoon.OLong.front)
		gIn.drawRect(lPontoon.OLong.side2)
		gIn.drawRect(rPontoon.OLong.side2)
		
		gIn.drawRect(lPontoon.OBack.bSide1)
		gIn.drawRect(rPontoon.OBack.bSide1)
		
		gIn.drawRect(lPontoon.OBack.tSide1)
		gIn.drawRect(rPontoon.OBack.tSide1)
		
		gIn.drawRect(lPontoon.OBack.Bottom)
		gIn.drawRect(rPontoon.OBack.Bottom)
		
		gIn.drawRect(lPontoon.OBack.Top)
		gIn.drawRect(rPontoon.OBack.Top)
		
		gIn.drawRect(lPontoon.OBack.bside2)
		gIn.drawRect(rPontoon.OBack.bside2)
		
		gIn.drawRect(lPontoon.OBack.tside2)
		gIn.drawRect(rPontoon.OBack.tside2)
		
		gIn.drawRect(lPontoon.OMiddle.bSide1)
		gIn.drawRect(rPontoon.OMiddle.bSide1)
		
		gIn.drawRect(lPontoon.OMiddle.tSide1)
		gIn.drawRect(rPontoon.OMiddle.tSide1)
		
		gIn.drawRect(lPontoon.OMiddle.Bottom)
		gIn.drawRect(rPontoon.OMiddle.Bottom)
		
		gIn.drawRect(lPontoon.OMiddle.Top)
		gIn.drawRect(rPontoon.OMiddle.Top)
		
		gIn.drawRect(lPontoon.OMiddle.bside2)
		gIn.drawRect(rPontoon.OMiddle.bside2)
		
		gIn.drawRect(lPontoon.OMiddle.tside2)
		gIn.drawRect(rPontoon.OMiddle.tside2)
		
		gIn.drawRect(lPontoon.OFront.bSide1)
		gIn.drawRect(rPontoon.OFront.bSide1)
		
		gIn.drawRect(lPontoon.OFront.tSide1)
		gIn.drawRect(rPontoon.OFront.tSide1)
		
		gIn.drawRect(lPontoon.OFront.Bottom)
		gIn.drawRect(rPontoon.OFront.Bottom)
		
		gIn.drawRect(lPontoon.OFront.Top)
		gIn.drawRect(rPontoon.OFront.Top)
		
		gIn.drawRect(lPontoon.OFront.bside2)
		gIn.drawRect(rPontoon.OFront.bside2)
		
		gIn.drawRect(lPontoon.OFront.tside2)
		gIn.drawRect(rPontoon.OFront.tside2)
		
		gIn.drawRect(lPontoon.OLong.bSide1)
		gIn.drawRect(rPontoon.OLong.bSide1)
		
		gIn.drawRect(lPontoon.OLong.tSide1)
		gIn.drawRect(rPontoon.OLong.tSide1)
		
		gIn.drawRect(lPontoon.OLong.Bottom)
		gIn.drawRect(rPontoon.OLong.Bottom)
		
		gIn.drawRect(lPontoon.OLong.Top)
		gIn.drawRect(rPontoon.OLong.Top)
		
		gIn.drawRect(lPontoon.OLong.bside2)
		gIn.drawRect(rPontoon.OLong.bside2)
		
		gIn.drawRect(lPontoon.OLong.tside2)
		gIn.drawRect(rPontoon.OLong.tside2)
	End Sub
End Class