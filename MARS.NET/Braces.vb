Option Strict Off
Option Explicit On
Friend Class PltBraces
	' Braces should really be circular but we'll estimate
	' them with a box.
	
	' so a single brace will need 4 rectangles to make it up
	
	Public Bottom As threeDRect
	Public Top As threeDRect
	Public side1 As threeDRect
	Public side2 As threeDRect
	
	' this method is called to draw a brace
	
	Public Sub drawBraces(ByRef gIn As threeDGrapher)
		gIn.drawRect(Bottom)
		gIn.drawRect(Top)
		gIn.drawRect(side1)
		gIn.drawRect(side2)
	End Sub
	
	' used to create a brace.  You specify the center starting
	' and finishing coordinates.
	
	Public Sub createBrace(ByVal startx As Single, ByVal startY As Single, ByRef startZ As Single, ByRef finishX As Single, ByRef finishY As Single, ByRef finishZ As Single)
		Dim X0 As Single
		Dim Y0 As Single
		Dim z0 As Single
		Dim temp1 As Single
		Dim temp2 As Single

        Dim hAngle As Single
		Dim aWidth As Single
		Dim alength As Single
		
		aWidth = 2 * hBraceRadius
		
		alength = (startx - finishX) * (startx - finishX) + (startY - finishY) * (startY - finishY)
		alength = System.Math.Sqrt(alength)
		
		' right now lets ignore the vertical angle.
		
		' do the ends first.
		side1.createVerticalRect(2 * hBraceRadius, 2 * hBraceRadius, startx, startY, startZ, 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		side2.createVerticalRect(2 * hBraceRadius, 2 * hBraceRadius, finishX, finishY, finishZ, 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' now create the top and bottom.  Actually these are created
		' using two vertical triangles, doesn't matter, visual
		' effect is the same.
		temp1 = (startx + finishX) / 2
		X0 = temp1 '- hBraceRadius
		temp2 = (startY + finishY) / 2
		Y0 = temp2 - hBraceRadius
		z0 = (startZ + finishZ) / 2
		
		If Not (finishX - startx) = 0 Then
			hAngle = System.Math.Atan((finishY - startY) / (finishX - startx))
		Else
			hAngle = 90 * PI / 180
		End If
		hAngle = hAngle * 180 / PI
		
		Bottom.createVerticalRect(alength, aWidth, X0, Y0, z0, hAngle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		Y0 = temp2 + hBraceRadius
		Top.createVerticalRect(alength, aWidth, X0, Y0, z0, hAngle, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
	End Sub

    Public Sub New()
		MyBase.New()
        Bottom = New threeDRect
        Top = New threeDRect
        side1 = New threeDRect
        side2 = New threeDRect
    End Sub
End Class