Option Strict Off
Option Explicit On
Friend Class PltAnchorLines
	
	Private Structure Points
		Dim X As Single
		Dim Y As Single
		Dim z As Single
	End Structure
	Private Lines() As threeDLine
	Private anchors(12) As PltAnchor
	Private starting(12) As Points
	
	Private numAnchors As Short
	Private numberLines As Short
	Private sizeOfArray As Short
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		numberLines = 0
		sizeOfArray = 10
		ReDim Lines(sizeOfArray)
		
		numAnchors = 0
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	Private Sub CreateStartingPoints()
		' create line 1 and line 2 starting points (they are same)
		
		Dim counter As Short
		
		starting(0).X = -PDistance / 2 - PontWidth
		starting(0).Y = PontLength / 2 - PLength
		starting(0).z = 63
		
		starting(1).X = starting(0).X
		starting(1).Y = starting(0).Y
		starting(1).z = 63
		
		starting(2).X = -PDistance / 2 - PontWidth
		starting(2).Y = -PontLength / 2 + PLength
		starting(2).z = 63
		
		starting(3).X = starting(2).X
		starting(3).Y = starting(2).Y
		starting(3).z = 63
		
		starting(4).X = PDistance / 2 + PontWidth
		starting(4).Y = -PontLength / 2 + PLength
		starting(4).z = 63
		
		starting(5).X = starting(4).X
		starting(5).Y = starting(4).Y
		starting(5).z = 63
		
		starting(6).X = PDistance / 2 + PontWidth
		starting(6).Y = PontLength / 2 - PLength
		starting(6).z = 63
		
		starting(7).X = starting(6).X
		starting(7).Y = starting(6).Y
		starting(7).z = 63
		
		
	End Sub
	
	Public Sub ReadFile(ByVal FileName As String)
		
		Dim OpenFile As Boolean
		
		On Error GoTo ExitErr
		
		OpenFile = False
		
		Dim fNum As Short
		fNum = FreeFile
		
		FileOpen(fNum, FileName, OpenMode.Input)
		OpenFile = True
		numAnchors = 0
		numberLines = 0
		
		Dim x1 As Single
		Dim y1 As Single
		Dim z1 As Single
		
		Dim x2 As Single
		Dim y2 As Single
		Dim z2 As Single
		
		Dim Index As Short
		
		Index = 0
		
		Dim lineNumber As Short
		lineNumber = 0
		
		Dim tester As Object
		Do While Not EOF(fNum)
			Input(fNum, tester)
			If Not IsNumeric(tester) Then
				'Input #fNum, tester
				'lineNum = lineNum +
				'Input #fNum, y1, x1, z1
				If Not Index = 0 Then
					anchors(numAnchors) = New PltAnchor
					anchors(numAnchors).createAnchor(x2, y2, z2)
					numAnchors = numAnchors + 1
				End If
				'        y1 = starting(lineNumber).y
				'        x1 = starting(lineNumber).x
				'        z1 = starting(lineNumber).z
				Input(fNum, y1)
				Input(fNum, x1)
				Input(fNum, z1)
				x1 = -x1
				lineNumber = lineNumber + 1
				
				
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object tester. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				y2 = tester
				Input(fNum, x2)
				Input(fNum, z2)
				x2 = -x2
				If Index >= sizeOfArray Then
					sizeOfArray = sizeOfArray + 10
					ReDim Preserve Lines(sizeOfArray)
				End If
				Lines(Index) = New threeDLine
				Lines(Index).setCoords(x1, y1, z1, x2, y2, z2)
				
				Index = Index + 1
				numberLines = numberLines + 1
				x1 = x2
				y1 = y2
				z1 = z2
				
			End If
			
		Loop 
		
		If Not Index = 0 Then
			anchors(numAnchors) = New PltAnchor
			anchors(numAnchors).createAnchor(x2, y2, z2)
			numAnchors = numAnchors + 1
			
		End If
		
		FileClose(fNum)
		OpenFile = False
		
		Exit Sub
		
ExitErr: 
		If OpenFile Then FileClose(fNum)
		Exit Sub
		
	End Sub
	
	Public Sub drawAnchorLines(ByRef aGrapher As threeDGrapher)
		Dim counter As Short
		counter = 0
		
		Do While counter < numberLines
			aGrapher.drawTheLine(Lines(counter))
			counter = counter + 1
		Loop 
		
		counter = 0
		Do While counter < numAnchors
			anchors(counter).drawAnchor(aGrapher)
			counter = counter + 1
		Loop 
		
		
	End Sub
End Class