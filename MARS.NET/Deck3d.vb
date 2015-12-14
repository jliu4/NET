Option Strict Off
Option Explicit On
Friend Class PltDeck3d
	' this class create the deck.  It is simply a rectangular
	' box
	
	
	
	Private Const DeckHeight As Single = 8
	Private Const DeckLength As Single = 3.28 * 90
	Private Const DeckWidth As Single = 3.28 * (6 + 76 + 6)
	
	Private Const HeliDeckLength As Single = 3.28 * 25
	Private Const HeliDeckHeight As Single = 3.28 * 2.5
	Private Const helideckElev As Single = DeckElev + DeckHeight
	
	Private side1 As threeDRect
	Private side2 As threeDRect
	Private Top As threeDRect
	Private Bottom As threeDRect
	Private HeliDeck As PltOctagon3d
	Private theRig As PltDrillRig
	'Private deckShading(10) As PltBox3d
	
	' the class constructor.  It sets up instances of the four
	' member variables and also creates the rectangles which
	' will make up the deck
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		side1 = New threeDRect
		side2 = New threeDRect
		Top = New threeDRect
		Bottom = New threeDRect
		theRig = New PltDrillRig
		HeliDeck = New PltOctagon3d
		'Dim x As Integer
		'x = 1
		'Do While x <= 10
		'    Set deckShading(x) = New PltBox3d
		'    x = x + 1
		'Loop
		
		createdeck()
		createHelideck()
		
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	Private Sub createHelideck()
		Dim X As Single
		Dim Y As Single
		Dim z As Single
		
		X = PDistance / 2 + PontWidth + (PDistance - 23) + PontWidth / 2
		Y = -3.28 * 77 / 2
		z = helideckElev
		
		HeliDeck.createOctagon(HeliDeckLength, HeliDeckHeight, X, Y, z)
		
	End Sub
	
	
	' this method creates the deck from the rectangles using
	' the constants above for its dimensions
	
	Private Sub createdeck()
		Dim X As Single
		Dim Y As Single
		Dim z As Single
		
		' draw side1
		X = -DeckWidth / 2
		Y = 0
		z = DeckElev + DeckHeight / 2
		side1.createVerticalRect(DeckLength, DeckHeight, X, Y, z, 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' draw side2
		X = DeckWidth / 2
		Y = 0
		z = DeckElev + DeckHeight / 2
		side2.createVerticalRect(DeckLength, DeckHeight, X, Y, z, 90, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		'bottom
		X = 0
		Y = 0
		z = DeckElev
		Bottom.createHorizontalRect(DeckWidth, DeckLength, X, Y, z, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' top
		X = 0
		Y = 0
		z = DeckElev + DeckHeight
		Top.createHorizontalRect(DeckWidth, DeckLength, X, Y, z, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
		
		' create the shading
		'UPGRADE_NOTE: step was upgraded to step_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim step_Renamed As Single
		'UPGRADE_NOTE: Current was upgraded to Current_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Current_Renamed As Single
		step_Renamed = DeckLength / 10
		X = 0
		Y = 0
		z = DeckElev + DeckHeight / 2
		
		Current_Renamed = step_Renamed
		Dim Index As Short
		Index = 1
		'Do While index <= 10
		
		'   deckShading(index).createBox DeckWidth, current, _
		''   DeckHeight, x, y, z, vbRed
		'   index = index + 1
		'   current = current + step
		'Loop
		
	End Sub
	
	Public Sub drawDeck(ByRef gIn As threeDGrapher)
		
		gIn.drawRect(Bottom, True)
		gIn.drawRect(side1)
		gIn.drawRect(side2)
		
		gIn.drawRect(Top, True)
		gIn.drawOctagon(HeliDeck, True)
		theRig.drawRig(gIn)
		
		'Dim index As Integer
		'index = 1
		'Do While index <= 10
		'Debug.Print index
		'    gIn.drawBox deckShading(index)
		'    index = index + 1
		'Loop
		
	End Sub
End Class