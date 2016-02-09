Option Strict Off
Option Explicit On
Friend Class PltColumns
	' this class creates the columns that support the deck
	' and are found between the deck and pontoons
	
	
	Private Const BackCol1Length As Single = PontWidth '...back box column width
	Private Const BackCol1Width As Single = 17 '...back box column length
	Private Const BackCol2Length As Single = PontWidth '... rear octagon column
	Private Const SmallLength As Single = 3.28 * 5.8
	'...modify by neil
	Private Const BackCol3Length As Single = 17 '...back box column width (new pontoon)
	Private Const BackCol3Width As Single = PontWidth '...back box column length (new pontoon)
	
	Private LBackCol1 As PltBox3d
	Private RBackCol1 As PltBox3d
	Private LBackCol2 As PltOctagon3d
	Private RBackCol2 As PltOctagon3d
	Private RFrontCol1 As PltBox3d
	Private LFrontCol1 As PltBox3d
	Private LFrontCol2 As PltOctagon3d
	Private RFrontCol2 As PltOctagon3d
	Private LFrontSmall As PltOctagon3d
	Private RFrontSmall As PltOctagon3d
	Private LBackSmall As PltOctagon3d
	Private RBackSmall As PltOctagon3d
	'Private LMiddle As PltOctagon3d
	'Private RMiddle As PltOctagon3d
	'...modify by neil
	Private LBackCol3 As PltBox3d
	Private RBackCol3 As PltBox3d
	Private LBackCol4 As PltOctagon3d
	Private RBackCol4 As PltOctagon3d
	Private LFrontCol3 As PltBox3d
	Private RFrontCol3 As PltBox3d
	Private LFrontCol4 As PltOctagon3d
	Private RFrontCol4 As PltOctagon3d


    Public Sub New()
		MyBase.New()
        LBackCol1 = New PltBox3d
        RBackCol1 = New PltBox3d
        LBackCol2 = New PltOctagon3d
        RBackCol2 = New PltOctagon3d
        RFrontCol1 = New PltBox3d
        LFrontCol1 = New PltBox3d
        RFrontCol2 = New PltOctagon3d
        LFrontCol2 = New PltOctagon3d
        LFrontSmall = New PltOctagon3d
        RFrontSmall = New PltOctagon3d
        LBackSmall = New PltOctagon3d
        RBackSmall = New PltOctagon3d
        '...modify by neil
        LBackCol3 = New PltBox3d
        RBackCol3 = New PltBox3d
        LBackCol4 = New PltOctagon3d
        RBackCol4 = New PltOctagon3d
        RFrontCol3 = New PltBox3d
        LFrontCol3 = New PltBox3d
        RFrontCol4 = New PltOctagon3d
        LFrontCol4 = New PltOctagon3d
        'Set RMiddle = New PltOctagon3d
        'Set LMiddle = New PltOctagon3d
        createColumns()
    End Sub
	
	Private Sub createColumns()
		' create the back left and right columns
		Dim x1 As Single
		Dim y1 As Single
		Dim z1 As Single
		
		Dim Height As Single
		'============================================================
		'...left rear box column
		x1 = -PDistance / 2 - PontWidth / 2
		y1 = -(PontLength / 2 - PLength + BackCol1Width / 2)
		z1 = (DeckElev + PontHeight) / 2
		Height = DeckElev - PontHeight
		
		LBackCol1.createBox(BackCol1Length, BackCol1Width, Height, x1, y1, z1)
		
		'...right rear box column
		x1 = -x1
		
		RBackCol1.createBox(BackCol1Length, BackCol1Width, Height, x1, y1, z1)
		'============================================================
		'...left rear octagon column
		x1 = -x1
		y1 = PontLength / 2 - PLength - BackCol2Length / 2
		z1 = PontHeight
		Height = DeckElev - PontHeight
		
		LBackCol2.createOctagon(BackCol2Length, Height, x1, y1, z1)
		
		'...right rear octagon column
		x1 = -x1
		RBackCol2.createOctagon(BackCol2Length, Height, x1, y1, z1)
		'============================================================
		'...left front box column
		x1 = -PDistance / 2 - PontWidth / 2
		y1 = PontLength / 2 - PLength + BackCol1Width / 2
		z1 = (PontHeight + DeckElev) / 2
		Height = DeckElev - PontHeight
		
		LFrontCol1.createBox(BackCol1Length, BackCol1Width, Height, x1, y1, z1)
		
		'...right front box column
		x1 = -x1
		RFrontCol1.createBox(BackCol1Length, BackCol1Width, Height, x1, y1, z1)
		'============================================================
		'...left front octagon column
		x1 = -PDistance / 2 - PontWidth / 2
		y1 = -(PontLength / 2 - PLength - BackCol2Length / 2)
		z1 = PontHeight
		
		Height = DeckElev - PontHeight
		
		LFrontCol2.createOctagon(BackCol2Length, Height, x1, y1, z1)
		
		'...right front octagon column
		x1 = -x1
		RFrontCol2.createOctagon(BackCol2Length, Height, x1, y1, z1)
		'============================================================
		'...left front small column
		x1 = -PDistance / 2 - PontWidth / 2
		y1 = PontLength / 2 - 141
		z1 = PontHeight
		Height = DeckElev - PontHeight
		
		LFrontSmall.createOctagon(SmallLength, Height, x1, y1, z1)
		
		'...right front small column
		x1 = -x1
		RFrontSmall.createOctagon(SmallLength, Height, x1, y1, z1)
		'============================================================
		'...left rear small column
		x1 = -x1
		y1 = -y1
		
		LBackSmall.createOctagon(SmallLength, Height, x1, y1, z1)
		
		'...right rear small column
		x1 = -x1
		RBackSmall.createOctagon(SmallLength, Height, x1, y1, z1)
		'============================================================
		'...left rear box column (new pontoon)
		x1 = -PDistance / 2 - PontWidth * 2 - (PDistance - 23) - BackCol3Length / 2
		y1 = -(PontLength / 2 - 141)
		z1 = (DeckElev + PontHeight) / 2
		Height = DeckElev - PontHeight
		LBackCol3.createBox(BackCol3Length, BackCol3Width, Height, x1, y1, z1)
		
		'...right rear box column (new pontoon)
		x1 = -x1
		
		RBackCol3.createBox(BackCol3Length, BackCol3Width, Height, x1, y1, z1)
		'============================================================
		'...left front box column (new pontoon)
		x1 = -x1
		y1 = -y1
		
		LFrontCol3.createBox(BackCol3Length, BackCol3Width, Height, x1, y1, z1)
		
		'...right front box column (new pontoon)
		x1 = -x1
		
		RFrontCol3.createBox(BackCol3Length, BackCol3Width, Height, x1, y1, z1)
		'============================================================
		'...left rear octagon column (new pontoon)
		x1 = -PDistance / 2 - PontWidth - (PDistance - 23) - BackCol3Width / 2
		y1 = PontLength / 2 - 141
		z1 = PontHeight
		Height = DeckElev - PontHeight
		
		LBackCol4.createOctagon(BackCol2Length, Height, x1, y1, z1)
		
		'...right rear octagon column (new pontoon)
		x1 = -x1
		RBackCol4.createOctagon(BackCol2Length, Height, x1, y1, z1)
		'============================================================
		'...left front octagon column (new pontoon)
		x1 = -x1
		y1 = -y1
		
		LFrontCol4.createOctagon(BackCol2Length, Height, x1, y1, z1)
		
		'...right front octagon column (new pontoon)
		x1 = -x1
		RFrontCol4.createOctagon(BackCol2Length, Height, x1, y1, z1)
		'============================================================
		' do the middle one
		'x1 = -PDistance / 2 - PontWidth / 2
		'y1 = 0
		'z1 = PontHeight
		
		'LMiddle.createOctagon BackCol2Length, Height, x1, y1, z1
		'RMiddle.createOctagon BackCol2Length, Height, -x1, y1, z1
		
	End Sub
	
	Public Sub drawColumns(ByRef gIn As threeDGrapher)
		gIn.drawBox(LBackCol1)
		gIn.drawBox(RBackCol1)
		gIn.drawOctagon(LBackCol2)
		gIn.drawOctagon(RBackCol2)
		gIn.drawBox(LFrontCol1)
		gIn.drawBox(RFrontCol1)
		gIn.drawOctagon(LFrontCol2)
		gIn.drawOctagon(RFrontCol2)
		gIn.drawOctagon(LFrontSmall)
		gIn.drawOctagon(RFrontSmall)
		gIn.drawOctagon(LBackSmall)
		gIn.drawOctagon(RBackSmall)
		'gIn.drawOctagon RMiddle
		'gIn.drawOctagon LMiddle
		'...modify by neil
		gIn.drawBox(LBackCol3)
		gIn.drawBox(RBackCol3)
		gIn.drawBox(LFrontCol3)
		gIn.drawBox(RFrontCol3)
		gIn.drawOctagon(LBackCol4)
		gIn.drawOctagon(RBackCol4)
		gIn.drawOctagon(LFrontCol4)
		gIn.drawOctagon(RFrontCol4)
	End Sub
End Class