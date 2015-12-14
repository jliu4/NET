Option Strict Off
Option Explicit On
Friend Class PltDrillRig
	
	' this class creates a drilling rig
	
	Private Const GroundHeight As Single = DeckElev + 8
	
	Private Const BaseWidth As Single = 3.28 * 24
	Private Const BaseHeight As Single = 3.28 * 4
	Private Const BaseLength As Single = 3.28 * 15
	
	Private Const aboveBaseHeight As Single = 3.28 * 26
	Private Const aboveBaseLength As Single = 3.28 * 15
	Private Const aboveBaseWidth As Single = 3.28 * 16
	
	Private Const namePlateLength As Single = 15 * 3.28
	Private Const namePlateWidth As Single = 20 * 3.28
	Private Const namePlateHeight As Single = 7 * 3.28
	
	Private Const topPyrHeight As Single = 3.28 * 17.5
	Private Const topPyrBLength As Single = 3.28 * 15
	Private Const topPyrBWidth As Single = 3.28 * 16
	Private Const topPyrTLength As Single = 3.28 * 7
	Private Const topPyrTWidth As Single = 3.28 * 6
	
	Private Const theTopLength As Single = 3.28 * 7
	Private Const theTopWidth As Single = 3.28 * 6
	Private Const theTopHeight As Single = 3.28 * 13.5
	
	Private Base As PltBox3d
	Private AboveBase As PltBox3d
	Private namePlate As PltBox3d
	Private topPyr As PltBox3d
	Private theTop As PltBox3d
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		Base = New PltBox3d
		AboveBase = New PltBox3d
		namePlate = New PltBox3d
		topPyr = New PltBox3d
		theTop = New PltBox3d
		createBoxes()
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	Private Sub createBoxes()
		
		' create the base box
		Dim X As Single
		Dim Y As Single
		Dim z As Single
		X = 0
		Y = 0
		z = GroundHeight + BaseHeight / 2
		Base.createBox(BaseLength, BaseWidth, BaseHeight, X, Y, z)
		
		' create the above Base part
		X = 0
		Y = 0
		z = GroundHeight + BaseHeight + aboveBaseHeight / 2
		AboveBase.createBox(aboveBaseLength, aboveBaseWidth, aboveBaseHeight, X, Y, z)
		
		' create the name plate part
		X = 0
		Y = 0
		z = GroundHeight + BaseHeight + aboveBaseHeight + namePlateHeight / 2
		namePlate.createBox(namePlateLength, namePlateWidth, namePlateHeight, X, Y, z)
		
		
		' create that pyramidal looking box
		X = 0
		Y = 0
		z = GroundHeight + BaseHeight + aboveBaseHeight + namePlateHeight + topPyrHeight / 2
		topPyr.PyramidBox(topPyrBLength, topPyrBWidth, topPyrTLength, topPyrTWidth, topPyrHeight, X, Y, z)
		
		' create the top box and we are done!
		X = 0
		Y = 0
		z = GroundHeight + BaseHeight + aboveBaseHeight + namePlateHeight + topPyrHeight + theTopHeight / 2
		
		theTop.createBox(theTopLength, theTopWidth, theTopHeight, X, Y, z)
		
	End Sub
	
	Public Sub drawRig(ByRef gIn As threeDGrapher)
		gIn.drawBox(Base)
		gIn.drawBox(AboveBase)
		gIn.drawBox(namePlate)
		gIn.drawBox(topPyr)
		gIn.drawBox(theTop)
	End Sub
End Class