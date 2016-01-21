Option Strict Off
Option Explicit On
Friend Class Riser
	
	' riser properties
	
	' properties
	' Mass:         riser total mass including mud (lbs)
	' Length:       riser length (ft) = water depth
	' TopTen:       top tension (lbs)
	' FhLocl:       horizontal force in local coordinate (lbs)
	' Dia:          riser diameter (ft)
	' Cd:           drag coefficient 1.2
	' Cm:           mass coefficient 1.0
	
	Private msngMass As Single
	Private msngUnitMass As Single
	Private msngUnitAddMass As Single
	Private msngLength As Single
	Private msngTopTen As Single
	Private msngDia As Single
	Private msngCd As Single
	Private msngCm As Single

    Private mclsFhLocl As Force
    Public RiserGrapher As threeDGrapher

    Private theRiser As threeDLine



    'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub Class_Initialize_Renamed()

        theRiser = New threeDLine
        theRiser.setColor(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red))

        RiserGrapher = New threeDGrapher

    End Sub

    Public Sub drawRiser(ByRef WellSite As Well, ByRef ShipGlob As ShipGlobal, ByVal ShipDraft As Single)

        Dim x1 As Single
        Dim y1 As Single
        Dim z1 As Single
        Dim x2 As Single
        Dim y2 As Single
        Dim z2 As Single

        x1 = 0 '...ship water center x coords sys.
        y1 = 0 '...ship water center y coords sys.
        z1 = DeckElev

        With WellSite
            x2 = .Xg - ShipGlob.Xg '...referencr to ship sys. coords
            y2 = .Yg - ShipGlob.Yg '...referencr to ship sys. coords
            z2 = ShipDraft - .Depth
        End With

        theRiser.setCoords(x1, y1, z1, x2, y2, z2)
        RiserGrapher.drawTheLine(theRiser)

    End Sub

    Public Sub angleChange(ByVal hChange As Short, ByRef vChange As Short)
        RiserGrapher.newAngles(hChange, vChange)
    End Sub

    Public Sub setGraph(ByRef aGrapher As System.Windows.Forms.PictureBox, ByRef X As Single, ByRef Y As Single, ByRef Incremental As Boolean)
        RiserGrapher.setDrawSurface(aGrapher, X, Y, Incremental)
    End Sub

    Public Function ZoomGraph(ByVal aChange As Single, ByRef WellSite As Well, ByRef ShipGlob As ShipGlobal, ByVal ShipDraft As Single) As Object

        Dim zoomVar As Boolean

        zoomVar = RiserGrapher.changeScale(aChange)
        If zoomVar = True Then
            drawRiser(WellSite, ShipGlob, ShipDraft)
        End If
        ZoomGraph = zoomVar

    End Function

    Public Sub New()
        MyBase.New()

        Class_Initialize_Renamed()
        mclsFhLocl = New Force
        msngCd = 1.2
        msngCm = 1.0#
	End Sub
	
	' properties	
	Public Property mass() As Single
		Get
			
			mass = msngMass
			
		End Get
		Set(ByVal Value As Single)
			
			msngMass = Value
			
		End Set
	End Property
	
	
	Public Property Length() As Single
		Get
			
			Length = msngLength
			
		End Get
		Set(ByVal Value As Single)
			
			msngLength = Value
			
		End Set
	End Property
	
	
	Public Property TopTen() As Single
		Get
			
			TopTen = msngTopTen
			
		End Get
		Set(ByVal Value As Single)
			
			msngTopTen = Value
			
		End Set
	End Property
	
	
	Public Property Dia() As Single
		Get
			
			Dia = msngDia
			
		End Get
		Set(ByVal Value As Single)
			
			msngDia = Value
			
		End Set
	End Property
	
	Public ReadOnly Property Cd() As Single
		Get
			
			Cd = msngCd
			
		End Get
	End Property
	
	Public ReadOnly Property Cm() As Single
		Get
			
			Cm = msngCm
			
		End Get
	End Property
	
	Public ReadOnly Property FhLocl(ByVal ShipPos As ShipGlobal, ByVal ShipVel As Motion, ByVal ShipAcc As Motion, ByVal Current As Current) As Force
		Get
			
			Dim i, NumCurrent, j As Short
            Dim CurrentDepth(2) As Single
            Dim CurrentApp(2, 2) As Single

			Dim EnvDir As Single
            Dim TotalDrag(2) As Single
            Dim TotalMoment(2) As Single
			Dim Drag As Single
			Dim Moment As Single
			
			Dim FRiser As Force
			Dim FfrTopTen As Force
			Dim Finertia As Force
			Dim ShipPosLocl As ShipLocal
			
			FRiser = mclsFhLocl
			FfrTopTen = New Force
			Finertia = New Force
			ShipPosLocl = New ShipLocal
			
			If msngLength > 0# Then
				msngUnitMass = msngMass / msngLength
			Else
				msngUnitMass = msngMass
			End If
			msngUnitAddMass = 0.25 * PI * msngDia ^ 2 * 64# * msngCm
			
			Call ShipLoclFrmGlob(ShipPosLocl, ShipPos)
			
			With FfrTopTen
				.Fx = -msngTopTen * ShipPosLocl.Surge / msngLength
				.Fy = -msngTopTen * ShipPosLocl.Sway / msngLength
			End With
			
			With Finertia
				.Fx = -1 / 3 * ((msngUnitMass + msngUnitAddMass) / 32.18) * ShipAcc.Surge * msngLength
				.Fy = -1 / 3 * ((msngUnitMass + msngUnitAddMass) / 32.18) * ShipAcc.Sway * msngLength
			End With
			
			With Current
				NumCurrent = .ProfileCount
				TotalDrag(1) = 0#
				TotalDrag(2) = 0#
				TotalMoment(1) = 0#
				TotalMoment(2) = 0#
				
				For i = 1 To NumCurrent - 1
					CurrentDepth(1) = .Profile(i).Depth
					CurrentDepth(2) = .Profile(i + 1).Depth
					EnvDir = ShipPos.Heading - .Heading
					If CurrentDepth(1) <= msngLength Then
						CurrentApp(1, 1) = .Profile(i).Velocity * System.Math.Cos(EnvDir) - ShipVel.Surge * (1# - CurrentDepth(1) / msngLength)
						CurrentApp(1, 2) = .Profile(i).Velocity * System.Math.Sin(EnvDir) - ShipVel.Sway * (1# - CurrentDepth(1) / msngLength)
					Else
						CurrentApp(1, 1) = 0#
						CurrentApp(1, 2) = 0#
					End If
					If CurrentDepth(2) <= msngLength Then
						CurrentApp(2, 1) = .Profile(i + 1).Velocity * System.Math.Cos(EnvDir) - ShipVel.Surge * (1# - CurrentDepth(2) / msngLength)
						CurrentApp(2, 2) = .Profile(i + 1).Velocity * System.Math.Sin(EnvDir) - ShipVel.Sway * (1# - CurrentDepth(2) / msngLength)
					Else
						CurrentApp(1, 1) = 0#
						CurrentApp(1, 2) = 0#
					End If
					
					For j = 1 To 2
						Call DragFnM(CurrentDepth(1), CurrentDepth(2), CurrentApp(1, j), CurrentApp(2, j), Drag, Moment)
						TotalDrag(j) = TotalDrag(j) + Drag
						TotalMoment(j) = TotalMoment(j) + Moment
					Next j
				Next i
			End With
			
			With FRiser
				.Fx = FfrTopTen.Fx + Finertia.Fx + (TotalDrag(1) - TotalMoment(1) / msngLength) * msngDia * msngCd * 1.99 / 2#
				.Fy = FfrTopTen.Fy + Finertia.Fy + (TotalDrag(2) - TotalMoment(2) / msngLength) * msngDia * msngCd * 1.99 / 2#
			End With
			
			FhLocl = FRiser

		End Get
	End Property
	
	Private Sub ShipLoclFrmGlob(ByRef ShipPosLocl As ShipLocal, ByRef ShipPosGlob As ShipGlobal)
		
		Dim Alpha As Single
		
		With ShipPosGlob
			Alpha = .Heading
			
			ShipPosLocl.Heading = Alpha
			ShipPosLocl.Surge = .Yg * System.Math.Cos(Alpha) + .Xg * System.Math.Sin(Alpha)
			ShipPosLocl.Sway = .Yg * System.Math.Sin(Alpha) - .Xg * System.Math.Cos(Alpha)
			ShipPosLocl.Yaw = 0#
		End With
		
	End Sub
	
	Private Sub DragFnM(ByVal x1 As Single, ByVal x2 As Single, ByVal v1 As Single, ByVal v2 As Single, ByRef F As Single, ByRef M As Single)
		
		Dim b, a, x3 As Single
		Dim m1, f1, f2, m2 As Single
		
		If x1 = x2 Then
			F = 0#
			M = 0#
			
		ElseIf v2 * v1 >= 0# Then 
			If v1 = 0# And v2 = 0# Then
				F = 0#
				M = 0#
			Else
				a = (v2 - v1) / (x2 - x1)
				b = (v1 * x2 - v2 * x1) / (x2 - x1)
				F = a ^ 2 * (x2 ^ 3 - x1 ^ 3) / 3# + a * b * (x2 ^ 2 - x1 ^ 2) + b ^ 2 * (x2 - x1)
				M = a ^ 2 * (x2 ^ 4 - x1 ^ 4) / 4# + 2 * a * b * (x2 ^ 3 - x1 ^ 3) / 3# + b ^ 2 * (x2 ^ 2 - x1 ^ 2) / 2#
				
				If v1 <> 0# Then
					F = F * v1 / System.Math.Abs(v1)
					M = M * v1 / System.Math.Abs(v1)
				Else
					F = F * v2 / System.Math.Abs(v2)
					M = M * v2 / System.Math.Abs(v2)
				End If
			End If
		Else
			x3 = -(x2 - x1) / (v2 - v1) * v1 + x1
			
			a = -v1 / (x3 - x1)
			b = v1 * x3 / (x3 - x1)
			f1 = a ^ 2 * (x3 ^ 3 - x1 ^ 3) / 3# + a * b * (x3 ^ 2 - x1 ^ 2) + b ^ 2 * (x3 - x1)
			m1 = a ^ 2 * (x3 ^ 4 - x1 ^ 4) / 4# + 2 * a * b * (x3 ^ 3 - x1 ^ 3) / 3# + b ^ 2 * (x3 ^ 2 - x1 ^ 2) / 2#
			f1 = f1 * v1 / System.Math.Abs(v1)
			m1 = m1 * v1 / System.Math.Abs(v1)
			
			a = v2 / (x2 - x3)
			b = -v2 * x3 / (x2 - x3)
			f2 = a ^ 2 * (x2 ^ 3 - x3 ^ 3) / 3# + a * b * (x2 ^ 2 - x3 ^ 2) + b ^ 2 * (x2 - x3)
			m2 = a ^ 2 * (x2 ^ 4 - x3 ^ 4) / 4# + 2 * a * b * (x2 ^ 3 - x3 ^ 3) / 3# + b ^ 2 * (x2 ^ 2 - x3 ^ 2) / 2#
			f2 = f2 * v2 / System.Math.Abs(v2)
			m2 = m2 * v2 / System.Math.Abs(v2)
			
			F = f1 + f2
			M = m1 + m2
		End If
		
	End Sub
End Class