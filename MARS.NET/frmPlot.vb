Option Strict Off
Option Explicit On
Friend Class frmPlot
	Inherits System.Windows.Forms.Form
	
	Private anObj2 As PltVessel
	Private anRiser As Riser
	Private anSeaBed As PltSeabed
	Private anPipeLines As PltPipeLines
	Private anSurface As PltWaterSurface
	Private MDown As Boolean
	Private MX0, MY0 As Single
	Private H, V As Short
	
	Public PlotIn3D As Boolean

    Private Sub frmPlot_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        '   Set up the graphics
        anObj2 = New PltVessel
        anRiser = New Riser
		anSeaBed = New PltSeabed
		anPipeLines = New PltPipeLines
		anSurface = New PltWaterSurface
		anSeaBed.WaterDepth = CurVessel.WaterDepth
		lblVessel.Text = CurVessel.Name
		
		H = 0
		V = 0
		
		txtPicControls(0).Text = VB6.Format(H, "##0")
		txtPicControls(1).Text = VB6.Format(V, "##0")
		
		frmMain.PlotShow = True
		
	End Sub
	
	Private Sub frmPlot_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		
		frmMain.PlotShow = False
        anObj2 = Nothing
        anRiser = Nothing
        anSeaBed = Nothing
        anPipeLines = Nothing
        anSurface = Nothing

    End Sub
	
	' buttons
	
	Private Sub btnClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnClose.Click
		
		Me.Close()
		
	End Sub
	
	Private Sub btnPlot_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnPlot.Click
		Dim Index As Short = btnPlot.GetIndex(eventSender)
		
		Select Case Index
			Case 0
				PlotIn3D = True
			Case 1
				PlotIn3D = False
		End Select
		UpdatePicture(True)
		
	End Sub
	
	Private Sub btnZoom_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnZoom.Click
		Dim Index As Short = btnZoom.GetIndex(eventSender)
		
		Dim ZoomRatio As Single

        btnZoom(1 - Index).Enabled = True
		Select Case Index
			Case 0
				ZoomRatio = 1.25
			Case 1
				ZoomRatio = 0.75
		End Select
		
		If Not Zoom(ZoomRatio) Then btnZoom(Index).Enabled = False
		AddNorthIndicator((-hScroll_Renamed.Value * Degrees2Radians))
		
	End Sub
	
	' picture box
	
	Private Sub ThreeD_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles ThreeD.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Integer = eventArgs.X
        Dim Y As Single = eventArgs.Y

        If Not PlotIn3D Then Exit Sub
		MDown = True
		
	End Sub
	
	Private Sub ThreeD_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles ThreeD.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = eventArgs.X
        Dim Y As Single = eventArgs.Y

        Dim V, Vmax As Short
		Dim H, Hmax As Short
		
		If Not PlotIn3D Then Exit Sub
		
		If MDown Then
			Hmax = (hPan.Maximum - hPan.LargeChange + 1)
			Vmax = (vPan.Maximum - vPan.LargeChange + 1)
			H = hPan.Value
			V = vPan.Value
			
			With ThreeD
                H = H + (X - .Width / 2) * Hmax / .Width
                V = V + (Y - .Height / 2) * Vmax / .Height
            End With
			
			If H > Hmax Then H = Hmax
			If H < 0 Then H = 0
			
			If V > Vmax Then V = Vmax
			If V < 0 Then V = 0
			
			hPan.Value = H
			vPan.Value = V
			
			MDown = False
		End If
		
	End Sub
	
	' scroll bars
	
	'UPGRADE_NOTE: hScroll.Change was changed from an event to a procedure. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="4E2DC008-5EDA-4547-8317-C9316952674F"'
	'UPGRADE_WARNING: HScrollBar event hScroll.Change has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub hScroll_Change(ByVal newScrollValue As Integer)
		
		Dim CurWell As Well
		
		H = hScroll_Renamed.Value
		With CurField
			CurWell = .Item(.CurWellNo)
		End With
		
		anObj2.angleChange(H, V)
		anObj2.drawBoat()
		With CurVessel
			anRiser.angleChange(H, V)
			anRiser.drawRiser(CurWell, .ShipCurGlob, .ShipDraft)
			
			anSeaBed.angleChange(H - .ShipCurGlob.Heading * Radians2Degrees, V)
			anSeaBed.drawSeabed(.ShipCurGlob, .ShipDraft)
			
			anPipeLines.angleChange(H - .ShipCurGlob.Heading * Radians2Degrees, V)
			anPipeLines.drawPipeLines(.ShipCurGlob, .ShipDraft)
			
			anSurface.angleChange(H - .ShipCurGlob.Heading * Radians2Degrees, V)
			anSurface.drawSurface(.ShipDraft)
		End With
		
		txtPicControls(0).Text = VB6.Format(H, "##0")
		AddNorthIndicator((-H * Degrees2Radians))
		
	End Sub
	
	'UPGRADE_NOTE: vScroll.Change was changed from an event to a procedure. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="4E2DC008-5EDA-4547-8317-C9316952674F"'
	'UPGRADE_WARNING: VScrollBar event vScroll.Change has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub vScroll_Change(ByVal newScrollValue As Integer)
		
		Dim CurWell As Well
		
		V = 90 - vScroll_Renamed.Value
		With CurField
			CurWell = .Item(.CurWellNo)
		End With
		
		anObj2.angleChange(H, V)
		anObj2.drawBoat()
		With CurVessel
			anRiser.angleChange(H, V)
			anRiser.drawRiser(CurWell, .ShipCurGlob, .ShipDraft)
			
			anSeaBed.angleChange(H - .ShipCurGlob.Heading * Radians2Degrees, V)
			anSeaBed.drawSeabed(.ShipCurGlob, .ShipDraft)
			
			anPipeLines.angleChange(H - .ShipCurGlob.Heading * Radians2Degrees, V)
			anPipeLines.drawPipeLines(.ShipCurGlob, .ShipDraft)
			
			anSurface.angleChange(H - .ShipCurGlob.Heading * Radians2Degrees, V)
			anSurface.drawSurface(.ShipDraft)
		End With
		
		txtPicControls(1).Text = VB6.Format(V, "##0")
		AddNorthIndicator((-H * Degrees2Radians))
		
	End Sub
	
	'UPGRADE_NOTE: vPan.Change was changed from an event to a procedure. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="4E2DC008-5EDA-4547-8317-C9316952674F"'
	'UPGRADE_WARNING: VScrollBar event vPan.Change has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub vPan_Change(ByVal newScrollValue As Integer)
		
		PanChange()
		
	End Sub
	
	'UPGRADE_NOTE: hPan.Change was changed from an event to a procedure. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="4E2DC008-5EDA-4547-8317-C9316952674F"'
	'UPGRADE_WARNING: HScrollBar event hPan.Change has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub hPan_Change(ByVal newScrollValue As Integer)
		
		PanChange()
		
	End Sub
	
	' text boxes
	
	Private Sub txtPicControls_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtPicControls.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = txtPicControls.GetIndex(eventSender)
		
		Select Case KeyCode
			Case System.Windows.Forms.Keys.Return
				txtPicControls_Leave(txtPicControls.Item(Index), New System.EventArgs())
		End Select
		
	End Sub
	
	Private Sub txtPicControls_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPicControls.Leave
		Dim Index As Short = txtPicControls.GetIndex(eventSender)
		
		Dim SAng As Short
		
		SAng = CShort(CheckData(txtPicControls(Index).Text))
		Select Case Index
			Case 0
				With hScroll_Renamed
					If SAng > (.Maximum - .LargeChange + 1) Then
						SAng = (.Maximum - .LargeChange + 1)
					ElseIf SAng < .Minimum Then 
						SAng = .Minimum
					End If
					.Value = SAng
				End With
			Case 1
				With vScroll_Renamed
					If SAng > (.Maximum - .LargeChange + 1) Then
						SAng = (.Maximum - .LargeChange + 1)
					ElseIf SAng < .Minimum Then 
						SAng = .Minimum
					End If
					.Value = 90 - SAng
				End With
		End Select
		txtPicControls(Index).Text = CStr(SAng)
		
	End Sub

    ' operations

    Public Sub UpdatePicture(Optional ByVal Refresh_Renamed As Boolean = False)


        Dim CatFile As String

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        '   Based upon which mode we're in (decided by which value was last changed),
        '   set the new position in Earth Global coordinates

        '   Compute the tensions, etc.

        '   Fill in the remaining values for the textboxes & grid of results

        If Not Refresh_Renamed Then
            CatFile = MarsDir & CatenaryFile
            If Not CurVessel.MoorSystem.GenCatenaryFile(CatFile) Then
                If PlotIn3D Then Exit Sub
            Else
                anObj2.UpdateCatenary(CatFile)
            End If
        End If
        If PlotIn3D Then
            Plot3D()
        Else
            PlotRadar()
        End If

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

    End Sub

    Private Sub AddNorthIndicator(ByVal NorthAngle As Single)
		
		NorthAngle = NorthAngle + CurVessel.ShipCurGlob.Heading
        Call DrawNorthArrow(NorthAngle, ThreeD)

    End Sub
	
	Private Sub Plot3D()

        UpdateControls(True)

        With CurField
			With .Item(.CurWellNo)
				If .Depth <= 0# Then .Depth = CurVessel.WaterDepth
			End With
		End With
		
		vPan.Value = (vPan.Maximum - vPan.LargeChange + 1) / 2
		hPan.Value = (hPan.Maximum - hPan.LargeChange + 1) / 2

        'ThreeD.AutoRedraw = True
        PanChange()
		
		anObj2.angleChange(H, V)
		With CurVessel
			anRiser.angleChange(H, V)
			anSeaBed.angleChange(H - .ShipCurGlob.Heading * Radians2Degrees, V)
			anPipeLines.angleChange(H - .ShipCurGlob.Heading * Radians2Degrees, V)
			anSurface.angleChange(H - .ShipCurGlob.Heading * Radians2Degrees, V)
		End With
		
		AddNorthIndicator((-H * Degrees2Radians))
		
	End Sub
	
	Private Sub PlotRadar()

        Dim Y0, X0, Head0 As Single
        Dim WY, WX, WD As Single

        UpdateControls(False)

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        With CurVessel.ShipCurGlob
			X0 = .Xg
			Y0 = .Yg
			Head0 = .Heading
		End With
		
		With CurField
			WX = .Item(.CurWellNo).Xg
			WY = .Item(.CurWellNo).Yg
			WD = .Item(.CurWellNo).Depth
			If WD = 0# Then WD = CurVessel.WaterDepth
		End With

        Call DrawPolarPlot(WX, WY, WD, X0, Y0, Head0, ThreeD)

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

    End Sub
	
	Private Sub PanChange()
		
		Dim V, Vmax As Short
		Dim H, Hmax As Short
		Dim CurWell As Well
		
		With CurField
			CurWell = .Item(.CurWellNo)
		End With
		
		H = hPan.Value
		V = vPan.Value
		Hmax = (hPan.Maximum - hPan.LargeChange + 1)
		Vmax = (vPan.Maximum - vPan.LargeChange + 1)

        anObj2.setGraph(ThreeD, ThreeD.Width * (Hmax - H) / Hmax, ThreeD.Height * (Vmax - V) / Vmax, False)

        anRiser.setGraph(ThreeD, ThreeD.Width * (Hmax - H) / Hmax, ThreeD.Height * (Vmax - V) / Vmax, False)
        anSeaBed.setGraph(ThreeD, ThreeD.Width * (Hmax - H) / Hmax, ThreeD.Height * (Vmax - V) / Vmax, False)
        anPipeLines.setGraph(ThreeD, ThreeD.Width * (Hmax - H) / Hmax, ThreeD.Height * (Vmax - V) / Vmax, False)
        anSurface.setGraph(ThreeD, ThreeD.Width * (Hmax - H) / Hmax, ThreeD.Height * (Vmax - V) / Vmax, False)
        anObj2.drawBoat()
		
		With CurVessel
			anRiser.drawRiser(CurWell, .ShipCurGlob, .ShipDraft)
			anSeaBed.drawSeabed(.ShipCurGlob, .ShipDraft)
			anPipeLines.drawPipeLines(.ShipCurGlob, .ShipDraft)
			anSurface.drawSurface(.ShipDraft)
		End With
		
		AddNorthIndicator((-hScroll_Renamed.Value * Degrees2Radians))
		
	End Sub
	
	Private Sub UpdateControls(ByVal SetVal As Boolean)

        btnZoom(0).Visible = SetVal
        btnZoom(1).Visible = SetVal

        hScroll_Renamed.Visible = SetVal
        vScroll_Renamed.Visible = SetVal

        hPan.Enabled = SetVal
        vPan.Enabled = SetVal

        txtPicControls(0).Visible = SetVal
        txtPicControls(1).Visible = SetVal

        'lblUnit(0).Visible = SetVal
        ' lblUnit(1).Visible = SetVal

        btnPlot(0).Enabled = Not SetVal
        btnPlot(1).Enabled = SetVal

    End Sub
	
	Private Function Zoom(ByVal ZoomRatio As Single) As Boolean
		
		Dim Zoom1 As Boolean
		Dim CurWell As Well
		
		With CurField
			CurWell = .Item(.CurWellNo)
		End With
		
		'UPGRADE_WARNING: Couldn't resolve default property of object anObj2.ZoomGraph(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Zoom = anObj2.ZoomGraph(ZoomRatio)
		With CurVessel
			'UPGRADE_WARNING: Couldn't resolve default property of object anRiser.ZoomGraph(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Zoom1 = anRiser.ZoomGraph(ZoomRatio, CurWell, .ShipCurGlob, .ShipDraft)
			'UPGRADE_WARNING: Couldn't resolve default property of object anSeaBed.ZoomGraph(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Zoom1 = anSeaBed.ZoomGraph(ZoomRatio, .ShipCurGlob, .ShipDraft)
			Zoom1 = anPipeLines.ZoomGraph(ZoomRatio, .ShipCurGlob, .ShipDraft)
			'UPGRADE_WARNING: Couldn't resolve default property of object anSurface.ZoomGraph(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Zoom1 = anSurface.ZoomGraph(ZoomRatio, .ShipDraft)
		End With
		
	End Function
	Private Sub hScroll_Renamed_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ScrollEventArgs) Handles hScroll_Renamed.Scroll
		Select Case eventArgs.type
			Case System.Windows.Forms.ScrollEventType.EndScroll
				hScroll_Change(eventArgs.newValue)
		End Select
	End Sub
	Private Sub vScroll_Renamed_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ScrollEventArgs) Handles vScroll_Renamed.Scroll
		Select Case eventArgs.type
			Case System.Windows.Forms.ScrollEventType.EndScroll
				vScroll_Change(eventArgs.newValue)
		End Select
	End Sub
	Private Sub vPan_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ScrollEventArgs) Handles vPan.Scroll
		Select Case eventArgs.type
			Case System.Windows.Forms.ScrollEventType.EndScroll
				vPan_Change(eventArgs.newValue)
		End Select
	End Sub
	Private Sub hPan_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ScrollEventArgs) Handles hPan.Scroll
		Select Case eventArgs.type
			Case System.Windows.Forms.ScrollEventType.EndScroll
				hPan_Change(eventArgs.newValue)
		End Select
	End Sub
End Class