Option Strict Off
Option Explicit On
Friend Class frmMove
	Inherits System.Windows.Forms.Form
	
	Private CheckingGrid As Boolean
	
	Private PayoutLabels(5) As String
	Private LCLabels(2) As String
	
	Private NumLines As Short
	Private MoorLines As MoorSystem
	Private ShipTarLoc As ShipGlobal
	
	'UPGRADE_WARNING: Lower bound of array Payout was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Private Payout(MaxNumLines) As Object
	
	Private Sub btnReport_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnReport.Click
		' check if has data
		If CurVessel.MoorSystem.MoorLineCount = 0 Then
			MsgBox("Mooring system is not yet defined.", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Error")
			Exit Sub
		End If
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		Dim ox As New ExcelReporter
		ox.ReportMooringLayout(txtClientName.Text, txtLocationName.Text, CurVessel, MoorLines, ShipTarLoc)
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		
	End Sub
	
	' form load and unload
	
	Private Sub frmMove_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		Dim i As Short
		
		Text = Text & " - " & CurProj.Title
		
		LoadMoorLines()
		RefreshData()
		btnPosition.Enabled = False
		
		txtClientName.Text = frmMain.txtClientName.Text
		txtLocationName.Text = frmMain.txtLocationName.Text
	End Sub
    Private Sub SetUnitLabels()
        If IsMetricUnit Then
            _lblLengthUnit_0.Text = "m"
            _lblLengthUnit_1.Text = "m"
            _lblLengthUnit_2.Text = "m"
            _lblLengthUnit_3.Text = "m"
            _lblLengthUnit_4.Text = "m"
            _lblLengthUnit_5.Text = "m"
            _lblLengthUnit_6.Text = "m"
        Else
            _lblLengthUnit_0.Text = "ft"
            _lblLengthUnit_1.Text = "ft"
            _lblLengthUnit_2.Text = "ft"
            _lblLengthUnit_3.Text = "ft"
            _lblLengthUnit_4.Text = "ft"
            _lblLengthUnit_5.Text = "ft"
            _lblLengthUnit_6.Text = "ft"
        End If
    End Sub

    Private Sub RefreshData()
		InitiateGrid()
		LoadData((True))
        SetUnitLabels()
    End Sub
	
	Private Sub frmMove_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		
		'UPGRADE_NOTE: Object MoorLines may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		MoorLines = Nothing
		'UPGRADE_NOTE: Object ShipTarLoc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		ShipTarLoc = Nothing
		
	End Sub
	
	' command buttons
	
	Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click
		
		Me.Close()
		
	End Sub
	
	Private Sub btnSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnSave.Click
		
		SaveMoorLines()
		CurProj.Saved = False
		frmMain.LoadData()
		Me.Close()
		
	End Sub
	
	Private Sub btnPayout_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnPayout.Click
		Dim Index As Short = btnPayout.GetIndex(eventSender)
		If IsMetricUnit Then
			LFactor = 0.3048 ' ft -> m
			FrcFactor = 4.448222 ' kips -> KN
			StressFactor = 6.894757 ' ksi -> MPa
			DiameterFactor = 2.54 ' in -> cm
			LUnit = "m"
			FrcUnit = " (KN)"
		Else
			LFactor = 1
			FrcFactor = 1
			StressFactor = 1
			DiameterFactor = 1
			LUnit = "ft"
			FrcUnit = " (kips)"
		End If
		
		Dim TmpStr As String
        Dim TopTension As Single

        Select Case Index
            Case 0
                SaveLC()
                '       LoadData
                MoveVessel()
                UpdateData()
            Case 1
                TmpStr = CStr(CDbl(InputBox("Target Tension" & FrcUnit, "MARS")) / FrcFactor)
                TopTension = CDbl(CheckData(TmpStr)) * 1000.0#
                If TopTension = 0# Then Exit Sub
                SaveLC()
                MoveVessel1(TopTension)
                UpdateData()
        End Select

        btnPosition.Enabled = True
    End Sub

    Private Sub btnPosition_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnPosition.Click

        ReLoadData() ' only load current vessel location

    End Sub

    ' grids operation

    '  Private Sub grdLC_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles grdLC.DblClick

    '  With grdLC
    ' Select Case .Col
    'Case 1
    'If .Text = "Yes" Then
    '.Text = "No"
    'Else
    '.Text = "Yes"
    'End If
    '               MoorLines.MoorLines(.Row).Connected = Not MoorLines.MoorLines(.Row).Connected
    'Case 2
    'If .Text = "Yes" Then
    '.Text = "No"
    'Else
    '.Text = "Yes"
    'End If
    '               MoorLines.MoorLines(.Row).Winched = Not MoorLines.MoorLines(.Row).Winched
    'End Select
    'End With
    '
    ' ************  Temporarily Disabled 9/21/99   *******************
    '        Exit Sub
    '        If grdLC.Text = "Yes" Then
    '            grdLC.Text = "No"
    '        ElseIf grdLC.Text = "No" Then
    '            grdLC.Text = "Yes"
    '        End If
    '        TensionFixed(grdLC.Row) = Not TensionFixed(grdLC.Row)

    '  End Sub

    ' combo box

    'UPGRADE_WARNING: Event cboWells.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub cboWells_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboWells.SelectedIndexChanged
        Dim Index As Short = cboWells.GetIndex(eventSender)

        With CurField
            .CurWellNo = cboWells(Index).SelectedIndex + 1
            cboWells(1 - Index).SelectedIndex = .CurWellNo - 1
        End With

    End Sub

    ' operation subroutines
    ' initiating

    Private Sub InitiateGrid()

        Dim r, c As Short
        Dim ColW As Single
        Dim RShown, RowH, BarSize As Short

        CheckingGrid = True

        Call SetLabels()
        If NumLines > 8 Then
            'BarSize = SysInfo1.ScrollBarSize
        Else
            BarSize = 0
        End If

        'With grdLC
        ' .Rows.count
        '.Rows = Max(9, NumLines + 1)
        '.WordWrap = True

        'ColW = (VB6.PixelsToTwipsX(.Width) - BarSize) / .Cols - .GridLineWidth
        'RowH = VB6.PixelsToTwipsY(.Height) / 11 - .GridLineWidth
        '.set_ColWidth(0, VB6.PixelsToTwipsX(.Width) - BarSize - ColW * (.Cols - 1) - 100)
        '.set_RowHeight(0, VB6.PixelsToTwipsY(.Height) - RowH * 8 - 90)

        'For c = 1 To .Cols - 1
        '.set_ColWidth(c, ColW)
        'Next c
        'For r = .FixedRows To .Rows - 1
        '.set_RowHeight(r, RowH)
        'Next r

        '.Col = 0
        'For r = .FixedRows To .Rows - 1
        '.Row = r
        '.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
        '.Text = "Line " & (r - .FixedRows + 1)
        'Next r
        '.Row = 0
        'For c = 0 To .Cols - 1
        '.Col = c
        '.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
        '.Text = LCLabels(c)
        'Next c

        '.Col = 0
        '.Row = 0
        'End With

        'With grdPayout
        '.Rows = Max(9, NumLines + 1)
        '.WordWrap = True
        '
        'ColW = (VB6.PixelsToTwipsX(.Width) - BarSize) / .Cols - .GridLineWidth
        'RowH = VB6.PixelsToTwipsY(.Height) / 11 - .GridLineWidth
        '.set_ColWidth(0, VB6.PixelsToTwipsX(.Width) - BarSize - ColW * (.Cols - 1) - 80)
        '.set_RowHeight(0, VB6.PixelsToTwipsY(.Height) - RowH * 8 - 90)
        '
        'For c = 1 To .Cols - 1
        '.set_ColWidth(c, ColW)
        'Next c
        'For r = .FixedRows To .Rows - 1
        '.set_RowHeight(r, RowH)
        'Next r

        '.Col = 0
        'For r = .FixedRows To .Rows - 1
        '.Row = r
        '.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
        '.Text = "Line " & (r - .FixedRows + 1)
        'Next r
        '.Row = 0
        'For c = 0 To .Cols - 1
        '.Col = c
        '.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
        '.Text = PayoutLabels(c)
        'Next c

        '.Col = 0
        '.Row = 0
        'End With

        CheckingGrid = False

    End Sub

    Private Sub SetLabels()
        Dim LUnit, FrcUnit As String
        If IsMetricUnit Then
            LUnit = "(m)"
            FrcUnit = "(KN)"
        Else
            LUnit = "(ft)"
            FrcUnit = "(kips)"
        End If

        LCLabels(0) = "Line No."
        LCLabels(1) = "Line Connected"
        LCLabels(2) = "Winch Functional"

        PayoutLabels(0) = "Line No."
        PayoutLabels(1) = "Current Payout " & LUnit
        PayoutLabels(2) = "New Payout " & LUnit
        PayoutLabels(3) = "Changes " & LUnit
        PayoutLabels(4) = "Current Top Tension " & FrcUnit
        PayoutLabels(5) = "New Top Tension " & FrcUnit

    End Sub

    ' operation subroutines
    ' load to and update from form

    Private Sub LoadData(Optional ByVal FirstTime As Boolean = False)

        If IsMetricUnit Then
            LFactor = 0.3048 ' ft -> m
            FrcFactor = 4.448222 ' kips -> KN
        Else
            LFactor = 1
            FrcFactor = 1
        End If


        Dim r As Short
        Dim TmpStr As String
        Dim FMGlob As Force
        Dim Distance, Bearing As Single

        If FirstTime Then
            With MoorLines
                txtVslSt(0).Text = Format(.ShipGlob.Xg * LFactor, "0.0")
                txtVslSt(1).Text = Format(.ShipGlob.Yg * LFactor, "0.0")
                txtVslSt(4).Text = Format(.ShipGlob.Heading * Radians2Degrees, "0.00")
                txtVslSt(5).Text = Format(.ShipDraft * LFactor, "0.00")

                Coord2Bear(.ShipGlob, Distance, Bearing)
                txtVslSt(2).Text = Format(Distance * LFactor, "0.0")
                txtVslSt(3).Text = Format(Bearing * Radians2Degrees, "0.00")

                txtVslSt(6).Text = txtVslSt(0).Text
                txtVslSt(7).Text = txtVslSt(1).Text
                txtVslSt(10).Text = txtVslSt(4).Text

                txtVslSt(8).Text = txtVslSt(2).Text
                txtVslSt(9).Text = txtVslSt(3).Text
            End With
            optInputSystem(0).Checked = True
            optInputSystem(2).Checked = True

            If MoorLines.WinchCap <= 0# Then
                txtConditions.Text = " "
            Else
                txtConditions.Text = Format(MoorLines.WinchCap * 0.0018 * FrcFactor, "0.0")
            End If
        End If

        FMGlob = MoorLines.FMoorGlob
        With FMGlob
            txtExtLoad(0).Text = Format(.Fx * 0.001 * FrcFactor, "0.0")
            txtExtLoad(1).Text = Format(.Fy * 0.001 * FrcFactor, "0.0")
            txtExtLoad(2).Text = Format(.MYaw * 0.001 * FrcFactor, "0.0")
        End With

        ' With grdLC
        'For r = 1 To NumLines
        '        .Row = r
        '       .Col = 1
        '      If MoorLines.MoorLines(r).Connected Then
        '     .CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
        '    .Text = "Yes"
        '   Else
        '  .CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
        ' .Text = "No"
        'End If
        '.Col = 2
        'If MoorLines.MoorLines(r).Winched Then
        '.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
        '.Text = "Yes"
        'Else
        '.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
        '.Text = "No"
        'End If
        'Next r
        ' End With

        ' With grdPayout
        'For r = 1 To NumLines
        ''.Row = r
        'If MoorLines.MoorLines(r).Connected Then
        'TmpStr = Format(MoorLines.MoorLines(r).Payout * LFactor, "0.0")
        'Else
        'TmpStr = "--"
        'End If
        '.Col = 1
        '.Text = TmpStr
        '.Col = 2
        '.Text = TmpStr
        '.Col = 3
        'If MoorLines.MoorLines(r).Connected Then
        '.Text = "0.0"
        'TmpStr = Format(MoorLines.MoorLines(r).TopTension * 0.001 * FrcFactor, "0.00")
        'Else
        '.Text = "--"
        'End If
        '.Col = 4
        '.Text = TmpStr
        '.Col = 5
        '.Text = TmpStr
        'Next r
        'End With

        'UPGRADE_NOTE: Object FMGlob may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'FMGlob = Nothing

    End Sub

    Private Sub ReLoadData()

        Dim r, c As Short
        Dim TmpStr As String
        Dim FMGlob As Force

        Dim Distance, Bearing As Single

        With MoorLines
            Coord2Bear(.ShipGlob, Distance, Bearing)

            With .ShipGlob
                .Xg = ShipTarLoc.Xg
                .Yg = ShipTarLoc.Yg
                .Heading = ShipTarLoc.Heading

                txtVslSt(0).Text = Format(.Xg * LFactor, "0.0")
                txtVslSt(1).Text = Format(.Yg * LFactor, "0.0")
                txtVslSt(2).Text = Format(Distance * LFactor, "0.0")
                txtVslSt(3).Text = Format(Bearing * Radians2Degrees, "0.00")
                txtVslSt(4).Text = Format(.Heading * Radians2Degrees, "0.0")
            End With

            For r = 1 To NumLines
                'UPGRADE_WARNING: Couldn't resolve default property of object Payout(r). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Payout(r) = .MoorLines(r).Payout
            Next r
        End With

        FMGlob = MoorLines.FMoorGlob
        With FMGlob
            txtExtLoad(0).Text = Format(.Fx * 0.001 * FrcFactor, "0.0")
            txtExtLoad(1).Text = Format(.Fy * 0.001 * FrcFactor, "0.0")
            txtExtLoad(2).Text = Format(.MYaw * 0.001 * FrcFactor * LFactor, "0.0")
        End With

        'UPGRADE_NOTE: Object FMGlob may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        FMGlob = Nothing

    End Sub

    Private Sub SaveLC()

        Dim r As Short
        Dim Distance, Bearing As Single



        With MoorLines
            If optInputSystem(0).Checked Then
                With .ShipGlob
                    .Xg = CDbl(CheckData(txtVslSt(0).Text,  , True)) / LFactor
                    .Yg = CDbl(CheckData(txtVslSt(1).Text,  , True)) / LFactor
                    .Heading = CDbl(CheckData(txtVslSt(4).Text,  , True)) * Degrees2Radians
                End With
                Coord2Bear(.ShipGlob, Distance, Bearing)
                txtVslSt(2).Text = Format(Distance / LFactor, "0.0")
                txtVslSt(3).Text = Format(Bearing * Radians2Degrees, "0.00")
            Else
                Distance = CDbl(CheckData(txtVslSt(2).Text,  , True)) / LFactor
                Bearing = CDbl(CheckData(txtVslSt(3).Text,  , True)) * Degrees2Radians
                Bear2Coord(.ShipGlob, Distance, Bearing)
                With .ShipGlob
                    txtVslSt(0).Text = Format(.Xg * LFactor, "0.0")
                    txtVslSt(1).Text = Format(.Yg * LFactor, "0.0")
                    .Heading = CDbl(CheckData(txtVslSt(4).Text,  , True)) * Degrees2Radians
                End With
            End If

            .WinchCap = CDbl(CheckData(txtConditions.Text,  , True)) * 1000.0# / FrcFactor
        End With

        If optInputSystem(2).Checked Then
            With ShipTarLoc
                .Xg = CDbl(CheckData(txtVslSt(6).Text,  , True)) / LFactor
                .Yg = CDbl(CheckData(txtVslSt(7).Text,  , True)) / LFactor
                .Heading = CDbl(CheckData(txtVslSt(10).Text,  , True)) * Degrees2Radians
            End With
            Coord2Bear(ShipTarLoc, Distance, Bearing)
            txtVslSt(8).Text = Format(Distance * LFactor, "0.0")
            txtVslSt(9).Text = Format(Bearing * Radians2Degrees, "0.00")
        Else
            Distance = CDbl(CheckData(txtVslSt(8).Text,  , True)) / LFactor
            Bearing = CDbl(CheckData(txtVslSt(9).Text,  , True)) * Degrees2Radians
            Bear2Coord(ShipTarLoc, Distance, Bearing)
            With ShipTarLoc
                txtVslSt(6).Text = Format(.Xg * LFactor, "0.0")
                txtVslSt(7).Text = Format(.Yg * LFactor, "0.0")
                .Heading = CDbl(CheckData(txtVslSt(10).Text,  , True)) * Degrees2Radians
            End With
        End If

    End Sub

    Private Sub UpdateData()

        Dim r, c As Short
        Dim TmpStr As String
        Dim OldPOL As Single
        Dim FMGlob As New Force


        MoorLines.MoorForce(FMGlob, ShipTarLoc)
        With FMGlob
            txtExtLoad(3).Text = Format(.Fx * 0.001 * FrcFactor, "0.0")
            txtExtLoad(4).Text = Format(.Fy * 0.001 * FrcFactor, "0.0")
            txtExtLoad(5).Text = Format(.MYaw * 0.001 * FrcFactor * LFactor, "0.0")
        End With

        'UPGRADE_NOTE: Object FMGlob may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        FMGlob = Nothing

    End Sub


    Private Sub LoadMoorLines()

        'UPGRADE_NOTE: Segment was upgraded to Segment_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Dim Segment_Renamed As Segment
        'UPGRADE_NOTE: MoorLine was upgraded to MoorLine_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Dim MoorLine_Renamed As MoorLine
        Dim i As Short
        Dim NumWells As Short

        Dim SegTp As String
        Dim Lg, dia As Single
        Dim E1, BS, E2 As Single
        Dim Buoy, DryWt, WetWt, BuoyL As Single
        Dim FrCoef As Single

        MoorLines = New MoorSystem
        ShipTarLoc = New ShipGlobal

        With CurVessel
            NumLines = .MoorSystem.MoorLineCount
            MoorLines.ShipDraft = .ShipDraft
        End With

        NumWells = CurField.Count
        For i = 1 To NumWells
            With CurField.Item(i)
                cboWells(0).Items.Add(.NameID)
                cboWells(1).Items.Add(.NameID)
            End With
        Next i
        cboWells(0).SelectedIndex = CurField.CurWellNo - 1
        cboWells(1).SelectedIndex = CurField.CurWellNo - 1

        With MoorLines.ShipGlob
            .Xg = CurVessel.ShipCurGlob.Xg
            .Yg = CurVessel.ShipCurGlob.Yg
            .Heading = CurVessel.ShipCurGlob.Heading
        End With

        For i = 1 To NumLines Step 1
            MoorLines.MoorLineAdd()
            MoorLine_Renamed = CurVessel.MoorSystem.MoorLines(i)

            With MoorLines.MoorLines(i)
                .BottomSlope = MoorLine_Renamed.BottomSlope
                .Connected = True
                .DesScope = MoorLine_Renamed.DesScope

                'UPGRADE_WARNING: Couldn't resolve default property of object Payout(i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Payout(i) = MoorLine_Renamed.Payout
                'UPGRADE_WARNING: Couldn't resolve default property of object Payout(i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Payout = Payout(i)
                .PayoutSur = MoorLine_Renamed.PayoutSur
                .PayoutOpr = MoorLine_Renamed.PayoutOpr
                .PretensionSur = MoorLine_Renamed.PretensionSur
                .PretensionOpr = MoorLine_Renamed.PretensionOpr

                .FairLead.SprdAngle = MoorLine_Renamed.FairLead.SprdAngle
                .FairLead.Xs = MoorLine_Renamed.FairLead.Xs
                .FairLead.Ys = MoorLine_Renamed.FairLead.Ys
                .FairLead.z = MoorLine_Renamed.FairLead.z

                .Anchor.Xg = MoorLine_Renamed.Anchor.Xg
                .Anchor.Yg = MoorLine_Renamed.Anchor.Yg
                .WaterDepth = MoorLine_Renamed.WaterDepth
            End With

            For Each Segment_Renamed In MoorLine_Renamed
                With Segment_Renamed
                    SegTp = .SegType
                    Lg = .Length
                    dia = .Diameter
                    BS = .BS
                    E1 = .E1
                    E2 = .E2
                    DryWt = .UnitDryWeight
                    WetWt = .UnitWetWeight
                    Buoy = .Buoy
                    BuoyL = .BuoyLength
                    FrCoef = .FrictionCoef
                End With
                Call MoorLines.MoorLines(i).SegmentAdd(SegTp, Lg, 100, dia, BS, E1, E2, DryWt, WetWt, Buoy, BuoyL, FrCoef)
            Next Segment_Renamed
        Next

    End Sub

    Private Sub SaveMoorLines()

        'UPGRADE_NOTE: Segment was upgraded to Segment_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Dim Segment_Renamed As Segment
        'UPGRADE_NOTE: MoorLine was upgraded to MoorLine_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Dim MoorLine_Renamed As MoorLine
        Dim i As Short

        With CurVessel.ShipCurGlob
            .Xg = ShipTarLoc.Xg
            .Yg = ShipTarLoc.Yg
            .Heading = ShipTarLoc.Heading
        End With

        With CurVessel.MoorSystem
            For i = 1 To NumLines Step 1
                With .MoorLines(i)
                    .Payout = MoorLines.MoorLines(i).Payout
                    .Connected = MoorLines.MoorLines(i).Connected
                    .Winched = MoorLines.MoorLines(i).Winched
                End With
            Next i
        End With

    End Sub

    ' operational subroutine
    ' calculate payout adjustment by moving

    Public Sub CalcPayout(ByRef Index As Short, ByRef TargetLoc As ShipGlobal)

        If IsMetricUnit Then
            LFactor = 0.3048 ' ft -> m
            FrcFactor = 4.448222 ' kips -> KN
            StressFactor = 6.894757 ' ksi -> MPa
            DiameterFactor = 2.54 ' in -> cm
            LUnit = "m"
            FrcUnit = " (KN)"
        Else
            LFactor = 1
            FrcFactor = 1
            StressFactor = 1
            DiameterFactor = 1
            LUnit = "ft"
            FrcUnit = " (kips)"
        End If

        Dim TmpStr As String
        Dim TopTension As Single

        ShipTarLoc = TargetLoc

        Select Case Index
            Case 0
                LoadData()
                MoveVessel()
                UpdateData()
            Case 1
                TmpStr = InputBox("Target Tension" & FrcUnit & ":", "MARS", "0")
                TopTension = CDbl(CheckData(TmpStr)) * 1000.0# / FrcFactor
                If TopTension = 0# Then
                    Me.Close()
                Else
                    MoveVessel1(TopTension)
                    UpdateData()
                End If
        End Select

    End Sub

    Private Sub MoveVessel()

        Dim RunAgain As Boolean

        Dim j, Kab, i, k As Short
        Dim FMGlob As New Force

        Dim SpAngGlob As Single
        'UPGRADE_WARNING: Lower bound of array OldHorFrc was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
        Dim OldHorFrc(MaxNumLines) As Single
        'UPGRADE_WARNING: Lower bound of array NewHorFrc was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
        Dim NewHorFrc(MaxNumLines) As Single
        'UPGRADE_WARNING: Lower bound of array Cp was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
        Dim Cp(MaxNumLines) As Single
        'UPGRADE_WARNING: Lower bound of array Cn was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
        Dim Cn(MaxNumLines) As Single
        'UPGRADE_WARNING: Lower bound of array Sp was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
        Dim Sp(MaxNumLines) As Single
        'UPGRADE_WARNING: Lower bound of array Sn was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
        Dim Sn(MaxNumLines) As Single
        'UPGRADE_WARNING: Lower bound of array Zp was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
        Dim Zp(MaxNumLines) As Single
        'UPGRADE_WARNING: Lower bound of array Zn was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
        Dim Zn(MaxNumLines) As Single
        'UPGRADE_WARNING: Lower bound of array Dh was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
        Dim Dh(MaxNumLines) As Single
        'UPGRADE_WARNING: Lower bound of array Iuse was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
        Dim Iuse(MaxNumLines) As Boolean

        'UPGRADE_WARNING: Lower bound of array C33 was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
        Dim C33(3, 3) As Single
        'UPGRADE_WARNING: Lower bound of array Ff was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
        Dim Ff(3, 1) As Single
        'UPGRADE_WARNING: Lower bound of array Fm was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
        Dim Fm(3) As Single
        'UPGRADE_WARNING: Lower bound of array A3n was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
        Dim A3n(3, MaxNumLines) As Single
        'UPGRADE_WARNING: Lower bound of array An3 was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
        Dim An3(MaxNumLines, 3) As Single

        'UPGRADE_WARNING: Lower bound of array Scope was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
        Dim Scope(MaxNumLines) As Single
        'UPGRADE_WARNING: Lower bound of array Payout was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
        Dim Payout(MaxNumLines) As Single

        '   Initialize the "Iuse" array, whose purpose is to flag any lines for which
        '   a solution won't be sought.  This gives the ability to set the tension for
        '   one or more lines, then determine the other line tensions given those
        '   constraints
        For i = 1 To NumLines
            With MoorLines.MoorLines(i)
                OldHorFrc(i) = .TopTension * System.Math.Cos(.FLAngle)

                '           Calculate the previous matrix A
                SpAngGlob = PI * 0.5 - .SprdAngle((MoorLines.ShipGlob)) - MoorLines.ShipGlob.Heading
                Cp(i) = System.Math.Cos(SpAngGlob)
                Sp(i) = System.Math.Sin(SpAngGlob)
                Zp(i) = .FairLead.Xs * System.Math.Sin(- .SprdAngle((MoorLines.ShipGlob))) - .FairLead.Ys * System.Math.Cos(- .SprdAngle((MoorLines.ShipGlob)))

                If Not .Connected Then
                    Dh(i) = 0#
                    Iuse(i) = False
                ElseIf Not .Winched Then
                    Iuse(i) = False
                Else
                    Iuse(i) = True
                End If
            End With
        Next i

        '   for fixed horizontal tension. it works!
        '        FixedLine = 1
        '        Iuse(FixedLine) = False
        '        Dh(FixedLine) = 0#

        '   Remember the trig functions and the moment arms for later use
        Ff(1, 1) = 0#
        Ff(2, 1) = 0#
        Ff(3, 1) = 0#

Again:
        MoorLines.MoorForce(FMGlob, ShipTarLoc)

        For i = 1 To NumLines
            '       calculate new matrix A
            With MoorLines.MoorLines(i)
                NewHorFrc(i) = .TopTension * System.Math.Cos(.FLAngle)
                Scope(i) = .ScopeByVesselLocation(ShipTarLoc)

                SpAngGlob = PI * 0.5 - .SprdAngle(ShipTarLoc) - MoorLines.ShipGlob.Heading
                Cn(i) = System.Math.Cos(SpAngGlob)
                Sn(i) = System.Math.Sin(SpAngGlob)
                Zn(i) = .FairLead.Xs * System.Math.Sin(- .SprdAngle(ShipTarLoc)) - .FairLead.Ys * System.Math.Cos(- .SprdAngle(ShipTarLoc))

                If Not .Connected Then
                    '               calculate unbalanced forces caused by disconnecting a line
                    Ff(1, 1) = Ff(1, 1) + OldHorFrc(i) * Cp(i)
                    Ff(2, 1) = Ff(2, 1) + OldHorFrc(i) * Sp(i)
                    Ff(3, 1) = Ff(3, 1) + OldHorFrc(i) * Zp(i)
                ElseIf Not .Winched Then
                    '               calculate unbalanced forces caused by unfunctional winches
                    Ff(1, 1) = Ff(1, 1) + OldHorFrc(i) * Cp(i) - NewHorFrc(i) * Cn(i)
                    Ff(2, 1) = Ff(2, 1) + OldHorFrc(i) * Sp(i) - NewHorFrc(i) * Sn(i)
                    Ff(3, 1) = Ff(3, 1) + OldHorFrc(i) * Zp(i) - NewHorFrc(i) * Zn(i)
                    Dh(i) = NewHorFrc(i) - OldHorFrc(i)
                Else
                    '               calculate unbalanced forces caused by change of spreading angles
                    Ff(1, 1) = Ff(1, 1) + OldHorFrc(i) * (Cp(i) - Cn(i))
                    Ff(2, 1) = Ff(2, 1) + OldHorFrc(i) * (Sp(i) - Sn(i))
                    Ff(3, 1) = Ff(3, 1) + OldHorFrc(i) * (Zp(i) - Zn(i))
                End If
            End With
        Next i

        Fm(1) = Ff(1, 1)
        Fm(2) = Ff(2, 1)
        Fm(3) = Ff(3, 1)

        '   A3n is an intermediate array used in the building of the force matrix
        '   that will ultimately produce the changes in cable length which we seek

        Kab = 0
        For i = 1 To NumLines
            '       Make sure that we want to use this line
            If Iuse(i) Then
                '           If so, update the counter telling us how many unknowns there are
                Kab = Kab + 1
                '           Set each force component (X, Y, and Yaw Moment)
                A3n(1, Kab) = Cn(i)
                A3n(2, Kab) = Sn(i)
                A3n(3, Kab) = Zn(i)
            End If
        Next i

        '   Set up the array representing the matrix [A]*[A]Transpose
        For i = 1 To 3
            For j = 1 To 3
                C33(i, j) = 0#
            Next j
        Next i

        For i = 1 To NumLines
            If Iuse(i) Then
                C33(1, 1) = C33(1, 1) + Cn(i) * Cn(i)
                C33(1, 2) = C33(1, 2) + Cn(i) * Sn(i)
                C33(1, 3) = C33(1, 3) + Cn(i) * Zn(i)
                C33(2, 1) = C33(2, 1) + Sn(i) * Cn(i)
                C33(2, 2) = C33(2, 2) + Sn(i) * Sn(i)
                C33(2, 3) = C33(2, 3) + Sn(i) * Zn(i)
                C33(3, 1) = C33(3, 1) + Zn(i) * Cn(i)
                C33(3, 2) = C33(3, 2) + Zn(i) * Sn(i)
                C33(3, 3) = C33(3, 3) + Zn(i) * Zn(i)
            End If
        Next i

        '   Invert the coeficient matrix, solving for the unknown vector
        Call GaussJordan(C33, 3, Ff, 1)

        '   Compute the final coefficients
        Kab = 0
        For i = 1 To NumLines
            If Iuse(i) Then
                Kab = Kab + 1
                For j = 1 To 3
                    An3(Kab, j) = 0#
                    For k = 1 To 3
                        An3(Kab, j) = An3(Kab, j) + A3n(k, Kab) * C33(k, j)
                    Next k
                Next j
            End If
        Next i

        '   Calculate the differential forces for each line
        Kab = 0
        For i = 1 To NumLines
            If Iuse(i) Then
                Kab = Kab + 1
                Dh(i) = 0#
                For k = 1 To 3
                    Dh(i) = Dh(i) + An3(Kab, k) * Fm(k)
                Next k
            End If
            '       And the new horizontal force for each line
            NewHorFrc(i) = OldHorFrc(i) + Dh(i)
        Next i

        '   Compute the tension and payout length associated with these forces
        Ff(1, 1) = 0#
        Ff(2, 1) = 0#
        Ff(3, 1) = 0#
        RunAgain = False
        For i = 1 To NumLines
            With MoorLines.MoorLines(i)
                If .Winched Then
                    If Iuse(i) Then
                        If NewHorFrc(i) < 0# Then
                            .Connected = False
                            Iuse(i) = False
                            RunAgain = True
                        Else
                            Payout(i) = .POLByScopeFrcHor(Scope(i), NewHorFrc(i))
                            If Payout(i) < 0# Then
                                If Payout(i) = -1.0# Then
                                    .Payout = 0#
                                Else
                                    .Payout = -Payout(i)
                                End If
                                .Winched = False
                                Iuse(i) = False
                                RunAgain = True
                            Else
                                .Payout = Payout(i)
                                If Not MoorLines.WinchCap = 0# Then
                                    Payout(i) = .PayoutByScopeTopTension(Scope(i), (MoorLines.WinchCap))
                                    If Payout(i) > .Payout Then
                                        .Winched = False
                                        .Payout = Payout(i)
                                        Iuse(i) = False
                                        RunAgain = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End With
        Next i

        If RunAgain Then GoTo Again
        MoorLines.MoorForce(FMGlob, ShipTarLoc)

    End Sub

    Private Sub MoveVessel1(ByVal Pretension As Single)

        Dim i As Short
        Dim Payout, Scope As Single

        For i = 1 To NumLines
            With MoorLines.MoorLines(i)
                Scope = .ScopeByVesselLocation(ShipTarLoc, True)
                Payout = .PayoutByScopeTopTension(Scope, Pretension)
                If Payout < 0# Then
                    If Payout = -1.0# Then
                        .Payout = 0#
                    Else
                        .Payout = -Payout
                    End If
                Else
                    .Payout = Payout
                End If
            End With
        Next i

    End Sub

    Private Sub Coord2Bear(ByRef ShipGlob As ShipGlobal, ByRef Distance As Single, ByRef Bearing As Single)

        Dim dx, dy As Single

        With CurField
            dx = ShipGlob.Xg - .Item(.CurWellNo).Xg
            dy = ShipGlob.Yg - .Item(.CurWellNo).Yg
        End With

        Distance = System.Math.Sqrt(dx ^ 2 + dy ^ 2)
        Bearing = Atan(dy, dx)

    End Sub

    Private Sub Bear2Coord(ByRef ShipGlob As ShipGlobal, ByRef Distance As Single, ByRef Bearing As Single)

        Dim dx, dy As Single

        dx = Distance * System.Math.Sin(Bearing)
		dy = Distance * System.Math.Cos(Bearing)
		
		With CurField
			ShipGlob.Xg = .Item(.CurWellNo).Xg + dx
			ShipGlob.Yg = .Item(.CurWellNo).Yg + dy
		End With
		
	End Sub
	
	'UPGRADE_WARNING: Event txtVslSt.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtVslSt_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtVslSt.TextChanged
		Dim Index As Short = txtVslSt.GetIndex(eventSender)
		btnPosition.Enabled = False
	End Sub

    Private Sub frmMove_Load_1(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class