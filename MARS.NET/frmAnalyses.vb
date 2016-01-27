Option Strict Off
Option Explicit On
Friend Class frmAnalyses
	Inherits System.Windows.Forms.Form
	' frmAnalyses           analysis results (quasi-static)
	' Version 2.1
	' 2001, Copyright DTCEL, All Rights Reserved
	
	
	Private SaveLoc As Boolean
	Private CheckingGrid As Boolean
	
	Private MoveLabels(6) As String
	Private TenLabels(7) As String
	Private DirLabels(2) As String
	
	Private NumLines As Short
	Private ShipLoc As ShipGlobal
	Private Movement, LFM, WFM, MaxDM As Motion
    Private SgnY, SgnX, SgnZ As Single
    Private Tension(MaxNumLines) As Single
    Private TSigLF(MaxNumLines) As Single
    Private TSigWF(MaxNumLines) As Single
    Private PreConnect(MaxNumLines) As Boolean

    Private Stiff As Force
    Private WFMd, LFMd As Motion
    Dim FrcFactor, LFactor, VelFactor As Single
    Dim LUnit, VelUnit As String

    Private Sub btnReport_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnReport.Click
        On Error GoTo ErrHandler

        ' assume always perform no env first and save
        Dim InitPos As New Motion
        InitPos.Surge = 0
        InitPos.Sway = 0
        InitPos.Yaw = 0

        If LFM Is Nothing Then
            MsgBox("Must perform analysis first.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Error")
            Exit Sub
        End If
        ' display report
        Dim ox As New ExcelReporter
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Call ox.ReportMooringAnalysisResults(txtReportTitle.Text, (txtSubTitle.Text), 1, True, CurVessel, Movement, InitPos, Stiff, LFM, WFM, Tension, TSigLF, TSigWF)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Exit Sub
ErrHandler:
        MsgBox("Err creating report...")
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub

    ' form load and unload

    Private Sub frmAnalyses_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        Text = Text & " - " & CurProj.Title

        NumLines = CurVessel.MoorSystem.MoorLineCount

        ShipLoc = New ShipGlobal
        Stiff = New Force
        WFMd = New Motion
        LFMd = New Motion
        MaxDM = New Motion

        Cancelled = True
        SaveLoc = False

        btnDerp.Enabled = False
        btnReport.Enabled = False
        txtReportTitle.Text = CurVessel.Name & ", Draft " & Format(CurVessel.ShipDraft, CStr(0)) & " ft, " & CurProj.WellSites.Item(CurProj.WellSites.CurWellNo).NameID
        txtSubTitle.Text = txtEnvironment.Text & ", " & _txtPercentDamping_0.Text & "% Damping"
        RefreshData()
    End Sub

    Private Sub RefreshData()
        RefreshUnitLabels(Me)
        InitiateGrid()
        LoadData()
    End Sub
    Private Sub frmAnalyses_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        ShipLoc = Nothing
        Stiff = Nothing
        WFMd = Nothing
        LFMd = Nothing
        MaxDM = Nothing

        With frmMain
            .Enabled = True
            .Activate()
            If SaveLoc Then .LoadData()
        End With

    End Sub

    ' command buttons

    Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click

        Dim i As Short
        '
        With CurVessel.MoorSystem
            For i = 1 To NumLines
                .MoorLines(i).Connected = PreConnect(i)
            Next i
        End With
        ' set back to original ship location?		'
        Me.Close()

    End Sub

    Private Sub btnAnalysis_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnAnalysis.Click

        If System.Math.Abs(Val(_txtVslSt_3.Text)) < 5 Then
            MsgBox("Draft too small.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Error")
            Exit Sub
        End If


        On Error GoTo ErrorHandler

        Cursor = System.Windows.Forms.Cursors.WaitCursor
        Me.Enabled = False

        Cancelled = False
        With frmProgress
            .Show()
        End With

        With frmProgress
            .Text = "Quasi-static Mooring Analysis"
            .CurrentStage.Text = "Initializing..."
            .Progress = 0
        End With

        SaveLC()
        StaticAnalysis()

        If Cancelled = False Then
            With frmProgress
                .CurrentStage.Text = "Post-processing..."
                .Progress = 99
            End With
            UpdateData()

            With frmProgress
                .CurrentStage.Text = "Completed successfully."
                .Progress = 100
            End With
        End If
        frmProgress.Close()
        Me.Enabled = True
        Me.Visible = True
        Me.Activate()
        btnDerp.Enabled = True
        btnReport.Enabled = True
        Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub
ErrorHandler:
        Me.Enabled = True
        Cursor = System.Windows.Forms.Cursors.Default
        MsgBox("Error: " & Err.Description & " from Static Analysis: " & Err.Source, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Error")

    End Sub

    Private Sub btnSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnSave.Click

        With CurVessel.DampingPercent
            .Surge = Val(_txtPercentDamping_0.Text)
            .Sway = Val(_txtPercentDamping_1.Text)
            .Yaw = Val(_txtPercentDamping_2.Text)
        End With

        '    If MsgBox("Are you sure you want to replace the current vessel position?", _
        ''                vbOKCancel, "MARS - Change Vessel Position?") = vbOK Then

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        With CurVessel.ShipCurGlob
            .Xg = ShipLoc.Xg
            .Yg = ShipLoc.Yg
            .Heading = ShipLoc.Heading
        End With
        SaveLoc = True
        CurProj.Saved = False

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Me.Close()
        '    End If
    End Sub

    Private Sub btnEnvironment_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnEnvironment.Click

        frmEnviron.Show()
        txtEnvironment.Text = CurVessel.EnvLoad.EnvCur.Name
        '  RefreshButtonStatus
    End Sub

    Private Sub btnDerp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnDerp.Click

        If Not DerpData() Then
            MsgBox("Data export failure.", MsgBoxStyle.OkOnly, msgTitle)
        Else
            btnDerp.Enabled = False
        End If

    End Sub

    ' grids operation

    Private Sub grdTensions_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Dim r As Short

        With grdTensions
            For r = 0 To .RowCount - 1
                .Rows(r).Cells(1).Value = "Yes"

                CurVessel.MoorSystem.MoorLines(r + 1).Connected = Not CurVessel.MoorSystem.MoorLines(r + 1).Connected
            Next r
        End With

    End Sub

    Private Sub txtEnvironment_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEnvironment.TextChanged
        txtSubTitle.Text = txtEnvironment.Text & ", " & _txtPercentDamping_0.Text & "% Damping"
    End Sub

    ' text box

    ' Private Sub txtVslSt_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
    'Dim Index As Short = txtVslSt.GetIndex(eventSender)

    '     txtVslSt(Index).Text = CheckData(txtVslSt(Index).Text)

    ' End Sub

    ' operation subroutines
    ' initiating

    Private Sub InitiateGrid()

        Dim r, c As Short

        CheckingGrid = True

        Call SetLabels()

        With grdVM

            .ColumnCount = 7
            .RowCount = 3
            For c = 0 To .ColumnCount - 1
                .Columns(c).HeaderText = MoveLabels(c)
                .Columns(c).FillWeight = 100 / .ColumnCount
                .Columns(c).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            Next c

            For r = 0 To .RowCount - 1
                .Rows(r).Cells(0).Value = DirLabels(r)
            Next

        End With


        With grdTensions
            .RowCount = NumLines
            .ColumnCount = 8
            For c = 0 To .ColumnCount - 1
                .Columns(c).FillWeight = 100 / .ColumnCount
                .Columns(c).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

                .Columns(c).HeaderText = TenLabels(c)
            Next c
            For r = 0 To .RowCount - 1
                .Rows(r).Cells(0).Value = "Line " & r + 1
            Next

        End With

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

        MoveLabels(0) = "Direction"
        MoveLabels(1) = "Mean"
        MoveLabels(2) = "Sig LFM"
        MoveLabels(3) = "Sig WFM"
        MoveLabels(4) = "Max Dyn"
        MoveLabels(5) = "Maximum"
        MoveLabels(6) = "Minimum"


        DirLabels(0) = "Surge " & LUnit
        DirLabels(1) = "Sway " & LUnit
        DirLabels(2) = "Yaw (deg)"

        TenLabels(0) = "Line No."
        TenLabels(1) = "Connected"
        TenLabels(2) = "Mean Tension " & FrcUnit
        TenLabels(3) = "Sig LF Tension " & FrcUnit
        TenLabels(4) = "Sig WF Tension " & FrcUnit
        TenLabels(5) = "Max Tension " & FrcUnit
        TenLabels(6) = "Min Tension " & FrcUnit
        TenLabels(7) = "Min FOS"

    End Sub

    ' operation subroutines
    ' load to and update from form

    Private Sub LoadData()
        Dim LFactor, FrcFactor As Single

        If IsMetricUnit Then
            LFactor = 0.3048 ' ft -> m
            FrcFactor = 4.448222 ' kips -> KN
        Else
            LFactor = 1
            FrcFactor = 1
        End If

        Dim r As Short
        Dim Bearing, Distance, Depth As Single

        With CurVessel
            txtWaterDepth.Text = Format(.WaterDepth * LFactor, "0.0")
            With .ShipCurGlob
                _txtVslSt_0.Text = Format(.Xg * LFactor, "0.00")
                _txtVslSt_1.Text = Format(.Yg * LFactor, "0.00")
                _txtVslSt_2.Text = Format(.Heading * Radians2Degrees, "0.00")
                ShipLoc.Xg = .Xg
                ShipLoc.Yg = .Yg
                ShipLoc.Heading = .Heading
            End With
            _txtVslSt_3.Text = Format(.ShipDraft * LFactor, "0.00")

            If CurVessel.DampingPercent.Surge > 0 Then
                _txtPercentDamping_0.Text = Format(CurVessel.DampingPercent.Surge, "0")
                _txtPercentDamping_1.Text = Format(CurVessel.DampingPercent.Sway, "0")
                _txtPercentDamping_2.Text = Format(CurVessel.DampingPercent.Yaw, "0")
            Else
                _txtPercentDamping_0.Text = ""
                _txtPercentDamping_1.Text = ""
                _txtPercentDamping_2.Text = ""
            End If
        End With

        With CurField
            Depth = .Item(.CurWellNo).Depth
        End With
        If Depth <= 0# Then Depth = CurVessel.WaterDepth

        Coord2Bear(ShipLoc, Distance, Bearing)

        _txtMotion_2.Text = Format(Distance * LFactor, "0.0")
        _txtMotion_3.Text = Format(Distance / Depth * 100, "0.00")
        _txtMotion_4.Text = _txtVslSt_0.Text
        _txtMotion_5.Text = _txtVslSt_1.Text
        _txtMotion_6.Text = _txtVslSt_2.Text

        With grdTensions
            For r = 1 To NumLines
                PreConnect(r) = CurVessel.MoorSystem.MoorLines(r).Connected
                If PreConnect(r) Then
                    .Rows(r - 1).Cells(1).Value = "Yes"
                Else
                    .Rows(r - 1).Cells(1).Value = "No"
                End If

            Next r
        End With

        With CurField
            txtCurWell.Text = .Item(.CurWellNo).NameID
        End With

        txtEnvironment.Text = CurVessel.EnvLoad.EnvCur.Name
        '  RefreshButtonStatus
    End Sub

    'Sub RefreshButtonStatus()
    '    Dim i As Integer
    '    If InStr(txtEnvironment.Text, "No Environment") > 0 Then
    '        With txtPercentDamping
    '            For i = 0 To .Count - 1
    '                .Item(i).Enabled = True
    '                .Item(i).BackColor = vbWindowBackground
    '            Next
    '        End With
    '    Else
    '        With txtPercentDamping
    '            For i = 0 To .Count - 1
    '                .Item(i).Enabled = False
    '                .Item(i).BackColor = vbButtonFace
    '            Next
    '        End With
    '    End If
    'End Sub

    Private Sub SaveLC()
        If IsMetricUnit Then
            LFactor = 0.3048 ' ft -> m
            FrcFactor = 4.448222 ' kips -> KN
        Else
            LFactor = 1
            FrcFactor = 1
        End If

        Dim r As Short

        With grdTensions
            For r = 0 To NumLines - 1
                If .Rows(r).Cells(1).Value = "Yes" Then
                    CurVessel.MoorSystem.MoorLines(r + 1).Connected = True
                Else
                    CurVessel.MoorSystem.MoorLines(r + 1).Connected = False
                End If
            Next r
        End With


        With ShipLoc
            .Xg = CSng(CheckData(_txtVslSt_0.Text,  , True)) / LFactor
            .Yg = CSng(CheckData(_txtVslSt_1.Text,  , True)) / LFactor
            .Heading = CSng(CheckData(_txtVslSt_2.Text,  , True)) * Degrees2Radians
        End With

        With CurVessel.DampingPercent
            .Surge = Val(_txtPercentDamping_0.Text)
            .Sway = Val(_txtPercentDamping_1.Text)
            .Yaw = Val(_txtPercentDamping_2.Text)
        End With

    End Sub

    Private Sub UpdateData()
        Dim minFS, Bearing As Single

        _txtPercentDamping_0.Text = Format(CurVessel.DampingPercent.Surge, "0")
        _txtPercentDamping_1.Text = Format(CurVessel.DampingPercent.Sway, "0")
        _txtPercentDamping_2.Text = Format(CurVessel.DampingPercent.Yaw, "0")

        If IsMetricUnit Then
            LFactor = 0.3048 ' ft -> m
            FrcFactor = 4.448222 ' kips -> KN
        Else
            LFactor = 1
            FrcFactor = 1
        End If

        Dim r As Short
        Dim ToCurWell, ToCurLoc, Depth As Single

        With CurField
            Depth = .Item(.CurWellNo).Depth
        End With
        If Depth <= 0# Then Depth = CurVessel.WaterDepth

        Call MaxDynMotion(MaxDM)
        MaxDynamicMotion = MaxDM.Move
        Call MaxOffset(ToCurLoc, ToCurWell, Bearing)

        _txtMotion_0.Text = Format(ToCurLoc * LFactor, "0.0")
        _txtMotion_1.Text = Format(ToCurLoc / Depth * 100.0#, "0.00")
        _txtMotion_2.Text = Format(ToCurWell * LFactor, "0.0")
        _txtMotion_3.Text = Format(ToCurWell / Depth * 100.0#, "0.00")
        txtBearing.Text = Format(Bearing * Radians2Degrees, "0.0")

        With ShipLoc
            _txtMotion_4.Text = Format(.Xg * LFactor, "0.0")
            _txtMotion_5.Text = Format(.Yg * LFactor, "0.0")
            _txtMotion_6.Text = Format(.Heading * Radians2Degrees, "0.00")
        End With

        With grdVM

            .Rows(0).Cells(1).Value = Format(Movement.Surge * LFactor, "0.00")
            .Rows(0).Cells(2).Value = Format(LFM.Surge * LFactor, "0.00")
            .Rows(0).Cells(3).Value = Format(WFM.Surge * LFactor, "0.00")
            .Rows(0).Cells(4).Value = Format(MaxDM.Surge * LFactor, "0.00")
            .Rows(0).Cells(5).Value = Format((Movement.Surge + MaxDM.Surge) * LFactor, "0.00")
            .Rows(0).Cells(6).Value = Format((Movement.Surge - MaxDM.Surge) * LFactor, "0.00")

            .Rows(1).Cells(1).Value = Format(Movement.Sway * LFactor, "0.00")
            .Rows(1).Cells(2).Value = Format(LFM.Sway * LFactor, "0.00")
            .Rows(1).Cells(3).Value = Format(WFM.Sway * LFactor, "0.00")
            .Rows(1).Cells(4).Value = Format(MaxDM.Sway * LFactor, "0.00")
            .Rows(1).Cells(5).Value = Format((Movement.Sway + MaxDM.Sway) * LFactor, "0.00")
            .Rows(1).Cells(6).Value = Format((Movement.Sway - MaxDM.Sway) * LFactor, "0.00")

            .Rows(2).Cells(1).Value = Format(Movement.Yaw * Radians2Degrees, "0.00")
            .Rows(2).Cells(2).Value = Format(LFM.Yaw * Radians2Degrees, "0.00")
            .Rows(2).Cells(3).Value = Format(WFM.Yaw * Radians2Degrees, "0.00")
            .Rows(2).Cells(4).Value = Format(MaxDM.Yaw * Radians2Degrees, "0.00")
            .Rows(2).Cells(5).Value = Format((Movement.Yaw + MaxDM.Yaw) * Radians2Degrees, "0.00")
            .Rows(2).Cells(6).Value = Format((Movement.Yaw - MaxDM.Yaw) * Radians2Degrees, "0.00")

        End With

        Dim FrcHor As Single

        With grdTensions
            For r = 1 To NumLines
                .Rows(r - 1).Cells(2).Value = Format(Tension(r) * 0.001 * FrcFactor, "0.0")

                .Rows(r - 1).Cells(3).Value = Format(TSigLF(r) * 0.001 * FrcFactor, "0.0")

                .Rows(r - 1).Cells(4).Value = Format(TSigWF(r) * 0.001 * FrcFactor, "0.0")

                .Rows(r - 1).Cells(5).Value = Format((Tension(r) + MaxDynTen(r)) * 0.001 * FrcFactor, "0.0")

                .Rows(r - 1).Cells(6).Value = Format((Tension(r) - MaxDynTen(r)) * 0.001 * FrcFactor, "0.0")

                If Tension(r) + MaxDynTen(r) > 0 Then ' connected and converged
                    ' jguan 4/22/2004           '  update catenary using top tension
                    If CurVessel.MoorSystem.MoorLines(r).ScopeByTopTensionPOL(Tension(r) + MaxDynTen(r), FrcHor, (CurVessel.MoorSystem.MoorLines(r).Payout)) > 0 Then
                        minFS = CurVessel.MoorSystem.MoorLines(r).FOS
                        '                    ' found converged catenary solution
                    Else
                        minFS = CurVessel.MoorSystem.MoorLines(r).Segments(1).BS / (Tension(r) + MaxDynTen(r))
                    End If
                    '   Debug.Print "FS=" & CurVessel.MoorSystem.MoorLines(r).Segments(1).BS / (Tension(r) + MaxDynTen(r)) & ", minFS = " & CurVessel.MoorSystem.MoorLines(r).FOS
                    .Rows(r - 1).Cells(7).Value = Format(minFS, "0.00")
                Else
                    .Rows(r - 1).Cells(7).Value = "Damaged"
                End If
            Next r
        End With

    End Sub

    ' operation subroutines
    ' analyses

    Private Sub StaticAnalysis()

        Dim FWv, FEv As Force

        If Cancelled Then Exit Sub
        With frmProgress
            .CurrentStage.Text = "Finding equilibrium position..."
            .Progress = 5
        End With

        ShipLoc = CurVessel.FindEquilibriumPosition(Cancelled, frmProgress)
        With CurVessel.EnvLoad
            .ShipHead = ShipLoc.Heading
            FEv = .FEnvLocl
            FWv = .FWaveLocl(, False)
            .ShipHead = CurVessel.ShipCurGlob.Heading
        End With
        With FWv
            SgnX = System.Math.Sign(.Fx)
            SgnY = System.Math.Sign(.Fy)

            '       SgnZ = Sgn(.MYaw)

            '       ShipYaw = CurVessel.ShipCurGlob.Heading - ShipLoc.Heading
            '       If ShipYaw > PI Then ShipYaw = ShipYaw - 2 * PI
            '       If ShipYaw < -PI Then ShipYaw = ShipYaw + 2 * PI
            '       SgnZ = Sgn(ShipYaw)

            SgnZ = 0#

            If SgnX = 0# Then SgnX = 1.0#
            If SgnY = 0# Then SgnY = 1.0#
            '       If SgnZ = 0# Then SgnZ = 1#
        End With

        If Cancelled Then Exit Sub
        With frmProgress
            .CurrentStage.Text = "Calculating low frequency motion..."
            .Progress = 45
        End With

        Call CurVessel.MoorSystem.MoorStiff(Stiff, ShipLoc, False)

        Movement = CurVessel.ShipMotion((CurVessel.ShipCurGlob), ShipLoc)
        LFM = CurVessel.GetSigLFM(Stiff, Cancelled, frmProgress, ShipLoc)

        If Cancelled Then Exit Sub
        With frmProgress
            .CurrentStage.Text = "Calculating wave frequency motion..."
            .Progress = 75
        End With

        WFM = CurVessel.GetSigWFM(Cancelled, frmProgress, ShipLoc.Heading)
        With LFM
            LFMd.Surge = .Surge * SgnX
            LFMd.Sway = .Sway * SgnY
            LFMd.Yaw = .Yaw * SgnZ
        End With
        With WFM
            WFMd.Surge = .Surge * SgnX
            WFMd.Sway = .Sway * SgnY
            WFMd.Yaw = .Yaw * SgnZ
        End With

        If Cancelled Then Exit Sub
        With frmProgress
            .CurrentStage.Text = "Calculating low frequency tensions..."
            .Progress = 90
        End With

        Call CurVessel.FindSigDynMLTen(Tension, TSigLF, ShipLoc, LFMd)

        If Cancelled Then Exit Sub
        With frmProgress
            .CurrentStage.Text = "Calculating wave frequency tensions..."
            .Progress = 95
        End With

        Call CurVessel.FindSigDynMLTen(Tension, TSigWF, ShipLoc, WFMd)

        FEv = Nothing
        FWv = Nothing

    End Sub

    Private Sub MaxDynMotion(ByRef MaxDM As Motion)

        Dim Dyn1, Dyn2 As Single

        Dyn1 = LFM.Move * 1.5 + WFM.Move
        Dyn2 = WFM.Move * 1.86 + LFM.Move

        With MaxDM
            If Dyn1 >= Dyn2 Then
                .Surge = LFM.Surge * 1.5 + WFM.Surge
                .Sway = LFM.Sway * 1.5 + WFM.Sway
                .Yaw = LFM.Yaw * 1.5 + WFM.Yaw
            Else
                .Surge = WFM.Surge * 1.86 + LFM.Surge
                .Sway = WFM.Sway * 1.86 + LFM.Sway
                .Yaw = WFM.Yaw * 1.86 + LFM.Yaw
            End If
        End With

    End Sub

    Private Function MaxDynTen(ByVal Line As Short) As Single

        Dim Dyn1, Dyn2 As Single

        Dyn1 = TSigLF(Line) * 1.5 + TSigWF(Line)
        Dyn2 = TSigWF(Line) * 1.86 + TSigLF(Line)

        MaxDynTen = Max(Dyn1, Dyn2)

    End Function

    Private Sub MaxOffset(ByRef ToCurLoc As Single, ByRef ToCurWell As Single, ByRef Bearing As Single)

        Dim MaxDMd As New Motion
        Dim Off1L, Off2L As Single
        Dim Off1W, Off2W As Single
        Dim CurWell As Well

        With CurField
            CurWell = .Item(.CurWellNo)
        End With

        With MaxDM
            MaxDMd.Surge = .Surge * SgnX + Movement.Surge
            MaxDMd.Sway = .Sway * SgnY + Movement.Sway
            MaxDMd.Yaw = .Yaw * SgnZ + Movement.Yaw
        End With

        Off1L = MaxDMd.Move

        With CurVessel
            Coord2Bear(.NewShipLoc(.ShipCurGlob, MaxDMd), Off1W, Bearing)
        End With

        With MaxDM
            MaxDMd.Surge = - .Surge * SgnX + Movement.Surge
            MaxDMd.Sway = - .Sway * SgnY + Movement.Sway
            MaxDMd.Yaw = - .Yaw * SgnZ + Movement.Yaw
        End With

        Off2L = MaxDMd.Move

        With CurVessel
            Coord2Bear(.NewShipLoc(.ShipCurGlob, MaxDMd), Off2W, Bearing)
        End With

        ToCurLoc = Max(Off1L, Off2L)
        ToCurWell = Max(Off1W, Off2W)

        MaxDMd = Nothing

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
	
	Private Function DerpData() As Boolean

        Dim r As Short
        Dim FileOpen_Renamed As Boolean
		Dim CurLine As MoorLine
		
		DerpData = False
		
		On Error GoTo ErrorHandler
		
		FileOpen_Renamed = False
		FileOpen(FileNumRes, DerpDir & DerpFile, OpenMode.Output)
		FileOpen_Renamed = True

        PrintLine(FileNumRes, _txtMotion_3.Text & 0#)

        FileClose(FileNumRes)
		FileOpen_Renamed = False
		DerpData = True
		
		Exit Function
		
ErrorHandler: 
		If FileOpen_Renamed Then FileClose(FileNumRes)
		Exit Function
		
	End Function
	
	Public Sub ChangeUnit()
		
	End Sub
End Class