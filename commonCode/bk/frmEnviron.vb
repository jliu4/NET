Option Strict Off
Option Explicit On
Friend Class frmEnviron
	Inherits System.Windows.Forms.Form
	
	Private Const InFile As Short = 1
	Private Const OutFile As Short = 2
	
	Private EnvBack As Metocean
    Private ReName As Boolean
    Private Updated As Boolean
	Private VslHead As Single

    Private FrcLabels(9) As String
    Private lblLengthUnit(5) As Label
    Private btrDuration(2) As Button

    Private Sub setLabel()
        On Error GoTo ErrHandler
        Dim i As Short
        For i = 0 To 4 'lblLengthUnit.unbound - 1
            If i <> 5 Then
                If IsMetricUnit Then
                    lblLengthUnit(i).Text = "m"
                Else
                    lblLengthUnit(i).Text = "ft"
                End If
            End If
        Next i

        Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub

ErrHandler:
        'User pressed Cancel button
        Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub
    End Sub ' form load and unload

    Private Sub frmEnviron_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        Text = Text & " - " & CurProj.Title

        EnvBack = New Metocean
        lblLengthUnit = New Label() {_lblLengthUnit_0, _lblLengthUnit_1, _lblLengthUnit_2}

        RefreshData()
        BackCurEnv()
        Exit Sub

ErrHandler:
        MsgBox("frmEnviron_Load:")
        Exit Sub
    End Sub

    Private Sub RefreshData()
        On Error GoTo ErrHandler
        InitiateCombo()
        InitiateGrid()
        LoadVesselStation()
        LoadCurEnv()
        setLabel()
        btnForce_Click(btnForce, New System.EventArgs())
        Exit Sub
ErrHandler:

        MsgBox("RefreshData:")

    End Sub

    ' command buttons
    ' form level

    Private Sub btnAdd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnAdd.Click
		
		Dim i As Short
		Dim SelID As Short
		Dim AddNew As Boolean
		Dim msg As String
		
		SelID = cboCurEnv.SelectedIndex
        If SelID < 0 Or ReName Then
            AddNew = True
            '       check if the name have been used
            With cboCurEnv
                For i = 0 To .Items.Count - 1
                    If VB6.GetItemString(cboCurEnv, i) = .Text Then
                        AddNew = False
                        Exit For
                    End If
                Next
            End With
        Else
            AddNew = False
		End If
		
		If AddNew Then
			'   add new metocean
			msg = "Do you want to add '" & cboCurEnv.Text & "' into your environment database?"
			If MsgBox(msg, MsgBoxStyle.YesNo, Me.Text) = MsgBoxResult.No Then Exit Sub
			With CurVessel.EnvLoad
				.EnvironmentAdd()
				SelID = .EnvironmentCount - 1
				cboCurEnv.Items.Add(cboCurEnv.Text)
			End With
		Else
			'   update current one
			msg = "Do you want to update the metocean criteria of '" & cboCurEnv.Text & "'?"
			If MsgBox(msg, MsgBoxStyle.YesNo, Me.Text) = MsgBoxResult.No Then Exit Sub
		End If
		
		'   update current metocean criteria
		If Not Updated Then UpdateCurEnv()
		SaveCurEnv((SelID + 1))
		
		CurProj.Saved = False
		
	End Sub

    'Private Sub btnProfile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnProfile.Click

    'VB6.ShowForm(frmCurrent, 1, Me)

    'End Sub

    Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click

        CopyCurEnv()
        Me.Close()
		
	End Sub
	
	Private Sub btnForce_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnForce.Click
        Dim i As Short
        If IsMetricUnit Then
            LFactor = 0.3048 ' ft -> m
            FrcFactor = 4.448222 ' kips -> KN
            StressFactor = 6.894757 ' ksi -> MPa
            DiaFactor = 25.4 ' in -> mm
            LUnit = "m"
            FrcUnit = "KN"
        Else
            LFactor = 1
			FrcFactor = 1
			StressFactor = 1
			DiaFactor = 1
			LUnit = "ft"
			FrcUnit = "kips"
		End If
		
        Dim FWv, FWd, FCr, FEv As Force
        Dim WaveDir, WindDir, CurrDir As Single
        '   update current metocean criteria
        If Not Updated Then UpdateCurEnv()

        '   calculate environment force
        With CurVessel.EnvLoad
            FEv = .FEnvLocl
            FWd = .FWindLocl(, False)
            FCr = .FCurrLocl(, False)
            FWv = .FWaveLocl(, False)
        End With

        '   update force grid

        With grdForce
            .ColumnCount = 4
            For i = 0 To .ColumnCount - 1
                .Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            Next
            .Columns(0).FillWeight = 100 / .ColumnCount
            .Columns(1).FillWeight = 100 / .ColumnCount
            .Columns(2).FillWeight = 100 / .ColumnCount
            .Columns(3).FillWeight = 100 / .ColumnCount
            .RowCount = 5
            .Rows(0).Cells(0).Value = Format(FWd.Fx / 1000.0# * FrcFactor, "##,##0.0")
            .Rows(0).Cells(1).Value = Format(FWv.Fx / 1000.0# * FrcFactor, "##,##0.0")
            .Rows(0).Cells(2).Value = Format(FCr.Fx / 1000.0# * FrcFactor, "##,##0.0")
            .Rows(0).Cells(3).Value = Format(FEv.Fx / 1000.0# * FrcFactor, "##,##0.0")
            .Rows(1).Cells(0).Value = Format(FWd.Fy / 1000.0# * FrcFactor, "##,##0.0")
            .Rows(1).Cells(1).Value = Format(FWv.Fy / 1000.0# * FrcFactor, "##,##0.0")
            .Rows(1).Cells(2).Value = Format(FCr.Fy / 1000.0# * FrcFactor, "##,##0.0")
            .Rows(1).Cells(3).Value = Format(FEv.Fy / 1000.0# * FrcFactor, "##,##0.0")
            .Rows(2).Cells(0).Value = Format(FWd.MYaw / 1000.0# * FrcFactor * LFactor, "##,##0.0")
            .Rows(2).Cells(1).Value = Format(FWv.MYaw / 1000.0# * FrcFactor * LFactor, "##,##0.0")
            .Rows(2).Cells(2).Value = Format(FCr.MYaw / 1000.0# * FrcFactor * LFactor, "##,##0.0")
            .Rows(2).Cells(3).Value = Format(FEv.MYaw / 1000.0# * FrcFactor * LFactor, "##,##0.0")
            .Rows(3).Cells(0).Value = Format(FWd.Ft / 1000.0# * FrcFactor, "##,##0.0")
            .Rows(3).Cells(1).Value = Format(FWv.Ft / 1000.0# * FrcFactor, "##,##0.0")
            .Rows(3).Cells(2).Value = Format(FCr.Ft / 1000.0# * FrcFactor, "##,##0.0")
            .Rows(3).Cells(3).Value = Format(FEv.Ft / 1000.0# * FrcFactor, "##,##0.0")
            WindDir = Atan((FWd.Fx), (FWd.Fy), True)

            .Rows(4).Cells(0).Value = Format(WindDir * Radians2Degrees, "##0.0")

            WaveDir = Atan((FWv.Fx), (FWv.Fy), True)
            .Rows(4).Cells(1).Value = Format(WaveDir * Radians2Degrees, "##0.0")

            CurrDir = Atan((FCr.Fx), (FCr.Fy), True)
            .Rows(4).Cells(2).Value = Format(CurrDir * Radians2Degrees, "##0.0")

            .Rows(4).Cells(3).Value = Format(Atan((FEv.Fx), (FEv.Fy), True) * Radians2Degrees, "##0.0")

        End With

        '   update picture
        'DrawEnvPlot(WindDir, WaveDir, CurrDir, VslHead, picEnviron)

    End Sub
	
	Private Sub btnOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnOK.Click

        'If Not Updated Then UpdateCurEnv()
        'CurProj.Saved = False
        Me.Close()
		
	End Sub
	
	Private Sub frmEnviron_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        EnvBack = Nothing
    End Sub
	
	' menus
	
	Public Sub mnuClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuClose.Click
		
		btnCancel_Click(btnCancel, New System.EventArgs())
		
	End Sub
	
	Public Sub mnuOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuOpen.Click

        Dim ReadFile, isFileOpen As Boolean
        Dim msg As String

        isFileOpen = False

        '   should the user cancel the dialog box, exit
        On Error GoTo ErrHandler

        dlgFileOpen.Filter = "All Files (*.*)|*.*|Env Files (*.env)|*.env"
        dlgFileSave.Filter = "All Files (*.*)|*.*|Env Files (*.env)|*.env"
		dlgFileOpen.FilterIndex = 2
		dlgFileSave.FilterIndex = 2
		dlgFileOpen.InitialDirectory = CurProj.Directory
		dlgFileSave.InitialDirectory = CurProj.Directory
        dlgFileOpen.CheckFileExists = True
		dlgFileOpen.CheckPathExists = True
		dlgFileSave.CheckPathExists = True
        dlgFileOpen.ShowReadOnly = False
		dlgFileOpen.ShowDialog()
		dlgFileSave.FileName = dlgFileOpen.FileName
		
		'   open the file
		FileOpen(InFile, dlgFileOpen.FileName, OpenMode.Input, OpenAccess.Read)
        isFileOpen = True

        On Error GoTo 0
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		ReadFile = CurVessel.EnvLoad.InputEnv(InFile)
		'   close the input file and return
		FileClose(InFile)
        isFileOpen = False
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		
		If ReadFile Then
			InitiateCombo()
		Else
			cboCurEnv.Items.Clear()
			msg = dlgFileOpen.FileName & " appears to be corrupted. Please verify " & "the available environmental data."
			MsgBox(msg, MsgBoxStyle.OKOnly, Me.Text)
		End If
		cboCurEnv.Text = "Current Environment"
		
		Exit Sub
		
ErrHandler:
        '   User pressed Cancel button, or some tragedy occurred
        If isFileOpen Then FileClose(InFile)
        Exit Sub
		
	End Sub
	
	Public Sub mnuSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSave.Click

        Dim WriteFile, isFileOpen As Boolean

        isFileOpen = False

        On Error GoTo ErrHandler
		
        dlgFileOpen.Filter = "All Files (*.*)|*.*|Env Files (*.env)|*.env"
		dlgFileSave.Filter = "All Files (*.*)|*.*|Env Files (*.env)|*.env"
		dlgFileOpen.FilterIndex = 2
		dlgFileSave.FilterIndex = 2
		dlgFileOpen.FileName = "*.env"
		dlgFileSave.FileName = "*.env"
		dlgFileOpen.InitialDirectory = CurProj.Directory
		dlgFileSave.InitialDirectory = CurProj.Directory
        dlgFileSave.OverwritePrompt = True
        dlgFileOpen.ShowReadOnly = False
		dlgFileSave.ShowDialog()
		dlgFileOpen.FileName = dlgFileSave.FileName
		
		'   open the file for output
		FileOpen(OutFile, dlgFileOpen.FileName, OpenMode.Output)
        isFileOpen = True

        On Error GoTo 0
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		WriteFile = CurVessel.EnvLoad.OutputEnv(OutFile)
		'   Close the output file
		FileClose(OutFile)
        isFileOpen = False
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		Exit Sub
		
ErrHandler:
        '   User pressed Cancel button, or some tragedy occurred
        If isFileOpen Then FileClose(OutFile)
        Exit Sub
		
	End Sub
	
	' combolist
	
    Private Sub cboCurEnv_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboCurEnv.SelectedIndexChanged

        Dim SelID As Short

        SelID = cboCurEnv.SelectedIndex
        If SelID >= 0 Then
            ChangeCurEnv((SelID + 1))
            LoadCurEnv()
            ReName = False
        End If
    End Sub

    Private Sub cboCurEnv_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboCurEnv.TextChanged

        ReName = True
        Updated = False

    End Sub

    ' option button

    Private Sub btrDuration_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        If eventSender.Checked Then
            Dim Index As Short = btrDuration.GetIndex(eventSender)

            Updated = False

        End If
    End Sub

    ' text boxes

    Private Sub txtWind_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWind.TextChanged
        Dim Index As Short = txtWind.GetIndex(eventSender)

        Updated = False
    End Sub

    Private Sub txtWind_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWind.Leave
		Dim Index As Short = txtWind.GetIndex(eventSender)

        txtWind(Index).Text = CheckData(txtWind(Index).Text, "0.00")

    End Sub
	
	'UPGRADE_WARNING: Event txtWave.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtWave_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWave.TextChanged
		Dim Index As Short = txtWave.GetIndex(eventSender)
		
		Updated = False
	End Sub
	
	Private Sub txtWave_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWave.Leave
		Dim Index As Short = txtWave.GetIndex(eventSender)

        txtWave(Index).Text = CheckData(txtWave(Index).Text, "0.00")

    End Sub
	
	'UPGRADE_WARNING: Event txtCurr.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtCurr_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCurr.TextChanged
		Dim Index As Short = txtCurr.GetIndex(eventSender)
		
		Updated = False
	End Sub
	
	Private Sub txtCurr_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCurr.Leave
		Dim Index As Short = txtCurr.GetIndex(eventSender)

        txtCurr(Index).Text = CheckData(txtCurr(Index).Text, "0.00")

    End Sub
	
	' operation subroutines
	' initiating
	
	Private Sub InitiateCombo()
		
		Dim i As Short
		
		cboCurEnv.Items.Clear()
		
		For i = 1 To CurVessel.EnvLoad.EnvironmentCount
			With CurVessel.EnvLoad.Environments(i)
				cboCurEnv.Items.Add(.Name)
			End With
		Next 
		cboCurEnv.SelectedIndex = 0
	End Sub
	
	Private Sub InitiateGrid()
		
		Dim r, c As Short

        Call SetLabels()

        With grdForce
            .ColumnCount = 4
            .RowCount = 5

            For c = 0 To .ColumnCount - 1
                .Columns(c).HeaderText = FrcLabels(c + 1)

            Next
            For r = 0 To .RowCount - 1
                .Rows(r).HeaderCell.Value = FrcLabels(r + .ColumnCount + 1)

            Next
            .Columns(0).Width = 40
            .Columns(1).Width = 40
            .Columns(2).Width = 40
            .Columns(3).Width = 40
        End With
		
	End Sub
	
	Private Sub SetLabels()
		Dim LUnit, FrcUnit As String
		If IsMetricUnit Then
			LUnit = "m"
			FrcUnit = "KN"
		Else
			LUnit = "ft"
			FrcUnit = "kips"
		End If
		
		FrcLabels(0) = "(" & FrcUnit & "-" & LUnit & ")"
		FrcLabels(1) = "Wind"
		FrcLabels(2) = "Wave"
		FrcLabels(3) = "Current"
		FrcLabels(4) = "Total"
		FrcLabels(5) = "Surge"
		FrcLabels(6) = "Sway"
		FrcLabels(7) = "Yaw"
		FrcLabels(8) = "Overall"
		FrcLabels(9) = "Direction"
		
	End Sub
	
	' operation subroutines
	' load form and update curenv
	
	Private Sub LoadCurEnv()
		
		If IsMetricUnit Then
			LFactor = 0.3048 ' ft -> m
			FrcFactor = 4.448222 ' kips -> KN
			StressFactor = 6.894757 ' ksi -> MPa
			VelFactor = 0.5144444 ' knots - > m/s
			DiaFactor = 25.4 ' in -> mm
			LUnit = "m"
			FrcUnit = "KN"
		Else
			VelFactor = 1
			LFactor = 1
			FrcFactor = 1
			StressFactor = 1
			DiaFactor = 1
			LUnit = "ft"
			FrcUnit = "kips"
		End If
		
		With CurVessel.EnvLoad.EnvCur.Wind
			txtWind(0).Text = Format(.Velocity * Ftps2Knots * VelFactor, "0.00")
			txtWind(1).Text = Format(.Elevation * LFactor, "0.00")
			txtWind(2).Text = Format(.Heading * Radians2Degrees, "0.00")
			
			Select Case .Duration
				Case 3600
					btrDuration(0).Checked = True
				Case 60
					btrDuration(1).Checked = True
				Case 3
					btrDuration(2).Checked = True
			End Select
		End With
		
		With CurVessel.EnvLoad.EnvCur.Wave
			txtWave(0).Text = Format(.Height * LFactor, "0.00")
			txtWave(1).Text = Format(.Period, "0.00")
			txtWave(2).Text = Format(.Heading * Radians2Degrees, "0.00")
		End With
		
		With CurVessel.EnvLoad.EnvCur.Current
			txtCurr(0).Text = Format(.Profile(1).Velocity * Ftps2Knots * VelFactor, "0.00")
			txtCurr(1).Text = Format(.Heading * Radians2Degrees, "0.00")
		End With
		
		If cboCurEnv.SelectedIndex < 0 Then
			If CurVessel.EnvLoad.EnvCur.Name = "" Then
				cboCurEnv.Text = "Current Environment"
			Else
				cboCurEnv.Text = CurVessel.EnvLoad.EnvCur.Name
			End If
		End If

        Updated = True
        btnForce_Click(btnForce, New System.EventArgs())

    End Sub
	
	Private Sub LoadVesselStation()
		
		VslHead = CurVessel.ShipCurGlob.Heading
		txtVslSt(0).Text = Format(VslHead * Radians2Degrees, "0.00")
		txtVslSt(1).Text = Format(CurVessel.ShipDraft * LFactor, "0.00")
		
	End Sub
	
	Private Sub UpdateCurEnv()
		If IsMetricUnit Then
			LFactor = 0.3048 ' ft -> m
			FrcFactor = 4.448222 ' kips -> KN
			StressFactor = 6.894757 ' ksi -> MPa
			DiaFactor = 25.4 ' in -> mm
			VelFactor = 0.5144444 ' knots - > m/s
			LUnit = "m"
			FrcUnit = "KN"
		Else
			VelFactor = 1
			LFactor = 1
			FrcFactor = 1
			StressFactor = 1
			DiaFactor = 1
			LUnit = "ft"
			FrcUnit = "kips"
		End If
		
		Dim i As Short
		
		If Updated Then Exit Sub
		
		CurVessel.EnvLoad.EnvCur.Name = cboCurEnv.Text
		
		With CurVessel.EnvLoad.EnvCur.Wind
			.Velocity = CDbl(txtWind(0).Text) * Knots2Ftps / VelFactor
			.Elevation = CDbl(txtWind(1).Text) / LFactor
			.Heading = CDbl(txtWind(2).Text) * Degrees2Radians
			
			For i = 0 To 2
				If btrDuration(i).Checked Then
					Select Case i
						Case 0
							.Duration = 3600
						Case 1
							.Duration = 60
						Case 2
							.Duration = 3
					End Select
					Exit For
				End If
			Next 
		End With
		
		With CurVessel.EnvLoad.EnvCur.Wave
			.Height = CDbl(txtWave(0).Text) / LFactor
			.Period = CDbl(txtWave(1).Text)
			.Heading = CDbl(txtWave(2).Text) * Degrees2Radians
		End With
		
		With CurVessel.EnvLoad.EnvCur.Current
			.Profile(1).Velocity = CDbl(txtCurr(0).Text) * Knots2Ftps / VelFactor
			.Heading = CDbl(txtCurr(1).Text) * Degrees2Radians
		End With
		
		Updated = True
		
	End Sub
	
	' operation subroutines
	' backup, change, reload and save envcur
	
	Private Sub BackCurEnv()
		
		Dim i As Short
        Dim Wind_Renamed As Wind
        Dim Wave_Renamed As Wave
		Dim Curr As Current
		
		EnvBack.Name = CurVessel.EnvLoad.EnvCur.Name
		
		Wind_Renamed = EnvBack.Wind
		With CurVessel.EnvLoad.EnvCur.Wind
			Wind_Renamed.Duration = .Duration
			Wind_Renamed.Elevation = .Elevation
			Wind_Renamed.Heading = .Heading
			Wind_Renamed.Velocity = .Velocity
		End With
		
		Wave_Renamed = EnvBack.Wave
		With CurVessel.EnvLoad.EnvCur.Wave
			Wave_Renamed.Heading = .Heading
			Wave_Renamed.Height = .Height
			Wave_Renamed.Period = .Period
		End With
		
		Curr = EnvBack.Current
		With CurVessel.EnvLoad.EnvCur.Current
			Curr.Heading = .Heading
			Curr.WaterDepth = .WaterDepth
			
			Do While Curr.ProfileCount > 0
				Curr.ProfileDelete(1)
			Loop 
			
			If .ProfileCount = 0 Then
				Curr.ProfileAdd(0#, 0#)
			Else
				For i = 1 To .ProfileCount
					Curr.ProfileAdd((.Profile(i).Depth), (.Profile(i).Velocity))
				Next 
			End If
		End With
		
	End Sub
	
	Private Sub CopyCurEnv()
		
		Dim i As Short
        Dim Wind_Renamed As Wind
        Dim Wave_Renamed As Wave
		Dim Curr As Current
		
		CurVessel.EnvLoad.EnvCur.Name = EnvBack.Name
		
		Wind_Renamed = CurVessel.EnvLoad.EnvCur.Wind
		With EnvBack.Wind
			Wind_Renamed.Duration = .Duration
			Wind_Renamed.Elevation = .Elevation
			Wind_Renamed.Heading = .Heading
			Wind_Renamed.Velocity = .Velocity
		End With
		
		Wave_Renamed = CurVessel.EnvLoad.EnvCur.Wave
		With EnvBack.Wave
			Wave_Renamed.Heading = .Heading
			Wave_Renamed.Height = .Height
			Wave_Renamed.Period = .Period
		End With
		
		Curr = CurVessel.EnvLoad.EnvCur.Current
		With EnvBack.Current
			Curr.Heading = .Heading
			Curr.WaterDepth = .WaterDepth
			
			Do While Curr.ProfileCount > 0
				Curr.ProfileDelete(1)
			Loop 
			
			If .ProfileCount = 0 Then
				Curr.ProfileAdd(0#, 0#)
			Else
				For i = 1 To .ProfileCount
					Curr.ProfileAdd((.Profile(i).Depth), (.Profile(i).Velocity))
				Next 
			End If
		End With
		
		Updated = True
		
	End Sub
	
	Private Sub ChangeCurEnv(ByRef EnvID As Short)
		
		Dim i As Short
        Dim Wind_Renamed As Wind
        Dim Wave_Renamed As Wave
		Dim Curr As Current
		Dim SelEnv As Metocean
		
		SelEnv = CurVessel.EnvLoad.Environments(EnvID)
		
		CurVessel.EnvLoad.EnvCur.Name = SelEnv.Name
		
		Wind_Renamed = CurVessel.EnvLoad.EnvCur.Wind
		With SelEnv.Wind
			Wind_Renamed.Duration = .Duration
			Wind_Renamed.Elevation = .Elevation
			Wind_Renamed.Heading = .Heading
			Wind_Renamed.Velocity = .Velocity
		End With
		
		Wave_Renamed = CurVessel.EnvLoad.EnvCur.Wave
		With SelEnv.Wave
			Wave_Renamed.Heading = .Heading
			Wave_Renamed.Height = .Height
			Wave_Renamed.Period = .Period
		End With
		
		Curr = CurVessel.EnvLoad.EnvCur.Current
		With SelEnv.Current
			Curr.Heading = .Heading
			Curr.WaterDepth = .WaterDepth
			
			Do While Curr.ProfileCount > 0
				Curr.ProfileDelete(1)
			Loop 
			
			If .ProfileCount = 0 Then
				Curr.ProfileAdd(0#, 0#)
			Else
				For i = 1 To .ProfileCount
					Curr.ProfileAdd((.Profile(i).Depth), (.Profile(i).Velocity))
				Next 
			End If
		End With
		
		Updated = True
		
	End Sub
	
	Private Sub SaveCurEnv(ByRef EnvID As Short)
		
		Dim i As Short
        Dim Wind1 As Wind
        Dim Wave1 As Wave
        Dim Curr As Current
		Dim SelEnv As Metocean
		
		SelEnv = CurVessel.EnvLoad.Environments(EnvID)
		
		SelEnv.Name = CurVessel.EnvLoad.EnvCur.Name

        Wind1 = SelEnv.Wind
        With CurVessel.EnvLoad.EnvCur.Wind
            Wind1.Duration = .Duration
            Wind1.Elevation = .Elevation
            Wind1.Heading = .Heading
            Wind1.Velocity = .Velocity
        End With

        Wave1 = SelEnv.Wave
        With CurVessel.EnvLoad.EnvCur.Wave
            Wave1.Heading = .Heading
            Wave1.Height = .Height
            Wave1.Period = .Period
        End With
		
		Curr = SelEnv.Current
		With CurVessel.EnvLoad.EnvCur.Current
			Curr.Heading = .Heading
			Curr.WaterDepth = .WaterDepth
			
			Do While Curr.ProfileCount > 0
				Curr.ProfileDelete(1)
			Loop 
			
			If .ProfileCount = 0 Then
				Curr.ProfileAdd(0#, 0#)
			Else
				For i = 1 To .ProfileCount
					Curr.ProfileAdd((.Profile(i).Depth), (.Profile(i).Velocity))
				Next 
			End If
		End With
		
	End Sub


End Class