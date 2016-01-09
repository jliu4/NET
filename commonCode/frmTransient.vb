Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System
Imports System.Diagnostics
Imports System.ComponentModel

Friend Class frmTransient
	Inherits System.Windows.Forms.Form
	
	' for messages
	Private msgNotSavedWarning As String
	Private Const msgInputWarning As String = "The input file has not been read properly. " & " It may be corrupted. Please check data closely."
	Private Const msgNoFileWarning As String = "Cannot find the specifiied input file."
	Private msgOutputWarning As String
	
	' for form operations
	Private NumProj As Short
	
	Private CheckingGrid As Boolean
	Private FirstLoad As Boolean
	
	Private Const NumLCCaptions As Short = 2
	Private SegLabels(11, 1) As String
	Private CurLine As Short
	Private ShipDesLoc As ShipGlobal
	Private CurDraft As Single
	Private SurDraft As Single
	Private OprDraft As Single
	
	Dim Yaw, X, Y, Head As Single
	
	Dim PlotX() As Single
	Dim PlotY() As Single
	Dim PlotHead() As Single
	Dim PlotIndex As Short
	Dim MaxOffset As Single ' use to save max offset distance
	Dim MaxTransY, MaxTransX, MaxTransTime As Single
	
	Dim LCColHead(NumLCCaptions - 1) As String
	Dim TMColHead(3) As String
	Dim TMRowHead() As String
	Dim LCRowHead() As String
	
	Dim TransientComputed As Boolean

    Public Sub setLabel()
        On Error GoTo ErrHandler
        Dim i As Short
        For i = 0 To lblLengthUnit.Count - 1
            If i <> 5 Then
                If IsMetricUnit Then
                    lblLengthUnit(i).Text = "m"
                Else
                    lblLengthUnit(i).Text = "ft"
                End If
            End If
        Next i
        For i = 0 To lblForceUnit.Count - 1
            If IsMetricUnit Then
                lblForceUnit(i).Text = "KN"
            Else
                lblForceUnit(i).Text = "kips"
            End If
        Next i

        For i = 0 To lblMassUnit.Count - 1
            If IsMetricUnit Then
                lblMassUnit(i).Text = "MT"
            Else
                lblMassUnit(i).Text = "kips"
            End If
        Next i
        For i = 0 To lblDiaUnit.Count - 1
            If IsMetricUnit Then
                lblDiaUnit(i).Text = "mm"
            Else
                lblDiaUnit(i).Text = "in"
            End If
        Next i
        Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub

ErrHandler:
        'User pressed Cancel button
        Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub
    End Sub

    Private Sub frmTransient_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
		IsMetricUnit = False
		InitProject()
      
		NumProj = 1
		With CurProj
			CurVessel = .Vessel
			CurVessel.Name = .Vessel.Name
			
			.Title = "Project" & NumProj
			Text = "DODO Console - " & .Title
			.Saved = True
		End With
		
			If Len(Dir(DODODir & IniFile)) > 0 Then
            FileOpen(10, DODODir & IniFile, OpenMode.Input, OpenAccess.Read)

            If Defaults.InputData(10) Then
                UpdateFileMenu()
                CurProj.Directory = Defaults.WorkDir
            End If
            FileClose(10)

        End If

        NumTimeSteps = 2400
        TimeStep = 0.25 ' sec
        MaxTimeSteps = 10000
		
		' Intialize the text boxes
		
		txtDuration.Text = VB6.Format(TimeStep * NumTimeSteps, "##,##0.0")
		txtInterval.Text = VB6.Format(TimeStep, "##0.00")
		
		TransientComputed = False
		
		RefreshData()
		txtEnvironment.Text = CurVessel.EnvLoad.EnvCur.Name
		
		UpdateFileMenu()
        setLabel()
        Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub

ErrHandler:
        'User pressed Cancel button
        Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub
		
	End Sub
	
	Private Sub frmTransient_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        Dim cancel As Boolean = eventArgs.Cancel
		Dim unloadmode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        Try
            If CurProj.Saved = False Then
                If CurProj.Title = "" Then
                    msgNotSavedWarning = "Do you want to save changes made to " & " the current project ?"
                Else
                    msgNotSavedWarning = "Do you want to save changes made to " & CurProj.Title & " ?"
                End If

                Select Case MsgBox(msgNotSavedWarning, MsgBoxStyle.YesNoCancel, msgTitle)
                    Case MsgBoxResult.Yes
                        '           save
                        mnuFileSave_Click(mnuFileSave, New System.EventArgs())
                        Defaults.AddPreFile(CurProj.Directory & CurProj.FileName)
                    Case MsgBoxResult.No
                        If CurProj.FileName <> "" Then Defaults.AddPreFile(CurProj.Directory & CurProj.FileName)
                    Case MsgBoxResult.Cancel
                        '           do nothing
                        cancel = True
                End Select
            Else
                With CurProj
                    If .FileName <> "" Then Defaults.AddPreFile(.Directory & .FileName)
                End With
            End If

            CurProj = Nothing

            eventArgs.Cancel = cancel
        Catch
            MsgBox("frmTransient.closing Error")
        End Try
    End Sub

	' buttons
	
	Private Sub btnEnvironment_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnEnvironment.Click

        VB6.ShowForm(frmEnviron, 1, Me)
        txtEnvironment.Text = CurVessel.EnvLoad.EnvCur.Name
		
	End Sub
	
	Private Sub btnReport_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnReport.Click
        On Error GoTo ErrHandler
		If Not TransientComputed Then
			MsgBox("Must perform analysis first.", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Error")
			Exit Sub
		End If

		Exit Sub
ErrHandler: 
		MsgBox("Err creating report...")
		
	End Sub
	
	Private Sub btnTransient_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnTransient.Click
        On Error GoTo ErrHandler

		Dim LFactor, FrcFactor As Single
        Dim FileOpen1, FileOpen2, FileOpen3 As Boolean
        Dim FileNum1, FileNum2, FileNum3 As Integer
		Dim AppCur(MaxCurrentPair, 8) As Single
		Dim Depth(MaxCurrentPair) As Single
		Dim NumCurrent As Short
		Dim WaveDir(8) As Single
		
		Dim k, i, r, c, j, L As Short
		Dim PWD, Dist, Dist0 As Single
        Dim CurX, CurDir_Renamed, CurY As Single
        Dim Pos() As ShipGlobal
		Dim Exx As Single
        Dim Vel() As Motion
        Dim Acc() As Motion
		Dim sMass As New Motion
		Dim sDamp As New Motion
		Dim sYRD As Single
		Dim FMoor As New Force
		Dim FEnv As New Force
		Dim FRiser As New Force
		
		Dim ZWD As Single
		
		If IsMetricUnit Then
			LFactor = 0.3048 ' ft -> m
			FrcFactor = 4.448222 ' kips -> KN
		Else
			LFactor = 1
			FrcFactor = 1
		End If
		
        ReDim PlotX(NumTimeSteps + 1)
        ReDim PlotY(NumTimeSteps + 1)
		ReDim PlotHead(NumTimeSteps + 1)
		ReDim TransX(NumTimeSteps + 1)
		ReDim TransY(NumTimeSteps + 1)
		ReDim TransYaw(NumTimeSteps + 1)
        ReDim Acc(2)
        ReDim Pos(2)
        ReDim Vel(2)
        For i = 0 To 2
            Acc(i) = New Motion
            Pos(i) = New ShipGlobal

            Vel(i) = New Motion
        Next

		FileOpen1 = False
		FileOpen2 = False
        FileOpen3 = False
        'appvel
        FileNum1 = FreeFile()
        FileOpen(FileNum1, CurProj.Directory & AppVelOutput, OpenMode.Output)
        FileOpen1 = True
        'offset
        FileNum2 = FreeFile()
        FileOpen(FileNum2, CurProj.Directory & OffsetOutput, OpenMode.Output)
        FileOpen2 = True
        'debug
        FileNum3 = FreeFile()
        FileOpen(FileNum3, CurProj.Directory & "debug.out", OpenMode.Output)
        FileOpen3 = True

        Me.Enabled = False
		Cancelled = False
		With frmProgress
			.Text = "Transient Analysis"
			.CurrentStage.Text = "Initializing..."
			.Progress = 0
			VB6.ShowForm(frmProgress, 0, Me)
		End With
		
		SaveLC()
		
		ZWD = CurVessel.WaterDepth
		
        TransX(0) = CurVessel.ShipCurGlob.Xg
        TransY(0) = CurVessel.ShipCurGlob.Yg
        TransYaw(0) = CurVessel.ShipCurGlob.Heading
        Dist0 = System.Math.Sqrt(TransX(0) ^ 2 + TransY(0) ^ 2)
		Dist = Dist0
		
		With CurVessel.EnvLoad.EnvCur.Current
			NumCurrent = .ProfileCount
			For L = 1 To NumCurrent
				Depth(L) = .Profile(L).Depth
				AppCur(L, 0) = .Profile(L).Velocity * Ftps2Knots
			Next L
		End With
		With CurVessel
			WaveDir(0) = RadTo360(.ShipCurGlob.Heading - .EnvLoad.EnvCur.Wave.Heading)
		End With
		
		' Compute the motion and make the plot
		
		'   initializing
		sMass = CurVessel.ShipMass(CurVessel.ShipDraft)
		sDamp = CurVessel.ShipDamp(CurVessel.ShipDraft)
        'sYRD = CurVessel.ShipYawRateDrag(CurVessel.ShipDraft)
        For i = 0 To 2
			Acc(i).Surge = 0
			Vel(i).Surge = 0
			Acc(i).Sway = 0
			Vel(i).Sway = 0
			Acc(i).Yaw = 0
			Vel(i).Yaw = 0
			With CurVessel.ShipCurGlob
				Pos(i).Xg = .Xg
				Pos(i).Yg = .Yg
				Pos(i).Heading = .Heading
			End With
		Next i
			
		Dim PlotIndex As Short
		
		For i = 1 To NumTimeSteps + 1
			PlotX(i) = 0
			PlotY(i) = 0
			PlotHead(i) = 0
            TransX(i) = 0
            TransY(i) = 0
            TransYaw(i) = CurVessel.ShipCurGlob.Heading
		Next 
		
		MaxOffset = 0
		PlotIndex = 1
        PlotX(PlotIndex) = TransX(1)
        PlotY(PlotIndex) = TransY(1)
        PlotHead(PlotIndex) = TransYaw(1)
		
		For i = 1 To NumTimeSteps
			If Cancelled Then GoTo ExitSub
			With frmProgress
				.CurrentStage.Text = "Transient Analysis for Time = " & i * TimeStep & " sec"
				.Progress = ((i) / (NumTimeSteps)) * 75
			End With
			
			With CurVessel
				FEnv = .EnvLoad.FEnvLocl(Pos(1).Heading, (Vel(1).Surge), (Vel(1).Sway))
                'FRiser = .Riser.FhLocl(Pos(1), Vel(1), Acc(1), .EnvLoad.EnvCur.Current)

                Call .NextPosnVel(TimeStep, Acc, Vel, Pos)
				
				Acc(2).Surge = (-sDamp.Surge * Vel(2).Surge + FEnv.Fx + FRiser.Fx) / sMass.Surge
				Acc(2).Sway = (-sDamp.Sway * Vel(2).Sway + FEnv.Fy + FRiser.Fy) / sMass.Sway
				Acc(2).Yaw = (-sDamp.Yaw * Vel(2).Yaw - sYRD * Vel(2).Yaw * System.Math.Abs(Vel(2).Yaw) + FEnv.MYaw) / sMass.Yaw
				
				'--------------------------------------------------------------
				' Update the current location
				'--------------------------------------------------------------
				
				Call .UpdatePosVelnAcc(Acc, Vel, Pos)
			End With
			
			' save current transient data into arrays for display in tables / plot
			
            TransX(i) = Pos(2).Xg
            TransY(i) = Pos(2).Yg
            TransYaw(i) = Pos(2).Heading
			
            Dist = System.Math.Sqrt((TransX(i) - TransX(0)) ^ 2 + (TransY(i) - TransY(0)) ^ 2)
			If Dist > Dist0 Then
				k = 1
				Do While Dist0 / ZWD * 100# - k * 2 > 0#
					k = k + 1
				Loop 
				If Dist / ZWD * 100 - k * 2 >= 0# And k < 9 Then
					With CurVessel.EnvLoad.EnvCur.Current
                        CurDir_Renamed = TransYaw(i) - .Heading
						For L = 1 To NumCurrent
							CurX = .Profile(L).Velocity * System.Math.Cos(CurDir_Renamed) - Vel(2).Surge * (1# - Depth(L) / ZWD)
							CurY = .Profile(L).Velocity * System.Math.Sin(CurDir_Renamed) - Vel(2).Sway * (1# - Depth(L) / ZWD)
							AppCur(L, k) = System.Math.Sqrt(CurX ^ 2 + CurY ^ 2) * Ftps2Knots
						Next L
					End With
					
                    WaveDir(k) = RadTo360(TransYaw(i) - CurVessel.EnvLoad.EnvCur.Wave.Heading)
				End If
			End If
            'offset
            WriteLine(FileNum2, TimeStep * (i - 1), TransX(i), TransY(i), RadTo360(TransYaw(i)), Vel(2).Surge * Ftps2Knots, Vel(2).Sway * Ftps2Knots, Vel(2).Move * Ftps2Knots, Acc(2).Move, FEnv.Ft * 0.001, FRiser.Ft * 0.001, Dist / ZWD * 100.0#)
			
			'debug
            WriteLine(FileNum3, TimeStep * (i - 1), TransYaw(i), Vel(2).Yaw, Acc(2).Yaw, FEnv.MYaw, FRiser.MYaw, sMass.Yaw, Pos(1).Heading)
			
			Dist0 = Dist
			
			If Dist > MaxOffset Then
				MaxOffset = Dist ' Max Transient Offset
                MaxTransX = TransX(i)
                MaxTransY = TransY(i)
				MaxTransTime = TimeStep * (i - 1)
			End If
			
			If (i Mod 20 = 0) Then
				PlotIndex = PlotIndex + 1
                PlotX(PlotIndex) = TransX(i)
                PlotY(PlotIndex) = TransY(i)
                PlotHead(PlotIndex) = TransYaw(i)
			End If
			
		Next i
        'appvel
        WriteLine(FileNum1, "Dir", WaveDir(0), WaveDir(1), WaveDir(2), WaveDir(3), WaveDir(4), WaveDir(5), WaveDir(6), WaveDir(7), WaveDir(8))
        For L = 1 To NumCurrent
            WriteLine(FileNum1, Depth(L), AppCur(L, 0), AppCur(L, 1), AppCur(L, 2), AppCur(L, 3), AppCur(L, 4), AppCur(L, 5), AppCur(L, 6), AppCur(L, 7), AppCur(L, 8))
        Next L

        Call DrawTransientPlot(TransX(0), TransY(0), PlotIndex, PlotX, PlotY, PlotHead, MaxTransX, MaxTransY, picPolar) 'JLIU 09182015

        With frmProgress
			.CurrentStage.Text = "Transient Analysis completed."
			.Progress = 100
		End With
		frmProgress.Close()
		
		'----------------------------------------------------------------
		' Update the text boxes
		Dim dy, dx, Bearing As Single
        dx = TransX(NumTimeSteps) - TransX(0)
        dy = TransY(NumTimeSteps) - TransY(0)
		
		Dist = System.Math.Sqrt(dx ^ 2 + dy ^ 2)
		PWD = Dist / CurVessel.WaterDepth
		
		txtOffset.Text = VB6.Format(Dist * LFactor, "##,##0") '  final offset
		txtOffsetPWD.Text = VB6.Format(PWD * 100#, "#0.0") ' final offset percent WD
		
		If dx = 0 And dy = 0 Then
			Bearing = 0#
		Else
            Bearing = Atn2(dx, dy)
			Bearing = PI / 2# - Bearing
		End If
		
		txtOffsetBearing.Text = VB6.Format(RadTo360(Bearing), "#0") '  Bearing is the angle of final position measured clockwise from TN ???
		
		PWD = MaxOffset / CurVessel.WaterDepth
		
		txtMaxOffset.Text = VB6.Format(MaxOffset * LFactor, "##,##0") ' Max Transient Offset
		txtMaxOffsetPWD.Text = VB6.Format(PWD * 100#, "#0.0") ' Max Transient Offset Percent WD
		txtMaxOffsetTime.Text = CStr(MaxTransTime)

        ' Update the data table

        With grdTM
            .RowCount = NumTimeSteps + 1
            For r = 0 To .RowCount - 1
                .Rows(r).Cells(0).Value = r * TimeStep
                .Rows(r).Cells(1).Value = VB6.Format(TransX(r) * LFactor, "#,##0.00")

                .Rows(r).Cells(2).Value = VB6.Format(TransY(r) * LFactor, "#,##0.00")

                .Rows(r).Cells(3).Value = VB6.Format(RadTo360(Val(TransYaw(r))), "##0.0")

            Next r
        End With

        TransientComputed = True

        FileClose(FileNum3)
        FileOpen3 = False
        FileClose(FileNum1)
        FileOpen1 = False
        FileClose(FileNum2)
        FileOpen2 = False

ExitSub:
        Cursor = System.Windows.Forms.Cursors.Default
        Me.Enabled = True

        Exit Sub

ErrHandler:
        If FileOpen3 Then FileClose(FileNum3)
        If FileOpen1 Then FileClose(FileNum1)
        If FileOpen2 Then FileClose(FileNum2)

        Cursor = System.Windows.Forms.Cursors.Default
        MsgBox("Error: " & Err.Description & ", Source: " & Err.Source, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Error")
        Me.Enabled = True
		
	End Sub

    Private Sub btnDisplayCurves_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

        'If TransientComputed Then
        'frmTransientCurves.Show()
        'Else
        '      MsgBox("You have not yet computed any transient time histories to display!", MsgBoxStyle.OKOnly, "DODO - No Data To Display")
        '      Exit Sub
        'End If

    End Sub

    ' menus

    Public Sub mnuFileNew_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileNew.Click
		
		If CurProj.Saved = False Then
			If CurProj.Title = "" Then
				msgNotSavedWarning = "Do you want to save changes made to " & " the current project ?"
			Else
				msgNotSavedWarning = "Do you want to save changes made to " & CurProj.Title & " ?"
			End If
			
			Select Case MsgBox(msgNotSavedWarning, MsgBoxStyle.YesNoCancel, msgTitle)
				Case MsgBoxResult.Yes
					'           save
					mnuFileSave_Click(mnuFileSave, New System.EventArgs())
					If CurProj.FileName <> "" Then Defaults.AddPreFile(CurProj.Directory & CurProj.FileName)
				Case MsgBoxResult.No
					If CurProj.FileName <> "" Then Defaults.AddPreFile(CurProj.Directory & CurProj.FileName)
				Case MsgBoxResult.Cancel
					'           do nothing
					Exit Sub
			End Select
		Else
			With CurProj
				If .FileName <> "" Then Defaults.AddPreFile(.Directory & .FileName)
			End With
		End If
		UpdateFileMenu()
		
			frmTransient_Load(Me, New System.EventArgs())
		
	End Sub
	
	
	Public Sub mnuFileOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileOpen.Click
		
        Dim isFileOpen As Boolean
		
        On Error GoTo ErrHandler
		
        isFileOpen = False
		
		'   save the current project
		If CurProj.Saved = False Then
			If CurProj.Title = "" Then
				msgNotSavedWarning = "Do you want to save changes made to " & " the current project ?"
			Else
				msgNotSavedWarning = "Do you want to save changes made to " & CurProj.Title & " ?"
			End If
			
			Select Case MsgBox(msgNotSavedWarning, MsgBoxStyle.YesNoCancel, msgTitle)
				Case MsgBoxResult.Yes
					'           save
					mnuFileSave_Click(mnuFileSave, New System.EventArgs())
					If CurProj.FileName <> "" Then Defaults.AddPreFile(CurProj.Directory & CurProj.FileName)
				Case MsgBoxResult.No
					If CurProj.FileName <> "" Then Defaults.AddPreFile(CurProj.Directory & CurProj.FileName)
				Case MsgBoxResult.Cancel
					'           do nothing
					Exit Sub
			End Select
		Else
			With CurProj
				If .FileName <> "" Then Defaults.AddPreFile(.Directory & .FileName)
			End With
		End If
        UpdateFileMenu()

        With dlgFileOpen

            .Filter = "All Files (*.*)|*.*|DODO Files (*.mrs)|*.mrs"
            .FilterIndex = 2


            If CurProj.Directory = "" Then
                .InitialDirectory = DODODir
            Else
                .InitialDirectory = CurProj.Directory
            End If

            .ShowReadOnly = False
            .CheckFileExists = True
            .CheckPathExists = True

            .ShowDialog()
        End With
        Cursor = System.Windows.Forms.Cursors.WaitCursor

        CurProj = New Project

        '        ClearScreenData

        With CurProj
            CurVessel = .Vessel

            .GetDirNFileName(dlgFileOpen.FileName)
            Text = "DODO Console - " & .Title & " (" & .FileName & ")"

            If .ImportData(dlgFileOpen.FileName) Then
                CurVessel.Name = .Vessel.Name
                '                lblVessel.Caption = .Vessel.Name
                .Saved = True
                If Not .VesselParticular Then
                    MsgBox("DODO is unable to load vessel particulars.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, msgTitle)
                    '   Unload Me
                End If
            Else
                MsgBox(msgInputWarning, MsgBoxStyle.OkOnly, msgTitle)
                .Saved = False
            End If
        End With
        isFileOpen = False

        LoadData(True)

        Cursor = System.Windows.Forms.Cursors.Default

        Exit Sub

ErrHandler:
        '   User pressed Cancel button
        Cursor = System.Windows.Forms.Cursors.Default

    End Sub
	
	Public Sub mnuFileSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileSave.Click
		
        Dim isFileOpen As Boolean
		Dim FileNum As Integer
		
		On Error GoTo ErrorHandler
		
		SaveLC()
		
        isFileOpen = False
		FileNum = FreeFile
		
		With CurProj
			If .FileName = "" Then
				mnuFileSaveAs_Click(mnuFileSaveAs, New System.EventArgs())
			Else
				FileOpen(FileNum, .Directory & .FileName, OpenMode.Output)
                isFileOpen = True
				Cursor = System.Windows.Forms.Cursors.WaitCursor
				If Not .ExportData(FileNum) Then
				End If
				Cursor = System.Windows.Forms.Cursors.Default
				FileClose(FileNum)
                isFileOpen = False
				.Saved = True
			End If
		End With
		
		Exit Sub
		
ErrorHandler: 
		FileClose(FileNum)
		Cursor = System.Windows.Forms.Cursors.Default
		Exit Sub
		
	End Sub
	
	Public Sub mnuFileSaveAs_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileSaveAs.Click
		
        Dim isFileOpen As Boolean
		
		'   should the user cancel the dialog box, exit
		On Error GoTo ErrHandler
		
		SaveLC()
		
        isFileOpen = False
		
		'   get file name
		Dim fnum As Integer

        With dlgFileOpen

            .Filter = "All Files (*.*)|*.*|DODO Files (*.mrs)|*.mrs"
            .FilterIndex = 2
            If CurProj.Directory = "" Then
                .InitialDirectory = DODODir
            Else
                .InitialDirectory = CurProj.Directory
            End If
            If CurProj.FileName <> "" Then
                .FileName = CurProj.FileName
            Else
                .FileName = CurProj.Title & ".mrs"
            End If
            .ShowReadOnly = False
            .ShowDialog()

            '       save data
            fnum = FreeFile()
            FileOpen(fnum, .FileName, OpenMode.Output)
            isFileOpen = True

            Cursor = System.Windows.Forms.Cursors.WaitCursor
            With CurProj
                If .FileName <> "" Then
                    Defaults.AddPreFile(.Directory & .FileName)
                    '                UpdateFileMenu
                End If

                .GetDirNFileName(dlgFileOpen.FileName)
                Text = "DODO Console - " & .Title & " (" & .FileName & ")"

                If Not .ExportData(fnum) Then
                End If
                Cursor = System.Windows.Forms.Cursors.Default
                .Saved = True
            End With
            FileClose(fnum)
            isFileOpen = False
        End With
		
		Exit Sub
		
ErrHandler: 
		'   User pressed Cancel button
        If isFileOpen Then FileClose(fnum)
		Cursor = System.Windows.Forms.Cursors.Default
		If Err.Number <> 32755 Then ' don't report cancel dialog error
			MsgBox("Error " & Err.Description & ",  Source: " & Err.Source, MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Error")
		End If
		
	End Sub
	
	Public Sub mnuFilePre_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFilePre.Click
		Dim Index As Short = mnuFilePre.GetIndex(eventSender)
		
        Dim InputFile As String
        Dim isFileOpen As Boolean
		
		'   should the user cancel the dialog box, exit
		On Error GoTo ErrHandler
		
        isFileOpen = False
		
		'   save the current project
		If CurProj.Saved = False Then
			If CurProj.Title = "" Then
				msgNotSavedWarning = "Do you want to save changes made to " & " the current project ?"
			Else
				msgNotSavedWarning = "Do you want to save changes made to " & CurProj.Title & " ?"
			End If
			
			Select Case MsgBox(msgNotSavedWarning, MsgBoxStyle.YesNoCancel, msgTitle)
				Case MsgBoxResult.Yes
					'           save
					mnuFileSave_Click(mnuFileSave, New System.EventArgs())
					If CurProj.FileName <> "" Then Defaults.AddPreFile(CurProj.Directory & CurProj.FileName)
				Case MsgBoxResult.No
					If CurProj.FileName <> "" Then Defaults.AddPreFile(CurProj.Directory & CurProj.FileName)
					'           do nothing
				Case MsgBoxResult.Cancel
					Exit Sub
			End Select
		Else
			With CurProj
				If .FileName <> "" Then Defaults.AddPreFile(.Directory & .FileName)
			End With
		End If
		
		'   input from file
		'  InputFile = Defaults.PreFile(Index + 1)
		InputFile = VB.Right(mnuFilePre(Index).Text, Len(mnuFilePre(Index).Text) - 3)
        If Len(Dir(InputFile)) <= 0 Then
            MsgBox(msgNoFileWarning, MsgBoxStyle.OkOnly, msgTitle)
            Defaults.DelPreFile(Index + 1)
            UpdateFileMenu()
            Exit Sub
        End If
		
		Cursor = System.Windows.Forms.Cursors.WaitCursor
        CurProj = New Project
		
		With CurProj
			.GetDirNFileName(InputFile)
			Text = "DODO Console - " & .Title & " (" & .FileName & ")"
			
			If .ImportData(InputFile) Then
				.Saved = True
				CurVessel = .Vessel
                'lblVessel.Text = .Vessel.Name
                CurVessel.Name = .Vessel.Name
                'lblVessel.Caption = CurVessel.Name
				If Not .VesselParticular Then
					MsgBox("DODO is unable to load vessel particulars.", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, msgTitle)
				End If
			Else
				MsgBox(msgInputWarning, MsgBoxStyle.OKOnly, msgTitle)
				.Saved = False
			End If
		End With
		Cursor = System.Windows.Forms.Cursors.Default
		
		LoadData(True)
		
		Exit Sub
		
ErrHandler: 
		'   User pressed Cancel button
		Cursor = System.Windows.Forms.Cursors.Default
		MsgBox("mnuFilePre: " & Err.Description & ", Source: " & Err.Source, MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Error")
		
	End Sub
	
	Public Sub mnuFileExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileExit.Click
		
		Me.Close()
		
	End Sub

    Public Sub mnuOption_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

        frmOptions.Show()
        RefreshData()

    End Sub

    Private Sub UpdateFileMenu()
		
		Dim i, NumPreFiles As Short
		
		NumPreFiles = Defaults.CountPreFile
		
		If NumPreFiles > 0 Then mnuLinePreFile.Visible = True
		For i = 1 To NumPreFiles Step 1
			With mnuFilePre(i - 1)
				.Text = i & ". " & Defaults.PreFile(i)
				.Enabled = True
				.Visible = True
			End With
		Next i
		For i = NumPreFiles + 1 To MaxNumPreFiles Step 1
			With mnuFilePre(i - 1)
				.Text = ""
				.Enabled = False
				.Visible = False
			End With
		Next i
		
	End Sub
	
	Private Sub InitGrid()
		
		Dim i As Short
		Dim LUnit As String
		
		' Set up the transient grid
		
		ReDim TMRowHead(NumTimeSteps + 2)
		For i = 0 To NumTimeSteps + 1
			TMRowHead(i) = VB6.Format((i - 1) * TimeStep, "#,##0.0")
		Next i

        If IsMetricUnit Then
            LUnit = "(m)"
        Else
            LUnit = "(ft)"
        End If
        TMColHead(0) = "Time (sec)"
        TMColHead(1) = "X (E) " & LUnit
        TMColHead(2) = "Y (N) " & LUnit
        TMColHead(3) = "Heading (deg) TN"

        With grdTM
            .RowCount = 1
            .ColumnCount = 4
            For i = 0 To .ColumnCount - 1
                .Columns(i).HeaderText = TMColHead(i)

            Next
            .Columns(0).Width = 58
            .Columns(1).Width = 88
            .Columns(2).Width = 88
            .Columns(3).Width = 88
        End With

        With grdTM

        End With

        'Call SetupLineGrid(grdTM, TMRowHead, TMColHead, 0, VB6.PixelsToTwipsX(fraTransientMotion.Width) - 50, 300, VB6.PixelsToTwipsY(fraTransientMotion.Height) - SysInfo1.ScrollBarSize - 1200, NumTimeSteps + 1, SysInfo1, True, 2#, Me)
        'With grdTM
        '.Col = 0
        'For i = 1 To NumTimeSteps
        '.Row = i
        '.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
        'Next i
        'End With

    End Sub
	
	Private Sub RefreshData()

        setLabel()
        InitGrid()

    End Sub
	
	Private Sub ChangeCurEnv(ByRef EnvID As Short)

		Dim SelEnv As Metocean
		
		SelEnv = CurVessel.EnvLoad.Environments(EnvID)
		
		CurVessel.EnvLoad.EnvCur.Name = SelEnv.Name
		
	End Sub
	
    Private Sub frmTransient_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize

        'ResizePicture(picPolar, VB6.PixelsToTwipsX(fraTransientMotion.Left) + VB6.PixelsToTwipsX(fraTransientMotion.Width) + 150, VB6.PixelsToTwipsX(Me.ClientRectangle.Width) - 350, Me.ClientRectangle.Top - 150, VB6.PixelsToTwipsY(Me.ClientRectangle.Height) - 150)

    End Sub
	
	Private Sub txtInterval_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtInterval.Leave
		Dim TS As Single
		If txtInterval.Text = "" Then
			Exit Sub
		Else
			TS = CDbl(txtInterval.Text)
			If TS <= 0 Then
				MsgBox("Time interval must be > 0!", MsgBoxStyle.OKOnly, "DODO - Invalid Value Entered")
				Exit Sub
			ElseIf Fix((CDbl(txtDuration.Text) / TS) + 0.5) > MaxTimeSteps Then 
				MsgBox("Duration / Time Interval must be less than " & CStr(MaxTimeSteps), MsgBoxStyle.OKOnly, "DODO - Too Many Time Steps")
				Exit Sub
			Else
				TimeStep = TS
				NumTimeSteps = Fix((CDbl(txtDuration.Text) / TimeStep) + 0.5)
			End If
		End If
	End Sub
	
	Private Sub txtDuration_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDuration.Leave
		
		Dim i As Short
		Dim Dur As Single
		
		If txtDuration.Text = "" Then
			Exit Sub
		Else
			Dur = CDbl(txtDuration.Text)
			If Dur <= 0 Then
				MsgBox("Duration must be >0!", MsgBoxStyle.OKOnly, "DODO - Invalid Value Entered")
				Exit Sub
			ElseIf Fix((Dur / TimeStep) + 0.5) > MaxTimeSteps Then 
				MsgBox("Duration / Time Interval must be less than " & CStr(MaxTimeSteps), MsgBoxStyle.OKOnly, "DODO - Too Many Time Steps")
				Exit Sub
			Else
				NumTimeSteps = Fix((Dur / TimeStep) + 0.5)
				ReDim TMRowHead(NumTimeSteps + 1)
				For i = 0 To NumTimeSteps + 1
					TMRowHead(i) = VB6.Format((i - 1) * TimeStep, "#,##0.0")
				Next i
                'grdTM.Rows = NumTimeSteps + 2
                'With grdTM
                '.Col = 0
                'For i = 1 To NumTimeSteps + 1
                '.Row = i
                '.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
                '.Text = TMRowHead(i)
                'Next i
                'End With
			End If
		End If
		
	End Sub
	
    ' Private Sub txtVslSt_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtVslSt.TextChanged
    'Dim Index As Short = txtVslSt.GetIndex(eventSender)

    '   If Not FirstLoad Then CurProj.Saved = False

    'End Sub

    '
    'Private Sub txtWell_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWell.TextChanged
    '
    '       If Not FirstLoad Then CurProj.Saved = False

    '  End Sub

    ' Private Sub txtRiser_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRiser.TextChanged
    'Dim Index As Short = txtRiser.GetIndex(eventSender)

    '   If Not FirstLoad Then CurProj.Saved = False

    'End Sub

    Public Sub LoadData(Optional ByRef FirstTime As Boolean = False)

        If IsMetricUnit Then
            LFactor = 0.3048 ' ft -> m
            FrcFactor = 4.448222 ' kips -> KN
            MassFactor = 0.4536 ' kips -> MT
            DiaFactor = 25.4 ' in -> mm
        Else
            LFactor = 1
            FrcFactor = 1
            MassFactor = 1
            DiaFactor = 1
        End If

        If FirstTime Then FirstLoad = True

        With CurVessel
            txtWell.Text = VB6.Format(.WaterDepth * LFactor, "0.0")

            With .ShipCurGlob
                txtVslSt(0).Text = VB6.Format(.Xg * LFactor, "0.00")
                txtVslSt(1).Text = VB6.Format(.Yg * LFactor, "0.00")
                txtVslSt(2).Text = VB6.Format(RadTo360(.Heading), "0")
            End With

            txtVslSt(3).Text = VB6.Format(.ShipDraft * LFactor, "0.00")

            'With .Riser
            'txtRiser(0).Text = VB6.Format(.TopTen / 1000.0# * FrcFactor, "0.0")
            'txtRiser(1).Text = VB6.Format(.mass / 1000.0# * MassFactor, "0.0")
            'txtRiser(2).Text = VB6.Format(.Dia * 12.0# * DiaFactor, "0.0")
            'End With
        End With

    End Sub

    Public Sub SaveLC()

        If IsMetricUnit Then
            LFactor = 0.3048 ' ft -> m
            FrcFactor = 4.448222 ' kips -> kN
            MassFactor = 0.4536 ' kips -> MT
            DiaFactor = 25.4 ' in -> mm
        Else
            LFactor = 1
            FrcFactor = 1
            MassFactor = 1
            DiaFactor = 1
        End If

        With CurVessel
            With .ShipCurGlob
                .Xg = CDbl(CheckData(txtVslSt(0).Text, , True)) / LFactor
                .Yg = CDbl(CheckData(txtVslSt(1).Text, , True)) / LFactor
                .Heading = CDbl(CheckData(txtVslSt(2).Text, , True)) * Degrees2Radians
            End With
            .ShipDraft = CDbl(CheckData(txtVslSt(3).Text, , True)) / LFactor
            .WaterDepth = CDbl(CheckData(txtWell.Text, , True)) / LFactor
            ' .Riser.Length = .WaterDepth

            '  With .Riser
            ' .TopTen = CDbl(CheckData(txtRiser(0).Text, , True)) / FrcFactor * 1000.0#
            ' .mass = CDbl(CheckData(txtRiser(1).Text, , True)) / MassFactor * 1000.0#
            ' .Dia = CDbl(CheckData(txtRiser(2).Text, , True)) / DiaFactor / 12.0#
            'End With

        End With

    End Sub
End Class