Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmAnalysesB
	Inherits System.Windows.Forms.Form
	' frmAnalyses           analysis results (quasi-static)
	' Version 1.0.1
	' 2001, Copyright DTCEL, All Rights Reserved
	
	Private NumRuns As Short
	'    FileOpen As Boolean
	Private HeadE, HeadS, HeadStep As Single
	Private NumHead As Short
	Private hFile As Single
	Private FileName As String
	Private Directory As String
	Private msgInputWarning As String
	Dim FrcFactor, LFactor, VelFactor As Single
	Dim LUnit, VelUnit As String
	
	Private Sub btnReport_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnReport.Click
		Dim InitPos As New Motion
		
		InitPos.Surge = CSng(txtInitSurge.Text)
		InitPos.Sway = CSng(txtInitSway.Text)
		InitPos.Yaw = CSng(txtInitYaw.Text)

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim ox As New ExcelReporter
		If NumRuns = 0 Then
            HeadS = CSng(_txtHeadings_0.Text)
            HeadE = CSng(_txtHeadings_1.Text)
            NumHead = CShort(_txtHeadings_2.Text)
            If NumHead <= 1 Then
				If HeadE <> HeadS Then
					NumHead = 2
				Else
					NumHead = 1
				End If
			End If
			If NumHead > 1 Then
				HeadStep = (HeadE - HeadS) / (NumHead - 1)
			Else
				HeadStep = 1
			End If
			NumRuns = 3 * ((HeadE - HeadS) / HeadStep + 1)
		End If
		
		Call ox.ReportBatchResults(txtReportTitle.Text, (txtSubTitle.Text), NumRuns, CurVessel, InitPos, txtFile.Text, IsMetricUnit)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    ' form load and unload

    Private Sub frmAnalysesB_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		Dim i As Short
		
		Text = Text & " - " & CurProj.Title
		Cancelled = False
		RefreshData()
		txtReportTitle.Text = CurVessel.Name & ", Draft " & Format(CurVessel.ShipDraft, CStr(0)) & " ft, " & CurProj.WellSites.Item(CurProj.WellSites.CurWellNo).NameID
		txtSubTitle.Text = txtEnvironment.Text & ", " & CurVessel.DampingPercent.Surge & "% Damping"
	End Sub
	
	Private Sub RefreshData()
		LoadData()
		RefreshUnitLabels(Me)
	End Sub
	
	Private Sub frmAnalysesB_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		
		frmMain.Enabled = True
		frmMain.LoadData()
		
	End Sub
	
	' command buttons
	
	Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click
		
		Me.Close()
		
	End Sub
	
	Private Sub btnAnalysis_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnAnalysis.Click
		
		
		If Len(txtFile.Text) = 0 Then
			msgInputWarning = "You haven't specified a valid output file name."
			MsgBox(msgInputWarning, MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Error")
			Exit Sub
		End If
		
		If Len(txtNumLinesPerGroup.Text) = 0 Then
			msgInputWarning = "Must enter grouping pattern."
			MsgBox(msgInputWarning, MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Error")
			Exit Sub
		Else
			If Not IsNumeric(txtNumLinesPerGroup.Text) Then
				msgInputWarning = "Lines per group must be numeric."
				MsgBox(msgInputWarning, MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Error")
				Exit Sub
			End If
		End If
		
		Dim NumLines As Short
		NumLines = CurVessel.MoorSystem.MoorLineCount
		
		If NumLines Mod CShort(txtNumLinesPerGroup.Text) <> 0 Then
			msgInputWarning = "Incorrect grouping pattern for a mooring lauoyt of " & NumLines & " lines."
			MsgBox(msgInputWarning, MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Error")
			Exit Sub
		End If
		
		On Error GoTo ErrorHandler
		FileName = txtFile.Text
		
		Cursor = System.Windows.Forms.Cursors.WaitCursor
		Me.Enabled = False

        HeadS = CSng(_txtHeadings_0.Text)
        HeadE = CSng(_txtHeadings_1.Text)
        NumHead = CShort(_txtHeadings_2.Text)
        If NumHead <= 1 Then
			If HeadE <> HeadS Then
				NumHead = 2
			Else
				NumHead = 1
			End If
		End If
		If NumHead > 1 Then
			HeadStep = (HeadE - HeadS) / (NumHead - 1)
		Else
			HeadStep = 1
		End If
		
		Cancelled = False
		With frmProgress
			.Text = "Collinear Multi-heading Analyses"
			.CurrentStage.Text = "Initializing..."
			.Progress = 0
			.Show()
		End With
		
		StaticAnalysesB(HeadS, HeadE, HeadStep)
		
		With frmProgress
			.CurrentStage.Text = "Completed successfully."
			.Progress = 100
		End With
		frmProgress.Close()
		
		Me.Enabled = True
		'    Me.SetFocus
		Cursor = System.Windows.Forms.Cursors.Default
		
ErrorHandler: 
		Me.Enabled = True
		Cursor = System.Windows.Forms.Cursors.Default
		
	End Sub
	
	Private Sub btnEnvironment_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnEnvironment.Click
        frmEnviron.Show()
        txtEnvironment.Text = CurVessel.EnvLoad.EnvCur.Name
		
	End Sub
	
	Private Sub btnBrowse_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnBrowse.Click
		
		'   should the user cancel the dialog box, exit
		On Error GoTo ErrHandler

        With dlgFileOpen
            .Filter = "All Files (*.*)|*.*|MARS Result Files (*.sta)|*.sta"
            .FilterIndex = 2
            .InitialDirectory = Directory
            .FileName = FileName
            .ShowReadOnly = False
            .ShowDialog()

            GetDirNFileName(.FileName, Directory, FileName)
            txtFile.Text = Directory & FileName
        End With
        Exit Sub
		
ErrHandler: 
		'   User pressed Cancel button
		Exit Sub
		
	End Sub

    Private Sub txtEnvironment_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEnvironment.TextChanged
        txtSubTitle.Text = txtEnvironment.Text & ", " & CurVessel.DampingPercent.Surge & "% Damping"
    End Sub

    ' text boxes

    'Private Sub txtHeadings_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
    'Dim Index As Short = txtHeadings.GetIndex(eventSender)

    '   txtHeadings(Index).Text = CheckData(txtHeadings(Index).Text)

    'End Sub

    ' operation subroutines
    ' load to and update from form

    Private Sub LoadData()
		
		If IsMetricUnit Then
			LFactor = 0.3048 ' ft -> m
			FrcFactor = 4.448222 ' kips -> KN
		Else
			LFactor = 1
			FrcFactor = 1
		End If
		
		With CurVessel
			With .ShipCurGlob
                _txtVslSt_0.Text = Format(.Xg * LFactor, "0.0")
                _txtVslSt_1.Text = Format(.Yg * LFactor, "0.0")
                _txtVslSt_2.Text = Format(.Heading * Radians2Degrees, "0.00")
            End With
            _txtVslSt_3.Text = Format(.ShipDraft * LFactor, "0.00")
        End With
		
		txtEnvironment.Text = CurVessel.EnvLoad.EnvCur.Name
		
		With CurProj
			If .Vessel.Name <> "" Then
				FileName = .Vessel.Name & ".sta"
			Else
				FileName = .Title & ".sta"
			End If
			If .Directory = "" Then
				Directory = CurDir()
			Else
				Directory = .Directory
			End If
			If VB.Right(Directory, 1) <> "\" Then Directory = Directory & "\"
			txtFile.Text = Directory & FileName
		End With
		
	End Sub
	
	' analyses
	
	Private Sub StaticAnalysesB(ByVal HeadS As Single, ByVal HeadE As Single, ByVal HeadStep As Single)
		' no data directly from user interface, so already all in english units stored in objects
		
		'   If IsMetricUnit Then
		'       LFactor = 0.3048                 ' ft -> m
		'       FrcFactor = 4.448222           ' kips -> KN
		'       VelFactor = 0.5144444           ' knots - > m/s
		'   Else
		LFactor = 1
		FrcFactor = 1
		VelFactor = 1
		'   End If
		
		' Batch Analyses for Env of Collinear Multi-headings
		Dim CurRun As Object
		Dim j As Short
		Dim i, NumLines As Short
        Dim ShipLoc As ShipGlobal
        Dim WFM, LFM, Move_Renamed As Motion
		Dim FEv, FCr, FWd, FWv, Stiff As Force
		Dim SgnY, SgnX, SgnZ As Single
		Dim Tension(20) As Single
		Dim TSigLF(20) As Single
		Dim TSigWF(20) As Single
		Dim TenTotal(20) As Single
		Dim MaxTen As Single
		Dim N2ndTen, NMaxTen, NID As Short
		Dim minFOS, BS As Single
		Dim NminFOS, N2ndMinFOS As Short
		Dim SID As String
		Dim windHead, EnvHead, waveHead, currHead As Single
		
		'dchen 1/26/04
		Dim NumGL, NinGroup As Short
		
		'    Dim ShipYaw as Single
		
		Dim FileNum As Integer
		
		On Error GoTo ErrHandler
		
		Stiff = New Force
		
		'dchen 1/26/04
		NumGL = CShort(txtNumLinesPerGroup.Text)

        CurRun = 1

        NumRuns = 3 * ((HeadE - HeadS) / HeadStep + 1)
		
		FileNum = FreeFile
		If Len(FileName) = 0 Then ' if user did not give name use default
			FileName = My.Application.Info.DirectoryPath & "\Data\" & CurVessel.Name & "sta"
		End If
		
		If hFile > 0 Then ' if file already open, close it
            'AppActivate(hFile)
            Call System.Windows.Forms.SendKeys.SendWait("%{F4}")
		End If
		
		' remember individual headings before overwriting with batch heading
		windHead = CurVessel.EnvLoad.EnvCur.Wind.Heading
		waveHead = CurVessel.EnvLoad.EnvCur.Wave.Heading
		currHead = CurVessel.EnvLoad.EnvCur.Current.Heading
		
		FileOpen(FileNum, FileName, OpenMode.Output, OpenAccess.Write)
		
		Dim BS1, BS2 As Single
		For j = 1 To Fix((NumRuns - 1) / 3) + 1 Step 1
			EnvHead = HeadS + (j - 1) * HeadStep
			For NID = 1 To 3
				
				Select Case NID
					Case 1
						SID = "Intact"
					Case 2
						SID = "Damaged Most Critical Line"
					Case 3
						SID = "Damaged 2nd Critical Line"
				End Select
				
				CurVessel.EnvLoad.EnvCur.Wind.Heading = EnvHead * Degrees2Radians
				CurVessel.EnvLoad.EnvCur.Wave.Heading = EnvHead * Degrees2Radians + (waveHead - windHead)
				CurVessel.EnvLoad.EnvCur.Current.Heading = EnvHead * Degrees2Radians + (currHead - windHead)
				
				With frmProgress
					.lblBatchStage.Text = Format(EnvHead, "0.0") & "Invalid_string_refer_to_original_code" & SID
                    .ProgressBatch = CurRun / NumRuns * 99
                End With
				If Cancelled Then GoTo UserCancelled
				frmProgress.CurrentStage.Text = "Finding Equilibrium Position..."
				frmProgress.Progress = 25
				ShipLoc = CurVessel.FindEquilibriumPosition(Cancelled, frmProgress)
				
				PrintLine(FileNum, New String("-", 60) & vbCrLf)
				
				' output env loads in global coordinate system
				With CurVessel.EnvLoad
					.ShipHead = CurVessel.ShipCurGlob.Heading
					FEv = .FEnvGlob
					FWd = .FWindGlob( , False)
					FCr = .FCurrGlob( , False)
					FWv = .FWaveGlob( , False)
					
					PrintLine(FileNum, Trim(.EnvCur.Name), Format(.EnvCur.Wind.Heading * Radians2Degrees, "0.0"), Format(.EnvCur.Wave.Heading * Radians2Degrees, "0.0"), Format(.EnvCur.Current.Heading * Radians2Degrees, "0.0"), " (deg TN) ", SID)
					PrintLine(FileNum, "Env ", " Fx ", "  Fy ", " Mz ", "Metocean ")
					PrintLine(FileNum, "Loads", "kips", "kips", "kips-ft", "   ")
					PrintLine(FileNum, "Wind", Format(FWd.Fx / 1000# / FrcFactor, "0.0"), Format(FWd.Fy / 1000# / FrcFactor, "0.0"), Format(FWd.MYaw / FrcFactor / LFactor / 1000#, "0.0"), Format(.EnvCur.Wind.VelCorr / Knots2Ftps / VelFactor, "0.00"), " kn ")
					PrintLine(FileNum, "Curr", Format(FCr.Fx / 1000# / FrcFactor, "0.0"), Format(FCr.Fy / 1000# / FrcFactor, "0.0"), Format(FCr.MYaw / 1000# / FrcFactor / LFactor, "0.0"), Format(.EnvCur.Current.Profile(1).Velocity / Knots2Ftps / VelFactor, "0.00"), " kn ")
					PrintLine(FileNum, "Wave", Format(FWv.Fx / 1000# / FrcFactor, "0.0"), Format(FWv.Fy / 1000# / FrcFactor, "0.0"), Format(FWv.MYaw / 1000# / FrcFactor / LFactor, "0.0"), Format(.EnvCur.Wave.Height / LFactor, "0.00"), " ft ", Format(.EnvCur.Wave.Period, "#0.00"), " sec ")
					PrintLine(FileNum, "Total", Format(FEv.Fx / FrcFactor / 1000#, "0.0"), Format(FEv.Fy / 1000# / FrcFactor, "0.0"), Format(FEv.MYaw / FrcFactor / LFactor / 1000#, "####0.0"))
					PrintLine(FileNum)
				End With
				
				' output env loads in vessel local coordinate system
				With CurVessel.EnvLoad
					'                .ShipHead = ShipLoc.Heading
					.ShipHead = CurVessel.ShipCurGlob.Heading
					FEv = .FEnvLocl
					FWd = .FWindLocl( , False)
					FCr = .FCurrLocl( , False)
					FWv = .FWaveLocl( , False)
					
					PrintLine(FileNum, vbCrLf) ' add a blank line
					PrintLine(FileNum, Trim(.EnvCur.Name), Format(.EnvCur.Wind.Heading, "0.0"), Format(.EnvCur.Wave.Heading, "0.0"), Format(.EnvCur.Current.Heading, "0.0"), " (deg TN) ", SID)
					PrintLine(FileNum, "Env ", "Surge", "Sway", "Yaw", "Metocean ")
					PrintLine(FileNum, "Loads", "kips", "kips", "kips-ft", "   ")
					PrintLine(FileNum, "Wind", Format(FWd.Fx / 1000# / FrcFactor, "0.0"), Format(FWd.Fy / 1000# / FrcFactor, "0.0"), Format(FWd.MYaw / FrcFactor / LFactor / 1000#, "0.0"), Format(.EnvCur.Wind.VelCorr / Knots2Ftps / VelFactor, "0.00"), " ft ")
					PrintLine(FileNum, "Curr", Format(FCr.Fx / 1000# / FrcFactor, "0.0"), Format(FCr.Fy / 1000# / FrcFactor, "0.0"), Format(FCr.MYaw / 1000# / FrcFactor / LFactor, "0.0"), Format(.EnvCur.Current.Profile(1).Velocity / Knots2Ftps / VelFactor, "0.00"), " kn ")
					PrintLine(FileNum, "Wave", Format(FWv.Fx / 1000# / FrcFactor, "0.0"), Format(FWv.Fy / 1000# / FrcFactor, "0.0"), Format(FWv.MYaw / 1000# / FrcFactor / LFactor, "0.0"), Format(.EnvCur.Wave.Height / LFactor, "0.00"), " ft ", Format(.EnvCur.Wave.Period, "0.00"), " sec ")
					PrintLine(FileNum, "Total", Format(FEv.Fx / FrcFactor / 1000#, "0.0"), Format(FEv.Fy / 1000# / FrcFactor, "0.0"), Format(FEv.MYaw * FrcFactor / LFactor / 1000#, "0.0"))
					PrintLine(FileNum)
				End With
				With FWv
					SgnX = System.Math.Sign(.Fx)
					SgnY = System.Math.Sign(.Fy)
					
					'                SgnZ = Sgn(.MYaw)
					
					'                ShipYaw = CurVessel.ShipCurGlob.Heading - ShipLoc.Heading
					'                If ShipYaw > PI Then ShipYaw = ShipYaw - 2 * PI
					'                If ShipYaw < -PI Then ShipYaw = ShipYaw + 2 * PI
					'                SgnZ = Sgn(ShipYaw)
					
					SgnZ = 0#
					
					If SgnX = 0# Then SgnX = 1#
					If SgnY = 0# Then SgnY = 1#
					'                If SgnZ = 0# Then SgnZ = 1#
				End With
				
				Call CurVessel.MoorSystem.MoorStiff(Stiff, ShipLoc, False)
				With Stiff
					PrintLine(FileNum, "    ", "Surge", "Sway", "Yaw")
					PrintLine(FileNum, "Moor ", "kips/ft", "kips/ft", "kips-ft/ft", " ")
					PrintLine(FileNum, "Stiff", Format(.Fx / 1000# / FrcFactor, "0.0"), Format(.Fy / 1000# / FrcFactor, "0.0"), Format(.MYaw / 1000# / FrcFactor / LFactor, "0.0"))
					PrintLine(FileNum)
				End With
				
				Move_Renamed = CurVessel.ShipMotion((CurVessel.ShipCurGlob), ShipLoc)
				With Move_Renamed
					PrintLine(FileNum, "Vessel", "Surge", "Sway", "Yaw")
					PrintLine(FileNum, "Move ", "ft", "ft", "deg TN", " ")
					PrintLine(FileNum, "Mean", Format(.Surge / LFactor, "0.00"), Format(.Sway / LFactor, "0.00"), Format(.Yaw * Radians2Degrees, "0.00"))
				End With
				
				If Cancelled Then GoTo UserCancelled
				frmProgress.CurrentStage.Text = "Finding LF Motions ..."
				frmProgress.Progress = 50
				LFM = CurVessel.GetSigLFM(Stiff, Cancelled, frmProgress, ShipLoc)
				
				If Cancelled Then GoTo UserCancelled
				frmProgress.CurrentStage.Text = "Finding WF Motions ..."
				frmProgress.Progress = 60
				WFM = CurVessel.GetSigWFM(Cancelled, frmProgress, ShipLoc.Heading)
				With LFM
					PrintLine(FileNum, "LFSig", Format(.Surge / LFactor, "0.00"), Format(.Sway / LFactor, "0.00"), Format(.Yaw * Radians2Degrees, "0.00"))
					.Surge = .Surge * SgnX
					.Sway = .Sway * SgnY
					.Yaw = .Yaw * SgnZ
				End With
				With WFM
					PrintLine(FileNum, "WFSig", Format(.Surge / LFactor, "0.00"), Format(.Sway / LFactor, "0.00"), Format(.Yaw * Radians2Degrees, "0.00"))
					PrintLine(FileNum)
					.Surge = .Surge * SgnX
					.Sway = .Sway * SgnY
					.Yaw = .Yaw * SgnZ
				End With
				
				If Cancelled Then GoTo UserCancelled
				frmProgress.CurrentStage.Text = "Finding LF Tensions ..."
				frmProgress.Progress = 70
				Call CurVessel.FindSigDynMLTen(Tension, TSigLF, ShipLoc, LFM)
				If Cancelled Then GoTo UserCancelled
				frmProgress.CurrentStage.Text = "Finding WF Tensions ..."
				frmProgress.Progress = 95
				Call CurVessel.FindSigDynMLTen(Tension, TSigWF, ShipLoc, WFM)
				PrintLine(FileNum, " No", "Tmean", "TSigLF", "TSigWF")
				PrintLine(FileNum, "    ", " kips", " kips", " kips")
				
				NumLines = CurVessel.MoorSystem.MoorLineCount
				
				For i = 1 To NumLines
					PrintLine(FileNum, i, Format(Tension(i) / 1000# / FrcFactor, "0.0"), Format(TSigLF(i) / 1000# / FrcFactor, "0.0"), Format(TSigWF(i) / 1000# / FrcFactor, "0.0"))
					
					TenTotal(i) = Tension(i) + Max(TSigLF(i) + TSigWF(i) * 1.86, TSigLF(i) * 1.5 + TSigWF(i))
				Next i
				
				Select Case NID
					Case 1 ' Intact Condition
						
						'                  Find most loaded line (lowest FOS when different BS used) number
						'                    MaxTen = -100000
						'                    For i = 1 To NumLines
						'                        If TenTotal(i) >= MaxTen Then
						'                            MaxTen = TenTotal(i)
						'                            NMaxTen = i
						'                        End If
						'                    Next
						
						' assume design ensures Fairlead FOS governs to simplify
						minFOS = 100000
						For i = 1 To NumLines
							BS = CurVessel.MoorSystem.MoorLines(i).Segments(1).BS
							If BS / TenTotal(i) <= minFOS Then
								minFOS = BS / TenTotal(i)
								NMaxTen = i
							End If
						Next 
						
						'  Find Most loaded Line (lowest FOS) group Number
						'  Find 2nd loaded line which is in the same group as the most loaded line
						'dchen 1/26/04
						NinGroup = NMaxTen Mod NumGL

                        If NMaxTen = 1 Then
                            BS1 = CurVessel.MoorSystem.MoorLines(NMaxTen + 1).Segments(1).BS
                            BS2 = CurVessel.MoorSystem.MoorLines(NumLines).Segments(1).BS
                        ElseIf NMaxTen = NumLines Then
                            BS1 = CurVessel.MoorSystem.MoorLines(1).Segments(1).BS
                            BS2 = CurVessel.MoorSystem.MoorLines(NMaxTen - 1).Segments(1).BS
                        Else
                            BS1 = CurVessel.MoorSystem.MoorLines(NMaxTen + 1).Segments(1).BS
							BS2 = CurVessel.MoorSystem.MoorLines(NMaxTen - 1).Segments(1).BS
						End If
						
						If NinGroup > 1 Then
							If BS1 / TenTotal(NMaxTen + 1) <= BS2 / TenTotal(NMaxTen - 1) Then
								N2ndTen = NMaxTen + 1
							Else
								N2ndTen = NMaxTen - 1
							End If
						ElseIf NinGroup = 1 Then 
							N2ndTen = NMaxTen + 1
						Else
							If NumGL > 1 Then
								N2ndTen = NMaxTen - 1
							Else
								If NMaxTen = 1 Then
									If BS1 / TenTotal(NMaxTen + 1) <= BS2 / TenTotal(NumLines) Then
										N2ndTen = NMaxTen + 1
									Else
										N2ndTen = NumLines
									End If
								ElseIf NMaxTen = NumLines Then 
									If BS1 / TenTotal(1) <= BS2 / TenTotal(NMaxTen - 1) Then
										N2ndTen = 1
									Else
										N2ndTen = NMaxTen - 1
									End If
								Else
									If BS1 / TenTotal(NMaxTen + 1) <= BS2 / TenTotal(NMaxTen - 1) Then
										N2ndTen = NMaxTen + 1
									Else
										N2ndTen = NMaxTen - 1
									End If
								End If
							End If
						End If
						
						' Since Case 1 Intact calc is complete,
						' now breaking most loaded line (lowest FOS) for Case 2
						CurVessel.MoorSystem.MoorLines(NMaxTen).Connected = False
						
					Case 2
						' after case 2 calc, now break 2nd loaded line (2nd lowest FOS) for case 3
						CurVessel.MoorSystem.MoorLines(NMaxTen).Connected = True
						CurVessel.MoorSystem.MoorLines(N2ndTen).Connected = False
						
					Case 3
						
						' after Case 3 calc, now restore all broken lines
						CurVessel.MoorSystem.MoorLines(N2ndTen).Connected = True
						
				End Select
                CurRun = CurRun + 1
            Next NID
			PrintLine(FileNum) '  add a blank line to separate results from different headings
		Next j
UserCancelled: 
		'dchen -- do not restore in case user wants them excluded on purpose
		' restore mooring to all connected state
		'    With CurVessel.MoorSystem
		'        For i = 1 To .MoorLineCount
		'            .MoorLines(i).Connected = True
		'        Next i
		'    End With
		
		With frmProgress
			.lblBatchStage.Text = "Analysis completed"
			.ProgressBatch = 100
		End With
		
		FileClose(FileNum)
        Stiff = Nothing
        ' open the result file for user to view
        '   hFile = Shell("notepad " & FileName, vbNormalFocus)
        Exit Sub
ErrHandler: 
		FileClose(FileNum)
        Stiff = Nothing
        If Len(Err.Description) > 0 Then
			MsgBox("Error: " & Err.Description & " from " & Err.Source, MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Error")
		End If
	End Sub
End Class