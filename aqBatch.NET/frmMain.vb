Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmMain
	Inherits System.Windows.Forms.Form
	Private ProgHandle As Integer
	
	Private Const InFile As Short = 1
	Private Const OutFile As Short = 2

    Private Const AQWAPath1 As String = "c:\progra~1\ansysi~1\"
    Private Const AQWAPath2 As String = "\aqwa\bin\winx64\aqwa /nowind"
    Private BaseDir, rBaseDir As String
	
	Private ABFile, AFFile As String

    Private HeaderLabels(4) As String
    Private HeaderLabelsIndex(4) As Short
    Private HeaderLabels1(17) As String
	Private HeaderLabels2(17) As String
	
	Private Changed As Boolean
	Private ExistingTxt As String
	Private CheckingGrid As Boolean
	Private JustEnterCell As Boolean
	
	Private AQWAUnitsUS As Boolean
	Private Collinear As Boolean
    Private UDEF As Boolean
    Private seaType As New DataGridViewComboBoxCell()
    Private SwellType As New DataGridViewComboBoxCell()

    Private Sub btnBrowseBaseDir_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnBrowseBaseDir.Click

        On Error GoTo Errhandler
        With frmBrowseDir
            .Text = "Set Base Directory"
            .InitPath = lblBaseDir.Text
            .ControlToUpdate = lblBaseDir
            VB6.ShowForm(frmBrowseDir, 1, Me)
            'frmBrowseDir.Show()

        End With
        If lblBaseDir.Text <> "" And Trim(lblTargetDir.Text) = Trim(lblBaseDir.Text) Then
            MsgBox("Target directory should be different from Base directory.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Target Dierctory")
            lblTargetDir.Text = ""
            WorkDir = ""
        End If
        If lblBaseDir.Text <> "" Then

            frmPostProc.ReadBaseResults()
            CheckRunButtonState()
        End If
        Exit Sub
Errhandler:
        MsgBox("Error encountered when pick up base directory:" & lblBaseDir.Text)
        lblBaseDir.Text = ""
    End Sub

    Private Sub btnBrowseDir_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnBrowseDir.Click
		Dim nPos As Short
        Dim Path1, Path2 As String

        On Error GoTo ErrHandler

        With frmBrowseDir
			.Text = "Set Target Directory"
			If Len(lblTargetDir.Text) = 0 And Len(lblBaseDir.Text) > 0 Then
				.InitPath = lblBaseDir.Text
			Else
				.InitPath = lblTargetDir.Text
			End If
            .ControlToUpdate = lblTargetDir
            frmBrowseDir.Show()
            If Trim(lblTargetDir.Text) = Trim(lblBaseDir.Text) Then
				MsgBox("Target directory should be different from Base directory.", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Target Dierctory")
				lblTargetDir.Text = ""
				WorkDir = ""
			End If
		End With
        WorkDir =lblTargetDir.Text
        ' assume BaseDir is at same level as WorkDir (e.g. 25YH) where the AQWA runs

        If Len(lblBaseDir.Text) > 0 And Len(lblTargetDir.Text) > 0 Then
			
			nPos = InStrRev(lblBaseDir.Text, "\", -1, CompareMethod.Text)
			Path1 = VB.Left(lblBaseDir.Text, nPos)
			nPos = InStrRev(lblTargetDir.Text, "\")
			Path2 = VB.Left(lblTargetDir.Text, nPos)
			
			If StrComp(Path1, Path2, CompareMethod.Text) <> 0 Then
				MsgBox("Base and Working Directory should be on the same level for relative path to work", MsgBoxStyle.Information, "Warning")
				Exit Sub
			End If
		End If
        CheckRunButtonState()
        Exit Sub
Errhandler:
        MsgBox("Error encountered when pick up working directory:" & lblTargetDir.Text)
        lblTargetDir.Text = ""
    End Sub
	
	Private Sub btnCopyAll_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCopyAll.Click
        Dim i, k, NumPairs As Short
        ' copy case 1 to all cases
        For k = 2 To NumCases
            oMet(k).Wind.Velocity = oMet(1).Wind.Velocity
            oMet(k).Wave.Height = oMet(1).Wave.Height
            oMet(k).Wave.Period = oMet(1).Wave.Period
            oMet(k).Wave.gamma = oMet(1).Wave.gamma
            oMet(k).Wave.SpectrumName = oMet(1).Wave.SpectrumName
            oMet(k).Current.SurfaceVel = oMet(1).Current.SurfaceVel
            NumPairs = oMet(k).Current.ProfileCount

            For i = NumPairs To 1 Step -1
                oMet(k).Current.ProfileDelete((1)) 'After deleting first element, second one will become the "first" one
            Next
            For i = 1 To oMet(1).Current.ProfileCount
                oMet(k).Current.ProfileAdd(oMet(1).Current.Profile(i).Depth, oMet(1).Current.Profile(i).Velocity)
            Next i
            If UDEF Then
                oMet(k).Wave.SwellHeight = oMet(1).Wave.SwellHeight
                oMet(k).Wave.SwellPeriod = oMet(1).Wave.SwellPeriod
                oMet(k).Wave.SwellSpectrumName = oMet(1).Wave.SwellSpectrumName
                oMet(k).Wave.Swellgamma = oMet(1).Wave.Swellgamma
            End If
        Next k
		LoadGrid()
	End Sub
	
	Private Sub btnCreateFiles_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCreateFiles.Click
		Dim s As String

        UpdateMetOceanObjects()
		UpdateBrokenLines()
		
		On Error GoTo ErrHandler
		
		BaseDir = lblBaseDir.Text
		
		If BaseDir = "" Then Exit Sub
		
		ABFile = BaseDir & "\ABRUN.DAT"
		AFFile = BaseDir & "\AFRUN.DAT"
		
		Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
		If CreateAqwaInputFiles Then
            If CreateAqwaBatchFile() Then
                If CreateSilentRunBatchFile() Then
                    MsgBox("AQWA input Files have been successfully created.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Create Files")
                End If
            End If
        End If
		
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Exit Sub
ErrHandler: 
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Select Case Err.Number
			Case CInt("53")
				s = " Error: " & Err.Description & vbCrLf & "Looking for files:" & vbCrLf & "ABRUN.DAT, AFRUN.DAT, ABRUN.RES, ABRUN.HYD, ABRUN.EQP" & vbCrLf & "in the base directory:" & vbCrLf & BaseDir
				MsgBox(s, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "File Error")
			Case Else
				MsgBox("Error encountered when creating AQWA files: " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "Create Files Error")
		End Select
	End Sub
	
	Private Sub btnIntactFile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnIntactFile.Click
		Dim oxApp As Microsoft.Office.Interop.Excel.Application
		Dim msg, fname As String
        '   should the user cancel the dialog box, exit
        On Error GoTo ErrHandler

        dlgFileOpen.Filter = "All Files (*.*)|*.*|Intact (*.xls)|*.xls"
        dlgFileSave.Filter = "All Files (*.*)|*.*|Intact (*.xls)|*.xls"
		dlgFileOpen.FilterIndex = 2
		dlgFileSave.FilterIndex = 2
		dlgFileOpen.InitialDirectory = WorkDir
		dlgFileSave.InitialDirectory = WorkDir
        dlgFileOpen.CheckFileExists = True
        dlgFileOpen.CheckPathExists = True
		dlgFileSave.CheckPathExists = True
        dlgFileOpen.ShowReadOnly = False
        dlgFileOpen.ShowDialog()
		dlgFileSave.FileName = dlgFileOpen.FileName
		
		fname = dlgFileOpen.FileName
		
		If fname <> "" Then
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            oxApp = ExcelGlobal_definst.Workbooks.Open(FileName:=fname).Application
			Call ReadCriticalLines(oxApp)
			oxApp.Quit()
            oxApp = Nothing
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Exit Sub
		
ErrHandler: 
		'   User pressed Cancel button, or some tragedy occurred
		FileClose(InFile)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Len(Err.Description) > 0 Then
			MsgBox(Err.Description)
		End If
		
	End Sub

    Private Function ReadCriticalLines(ByRef oxApp As Microsoft.Office.Interop.Excel.Application) As Object
        Dim NumLines, i, NCase As Short

        With oxApp.ActiveWorkbook.Sheets("Summary")
            ' determine NumLines from xls
            NumLines = .Range("D30").End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Row - 29

            NCase = oxApp.Sheets.Count - 2
            If NCase > 0 Then
                ReDim Preserve L1Tmax(NCase)
                ReDim Preserve L2Tmax(NCase)
            End If

            For i = 3 To oxApp.Sheets.Count - 1

                L1Tmax(i - 2) = .Cells(29 + NumLines + 2, i)
                L2Tmax(i - 2) = .Cells(29 + NumLines + 3, i)
            Next i
        End With
    End Function

    Private Sub btnOpenMatrix_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnOpenMatrix.Click
        mnuOpen_Click()
    End Sub

    Private Sub btnsaveMatrix_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnSaveMatrix.Click
        mnuSave_Click()
    End Sub

    Private Sub btnPostProcess_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnPostProcess.Click
		If CheckAQWAErrors() Then Exit Sub
		
		If optMoorState(0).Checked Then
			IsDamaged = 0
		Else
			If optDamageLineType(0).Checked Then ' damage most Critical
				IsDamaged = 1
			Else ' damage second Critical
				IsDamaged = 2
			End If
		End If
		
		UpdateBrokenLines()
		With frmPostProc
			If AQWAUnitsUS Then
				.optUnit(0).Checked = True
			Else
				.optUnit(1).Checked = True
			End If
			
			If .ReadBaseResults Then
                frmPostProc.Show()
                'VB6.ShowForm(frmPostProc, 1, Me)
            End If
		End With
	End Sub
	
	Private Sub btnRun_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnRun.Click

        ' run AQWA batch file	
        Dim s As String
		
		ChDir(WorkDir)
		s = "RunAqwaCases.bat"
		On Error GoTo ErrHandler
		ProgHandle = Shell(s)
		
		Exit Sub
ErrHandler: 
		ReportError(err, "Error", "Error launching AQWA Batch File.")
		
	End Sub

    Private Function CreateAqwaInputFiles() As Boolean
		Dim TmpStr, msg As String
		Dim tmpVal As Double
		
		If WaterDepth = 0 Then frmPostProc.ReadBaseResults()
        tmpVal = oMet(1).Current.Profile(oMet(1).Current.ProfileCount).Depth

        If oMet(1).Current.ProfileCount > 1 Then
            If WaterDepth - tmpVal > 0 Then
                msg = "Current Profile Depth must be greater than or equal to the deepest anchor water depth. " & vbCrLf & "water depth is " & WaterDepth & vbCrLf & "Profile Depth is " & tmpVal
                GoTo ErrHandler
            End If
        End If

        If optMoorState(0).Checked Then
			IsDamaged = 0
		Else
			If optDamageLineType(0).Checked Then ' damage most Critical
				IsDamaged = 1
			Else ' damage second Critical
				IsDamaged = 2
			End If
		End If
		
		CreateAqwaInputFiles = False
		
		Const ForReading As Short = 1
		Const ForWriting As Short = 2
		Const ForAppending As Short = 3
		Const TristateUseDefault As Short = -2
		Const TristateTrue As Short = -1
		Const TristateFalse As Short = 0
		Dim ts, fs, f, MyFile As Object
		Dim s, ss As String
		Dim i As Short
		
		fs = CreateObject("Scripting.FileSystemObject")
		
		On Error GoTo ErrHandler

        ' read ABFile
        f = fs.GetFile(ABFile)
        ts = f.OpenAsTextStream(ForReading, TristateFalse)
        s = ts.Readall
        ts.Close()

        Dim s5, s3, s1, s2, s4, s6 As String
		Dim pos6, pos4, pos2, pos1, pos3, pos5, pos7 As Integer
		
		'find deck 11 and 14 and 15
		pos1 = InStr(1, s, "    11    ") ' Environment
		
		' should only replace env deck for each case
		' using everything else from templates
		
		pos2 = InStr(pos1, s, "    14    ") ' Start of Mooring Deck
		pos6 = InStr(pos1, s, "    13    SPEC")
		
		' remember wind spectrum name in base file
		If pos6 > 0 Then pos6 = InStr(pos6 + 5, s, vbCrLf)
		If pos6 > 0 Then pos7 = InStr(pos6 + 1, s, vbCrLf)
		If pos6 > 0 And pos7 > 0 Then WSPEC = Mid(s, pos6 + 2, pos7 - pos6 - 2)
		
		pos4 = InStr(pos1, s, "    13UDWD") ' find user-defined spectrum
		If pos4 > 0 Then pos5 = InStr(pos1, s, "    13WIND")
		pos3 = InStr(pos1, s, "    15    ") ' End of Mooring Deck
		
		' chop file into 3 sections
		s1 = VB.Left(s, pos1 - 1) ' Beginning to environment decks
		s3 = VB.Right(s, Len(s) - pos3) ' after mooring deck to end of file
		
		' if user defined spectrum used, need to save that info
		If pos4 > 0 Then
			s6 = Mid(s, pos4, pos5 - pos4)
			If Len(s6) > 0 Then UDWS = s6
		End If
		
		' Formulate mooring deck
		s2 = Mid(s, pos2, pos3 - pos2 + 1) ' base mooring deck
		
		s4 = s2 '  Intact mooring deck from ABFile
		
		'    If InStr(1, s, "14NLID") > 0 Then
		'       s4 = Replace(s2, "14NLIN", "14NLID")     ' mooring deck
		'    Else
		'         s4 = Replace(s2, "14NLID", "14NLIN")     ' mooring deck
		'    End If
		
		s5 = s4

        ' read AFFile
        f = fs.GetFile(AFFile)
        ts = f.OpenAsTextStream(ForReading, TristateFalse)
        ss = ts.Readall
        ts.Close()

        Dim ss3, ss1, ss2, ss4 As String
		Dim pp2, pp1, pp3 As Integer
		
		'find deck 11 and 14
		pp1 = InStr(1, ss, "    11    ")
		pp3 = InStr(pp1, ss, "    14    MOOR")
		pp2 = InStr(pp3, ss, "    15    ")
		
		ss1 = VB.Left(ss, pp1 - 1) ' beginning to env deck
        ss2 = VB.Right(ss, Len(ss) - pp2) ' Deck 15 to end
        ss3 = Mid(ss, pp3, pp2 - pp3 + 1) ' Intact mooring deck from AFFile
		ss4 = ss3
		
		'    On Error GoTo 0
		For i = 1 To NumCases
			If CreateEnvFile(i) Then
				If optMoorState(1).Checked Then ' Damaged State
					s4 = CreateMoorDeckWithBrokenLine(s5, i)
					ss4 = CreateMoorDeckWithBrokenLine(ss3, i)
				End If

                ' get env file content
                f = fs.GetFile(My.Application.Info.DirectoryPath & "\tmp" & i & ".txt")
                ts = f.OpenAsTextStream(ForReading, TristateFalse)
                s = ts.Readall
                ts.Close()

                If optMoorState(0).Checked Then ' Intact
                    If Not fs.FolderExists(WorkDir & "\case" & i) Then
                        Call fs.CreateFolder(WorkDir & "\case" & i)
                    End If
                    MyFile = fs.CreateTextFile(WorkDir & "\case" & i & "\abrun.dat", True)
                    MyFile.Write(s1 & s & s4 & s3)
                    MyFile.Close()
                    If chkPFLH.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                        MyFile = fs.CreateTextFile(WorkDir & "\case" & i & "\afrun1.dat", True)
                        Mid(ss1, 17, 4) = "DRFT"
					Else
                        MyFile = fs.CreateTextFile(WorkDir & "\case" & i & "\afrun.dat", True)
                    End If
                    MyFile.Write(ss1 & s & ss4 & ss2)
                    MyFile.Close()

                    If chkPFLH.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                        MyFile = fs.CreateTextFile(WorkDir & "\case" & i & "\afrun2.dat", True)
                        Mid(ss1, 17, 4) = "WFRQ"
                        MyFile.Write(ss1 & s & ss4 & ss2)
                        MyFile.Close()
                    End If
				Else
					If optDamageLineType(0).Checked = True Then
						TmpStr = "D1"
					Else
						TmpStr = "D2"
					End If
                    If Not fs.FolderExists(WorkDir & "\case" & i & TmpStr) Then
                        Call fs.CreateFolder(WorkDir & "\case" & i & TmpStr)
                    End If

                    MyFile = fs.CreateTextFile(WorkDir & "\case" & i & TmpStr & "\abrun.dat", True)
                    MyFile.Write(s1 & s & s4 & s3)
                    MyFile.Close()

                    If chkPFLH.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                        MyFile = fs.CreateTextFile(WorkDir & "\case" & i & TmpStr & "\afrun1.dat", True)
                        Mid(ss1, 17, 4) = "DRFT"
					Else
                        MyFile = fs.CreateTextFile(WorkDir & "\case" & i & TmpStr & "\afrun.dat", True)
                    End If
                    MyFile.Write(ss1 & s & ss4 & ss2)
                    MyFile.Close()

                    If chkPFLH.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                        MyFile = fs.CreateTextFile(WorkDir & "\case" & i & TmpStr & "\afrun2.dat", True)
                        Mid(ss1, 17, 4) = "WFRQ"
                        MyFile.Write(ss1 & s & ss4 & ss2)
                        MyFile.Close()
                    End If
				End If
			Else
				' error creating env file
			End If
		Next i
		On Error GoTo 0
		On Error Resume Next
		Kill(My.Application.Info.DirectoryPath & "\tmp*.txt") ' cleanup tmp env files
		CreateAqwaInputFiles = True
        Exit Function

ErrHandler: 
		MsgBox("Create Input File Error: " & vbCrLf & msg & vbCrLf & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "Error")
	End Function
	
	Private Function CreateMoorDeckWithBrokenLine(ByVal s As String, ByVal CaseNo As Short) As String
		' s original mooring deck string, allows grouping of mooring lines using same 14COMP card
		' and END was used for the last mooring line  14NLI card
		' to simplify programming -- user needs to ensure this format in base file
		
		
		Dim NumLines As Short
		' get number of mooring lines
		Dim Fields() As String
		Dim MatchStr As String
		
		If InStr(1, s, "14NLIN") > 0 Then
			MatchStr = "14NLIN"
		Else
			MatchStr = "14NLID"
		End If

        Fields = Split(s, MatchStr)

        If UBound(Fields) > 0 Then
			NumLines = UBound(Fields)
		Else
			MsgBox("Total number of Lines must be greater than zero.")
			Exit Function
		End If

        Dim Count As Short
        Dim pos2, pos1, breakPos As Integer
		Dim FirstLineStart As Object
		Dim LastLineStart As Short
		
		' look for first vbcrlf after "14    MOOR" deck
		pos1 = InStr(1, s, "14    MOOR")

        FirstLineStart = InStr(pos1, s, "    14NLI")

        Count = 0
		pos1 = 1
		While pos1 > 0 And Count < NumLines - 1
			If InStr(pos1, s, "    14NLI") > 0 Then
				pos2 = InStr(pos1, s, "    14NLI")
				Count = Count + 1
			Else
				pos1 = -1 '  exit while
			End If
			If Count = BreakLine(CaseNo) Then
				breakPos = pos2
				pos1 = -1 '  broken line found, exit while
			Else
				pos1 = pos2 + 5 ' move to search next mooring line
			End If
		End While
		If BreakLine(CaseNo) = NumLines Then ' if last line broken
			breakPos = pos2
		End If
		If breakPos > 0 Then
			pos2 = InStrRev(s, vbCrLf, breakPos, CompareMethod.Binary) + 1
			If BreakLine(CaseNo) = 1 Then ' If first line broke
                ' if line 1 is the only line in the group to use the ECAT ECAH cards
                ' else keep COMP cards
                pos1 = InStr(FirstLineStart, s, vbCrLf, CompareMethod.Binary)
                pos2 = FirstLineStart - 4
            ElseIf BreakLine(CaseNo) = NumLines Then 
				LastLineStart = pos2 '  last found line before NumLines
				pos2 = InStr(1, s, " END14NLI")
				pos1 = InStr(pos2, s, vbCrLf, CompareMethod.Binary)
				pos2 = InStr(LastLineStart + 1, s, vbCrLf, CompareMethod.Binary)
			Else ' 1 < BreakLineBreakLine(CaseNo) < NumLines
				pos1 = InStr(breakPos + 1, s, vbCrLf, CompareMethod.Binary)
				' assume stand alone line to find end of previous line before 14COMP - 14COMP will be deleted as well with the broken line
				pos2 = InStr(InStrRev(s, "    14NLI", breakPos - 1) + 1, s, vbCrLf, CompareMethod.Binary)
				' check if group exists, if so keep 14COMP data
				If InStr(Mid(s, breakPos - 80, 80), "14NLI") > 0 Or InStr(Mid(s, breakPos + 30, 80), "14NLI") > 0 Then ' grouped lines exist above or below the line to be broken
					pos2 = InStrRev(s, vbCrLf, pos1 - 1, CompareMethod.Binary) ' get end of previous line
				End If
			End If
			s = Replace(s, Mid(s, pos2, pos1 - pos2 + 1), "")
			If BreakLine(CaseNo) = NumLines Then
				pos1 = InStrRev(s, "    14NLI")
				Mid(s, pos1, 9) = " END14NLI"
			End If
		End If
		CreateMoorDeckWithBrokenLine = s
	End Function

    Private Function CreateAqwaBatchFile() As Boolean
        CreateAqwaBatchFile = False
        Dim Path1, subDir, Path2 As String
        Dim nPos As Short

        If Len(BaseDir) = 0 Or Len(WorkDir) = 0 Then
            MsgBox("Root Level not recommended to run batch jobs" & vbCrLf & "Please select a directory")
            Exit Function
        End If
        ' assume BaseDir is at same level as WorkDir (e.g. 25YH) where the AQWA runs

        nPos = InStrRev(BaseDir, "\", -1, CompareMethod.Text)
        rBaseDir = ".." & "\" & VB.Right(BaseDir, Len(BaseDir) - nPos)
        Path1 = VB.Left(BaseDir, nPos)
        nPos = InStrRev(WorkDir, "\")
        rWorkDir = ".." & "\" & VB.Right(WorkDir, Len(WorkDir) - nPos)
        Path2 = VB.Left(WorkDir, nPos)

        If StrComp(Path1, Path2, CompareMethod.Text) <> 0 Then
            MsgBox("Base and Working Directory should be on the same level for relative path to work", MsgBoxStyle.Information, "Warning")
            Exit Function
        End If

        Const ForReading As Short = 1
        Const ForWriting As Short = 2
        Const ForAppending As Short = 3
        Const TristateUseDefault As Short = -2
        Const TristateTrue As Short = -1
        Const TristateFalse As Short = 0
        Dim MyFile, s, f, fs, ts, ss, i As Object

        On Error GoTo ErrHandler

        fs = CreateObject("Scripting.FileSystemObject")

        MyFile = fs.CreateTextFile(WorkDir & "\RunAqwaCases.bat", True)

        For i = 1 To NumCases

            If IsDamaged = 1 Then
                subDir = "case" & i & "D1"
            ElseIf IsDamaged = 2 Then
                subDir = "case" & i & "D2"
            Else
                subDir = "case" & i
            End If

            MyFile.WriteLine("COPY " & rBaseDir & "\ABRUN.RES ABRUN.RES")
            ' If AQWAVersion Then
            MyFile.WriteLine("COPY " & rBaseDir & "\ABRUN.EQP ABRUN.EQP")
            'Else
            'MyFile.WriteLine("COPY " & rBaseDir & "\ABRUN.POS ABRUN.POS")
            'MyFile.WriteLine("COPY " & rBaseDir & "\ABRUN.HYD ABRUN.HYD")
            ' End If
            MyFile.WriteLine()
            ' copy data file to workdir to run
            MyFile.WriteLine("COPY " & rWorkDir & "\" & subDir & "\ABRUN.DAT ABRUN.DAT")
            ' If AQWAVersion Then ' new AQWA
            If Len(txtAQWAVersion.Text) > 0 Then
                '     If V12 Then
                MyFile.WriteLine(AQWAPath1 & Trim(txtAQWAVersion.Text) & AQWAPath2 & " LIBRIUM RUN")
                'Else
                'MyFile.WriteLine(Trim(txtAQWAVersion.Text) & " LIBRIUM RUN")

            End If
            '           If V12 Then
            '   MyFile.WriteLine(AQWA12Path & "AQWA LIBRIUM RUN")
            '     Else
            '     MyFile.WriteLine("AQWA LIBRIUM RUN")
            '             End If
            '   End If
            '   Else
            '   MyFile.WriteLine("LIBRIUM RUN")
            'End If
            '  ' COPY SELECTED RESULTS BACK TO CASE DIRECTORY
            MyFile.WriteLine("COPY ABRUN.LIS " & rWorkDir & "\" & subDir & "\ABRUN.LIS" & vbCrLf)

            MyFile.WriteLine("COPY ABRUN.RES AFRUN.RES")

            MyFile.WriteLine("COPY ABRUN.EQP AFRUN.EQP")

            ' copy data file to workdir to run

            If chkPFLH.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                MyFile.WriteLine("COPY " & rWorkDir & "\" & subDir & "\AFRUN1.DAT AFRUN.DAT")
            Else
                MyFile.WriteLine("COPY " & rWorkDir & "\" & subDir & "\AFRUN.DAT AFRUN.DAT")
            End If


            If Len(txtAQWAVersion.Text) > 0 Then
                MyFile.WriteLine(AQWAPath1 & Trim(txtAQWAVersion.Text) & AQWAPath2 & " FER RUN")
            End If
            ' COPY SELECTED RESULTS BACK TO CASE DIRECTORY
            If chkPFLH.CheckState = System.Windows.Forms.CheckState.Checked Then
                MyFile.WriteLine("COPY AFRUN.LIS " & rWorkDir & "\" & subDir & "\AFRUN.LIS" & vbCrLf)
            Else
                MyFile.WriteLine("COPY AFRUN.LIS " & rWorkDir & "\" & subDir & "\AFRUN1.LIS" & vbCrLf)

                MyFile.WriteLine("COPY AFRUN.LIS " & rWorkDir & "\" & subDir & "\AFRUN2.LIS" & vbCrLf)
                MyFile.WriteLine()
            End If

        Next i
        MyFile.Close()
        CreateAqwaBatchFile = True
        Exit Function
ErrHandler:
        ReportError(Err, "Error", "Error Creating AQWA Batch File")
        If MyFile.Open Then MyFile.Close()
    End Function

    Function CreateSilentRunBatchFile() As Boolean
        CreateSilentRunBatchFile = False
        Const ForReading As Short = 1
        Const ForWriting As Short = 2
        Const ForAppending As Short = 3
        Const TristateUseDefault As Short = -2
        Const TristateTrue As Short = -1
        Const TristateFalse As Short = 0
        Dim MyFile, s, f, fs, ts, ss, i As Object

        On Error GoTo ErrHandler

        fs = CreateObject("Scripting.FileSystemObject")

        If fs.FileExists(WorkDir & "\stdtests.com") Then fs.DeleteFile(WorkDir & "\stdtests.com")
        If fs.FileExists(WorkDir & "\stdtests.com") Then fs.DeleteFile(WorkDir & "\stdtests.com")
        If fs.FileExists(WorkDir & "\RunSilent.bat") Then fs.DeleteFile(WorkDir & "\RunSilent.bat")

        On Error GoTo ErrHandler

        ' read bat file
        f = fs.GetFile(WorkDir & "\RunAqwaCases.bat")
        ts = f.OpenAsTextStream(ForReading, TristateFalse)
        s = ts.Readall
        ' replace content AQWA with RUN
        If Len(txtAQWAVersion.Text) > 0 Then
            ss = Replace(s, Trim(txtAQWAVersion.Text), "RUN")
        Else
            ss = Replace(s, "AQWA", "RUN")
        End If

        ' won't support long file name by having to remove quotes
        'ss = Replace(ss, Chr(34), "")

        ' must use full path name for copy command

        ' ss = Replace(ss, " A", " " & WorkDir & "\A")

        ts.Close()

        f = fs.CreateTextFile(WorkDir & "\stdtestscom.bak", True)
        f.Write(ss)
        f.Close()

        ' rename batch file
        fs.CopyFile(WorkDir & "\stdtestscom.bak", WorkDir & "\stdtests.com")

        ' create new silent batch
        f = fs.CreateTextFile(WorkDir & "\RunSilent.bat", True)
        If Len(txtAQWAVersion.Text) > 0 Then
            f.Write(Trim(txtAQWAVersion.Text) & " std")
        Else
            f.Write("aqwa std")
        End If
        f.Close()

        CreateSilentRunBatchFile = True
		Exit Function
ErrHandler: 
		ReportError(err, "Error", "Error Creating AQWA Silent Batch File")
        If f.Open Then f.Close()

    End Function
	
	Private Function CreateEnvFile(ByVal Index As Short) As Boolean
		CreateEnvFile = False
		
		Dim s, msg As String
		Dim NumPts, i As Short
		' write a formatted temp file for env decks
		' ASSUME ALL MUST BE IN ENGLISH UNITS
		
		Dim Fvel, Flen As Double
		Dim Omg, Spec As Double
		Dim iOmg, nOmg As Short
		Dim buff As String
		Dim strBuff As New VB6.FixedLengthString(10)
		
		On Error GoTo ErrHandler
		FileOpen(OutFile, My.Application.Info.DirectoryPath & "\tmp" & Index & ".txt", OpenMode.Output)
		FileWidth(OutFile, 80)
		
		If AQWAUnitsUS Then
			Fvel = 1
			Flen = 1
		Else
			Fvel = 0.5144
			Flen = 0.3048
		End If


        PrintLine(OutFile, "    11    ENVR")
			If optWind(1).Checked Then
            PrintLine(OutFile, "    11WIND", TAB(13), CStr(Format(oMet(Index).Wind.Velocity * Flen, "#0.00")), TAB(22), CStr(VB6.Format(oMet(Index).Wind.Heading, "#0.00")))
        End If
        ' CONVERT FROM KNOTS TO FT/S FOR AQWA
        PrintLine(OutFile, "    11CURR", TAB(13), CStr(Format(oMet(Index).Current.SurfaceVel * Flen, "#0.00")), TAB(22), CStr(VB6.Format(oMet(Index).Current.Heading, "#0.00")))

        NumPts = oMet(Index).Current.ProfileCount

        If NumPts > 1 Then
				For i = NumPts To 1 Step -1
                PrintLine(OutFile, "    11CPRF", TAB(13), Format(oMet(Index).Current.Profile(i).Depth * (-1) * Flen, "#0"), TAB(22), Format((oMet(Index).Current.SurfaceVel - oMet(Index).Current.Profile(i).Velocity) * Flen, "#0.00"))
            Next i
			End If
        PrintLine(OutFile, " END11CDRN", TAB(13), CStr(Format(Val(oMet(Index).Current.Heading - 180.0#), "#0.00"))) ' assume deg used
        PrintLine(OutFile, "    12    NONE")
			PrintLine(OutFile, "    13    SPEC")
			If optWind(0).Checked Then ' 1-hr wind using spectrum
				If Len(UDWS) > 0 Then ' if user-defined spectrum
					Print(OutFile, UDWS)
				ElseIf Len(WSPEC) > 0 Then 
					PrintLine(OutFile, WSPEC)
				End If
            PrintLine(OutFile, "    13WIND", TAB(22), CStr(VB6.Format(oMet(Index).Wind.Velocity * Flen, "#0.00")), TAB(32), CStr(oMet(Index).Wind.Heading), TAB(42), VB6.Format(32.8 * Flen, "#0.0")) ' assume 10m ref height, i.e. 33 ft
        End If
			If UDEF Then
				Print(OutFile, "    13SPDN")
            PrintLine(OutFile, TAB(22), CStr(VB6.Format(Val(oMet(Index).Wave.Heading), "#0.0")))
            s = Trim(oMet(Index).Wave.SwellSpectrumName)
            If InStr(s, "GAUS") > 0 Then
                PrintLine(OutFile, "    13XSWL GAUS", TAB(42), CStr(Format(Val(oMet(Index).Wave.Swellgamma), "#0.0000")), TAB(52), CStr(Format(Val(CStr(oMet(Index).Wave.SwellHeight * Flen)), "#0.0")), TAB(62), CStr(Format(Val(CStr(2 * 3.1415926 / oMet(Index).Wave.SwellPeriod)), "#0.000")), TAB(72), CStr(VB6.Format(Val(oMet(Index).Wave.SwellHeading), "#0.0")))
            ElseIf InStr(s, "JONH") Then
                PrintLine(OutFile, "    13XSWL JONH", TAB(42), CStr(Format(Val(oMet(Index).Wave.Swellgamma), "#0.0000")), TAB(52), CStr(Format(Val(CStr(oMet(Index).Wave.SwellHeight * Flen)), "#0.0")), TAB(62), CStr(Format(Val(CStr(2 * 3.1415926 / oMet(Index).Wave.SwellPeriod)), "#0.000")), TAB(72), CStr(VB6.Format(Val(oMet(Index).Wave.SwellHeading), "#0.0")))
            Else ' PSMZ
                PrintLine(OutFile, "    13XSWL PSMZ", TAB(42), CStr(Format(Val(oMet(Index).Wave.Swellgamma), "#0.0000")), TAB(52), CStr(Format(Val(CStr(oMet(Index).Wave.SwellHeight * Flen)), "#0.0")), TAB(62), CStr(Format(Val(CStr(2 * 3.1415926 / oMet(Index).Wave.SwellPeriod)), "#0.000")), TAB(72), CStr(VB6.Format(Val(oMet(Index).Wave.SwellHeading), "#0.0")))
            End If
			Else
				PrintLine(OutFile, "    13RADS")
				Print(OutFile, "    13SPDN")
            PrintLine(OutFile, TAB(22), CStr(Format(Val(oMet(Index).Wave.Heading), "#0.0")))
        End If

        If UDEF And 1 = 0 Then
			nOmg = (2# - 0.1) / 0.1 + 1
			For iOmg = 1 To nOmg
				Omg = 0.1 + (iOmg - 1) * 0.1
				Spec = oMet(Index).Wave.GetSpectrumValue(Omg)
				If Not AQWAUnitsUS Then
					Spec = Spec * 0.3048 ^ 2
				End If
				buff = "    13UDEF"
				strBuff.Value = ""
				buff = buff & strBuff.Value
				strBuff.Value = VB6.Format(Omg, "0.0000")
				buff = buff & strBuff.Value
				strBuff.Value = VB6.Format(Spec, "0.000E+00")
				buff = buff & strBuff.Value
				PrintLine(OutFile, buff)
			Next iOmg
			PrintLine(OutFile, " END13")
		Else
            s = Trim(oMet(Index).Wave.SpectrumName)
            If InStr(s, "JONH") > 0 Then
                PrintLine(OutFile, " END13" & s, TAB(22), "0.200", TAB(32), " 1.500", TAB(42), CStr(oMet(Index).Wave.gamma), TAB(52), CStr(Format(oMet(Index).Wave.Height * Flen, "#0.000")), TAB(62), CStr(Format(Val(CStr(2 * 3.1415926 / oMet(Index).Wave.Period)), "#0.000")))
            Else ' PSMZ
                PrintLine(OutFile, " END13" & s, TAB(22), "0.200", TAB(32), " 1.500", TAB(42), CStr(VB6.Format(oMet(Index).Wave.Height * Flen, "#0.000")), TAB(52), CStr(Format(oMet(Index).Wave.Period / 1.4, "#0.000")))
            End If
		End If
		FileClose(OutFile)
		CreateEnvFile = True
		Exit Function
ErrHandler: 
		FileClose(OutFile)
		ReportError(err, "Error", "Create Env File Error " & msg)
	End Function

    Private Sub btnSilentRun_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnSilentRun.Click
		Dim s As String
		ChDir(WorkDir)
        ' run AQWA batch file
        s = "RunSilent.bat"
		On Error GoTo ErrHandler
		ProgHandle = Shell(s)
		
		Exit Sub
ErrHandler: 
		ReportError(err, "Error", "Error launching AQWA Batch File.")
	End Sub

    Private Sub chkCollinear_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkCollinear.CheckStateChanged

        Select Case chkCollinear.CheckState
            Case 0
                Collinear = False
            Case 1
                Collinear = True
        End Select

    End Sub

    Private Sub chkUDEF_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkUDEF.CheckStateChanged

        With grdMatrix
            Select Case chkUDEF.CheckState
                Case 0
                    .ColumnCount = 11

                    UDEF = False
                Case 1
                    .ColumnCount = 16

                    UDEF = True
            End Select
        End With
        InitiateGrid()
    End Sub

    ' form load and unload

    Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		WaterDepth = 0
		
		CheckingGrid = True
		AQWAUnitsUS = True
		Collinear = True
        'UDEF = False

        chkCollinear.CheckState = System.Windows.Forms.CheckState.Checked
		optAQWAUnits(0).Checked = True
		btnRun.Enabled = False
		btnSilentRun.Enabled = False
		
		SetDefaults()
        InitiateCombo()

        InitiateGrid()
		
		SelectText(txtNumCases)
		
	End Sub
	
	Private Sub CheckRunButtonState()
		btnRun.Enabled = False
		btnSilentRun.Enabled = False
		If NumCases > 0 Then

            If lblBaseDir.Text <> "" And lblTargetDir.Text <> "" Then
                    btnRun.Enabled = True
                    btnSilentRun.Enabled = True
                End If

        End If
		RefreshBackColor()
		btnCreateFiles.Enabled = btnRun.Enabled
		
		' if result files exist, enable get results button
		Dim fs As Object
		fs = CreateObject("Scripting.FileSystemObject")
		Dim TmpStr As String
		If optDamageLineType(0).Checked Then
			TmpStr = "D1"
		ElseIf optDamageLineType(1).Checked Then 
			TmpStr = "D2"
		Else
			TmpStr = ""
		End If

        btnPostProcess.Enabled = fs.FileExists(WorkDir & "\case1" & TmpStr & "\abrun.lis")

    End Sub
	
	Private Function CheckAQWAErrors() As Boolean
		CheckAQWAErrors = False
		Const ForReading As Short = 1
		Const ForWriting As Short = 2
		Const ForAppending As Short = 3
		Const TristateUseDefault As Short = -2
		Const TristateTrue As Short = -1
		Const TristateFalse As Short = 0
		Dim ts, fs, f, MyFile As Object
		Dim s, ss As String
		Dim i As Short
		
		fs = CreateObject("Scripting.FileSystemObject")
		
		On Error GoTo ErrHandler
		
		Dim s4, s2, s1, s3, s5 As String
		Dim pos3, pos1, pos2, pos4 As Integer
		Dim s0 As String
		
		Dim subFolder As String
		
		For i = 1 To NumCases
			
			If optMoorState(0).Checked Then
				subFolder = "\case" & i
				IsDamaged = 0
			ElseIf optDamageLineType(0).Checked Then 
				subFolder = "\case" & i & "D1"
				IsDamaged = 1
			Else
				subFolder = "\case" & i & "D2"
				IsDamaged = 2
			End If

            ' read ABrun
            f = fs.GetFile(WorkDir & subFolder & "\abrun.lis")
            ts = f.OpenAsTextStream(ForReading, TristateFalse)

            s = ts.Readall
            ts.Close()

            '    If i = 1 Then
            '        s0 = s
            '    Else
            If StrComp(s0, s, CompareMethod.Binary) = 0 Then
				s2 = s2 & "Incorrect result file for " & WorkDir & subFolder & "\abrun.lis" & vbCrLf
			Else
				s0 = s
			End If
			'    End If
			
			
			'    ' check if security access bumped by looking at time card
			'    pos1 = InStr(1, s, "TIME:")
			'
			'    If i = 1 Then
			'        sTime = Mid(s, pos1 + 5, 8)
			'    Else
			'        If TimeEarlierThan(sTime, Mid(s, pos1 + 5, 8)) Then        ' not reliable since cases may be rerun
			'            sTime = Mid(s, pos1 + 5, 8)
			'        Else
			'            s2 = s2 & "Time stamp incorrect in file: " & WorkDir & subFolder & "\abrun.lis" & vbCrLf
			'        End If
			'    End If
			
			' search for errors
			pos1 = InStr(1, s, "*** ERROR")
			pos2 = InStr(1, s, "ABORT")
			'pos3 = InStr(1, s, "BELOW MINIMUM ")
			pos3 = InStr(1, s, " FAILED TO CONVERGE ")
			pos4 = InStr(1, s, "**** TERMINATED WITH ERRORS")
			
			If pos1 > 0 Then
				s2 = s2 & "'ERROR' found in file: " & WorkDir & subFolder & "\abrun.lis" & vbCrLf
			End If
			If pos2 > 0 Then
				s2 = s2 & "'ABORTED' found in file: " & WorkDir & subFolder & "\abrun.lis" & vbCrLf
			End If
			'If pos3 > 0 Then
			'    s2 = s2 & "' FAILED TO CONVERGE ' found in file: " & WorkDir & subFolder & "\abrun.lis" & vbCrLf
			'End If
			'If pos4 > 0 Then
			'    s2 = s2 & "'**** TERMINATED WITH ERRORS' found in file: " & WorkDir & subFolder & "\abrun.lis" & vbCrLf
			'End If
		Next i
		
		If Len(s2) > 0 Then
			CheckAQWAErrors = True
			MsgBox(s2, MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "AQWA Error Check")
		End If
		
		Exit Function
ErrHandler: 
		CheckAQWAErrors = True
		MsgBox("Error: " & Err.Description & vbCrLf & "Check AQWA Errors in file: " & WorkDir & subFolder & "\abrun.lis", MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "Error")
	End Function
	
	Private Sub frmMain_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		frmBrowseDir.Close()
		Cleanup()
		End
	End Sub
	
	Private Sub SetDefaults()
		WorkDir = CurDir()
		NumCases = 4
		txtNumCases.Text = CStr(NumCases)
        ReDim oMet(NumCases)
        Dim i As Short
		For i = 1 To NumCases
			oMet(i) = New Metocean
		Next i
	End Sub
    ' menus
    Private Sub grdMatrix_ClickEvent(ByVal eventSender As System.Object, ByVal e As DataGridViewCellEventArgs) _
        Handles grdMatrix.CellClick

        If e.ColumnIndex = 9 Then
            With frmCurrProfile
                .lblCaseNo.Text = grdMatrix.CurrentRow.HeaderCell.Value
                .txtNumProfilePts.Text = (grdMatrix.CurrentRow.Cells(9).Value)
            End With
            VB6.ShowForm(frmCurrProfile, 1, Me)
        ElseIf (e.columnIndex = 2 Or e.ColumnIndex = 13) And e.RowIndex > -1 Then
            Dim cell As New DataGridViewComboBoxCell()

            cell.MaxDropDownItems = 3
            cell.FlatStyle = FlatStyle.Flat
            cell.Items.Add("PSMZ")
            cell.Items.Add("JOHN")
            cell.Items.Add("GAUS")

            cell.Value = "PSMZ"
            grdMatrix(e.ColumnIndex, e.RowIndex) = cell

        End If
    End Sub

    Private Sub grdMatrix_EnterCell(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        ExistingTxt = grdMatrix.Text
        JustEnterCell = True
    End Sub

    Private Sub grdMatrix_LeaveCell(ByVal eventSender As System.Object, ByVal e As DataGridViewCellEventArgs) _
        Handles grdMatrix.CellLeave

        If Not CheckingGrid Then
            With grdMatrix
                If (e.ColumnIndex = 2 Or e.ColumnIndex = 13) Then

                End If

                If Trim(.Text) <> ExistingTxt And .ColumnCount > 1 Then
                    If Trim(.Text) <> "" Then .Text = CheckData(.Text)
                    Changed = True
                End If
            End With
        End If
    End Sub

    Private Sub grdMatrix_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        grdMatrix_LeaveCell(grdMatrix, New System.EventArgs())
    End Sub

    Private Sub lblTargetDir_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblTargetDir.TextChanged
        'TODO JLIU, never got called.
        WorkDir = lblTargetDir.Text
    End Sub


    Public Sub mnuAbout_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        VB6.ShowForm(frmAbout, 1, Me)
    End Sub

    Public Sub mnuCreateFiles_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        btnCreateFiles_Click(btnCreateFiles, New System.EventArgs())
    End Sub

    Public Sub mnuExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Me.Close()
    End Sub

    Public Sub mnuOpen_Click()

        Dim msg, fname As String

        '   should the user cancel the dialog box, exit
        On Error GoTo ErrHandler

        dlgFileOpen.Filter = "All Files (*.*)|*.*|Env Matrix (*.mat)|*.mat"
        dlgFileSave.Filter = "All Files (*.*)|*.*|Env Matrix (*.mat)|*.mat"
        dlgFileOpen.FilterIndex = 2
        dlgFileSave.FilterIndex = 2
        dlgFileOpen.CheckFileExists = True
        dlgFileOpen.CheckPathExists = True
        dlgFileSave.CheckPathExists = True
        dlgFileOpen.ShowReadOnly = False
        dlgFileOpen.ShowDialog()
        dlgFileSave.FileName = dlgFileOpen.FileName


        fname = dlgFileOpen.FileName
        dlgFileOpen.InitialDirectory = System.IO.Path.GetDirectoryName(fname)
        dlgFileSave.InitialDirectory = System.IO.Path.GetDirectoryName(fname)


        If fname <> "" Then
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            '     ReadOldFile (fname)
            If Not ReadDataFromFile(fname) Then GoTo ErrHandler
            txtMetoceanName.Text = FileRootName(fname)
            WorkDir = dlgFileOpen.InitialDirectory
            '  lblBaseDir.Text = dlgFileOpen.InitialDirectory
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        CheckRunButtonState()
        Exit Sub

ErrHandler:
        '   User pressed Cancel button, or some tragedy occurred
        FileClose(InFile)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Len(Err.Description) > 0 Then
            MsgBox(Err.Description)
        End If
    End Sub

    Public Sub mnuPostProcess_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        btnPostProcess_Click(btnPostProcess, New System.EventArgs())
    End Sub

    Public Sub mnuPreProcess_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        btnRun_Click(btnRun, New System.EventArgs())
    End Sub

    Public Sub mnuSave_Click()

        ' validate no empty fields

        Call UpdateMetOceanObjects()
        Call UpdateBrokenLines()

        If txtMetoceanName.Text <> "" Then
            dlgFileOpen.FileName = txtMetoceanName.Text & ".mat"
            dlgFileSave.FileName = txtMetoceanName.Text & ".mat"
        Else
            dlgFileOpen.FileName = "*.mat"
            dlgFileSave.FileName = "*.mat"
        End If
        dlgFileOpen.Filter = "All Files (*.*)|*.*|Env Matrix (*.mat)|*.mat"
        dlgFileSave.Filter = "All Files (*.*)|*.*|Env Matrix (*.mat)|*.mat"
        dlgFileOpen.FilterIndex = 2
        dlgFileSave.FilterIndex = 2
        'dlgFileOpen.InitialDirectory = CurDir()
        'dlgFileSave.InitialDirectory = CurDir()
        dlgFileSave.OverwritePrompt = True
        dlgFileOpen.ShowReadOnly = False
        dlgFileSave.ShowDialog()
        dlgFileOpen.FileName = dlgFileSave.FileName

        '   open the file for output
        Dim fname As String
        fname = dlgFileOpen.FileName
        If fname = "" Then Exit Sub
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        WriteDataToFile(fname)
        CheckRunButtonState()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    ' operation subroutines
    ' initiating

    Private Sub InitiateCombo()

        seaType.Items.Clear()
        seaType.MaxDropDownItems = 3

        seaType.Items.Add("PSMZ")
        seaType.Items.Add("JONH")
        seaType.Items.Add("GAUS")

        'seaType.v = "PSMZ"
        SwellType.Items.Clear()
        SwellType.MaxDropDownItems = 3
        SwellType.Items.Add("PSMZ")
        SwellType.Items.Add("JONH")
        SwellType.Items.Add("GAUS")

        'SwellType.Value = "PSMZ"
    End Sub

    '
    'Handle the DataGridView.CellPainting event to draw text for each header cell

    Private Sub DataGridView1_CellPainting(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles grdMatrix.CellPainting
        If e.RowIndex = -1 AndAlso e.ColumnIndex > -1 Then
            e.PaintBackground(e.CellBounds, False)
            Dim r2 As Rectangle = e.CellBounds

            r2.Y += e.CellBounds.Height / 3
            r2.Height = e.CellBounds.Height / 3
            e.PaintContent(r2)
            e.Handled = True
            End If
    End Sub

    'Handle the DataGridView.Paint event to draw "merged" header cells

    Private Sub grdMatrix_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles grdMatrix.Paint
        'Handle the DataGridView.Paint event to draw "merged" header cells
        Dim j1 As Short
        With grdMatrix
            j1 = 0
            For j As Integer = 0 To .ColumnCount - 1
                If j = 0 Or j = 2 Or j = 7 Or j = 10 Or j = 11 Then

                    Dim r1 As Rectangle = .GetCellDisplayRectangle(j, -1, True)

                    r1.X += 1
                    r1.Y += 1
                    r1.Width = r1.Width * HeaderLabelsIndex(j1) - 2
                    r1.Height = r1.Height / 3 - 3

                    Using br As SolidBrush = New SolidBrush(.ColumnHeadersDefaultCellStyle.BackColor)
                        e.Graphics.FillRectangle(br, r1)
                    End Using

                    Using p As Pen = New Pen(SystemColors.InactiveBorder)
                        e.Graphics.DrawLine(p, r1.X, r1.Bottom, r1.Right, r1.Bottom)
                    End Using

                    Using format As StringFormat = New StringFormat()
                        Using br As SolidBrush = New SolidBrush(.ColumnHeadersDefaultCellStyle.ForeColor)
                            format.Alignment = StringAlignment.Center
                            format.LineAlignment = StringAlignment.Center
                            e.Graphics.DrawString(HeaderLabels(j1), .ColumnHeadersDefaultCellStyle.Font, Brushes.Black, r1, format)
                        End Using
                    End Using
                    j1 += 1
                End If

            Next
        End With
    End Sub

    Private Sub InitiateGrid()
        Dim c As Short

        Call SetLabels()

        With grdMatrix

            If UDEF Then
                .ColumnCount = 16
            Else
                .ColumnCount = 11
            End If

            For c = 0 To .ColumnCount - 1
                .Columns(c).FillWeight = 100 / .ColumnCount
                .Columns(c).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Columns(c).HeaderText = HeaderLabels1(c + 1) & vbCrLf & HeaderLabels2(c + 1)
            Next
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        End With

        fraDamageLineType.Enabled = False
        optDamageLineType(0).Checked = False
        optDamageLineType(1).Checked = False
        optDamageLineType(0).Enabled = False
        optDamageLineType(1).Enabled = False
    End Sub

    Private Sub SetLabels()

        HeaderLabels(0) = "Wind"
        HeaderLabels(1) = "Wave"
        HeaderLabels(2) = "Current"
        HeaderLabels(3) = "Break"
        HeaderLabels(4) = "Swell"

        HeaderLabelsIndex(0) = 2
        HeaderLabelsIndex(1) = 5
        HeaderLabelsIndex(2) = 3
        HeaderLabelsIndex(3) = 1
        HeaderLabelsIndex(4) = 5

        HeaderLabels1(0) = ""
        HeaderLabels1(1) = "Velocity"
		HeaderLabels1(2) = "Heading"
		HeaderLabels1(3) = "Spectrum"
		HeaderLabels1(4) = "Hs"
		HeaderLabels1(5) = "Tp"
		HeaderLabels1(6) = "Gamma"
		HeaderLabels1(7) = "Heading"
		HeaderLabels1(8) = "Heading"
		HeaderLabels1(9) = "Surf. Vel"
		HeaderLabels1(10) = "Profile"
		HeaderLabels1(11) = "Mooring"
		HeaderLabels1(12) = "Hs"
		HeaderLabels1(13) = "Tp"
		HeaderLabels1(14) = "Spectrum"
		HeaderLabels1(15) = "Gamma"
		HeaderLabels1(16) = "Heading"
		
		HeaderLabels2(0) = ""
        HeaderLabels2(1) = "(kts)"
        HeaderLabels2(2) = "(deg)"
		HeaderLabels2(3) = "Name"
		HeaderLabels2(4) = "(ft)"
		HeaderLabels2(5) = "(sec)"
		HeaderLabels2(6) = ""
		HeaderLabels2(7) = "(deg)"
		HeaderLabels2(8) = "(deg)"
        HeaderLabels2(9) = "(kts)"
        HeaderLabels2(10) = "Points"
		HeaderLabels2(11) = "Line"
		HeaderLabels2(12) = "(ft)"
		HeaderLabels2(13) = "(sec)"
		HeaderLabels2(14) = ""
		HeaderLabels2(15) = ""
        HeaderLabels2(16) = "(deg)"
    End Sub


    'Private Sub optAQWAVersion_Click(Index As Integer)
    '    If Index = 1 Then
    '        chkPFLH.Value = vbUnchecked
    '        chkPFLH.Enabled = False
    '    Else
    '        chkPFLH.Value = vbChecked
    '        chkPFLH.Enabled = True
    '    End If
    'End Sub

    Private Sub optAQWAUnits_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optAQWAUnits.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Short = optAQWAUnits.GetIndex(eventSender)

            If Index = 0 Then
                AQWAUnitsUS = True
            Else
                AQWAUnitsUS = False
            End If

        End If
    End Sub

    Private Sub optDamageLineType_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDamageLineType.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Short = optDamageLineType.GetIndex(eventSender)
            Dim R As Short
            On Error Resume Next

            If Index = 0 Then ' damage Most Critical lines
                With grdMatrix
                    For R = 0 To .RowCount - 1
                        .Rows(R).Cells(10).Value = CStr(L1Tmax(R + 1))
                    Next
                End With
            Else ' damage 2nd loaded lines
                With grdMatrix
                    For R = 0 To .RowCount - 1
                        .Rows(R).Cells(10).Value = CStr(L2Tmax(R + 1))
                    Next

                End With
            End If
        End If
    End Sub

    Private Sub optMoorState_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMoorState.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Short = optMoorState.GetIndex(eventSender)
            RefreshBackColor()
        End If
    End Sub

    Private Sub RefreshBackColor()
		Dim R As Short
		
		If optMoorState(0).Checked Then ' Intact

            fraDamageLineType.Enabled = False
			optDamageLineType(0).Enabled = False
			optDamageLineType(1).Enabled = False
			optDamageLineType(0).Checked = False
			optDamageLineType(1).Checked = False
		Else ' Damaged

            fraDamageLineType.Enabled = True
			optDamageLineType(0).Enabled = True
			optDamageLineType(1).Enabled = True
		End If
		
	End Sub

    Private Sub txtEdit_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs)
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = Asc(vbCr) Then KeyAscii = 0
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub Cleanup()
		Dim i As Short
		For i = 1 To UBound(oMet)
            oMet(i) = Nothing
        Next i
	End Sub
	
	Private Sub InitiateObjects()
		Dim i, UB_old As Short
		
		UB_old = UBound(oMet)
		
		If NumCases > UB_old Then
            ReDim Preserve oMet(NumCases)
            For i = UB_old + 1 To NumCases
				oMet(i) = New Metocean
			Next i
		Else
			For i = UBound(oMet) To NumCases + 1 Step -1
                oMet(i) = Nothing
            Next i
            ReDim Preserve oMet(NumCases)
        End If
	End Sub
	
	Private Function UpdateMetOceanObjects() As Boolean
        'grdMatrix_LeaveCell(grdMatrix, New System.EventArgs())
        UpdateMetOceanObjects = False
		Dim R As Short
		With grdMatrix
            ' all non-empty fields  - now update objects
            ' ASSUME .FixedCols=1
            For R = 0 To .RowCount - 1
                oMet(R + 1).Wind.Velocity = .Rows(R).Cells(0).Value * Knots2Ftps
                oMet(R + 1).Wind.Heading = .Rows(R).Cells(1).Value
                oMet(R + 1).Wave.SpectrumName = .Rows(R).Cells(2).Value
                oMet(R + 1).Wave.Height = .Rows(R).Cells(3).Value
                oMet(R + 1).Wave.Period = .Rows(R).Cells(4).Value
                oMet(R + 1).Wave.gamma = .Rows(R).Cells(5).Value
                If Collinear Then
                    oMet(R + 1).Wave.Heading = .Rows(R).Cells(1).Value
                    oMet(R + 1).Current.Heading = .Rows(R).Cells(1).Value
                Else
                    oMet(R + 1).Wave.Heading = .Rows(R).Cells(6).Value
                    oMet(R + 1).Current.Heading = .Rows(R).Cells(7).Value
                End If
                oMet(R + 1).Current.SurfaceVel = .Rows(R).Cells(8).Value * Knots2Ftps

                If UDEF Then
                    oMet(R + 1).Wave.SwellHeight = .Rows(R).Cells(11).Value
                    oMet(R + 1).Wave.SwellPeriod = .Rows(R).Cells(12).Value
                    oMet(R + 1).Wave.SwellSpectrumName = .Rows(R).Cells(13).Value
                    oMet(R + 1).Wave.Swellgamma = .Rows(R).Cells(14).Value
                    oMet(R + 1).Wave.SwellHeading = .Rows(R).Cells(15).Value

                End If
            Next R
        End With
		
		UpdateMetOceanObjects = True
	End Function
	
	Private Sub UpdateBrokenLines()
        'grdMatrix_LeaveCell(grdMatrix, New System.EventArgs())

        ReDim BreakLine(NumCases)
        Dim R As Short
		With grdMatrix
            ' all non-empty fields  - now update objects
            For R = 0 To .RowCount - 1
                BreakLine(R + 1) = CShort(.Rows(R).Cells(10).Value)
            Next R
        End With
		
	End Sub
	
	Private Function WriteDataToFile(ByVal fname As String) As Boolean
		
		WriteDataToFile = False
		Dim R As Short
		
		Dim fnum As Integer
		
		On Error GoTo ErrHandler
		fnum = FreeFile
		FileOpen(fnum, fname, OpenMode.Output)
		If UDEF Then
			PrintLine(fnum, NumCases & " UDEF")
			If optWind(1).Checked = True Then
				PrintLine(fnum, "Case", TAB, "1MinWindV", TAB, "WindHdg", TAB, "WaveName", TAB, "WaveHs", TAB, "WaveTp", TAB, "WaveGamma", TAB, "WaveHdg", TAB, "CurHdg", TAB, "SurVel" & "   " & "CurPts", TAB, "SwellHs", TAB, "SwellTp", TAB, "SwellWave", TAB, "SwellGamma", TAB, "SwellHeading")
			Else
				PrintLine(fnum, "Case", TAB, "1hrWindV", TAB, "WindHdg", TAB, "WaveName", TAB, "WaveHs", TAB, "WaveTp", TAB, "WaveGamma", TAB, "WaveHdg", TAB, "CurHdg", TAB, "SurVel" & "   " & "CurPts", TAB, "SwellHs", TAB, "SwellTp", TAB, "SwellWave", TAB, "SwellGamma", TAB, "SwellHeading")
			End If
		Else
			PrintLine(fnum, NumCases)
			If optWind(1).Checked = True Then
				PrintLine(fnum, "Case", TAB, "1MinWindV", TAB, "WindHdg", TAB, "WaveName", TAB, "WaveHs", TAB, "WaveTp", TAB, "WaveGamma", TAB, "WaveHdg", TAB, "CurHdg", TAB, "SurVel" & "   " & "CurPts")
			Else
				PrintLine(fnum, "Case", TAB, "1hrWindV", TAB, "WindHdg", TAB, "WaveName", TAB, "WaveHs", TAB, "WaveTp", TAB, "WaveGamma", TAB, "WaveHdg", TAB, "CurHdg", TAB, "SurVel" & "   " & "CurPts")
			End If
		End If
		
		For R = 1 To NumCases
			Print(fnum, R, TAB)
			Call oMet(R).WriteData(fnum, UDEF)
		Next R
		FileClose(fnum)
		WriteDataToFile = True
		Exit Function
ErrHandler: 
		FileClose(fnum)
		Debug.Print(Err.Description)
		
	End Function
	
	Private Function ReadDataFromFile(ByVal fname As String) As Boolean
		ReadDataFromFile = False
		Dim aline As String
		Dim R As Short
		Dim Fields() As String
		On Error GoTo ReadError
        '   open the file

        FileOpen(InFile, dlgFileOpen.FileName, OpenMode.Input, OpenAccess.Read)

        aline = LineInput(InFile) ' read numCases

        Fields = Split_Renamed(aline, " ")

        chkUDEF.CheckState = System.Windows.Forms.CheckState.Unchecked
		If UBound(Fields) >= 1 Then
			If Fields(1) = "UDEF" Then
				chkUDEF.CheckState = System.Windows.Forms.CheckState.Checked
			End If
		End If
		
		If CShort(Fields(0)) > 0 Then
			NumCases = CShort(Fields(0))
			txtNumCases.Text = CStr(NumCases)
			InitiateObjects()
			InitiateGrid()
		End If
        aline = LineInput(InFile) ' read if 1-min wind or 1-hr wind from header line
        If InStr(1, aline, "1hr") > 0 Then
			optWind(0).Checked = True
		Else
			optWind(1).Checked = True
		End If
		For R = 1 To NumCases
			Call oMet(R).ReadData(InFile, UDEF)
		Next R
		
		'   close the input file and return
		FileClose(InFile)
		LoadGrid()
		ReadDataFromFile = True
		Exit Function
ReadError: 
		MsgBox("Error Reading Env File. " & Err.Description, MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Error")
		ReadDataFromFile = False
		FileClose(InFile)
	End Function

    Friend Sub LoadGrid()

        Dim R As Short

        With grdMatrix
            .RowCount = NumCases

            For R = 0 To NumCases - 1
                ' load data

                .Rows(R).HeaderCell.Value = (R + 1).ToString
                .Rows(R).Cells(0).Value = Format(oMet(R + 1).Wind.Velocity * Ftps2Knots, "0.00")
                .Rows(R).Cells(1).Value = oMet(R + 1).Wind.Heading
                .Rows(R).Cells(2).Value = oMet(R + 1).Wave.SpectrumName
                .Rows(R).Cells(3).Value = oMet(R + 1).Wave.Height

                .Rows(R).Cells(4).Value = oMet(R + 1).Wave.Period

                .Rows(R).Cells(5).Value = oMet(R + 1).Wave.gamma

                .Rows(R).Cells(6).Value = oMet(R + 1).Wave.Heading

                .Rows(R).Cells(7).Value = oMet(R + 1).Current.Heading

                .Rows(R).Cells(8).Value = Format(oMet(R + 1).Current.SurfaceVel * Ftps2Knots, "0.00")

                .Rows(R).Cells(9).Value = oMet(R + 1).Current.ProfileCount
                If UDEF Then
                    .Rows(R).Cells(11).Value = oMet(R + 1).Wave.SwellHeight
                    .Rows(R).Cells(12).Value = oMet(R + 1).Wave.SwellPeriod
                    .Rows(R).Cells(13).Value = oMet(R + 1).Wave.SwellSpectrumName
                    .Rows(R).Cells(14).Value = oMet(R + 1).Wave.Swellgamma
                    .Rows(R).Cells(15).Value = oMet(R + 1).Wave.SwellHeading

                End If
            Next R

        End With

    End Sub

    ' Public Sub mnuAddRow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAddrow.Click

    '  AddRow(grdMatrix, CheckingGrid)
    '  End Sub



    ' Public Sub mnuDeleteRow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDeleteRow.Click

    '   DeleteRow(grdMatrix, CheckingGrid)

    'End Sub
    Private Sub txtNumCases_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNumCases.TextChanged
        If Val(txtNumCases.Text) > 0 Then
            ReDim Preserve L1Tmax(Val(txtNumCases.Text))
            ReDim Preserve L2Tmax(Val(txtNumCases.Text))
        End If
    End Sub

    Private Sub txtNumCases_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtNumCases.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		If KeyCode = System.Windows.Forms.Keys.Return Then
			NumCases = CShort(txtNumCases.Text)
			If NumCases > 0 Then
				InitiateObjects()
				InitiateGrid()
				LoadGrid()
				RefreshBackColor()
				ReDim Preserve L1Tmax(NumCases)
				ReDim Preserve L2Tmax(NumCases)
			End If
		End If
	End Sub
	
	Private Sub txtNumCases_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNumCases.Leave
		NumCases = CShort(txtNumCases.Text)
		If NumCases > 0 Then
			InitiateObjects()
			InitiateGrid()
			LoadGrid()
			RefreshBackColor()
			ReDim Preserve L1Tmax(NumCases)
			ReDim Preserve L2Tmax(NumCases)
		End If
	End Sub
End Class