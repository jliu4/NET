Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmMain
	Inherits System.Windows.Forms.Form
	Private ProgHandle As Integer
	
	Private Const InFile As Short = 1
	Private Const OutFile As Short = 2
	Private Const AQWAVersion As Boolean = True
	Private Const AQWA12Path As String = "c:\progra~1\ansysi~1\v150\aqwa\bin\winx64\"
	
	Private BaseDir, rBaseDir As String
	
	Private ABFile, AFFile As String
	
	Private HeaderLabels(17) As String
	Private HeaderLabels1(17) As String
	Private HeaderLabels2(17) As String
	
	Private Changed As Boolean
	Private ExistingTxt As String
	Private CheckingGrid As Boolean
	Private JustEnterCell As Boolean
	
	Private AQWAUnitsUS As Boolean
	Private Collinear As Boolean
	Private UDEF As Boolean
	
	'UPGRADE_NOTE: TimeString was upgraded to TimeString_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function TimeEarlierThan(ByRef TimeString_Renamed As String, ByRef CompareTo As String) As Boolean
		'Takes two time strings as returns true if the
		'first is earlier than the second
		
		'Example Usage:
		'Checks if it is before 6:00 PM
		'If TimeEarlierThan(Format(Now, "Short Time"), "6:00 PM") Then
		'    MsgBox "Good Afternoon (or morning)"
		'Else
		'    MsgBox "Good Evening"
		'End If
		
		If Not IsTime(TimeString_Renamed) Or Not IsTime(CompareTo) Then Exit Function
		
		TimeEarlierThan = CDate(TimeString_Renamed) < CDate(CompareTo)
		
	End Function
	
	Private Function IsTime(ByRef sTime As String) As Boolean
		
		'http://www.freevbcode.com/ShowCode.Asp?ID=1321
		'by Phil Fresle
		
		If VB.Left(Trim(sTime), 1) Like "#" Then
			IsTime = IsDate(Today & " " & sTime)
		End If
	End Function
	
	
	Private Sub btnBrowseBaseDir_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnBrowseBaseDir.Click
		On Error GoTo Hell
		With frmBrowseDir
			.Text = "Set Base Directory"
			.InitPath = lblBaseDir.Text
			.ControlToUpdate = lblBaseDir
			VB6.ShowForm(frmBrowseDir, 1, Me)
		End With
		If Trim(lblTargetDir.Text) = Trim(lblBaseDir.Text) Then
			MsgBox("Target directory should be different from Base directory.", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Target Dierctory")
			lblTargetDir.Text = ""
			WorkDir = ""
		End If
		
		frmPostProc.ReadBaseResults()
		CheckRunButtonState()
		Exit Sub
Hell: 
		lblBaseDir.Text = ""
	End Sub
	
	Private Sub ChangeCellColor()
		Dim CurCol, CurRow As Integer
		Dim R As Integer
		
		With grdMatrix
			' remember active cell
			CurCol = .Col
			CurRow = .Row
			.Col = 6
			For R = .FixedRows To .Rows - 1
				If InStr(.get_TextMatrix(.Row, 3), "PSMZ") > 0 Then
					.CellBackColor = System.Drawing.SystemColors.Control ' disable column
				Else
					'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
					.CellBackColor = System.Drawing.ColorTranslator.FromOle(vbNormal)
				End If
			Next R
			' restore active cell
			.Col = CurCol
			.Row = CurRow
		End With
	End Sub
	
	Private Sub btnBrowseDir_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnBrowseDir.Click
		Dim nPos As Short
		Dim Path1, Path2 As String
		
		With frmBrowseDir
			.Text = "Set Target Directory"
			If Len(lblTargetDir.Text) = 0 And Len(lblBaseDir.Text) > 0 Then
				.InitPath = lblBaseDir.Text
			Else
				.InitPath = lblTargetDir.Text
			End If
			.ControlToUpdate = lblTargetDir
			VB6.ShowForm(frmBrowseDir, 1, Me)
			If Trim(lblTargetDir.Text) = Trim(lblBaseDir.Text) Then
				MsgBox("Target directory should be different from Base directory.", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Target Dierctory")
				lblTargetDir.Text = ""
				WorkDir = ""
			End If
		End With
		
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
	End Sub
	
	Private Sub btnCopyAll_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCopyAll.Click
		Dim i, k As Short
		' copy case 1 to all cases
		For k = 2 To NumCases
			'UPGRADE_WARNING: Couldn't resolve default property of object oMet().Wind.Velocity.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oMet(k).Wind.Velocity.Value = oMet(1).Wind.Velocity.Value
			'UPGRADE_WARNING: Couldn't resolve default property of object oMet().Wave.Height.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oMet(k).Wave.Height.Value = oMet(1).Wave.Height.Value
			'UPGRADE_WARNING: Couldn't resolve default property of object oMet().Wave.Period.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oMet(k).Wave.Period.Value = oMet(1).Wave.Period.Value
			'UPGRADE_WARNING: Couldn't resolve default property of object oMet().Wave.gamma.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oMet(k).Wave.gamma.Value = oMet(1).Wave.gamma.Value
			'UPGRADE_WARNING: Couldn't resolve default property of object oMet().Wave.SpectrumName.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oMet(k).Wave.SpectrumName.Value = oMet(1).Wave.SpectrumName.Value
			oMet(k).Current.SurfaceVel = oMet(1).Current.SurfaceVel
			oMet(k).Current.Profile.Clear()
			For i = 1 To oMet(1).Current.Profile.Count
				oMet(k).Current.Profile.Add(oMet(1).Current.Profile.Item(i).Depth.Value, oMet(1).Current.Profile.Item(i).Velocity.Value)
			Next i
			If UDEF Then
				'UPGRADE_WARNING: Couldn't resolve default property of object oMet().Wave.SwellHeight.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				oMet(k).Wave.SwellHeight.Value = oMet(1).Wave.SwellHeight.Value
				'UPGRADE_WARNING: Couldn't resolve default property of object oMet().Wave.SwellPeriod.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				oMet(k).Wave.SwellPeriod.Value = oMet(1).Wave.SwellPeriod.Value
				'UPGRADE_WARNING: Couldn't resolve default property of object oMet().Wave.SwellSpectrumName.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				oMet(k).Wave.SwellSpectrumName.Value = oMet(1).Wave.SwellSpectrumName.Value
				
				'UPGRADE_WARNING: Couldn't resolve default property of object oMet().Wave.SwellGamma.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				oMet(k).Wave.SwellGamma.Value = oMet(1).Wave.SwellGamma.Value
				' oMet(k).Wave.SwellHeading.Value = oMet(1).Wave.SwellHeading.Value
				
			End If
		Next k
		LoadGrid()
	End Sub
	
	Private Sub btnCreateFiles_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCreateFiles.Click
		Dim s As String
		If HasEmptyFields(grdMatrix) Then Exit Sub
		UpdateMetOceanObjects()
		UpdateBrokenLines()
		
		On Error GoTo ErrHandler
		
		BaseDir = lblBaseDir.Text
		
		If BaseDir = "" Then Exit Sub
		
		ABFile = BaseDir & "\ABRUN.DAT"
		AFFile = BaseDir & "\AFRUN.DAT"
		
		Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
		If CreateAqwaInputFiles Then
			If CreateAqwaBatchFile((chkVersion.CheckState)) Then
				If CreateSilentRunBatchFile Then
					MsgBox("AQWA input Files have been successfully created.", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Create Files")
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
		
		'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		dlgFileOpen.Filter = "All Files (*.*)|*.*|Intact (*.xls)|*.xls"
		dlgFileSave.Filter = "All Files (*.*)|*.*|Intact (*.xls)|*.xls"
		dlgFileOpen.FilterIndex = 2
		dlgFileSave.FilterIndex = 2
		dlgFileOpen.InitialDirectory = WorkDir
		dlgFileSave.InitialDirectory = WorkDir
		'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgFile.Flags was upgraded to dlgFileOpen.CheckFileExists which has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		dlgFileOpen.CheckFileExists = True
		dlgFileOpen.CheckPathExists = True
		dlgFileSave.CheckPathExists = True
		'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgFile.Flags was upgraded to dlgFileOpen.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		dlgFileOpen.ShowReadOnly = False
		dlgFileOpen.ShowDialog()
		dlgFileSave.FileName = dlgFileOpen.FileName
		
		fname = dlgFileOpen.FileName
		
		If fname <> "" Then
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
			oxApp = ExcelGlobal_definst.Workbooks.Open(FileName:=fname).Application
			Call ReadCriticalLines(oxApp)
			oxApp.Quit()
			'UPGRADE_NOTE: Object oxApp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oxApp = Nothing
		End If
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		Exit Sub
		
ErrHandler: 
		'   User pressed Cancel button, or some tragedy occurred
		FileClose(InFile)
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		If Len(Err.Description) > 0 Then
			MsgBox(Err.Description)
		End If
		
	End Sub
	
	Private Function ReadCriticalLines(ByRef oxApp As Microsoft.Office.Interop.Excel.Application) As Object
		Dim NumLines, i, NCase As Short
		
		With oxApp.ActiveWorkbook.Sheets("Summary")
			' determine NumLines from xls
			'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			NumLines = .Range("D30").End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Row - 29
			
			NCase = oxApp.Sheets.Count - 2
			If NCase > 0 Then
				ReDim Preserve L1Tmax(NCase)
				ReDim Preserve L2Tmax(NCase)
			End If
			
			For i = 3 To oxApp.Sheets.Count - 1
				
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.Sheets().Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				L1Tmax(i - 2) = .Cells(29 + NumLines + 2, i)
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.Sheets().Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				L2Tmax(i - 2) = .Cells(29 + NumLines + 3, i)
			Next i
		End With
	End Function
	
	Private Sub btnNewAQWA_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnNewAQWA.Click
		RestoreAQWADirectory()
		MsgBox("Your system is ready to run NEW AQWA.", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "AQWA Directory Change")
	End Sub
	
	Private Sub btnOldAQWA_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnOldAQWA.Click
		RenameAQWADirectory()
		MsgBox("Your system is ready to run OLD AQWA.", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "AQWA Directory Change")
	End Sub
	
	Private Sub btnOpenMatrix_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnOpenMatrix.Click
		mnuOpen_Click(mnuOpen, New System.EventArgs())
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
				
				VB6.ShowForm(frmPostProc, 1, Me)
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
	
	Private Sub RenameAQWADirectory()
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
		
		' check current AQWA version
		'UPGRADE_WARNING: Couldn't resolve default property of object fs.FolderExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If fs.FolderExists("c:\asash-new") Then '  already running old version - do nothing
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object fs.FolderExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If fs.FolderExists("c:\asash") Then
				'UPGRADE_WARNING: Couldn't resolve default property of object fs.MoveFolder. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				fs.MoveFolder("c:\asash", "c:\asash-new")
				'UPGRADE_WARNING: Couldn't resolve default property of object fs.FolderExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If fs.FolderExists("c:\asash-old") Then
					'UPGRADE_WARNING: Couldn't resolve default property of object fs.MoveFolder. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					fs.MoveFolder("c:\asash-old", "c:\asash")
				End If
			Else
				MsgBox("Failed to rename AQWA Directory in c:\AQWA.")
			End If
		End If
		
	End Sub
	
	Private Sub RestoreAQWADirectory()
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
		
		' check current AQWA version
		'UPGRADE_WARNING: Couldn't resolve default property of object fs.FolderExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If fs.FolderExists("c:\asash-old") Then '  already running new version - do nothing
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object fs.FolderExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If fs.FolderExists("c:\asash") Then
				'UPGRADE_WARNING: Couldn't resolve default property of object fs.MoveFolder. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				fs.MoveFolder("c:\asash", "c:\asash-old")
				'UPGRADE_WARNING: Couldn't resolve default property of object fs.FolderExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If fs.FolderExists("c:\asash-new") Then
					'UPGRADE_WARNING: Couldn't resolve default property of object fs.MoveFolder. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					fs.MoveFolder("c:\asash-new", "c:\asash")
				End If
			End If
		End If
	End Sub
	
	Private Function CreateAqwaInputFiles() As Boolean
		Dim TmpStr, msg As String
		Dim tmpVal As Double
		
		If WaterDepth = 0 Then frmPostProc.ReadBaseResults()
		'UPGRADE_WARNING: Couldn't resolve default property of object oMet().Current.Profile().Depth.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		tmpVal = oMet(1).Current.Profile.Item(oMet(1).Current.Profile.Count).Depth.Value
		
		If oMet(1).Current.Profile.Count > 1 Then
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
		'UPGRADE_WARNING: Couldn't resolve default property of object fs.GetFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		f = fs.GetFile(ABFile)
		'UPGRADE_WARNING: Couldn't resolve default property of object f.OpenAsTextStream. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ts = f.OpenAsTextStream(ForReading, TristateFalse)
		'UPGRADE_WARNING: Couldn't resolve default property of object ts.Readall. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		s = ts.Readall
		'UPGRADE_WARNING: Couldn't resolve default property of object ts.Close. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
		'UPGRADE_WARNING: Couldn't resolve default property of object fs.GetFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		f = fs.GetFile(AFFile)
		'UPGRADE_WARNING: Couldn't resolve default property of object f.OpenAsTextStream. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ts = f.OpenAsTextStream(ForReading, TristateFalse)
		'UPGRADE_WARNING: Couldn't resolve default property of object ts.Readall. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ss = ts.Readall
		'UPGRADE_WARNING: Couldn't resolve default property of object ts.Close. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
				'UPGRADE_WARNING: Couldn't resolve default property of object fs.GetFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				f = fs.GetFile(My.Application.Info.DirectoryPath & "\tmp" & i & ".txt")
				'UPGRADE_WARNING: Couldn't resolve default property of object f.OpenAsTextStream. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ts = f.OpenAsTextStream(ForReading, TristateFalse)
				'UPGRADE_WARNING: Couldn't resolve default property of object ts.Readall. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				s = ts.Readall
				'UPGRADE_WARNING: Couldn't resolve default property of object ts.Close. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ts.Close()
				
				If optMoorState(0).Checked Then ' Intact
					'UPGRADE_WARNING: Couldn't resolve default property of object fs.FolderExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If Not fs.FolderExists(WorkDir & "\case" & i) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object fs.CreateFolder. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Call fs.CreateFolder(WorkDir & "\case" & i)
					End If
					
					'UPGRADE_WARNING: Couldn't resolve default property of object fs.CreateTextFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					MyFile = fs.CreateTextFile(WorkDir & "\case" & i & "\abrun.dat", True)
					'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.Write. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					MyFile.Write(s1 & s & s4 & s3)
					'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.Close. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					MyFile.Close()
					If chkPFLH.CheckState = System.Windows.Forms.CheckState.Unchecked Then
						'UPGRADE_WARNING: Couldn't resolve default property of object fs.CreateTextFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						MyFile = fs.CreateTextFile(WorkDir & "\case" & i & "\afrun1.dat", True)
						Mid(ss1, 17, 4) = "DRFT"
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object fs.CreateTextFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						MyFile = fs.CreateTextFile(WorkDir & "\case" & i & "\afrun.dat", True)
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.Write. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					MyFile.Write(ss1 & s & ss4 & ss2)
					'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.Close. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					MyFile.Close()
					
					If chkPFLH.CheckState = System.Windows.Forms.CheckState.Unchecked Then
						'UPGRADE_WARNING: Couldn't resolve default property of object fs.CreateTextFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						MyFile = fs.CreateTextFile(WorkDir & "\case" & i & "\afrun2.dat", True)
						Mid(ss1, 17, 4) = "WFRQ"
						'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.Write. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						MyFile.Write(ss1 & s & ss4 & ss2)
						'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.Close. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						MyFile.Close()
					End If
				Else
					If optDamageLineType(0).Checked = True Then
						TmpStr = "D1"
					Else
						TmpStr = "D2"
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object fs.FolderExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If Not fs.FolderExists(WorkDir & "\case" & i & TmpStr) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object fs.CreateFolder. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Call fs.CreateFolder(WorkDir & "\case" & i & TmpStr)
					End If
					
					'UPGRADE_WARNING: Couldn't resolve default property of object fs.CreateTextFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					MyFile = fs.CreateTextFile(WorkDir & "\case" & i & TmpStr & "\abrun.dat", True)
					'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.Write. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					MyFile.Write(s1 & s & s4 & s3)
					'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.Close. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					MyFile.Close()
					
					If chkPFLH.CheckState = System.Windows.Forms.CheckState.Unchecked Then
						'UPGRADE_WARNING: Couldn't resolve default property of object fs.CreateTextFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						MyFile = fs.CreateTextFile(WorkDir & "\case" & i & TmpStr & "\afrun1.dat", True)
						Mid(ss1, 17, 4) = "DRFT"
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object fs.CreateTextFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						MyFile = fs.CreateTextFile(WorkDir & "\case" & i & TmpStr & "\afrun.dat", True)
					End If
					
					'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.Write. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					MyFile.Write(ss1 & s & ss4 & ss2)
					'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.Close. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					MyFile.Close()
					
					If chkPFLH.CheckState = System.Windows.Forms.CheckState.Unchecked Then
						'UPGRADE_WARNING: Couldn't resolve default property of object fs.CreateTextFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						MyFile = fs.CreateTextFile(WorkDir & "\case" & i & TmpStr & "\afrun2.dat", True)
						Mid(ss1, 17, 4) = "WFRQ"
						'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.Write. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						MyFile.Write(ss1 & s & ss4 & ss2)
						'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.Close. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Split_Renamed(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Fields = Split_Renamed(s, MatchStr)
		
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
		'UPGRADE_WARNING: Couldn't resolve default property of object FirstLineStart. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
				'UPGRADE_WARNING: Couldn't resolve default property of object FirstLineStart. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pos1 = InStr(FirstLineStart, s, vbCrLf, CompareMethod.Binary)
				'UPGRADE_WARNING: Couldn't resolve default property of object FirstLineStart. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
	
	Private Function CreateAqwaBatchFile(ByRef V12 As Boolean) As Boolean
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
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fs.CreateTextFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		MyFile = fs.CreateTextFile(WorkDir & "\RunAqwaCases.bat", True)
		
		For i = 1 To NumCases
			
			If IsDamaged = 1 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				subDir = "case" & i & "D1"
			ElseIf IsDamaged = 2 Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				subDir = "case" & i & "D2"
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				subDir = "case" & i
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			MyFile.WriteLine("COPY " & rBaseDir & "\ABRUN.RES ABRUN.RES")
			If AQWAVersion Then
				'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MyFile.WriteLine("COPY " & rBaseDir & "\ABRUN.EQP ABRUN.EQP")
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MyFile.WriteLine("COPY " & rBaseDir & "\ABRUN.POS ABRUN.POS")
				'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MyFile.WriteLine("COPY " & rBaseDir & "\ABRUN.HYD ABRUN.HYD")
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			MyFile.WriteLine()
			' copy data file to workdir to run
			'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			MyFile.WriteLine("COPY " & rWorkDir & "\" & subDir & "\ABRUN.DAT ABRUN.DAT")
			If AQWAVersion Then ' new AQWA
				If Len(txtAQWAdir.Text) > 0 Then
					If V12 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						MyFile.WriteLine(AQWA12Path & Trim(txtAQWAdir.Text) & " LIBRIUM RUN")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						MyFile.WriteLine(Trim(txtAQWAdir.Text) & " LIBRIUM RUN")
					End If
				Else
					If V12 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						MyFile.WriteLine(AQWA12Path & "AQWA LIBRIUM RUN")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						MyFile.WriteLine("AQWA LIBRIUM RUN")
					End If
				End If
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MyFile.WriteLine("LIBRIUM RUN")
			End If
			' COPY SELECTED RESULTS BACK TO CASE DIRECTORY
			'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			MyFile.WriteLine("COPY ABRUN.LIS " & rWorkDir & "\" & subDir & "\ABRUN.LIS" & vbCrLf)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			MyFile.WriteLine("COPY ABRUN.RES AFRUN.RES")
			If AQWAVersion Then
				'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MyFile.WriteLine("COPY ABRUN.EQP AFRUN.EQP")
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MyFile.WriteLine("COPY ABRUN.POS AFRUN.POS")
				'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MyFile.WriteLine("COPY ABRUN.HYD AFRUN.HYD")
			End If
			' copy data file to workdir to run
			
			If chkPFLH.CheckState = System.Windows.Forms.CheckState.Unchecked Then
				'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MyFile.WriteLine("COPY " & rWorkDir & "\" & subDir & "\AFRUN1.DAT AFRUN.DAT")
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MyFile.WriteLine("COPY " & rWorkDir & "\" & subDir & "\AFRUN.DAT AFRUN.DAT")
			End If
			
			If AQWAVersion Then
				If Len(txtAQWAdir.Text) > 0 Then
					If V12 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						MyFile.WriteLine(AQWA12Path & Trim(txtAQWAdir.Text) & " FER RUN")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						MyFile.WriteLine(Trim(txtAQWAdir.Text) & " FER RUN")
					End If
				Else
					If V12 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						MyFile.WriteLine(AQWA12Path & "AQWA FER RUN")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						MyFile.WriteLine("AQWA FER RUN")
					End If
				End If
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MyFile.WriteLine("FER RUN")
			End If
			' COPY SELECTED RESULTS BACK TO CASE DIRECTORY
			If chkPFLH.CheckState = System.Windows.Forms.CheckState.Checked Then
				'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MyFile.WriteLine("COPY AFRUN.LIS " & rWorkDir & "\" & subDir & "\AFRUN.LIS" & vbCrLf)
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MyFile.WriteLine("COPY AFRUN.LIS " & rWorkDir & "\" & subDir & "\AFRUN1.LIS" & vbCrLf)
				
				' continue for AFRUN2
				'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MyFile.WriteLine("COPY ABRUN.RES AFRUN.RES")
				If AQWAVersion Then
					'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					MyFile.WriteLine("COPY ABRUN.EQP AFRUN.EQP")
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					MyFile.WriteLine("COPY ABRUN.POS AFRUN.POS")
					'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					MyFile.WriteLine("COPY ABRUN.HYD AFRUN.HYD")
				End If
				' copy data file to workdir to run
				'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MyFile.WriteLine("COPY " & rWorkDir & "\" & subDir & "\AFRUN2.DAT AFRUN.DAT")
				If AQWAVersion Then
					If Len(txtAQWAdir.Text) > 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						MyFile.WriteLine(Trim(txtAQWAdir.Text) & " FER RUN")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						MyFile.WriteLine("AQWA FER RUN")
					End If
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					MyFile.WriteLine("FER RUN")
				End If
				' COPY SELECTED RESULTS BACK TO CASE DIRECTORY
				'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MyFile.WriteLine("COPY AFRUN.LIS " & rWorkDir & "\" & subDir & "\AFRUN2.LIS" & vbCrLf)
				'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MyFile.WriteLine()
			End If
			
		Next i
		'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.Close. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		MyFile.Close()
		CreateAqwaBatchFile = True
		Exit Function
ErrHandler: 
		ReportError(err, "Error", "Error Creating AQWA Batch File")
		'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.Open. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object MyFile.Close. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fs.FileExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object fs.DeleteFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If fs.FileExists(WorkDir & "\stdtests.com") Then fs.DeleteFile(WorkDir & "\stdtests.com")
		'UPGRADE_WARNING: Couldn't resolve default property of object fs.FileExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object fs.DeleteFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If fs.FileExists(WorkDir & "\stdtests.com") Then fs.DeleteFile(WorkDir & "\stdtests.com")
		'UPGRADE_WARNING: Couldn't resolve default property of object fs.FileExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object fs.DeleteFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If fs.FileExists(WorkDir & "\RunSilent.bat") Then fs.DeleteFile(WorkDir & "\RunSilent.bat")
		
		On Error GoTo ErrHandler
		
		' read bat file
		'UPGRADE_WARNING: Couldn't resolve default property of object fs.GetFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		f = fs.GetFile(WorkDir & "\RunAqwaCases.bat")
		'UPGRADE_WARNING: Couldn't resolve default property of object f.OpenAsTextStream. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ts = f.OpenAsTextStream(ForReading, TristateFalse)
		'UPGRADE_WARNING: Couldn't resolve default property of object ts.Readall. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object s. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		s = ts.Readall
		' replace content AQWA with RUN
		If Len(txtAQWAdir.Text) > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object s. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ss. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ss = Replace(s, Trim(txtAQWAdir.Text), "RUN")
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object s. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ss. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ss = Replace(s, "AQWA", "RUN")
		End If
		
		' won't support long file name by having to remove quotes
		'ss = Replace(ss, Chr(34), "")
		
		' must use full path name for copy command
		
		' ss = Replace(ss, " A", " " & WorkDir & "\A")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object ts.Close. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ts.Close()
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fs.CreateTextFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		f = fs.CreateTextFile(WorkDir & "\stdtestscom.bak", True)
		'UPGRADE_WARNING: Couldn't resolve default property of object f.Write. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		f.Write(ss)
		'UPGRADE_WARNING: Couldn't resolve default property of object f.Close. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		f.Close()
		
		' rename batch file
		'UPGRADE_WARNING: Couldn't resolve default property of object fs.CopyFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fs.CopyFile(WorkDir & "\stdtestscom.bak", WorkDir & "\stdtests.com")
		
		' create new silent batch
		'UPGRADE_WARNING: Couldn't resolve default property of object fs.CreateTextFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		f = fs.CreateTextFile(WorkDir & "\RunSilent.bat", True)
		If Len(txtAQWAdir.Text) > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object f.Write. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			f.Write(Trim(txtAQWAdir.Text) & " std")
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object f.Write. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			f.Write("aqwa std")
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object f.Close. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		f.Close()
		
		CreateSilentRunBatchFile = True
		Exit Function
ErrHandler: 
		ReportError(err, "Error", "Error Creating AQWA Silent Batch File")
		'UPGRADE_WARNING: Couldn't resolve default property of object f.Open. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object f.Close. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
		
		If AQWAVersion Then
			PrintLine(OutFile, "    11    ENVR")
			If optWind(1).Checked Then
				PrintLine(OutFile, "    11WIND", TAB(13), CStr(VB6.Format(oMet(Index).Wind.Velocity.Value * Flen, "#0.00")), TAB(22), CStr(VB6.Format(oMet(Index).Wind.Heading.Value, "#0.00")))
			End If
			' CONVERT FROM KNOTS TO FT/S FOR AQWA
			PrintLine(OutFile, "    11CURR", TAB(13), CStr(VB6.Format(oMet(Index).Current.SurfaceVel * Flen, "#0.00")), TAB(22), CStr(VB6.Format(oMet(Index).Current.Heading.Value, "#0.00")))
			
			NumPts = oMet(Index).Current.Profile.Count
			
			If NumPts > 1 Then
				For i = NumPts To 1 Step -1
					PrintLine(OutFile, "    11CPRF", TAB(13), VB6.Format(oMet(Index).Current.Profile.Item(i).Depth.Value * (-1) * Flen, "#0"), TAB(22), VB6.Format((oMet(Index).Current.SurfaceVel - oMet(Index).Current.Profile.Item(i).Velocity.Value) * Flen, "#0.00"))
				Next i
			End If
			PrintLine(OutFile, " END11CDRN", TAB(13), CStr(VB6.Format(Val(oMet(Index).Current.Heading.Value - 180#), "#0.00"))) ' assume deg used
			PrintLine(OutFile, "    12    NONE")
			PrintLine(OutFile, "    13    SPEC")
			If optWind(0).Checked Then ' 1-hr wind using spectrum
				If Len(UDWS) > 0 Then ' if user-defined spectrum
					Print(OutFile, UDWS)
				ElseIf Len(WSPEC) > 0 Then 
					PrintLine(OutFile, WSPEC)
				End If
				PrintLine(OutFile, "    13WIND", TAB(22), CStr(VB6.Format(oMet(Index).Wind.Velocity.Value * Flen, "#0.00")), TAB(32), CStr(oMet(Index).Wind.Heading.Value), TAB(42), VB6.Format(32.8 * Flen, "#0.0")) ' assume 10m ref height, i.e. 33 ft
			End If
			If UDEF Then
				Print(OutFile, "    13SPDN")
				PrintLine(OutFile, TAB(22), CStr(VB6.Format(Val(oMet(Index).Wave.Heading.Value), "#0.0")))
				s = Trim(oMet(Index).Wave.SwellSpectrumName.Value)
				If InStr(s, "GAUS") > 0 Then
					PrintLine(OutFile, "    13XSWL GAUS", TAB(42), CStr(VB6.Format(Val(oMet(Index).Wave.SwellGamma.Value), "#0.0000")), TAB(52), CStr(VB6.Format(Val(CStr(oMet(Index).Wave.SwellHeight.Value * Flen)), "#0.0")), TAB(62), CStr(VB6.Format(Val(CStr(2 * 3.1415926 / oMet(Index).Wave.SwellPeriod.Value)), "#0.000")), TAB(72), CStr(VB6.Format(Val(oMet(Index).Wave.SwellHeading.Value), "#0.0")))
				ElseIf InStr(s, "JONH") Then 
					PrintLine(OutFile, "    13XSWL JONH", TAB(42), CStr(VB6.Format(Val(oMet(Index).Wave.SwellGamma.Value), "#0.0000")), TAB(52), CStr(VB6.Format(Val(CStr(oMet(Index).Wave.SwellHeight.Value * Flen)), "#0.0")), TAB(62), CStr(VB6.Format(Val(CStr(2 * 3.1415926 / oMet(Index).Wave.SwellPeriod.Value)), "#0.000")), TAB(72), CStr(VB6.Format(Val(oMet(Index).Wave.SwellHeading.Value), "#0.0")))
				Else ' PSMZ
					PrintLine(OutFile, "    13XSWL PSMZ", TAB(42), CStr(VB6.Format(Val(oMet(Index).Wave.SwellGamma.Value), "#0.0000")), TAB(52), CStr(VB6.Format(Val(CStr(oMet(Index).Wave.SwellHeight.Value * Flen)), "#0.0")), TAB(62), CStr(VB6.Format(Val(CStr(2 * 3.1415926 / oMet(Index).Wave.SwellPeriod.Value)), "#0.000")), TAB(72), CStr(VB6.Format(Val(oMet(Index).Wave.SwellHeading.Value), "#0.0")))
				End If
			Else
				PrintLine(OutFile, "    13RADS")
				Print(OutFile, "    13SPDN")
				PrintLine(OutFile, TAB(22), CStr(VB6.Format(Val(oMet(Index).Wave.Heading.Value), "#0.0")))
			End If
		Else
			PrintLine(OutFile, "    11    NONE")
			PrintLine(OutFile, "    12    NONE")
			PrintLine(OutFile, "    13    SPEC")
			PrintLine(OutFile, "    13WIND", TAB(22), CStr(VB6.Format(oMet(Index).Wind.Velocity.Value * Flen, "#0.00")), TAB(32), CStr(VB6.Format(oMet(Index).Wind.Heading.Value, "#0.00")))
			' CONVERT FROM KNOTS TO FT/S FOR AQWA
			PrintLine(OutFile, "    13CURR", TAB(22), CStr(VB6.Format(oMet(Index).Current.SurfaceVel * Flen, "#0.000")), TAB(32), CStr(VB6.Format(oMet(Index).Current.Heading.Value, "#0.00")))
			PrintLine(OutFile, "    13SPDN", TAB(22), CStr(VB6.Format(oMet(Index).Wave.Heading.Value, "#0.0")))
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
			s = Trim(oMet(Index).Wave.SpectrumName.Value)
			If InStr(s, "JONH") > 0 Then
				PrintLine(OutFile, " END13" & s, TAB(22), "0.200", TAB(32), " 1.500", TAB(42), CStr(oMet(Index).Wave.gamma.Value), TAB(52), CStr(VB6.Format(oMet(Index).Wave.Height.Value * Flen, "#0.000")), TAB(62), CStr(VB6.Format(Val(CStr(2 * 3.1415926 / oMet(Index).Wave.Period.Value)), "#0.000")))
			Else ' PSMZ
				PrintLine(OutFile, " END13" & s, TAB(22), "0.200", TAB(32), " 1.500", TAB(42), CStr(VB6.Format(oMet(Index).Wave.Height.Value * Flen, "#0.000")), TAB(52), CStr(VB6.Format(oMet(Index).Wave.Period.Value / 1.4, "#0.000")))
			End If
		End If
		FileClose(OutFile)
		CreateEnvFile = True
		Exit Function
ErrHandler: 
		FileClose(OutFile)
		ReportError(err, "Error", "Create Env File Error " & msg)
	End Function
	
	Private Sub btnSaveMatrix_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnSaveMatrix.Click
		mnuSave_Click(mnuSave, New System.EventArgs())
	End Sub
	
	
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
	
	'UPGRADE_WARNING: Event chkCollinear.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkCollinear_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkCollinear.CheckStateChanged
		
		Select Case chkCollinear.CheckState
			Case 0
				Collinear = False
			Case 1
				Collinear = True
		End Select
		
	End Sub
	
	'UPGRADE_WARNING: Event chkUDEF.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkUDEF_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkUDEF.CheckStateChanged
		Dim c, R As Short
		
		With grdMatrix
			Select Case chkUDEF.CheckState
				Case 0
					.Cols = 12
					
					UDEF = False
				Case 1
					.Cols = 17
					
					'            For c = .Cols - 3 To .Cols - 1
					'                .Col = c
					'                .Row = 0
					'                .CellAlignment = flexAlignCenterCenter
					'                .Text = HeaderLabels(c)
					'                .Row = 1
					'                .CellAlignment = flexAlignCenterCenter
					'                .Text = HeaderLabels1(c)
					'                .Row = 2
					'                .CellAlignment = flexAlignCenterCenter
					'                .Text = HeaderLabels2(c)
					'                ' set cell alignment
					'                For R = .FixedRows + 1 To .Rows - 1
					'                    .Row = R
					'                    .CellAlignment = flexAlignRightCenter
					'                Next R
					'            Next c
					
					UDEF = True
			End Select
		End With
		
	End Sub
	
	'UPGRADE_WARNING: Form event frmMain.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmMain_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		'    RefreshBackColor
	End Sub
	
	' form load and unload
	
	Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		WaterDepth = 0
		
		CheckingGrid = True
		AQWAUnitsUS = True
		Collinear = True
		UDEF = False
		
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
			If Not oMet(NumCases).Wind.Heading.IsEmpty Then
				If lblBaseDir.Text <> "" And lblTargetDir.Text <> "" Then
					btnRun.Enabled = True
					btnSilentRun.Enabled = True
				End If
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
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fs.FileExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
			'UPGRADE_WARNING: Couldn't resolve default property of object fs.GetFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			f = fs.GetFile(WorkDir & subFolder & "\abrun.lis")
			'UPGRADE_WARNING: Couldn't resolve default property of object f.OpenAsTextStream. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ts = f.OpenAsTextStream(ForReading, TristateFalse)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object ts.Readall. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			s = ts.Readall
			'UPGRADE_WARNING: Couldn't resolve default property of object ts.Close. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
		'UPGRADE_WARNING: Lower bound of array oMet was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim oMet(NumCases)
		Dim i As Short
		For i = 1 To NumCases
			oMet(i) = New Metocean
		Next i
	End Sub
	' menus
	
	Private Sub grdMatrix_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles grdMatrix.ClickEvent
		If grdMatrix.CellBackColor.equals(System.Drawing.SystemColors.Control) Then
		ElseIf grdMatrix.Col = 10 Then 
			With frmCurrProfile
				.lblCaseNo.Text = grdMatrix.get_TextMatrix(grdMatrix.Row, 0)
				.grdProfile.set_TextMatrix(1, 0, "0")
				.txtNumProfilePts.Text = grdMatrix.get_TextMatrix(grdMatrix.Row, 10)
				SelectText(.txtNumProfilePts)
				.grdProfile.set_TextMatrix(1, 1, grdMatrix.get_TextMatrix(grdMatrix.Row, 9))
				VB6.ShowForm(frmCurrProfile, 1, Me)
			End With
		ElseIf grdMatrix.Col = 3 Then 
			MSFlexGridCombo(grdMatrix, cboWaveType, True)
		Else
			MSFlexGridEdit(grdMatrix, txtEdit, System.Windows.Forms.Keys.F10)
		End If
	End Sub
	
	Private Sub grdMatrix_EnterCell(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles grdMatrix.EnterCell
		ExistingTxt = grdMatrix.Text
		JustEnterCell = True
	End Sub
	
	Private Sub grdMatrix_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSFlexGridLib.DMSFlexGridEvents_KeyDownEvent) Handles grdMatrix.KeyDownEvent
		KeyHandler(grdMatrix, txtEdit, eventArgs.KeyCode, eventArgs.Shift, JustEnterCell, ExistingTxt)
	End Sub
	
	Private Sub grdMatrix_LeaveCell(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles grdMatrix.LeaveCell
		
		If txtEdit.Visible = True Then
			grdMatrix.Text = txtEdit.Text
			txtEdit.Visible = False
		End If
		
		If cboWaveType.Visible Then
			grdMatrix.Text = cboWaveType.Text
			cboWaveType.Visible = False
			ChangeCellColor()
		End If
		
		If Not CheckingGrid Then
			With grdMatrix
				If Trim(.Text) <> ExistingTxt And .Col > 1 Then
					If Trim(.Text) <> "" Then .Text = CheckData(.Text)
					Changed = True
				End If
			End With
		End If
	End Sub
	
	Private Sub grdMatrix_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles grdMatrix.Scroll
		grdMatrix_LeaveCell(grdMatrix, New System.EventArgs())
	End Sub
	
	'UPGRADE_ISSUE: Label event lblTargetDir.Change was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub lblTargetDir_Change()
		WorkDir = lblTargetDir.Text
	End Sub
	
	
	Public Sub mnuAbout_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAbout.Click
		VB6.ShowForm(frmAbout, 1, Me)
	End Sub
	
	Public Sub mnuCreateFiles_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCreateFiles.Click
		btnCreateFiles_Click(btnCreateFiles, New System.EventArgs())
	End Sub
	
	Public Sub mnuExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuExit.Click
		Me.Close()
	End Sub
	
	Public Sub mnuNewAQWA_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNewAQWA.Click
		btnNewAQWA_Click(btnNewAQWA, New System.EventArgs())
	End Sub
	
	Public Sub mnuOldAQWA_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuOldAQWA.Click
		btnOldAQWA_Click(btnOldAQWA, New System.EventArgs())
	End Sub
	
	Public Sub mnuOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuOpen.Click
		
		Dim msg, fname As String
		
		'   should the user cancel the dialog box, exit
		On Error GoTo ErrHandler
		
		'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		dlgFileOpen.Filter = "All Files (*.*)|*.*|Env Matrix (*.mat)|*.mat"
		dlgFileSave.Filter = "All Files (*.*)|*.*|Env Matrix (*.mat)|*.mat"
		dlgFileOpen.FilterIndex = 2
		dlgFileSave.FilterIndex = 2
		dlgFileOpen.InitialDirectory = CurDir()
		dlgFileSave.InitialDirectory = CurDir()
		'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgFile.Flags was upgraded to dlgFileOpen.CheckFileExists which has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		dlgFileOpen.CheckFileExists = True
		dlgFileOpen.CheckPathExists = True
		dlgFileSave.CheckPathExists = True
		'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgFile.Flags was upgraded to dlgFileOpen.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		dlgFileOpen.ShowReadOnly = False
		dlgFileOpen.ShowDialog()
		dlgFileSave.FileName = dlgFileOpen.FileName
		
		fname = dlgFileOpen.FileName
		
		If fname <> "" Then
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
			'     ReadOldFile (fname)
			If Not ReadDataFromFile(fname) Then GoTo ErrHandler
			'UPGRADE_WARNING: Couldn't resolve default property of object FileRootName(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			txtMetoceanName.Text = FileRootName(fname)
		End If
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		CheckRunButtonState()
		Exit Sub
		
ErrHandler: 
		'   User pressed Cancel button, or some tragedy occurred
		FileClose(InFile)
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		If Len(Err.Description) > 0 Then
			MsgBox(Err.Description)
		End If
	End Sub
	
	Public Sub mnuOpenOldMatrix_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuOpenOldMatrix.Click
		Dim msg, fname As String
		
		'   should the user cancel the dialog box, exit
		On Error GoTo ErrHandler
		
		'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		dlgFileOpen.Filter = "All Files (*.*)|*.*|Env Matrix (*.mat)|*.mat"
		dlgFileSave.Filter = "All Files (*.*)|*.*|Env Matrix (*.mat)|*.mat"
		dlgFileOpen.FilterIndex = 2
		dlgFileSave.FilterIndex = 2
		dlgFileOpen.InitialDirectory = My.Application.Info.DirectoryPath
		dlgFileSave.InitialDirectory = My.Application.Info.DirectoryPath
		'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgFile.Flags was upgraded to dlgFileOpen.CheckFileExists which has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		dlgFileOpen.CheckFileExists = True
		dlgFileOpen.CheckPathExists = True
		dlgFileSave.CheckPathExists = True
		'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgFile.Flags was upgraded to dlgFileOpen.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		dlgFileOpen.ShowReadOnly = False
		dlgFileOpen.ShowDialog()
		dlgFileSave.FileName = dlgFileOpen.FileName
		
		fname = dlgFileOpen.FileName
		
		If fname <> "" Then
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
			Call ReadOldFile(fname)
			'UPGRADE_WARNING: Couldn't resolve default property of object FileRootName(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			txtMetoceanName.Text = FileRootName(fname)
		End If
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		CheckRunButtonState()
		Exit Sub
		
ErrHandler: 
		'   User pressed Cancel button, or some tragedy occurred
		FileClose(InFile)
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		If Len(Err.Description) > 0 Then
			MsgBox(Err.Description)
		End If
		
	End Sub
	
	Public Sub mnuPostProcess_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPostProcess.Click
		btnPostProcess_Click(btnPostProcess, New System.EventArgs())
	End Sub
	
	Public Sub mnuPreProcess_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPreProcess.Click
		btnRun_Click(btnRun, New System.EventArgs())
	End Sub
	
	Public Sub mnuSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSave.Click
		
		' validate no empty fields
		Dim R, c As Short
		
		
		Call UpdateMetOceanObjects()
		Call UpdateBrokenLines()
		
		If txtMetoceanName.Text <> "" Then
			dlgFileOpen.FileName = txtMetoceanName.Text & ".mat"
			dlgFileSave.FileName = txtMetoceanName.Text & ".mat"
		Else
			dlgFileOpen.FileName = "*.mat"
			dlgFileSave.FileName = "*.mat"
		End If
		'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		dlgFileOpen.Filter = "All Files (*.*)|*.*|Env Matrix (*.mat)|*.mat"
		dlgFileSave.Filter = "All Files (*.*)|*.*|Env Matrix (*.mat)|*.mat"
		dlgFileOpen.FilterIndex = 2
		dlgFileSave.FilterIndex = 2
		dlgFileOpen.InitialDirectory = CurDir()
		dlgFileSave.InitialDirectory = CurDir()
		'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgFile.Flags was upgraded to dlgFileSave.OverwritePrompt which has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		dlgFileSave.OverwritePrompt = True
		'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgFile.Flags was upgraded to dlgFileOpen.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		dlgFileOpen.ShowReadOnly = False
		dlgFileSave.ShowDialog()
		dlgFileOpen.FileName = dlgFileSave.FileName
		
		'   open the file for output
		Dim fname As String
		fname = dlgFileOpen.FileName
		If fname = "" Then Exit Sub
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		
		WriteDataToFile(fname)
		CheckRunButtonState()
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
	End Sub
	
	
	
	' operation subroutines
	' initiating
	
	Private Sub InitiateCombo()
		
		Dim i As Short
		
		cboWaveType.Items.Clear()
		
		cboWaveType.Items.Add("PSMZ")
		cboWaveType.Items.Add("JONH")
		cboWaveType.Items.Add("GAUS")
		
		cboWaveType.SelectedIndex = 0
	End Sub
	
	Private Sub InitiateGrid()
		
		Dim R, c As Short
		Dim ColW As Single
		Dim RowH As Short
		
		Call SetLabels()
		
		With grdMatrix
			If UDEF Then
				.Cols = 17
			Else
				.Cols = 12
			End If
			
			.Rows = NumCases + grdMatrix.FixedRows
			.MergeCells = MSFlexGridLib.MergeCellsSettings.flexMergeRestrictColumns
			.set_ColWidth(0, 800)
			ColW = VB6.PixelsToTwipsX(.Width) / .Cols - .GridLineWidth - 20
			For c = 1 To .Cols - 1
				.set_ColWidth(c, ColW)
			Next 
			.Row = 0
			For c = 0 To .Cols - 1
				.Col = c
				.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
				.Text = HeaderLabels(c)
			Next 
			.Row = 1
			For c = 0 To .Cols - 1
				.Col = c
				.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
				.Text = HeaderLabels1(c)
			Next 
			.Row = 2
			For c = 0 To .Cols - 1
				.Col = c
				.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
				.Text = HeaderLabels2(c)
			Next 
			For R = .FixedRows To .Rows - 1
				.set_TextMatrix(R, 0, "Case " & R - .FixedRows + 1)
				.set_RowHeight(R, VB6.PixelsToTwipsY(txtEdit.Height))
				.set_TextMatrix(R, 10, "1")
			Next 
			' set cell alignment
			For R = .FixedRows To .Rows - 1
				.Row = R
				For c = .FixedCols To .Cols - 1
					.Col = c
					.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
				Next c
			Next 
			
			.set_MergeRow(0, True)
			
			.Col = .FixedCols
			.Row = .FixedRows
			
		End With
		
		fraDamageLineType.Enabled = False
		optDamageLineType(0).Checked = False
		optDamageLineType(1).Checked = False
		optDamageLineType(0).Enabled = False
		optDamageLineType(1).Enabled = False
	End Sub
	
	
	Private Sub SetLabels()
		
		HeaderLabels(0) = ""
		HeaderLabels(1) = "Wind"
		HeaderLabels(2) = "Wind"
		HeaderLabels(3) = "Wave"
		HeaderLabels(4) = "Wave"
		HeaderLabels(5) = "Wave"
		HeaderLabels(6) = "Wave"
		HeaderLabels(7) = "Wave"
		HeaderLabels(8) = "Current"
		HeaderLabels(9) = "Current"
		HeaderLabels(10) = "Current"
		HeaderLabels(11) = "Break"
		HeaderLabels(12) = "Swell"
		HeaderLabels(13) = "Swell"
		HeaderLabels(14) = "Swell"
		HeaderLabels(15) = "Swell"
		HeaderLabels(16) = "Swell"
		
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
		HeaderLabels2(1) = "(knots)"
		HeaderLabels2(2) = "(deg)"
		HeaderLabels2(3) = "Name"
		HeaderLabels2(4) = "(ft)"
		HeaderLabels2(5) = "(sec)"
		HeaderLabels2(6) = ""
		HeaderLabels2(7) = "(deg)"
		HeaderLabels2(8) = "(deg)"
		HeaderLabels2(9) = "(knots)"
		HeaderLabels2(10) = "Points"
		HeaderLabels2(11) = "Line"
		HeaderLabels2(12) = "(ft)"
		HeaderLabels2(13) = "(sec)"
		HeaderLabels2(14) = ""
		HeaderLabels2(15) = ""
		HeaderLabels2(16) = ""
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
	
	'UPGRADE_WARNING: Event optAQWAUnits.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
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
	
	'UPGRADE_WARNING: Event optDamageLineType.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optDamageLineType_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDamageLineType.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optDamageLineType.GetIndex(eventSender)
			Dim R As Short
			
			On Error Resume Next
			
			If Index = 0 Then ' damage Most Critical lines
				With grdMatrix
					.Col = 11
					For R = .FixedRows To .Rows - 1
						.Row = R
						If L1Tmax(R - .FixedRows + 1) > 0 Then
							.Text = CStr(L1Tmax(R - .FixedRows + 1))
						Else
							.Text = ""
						End If
					Next R
				End With
			Else ' damage 2nd loaded lines
				With grdMatrix
					.Col = 11
					For R = .FixedRows To .Rows - 1
						.Row = R
						If L2Tmax(R - .FixedRows + 1) > 0 Then
							.Text = CStr(L2Tmax(R - .FixedRows + 1))
						Else
							.Text = ""
						End If
					Next R
				End With
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optMoorState.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optMoorState_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMoorState.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optMoorState.GetIndex(eventSender)
			RefreshBackColor()
		End If
	End Sub
	
	Private Sub RefreshBackColor()
		Dim R As Short
		
		If optMoorState(0).Checked Then ' Intact
			With grdMatrix
				.Col = 11
				For R = .FixedRows To .Rows - 1
					.Row = R
					.CellBackColor = System.Drawing.SystemColors.Control
					.Text = ""
				Next R
			End With
			fraDamageLineType.Enabled = False
			optDamageLineType(0).Enabled = False
			optDamageLineType(1).Enabled = False
			optDamageLineType(0).Checked = False
			optDamageLineType(1).Checked = False
		Else ' Damaged
			With grdMatrix
				.Col = 11
				For R = .FixedRows To .Rows - 1
					.Row = R
					'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
					.CellBackColor = System.Drawing.ColorTranslator.FromOle(vbNormal)
				Next R
			End With
			fraDamageLineType.Enabled = True
			optDamageLineType(0).Enabled = True
			optDamageLineType(1).Enabled = True
		End If
		
	End Sub
	
	
	'Private Sub Timer1_Timer()
	'   'Test to see if the program is running...
	'   If IsActive(ProgHandle) Then
	'       'THE SHELLED PROGRAM IS ACTIVE
	'       Debug.Print "Running"
	'   Else
	'       'THE SHELLED PROGRAM IS NOT ACTIVE
	'       Debug.Print "Not Running"
	'       RestoreAQWADirectory
	'   End If
	'End Sub
	
	Private Sub txtEdit_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtEdit.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		EditKeyCode(grdMatrix, txtEdit, KeyCode, Shift)
	End Sub
	
	Private Sub txtEdit_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtEdit.KeyPress
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
			'UPGRADE_NOTE: Object oMet() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oMet(i) = Nothing
		Next i
	End Sub
	
	Private Sub InitiateObjects()
		Dim i, UB_old As Short
		
		UB_old = UBound(oMet)
		
		If NumCases > UB_old Then
			'UPGRADE_WARNING: Lower bound of array oMet was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim Preserve oMet(NumCases)
			For i = UB_old + 1 To NumCases
				oMet(i) = New Metocean
			Next i
		Else
			For i = UBound(oMet) To NumCases + 1 Step -1
				'UPGRADE_NOTE: Object oMet() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				oMet(i) = Nothing
			Next i
			'UPGRADE_WARNING: Lower bound of array oMet was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim Preserve oMet(NumCases)
		End If
	End Sub
	
	Private Function UpdateMetOceanObjects() As Boolean
		grdMatrix_LeaveCell(grdMatrix, New System.EventArgs())
		UpdateMetOceanObjects = False
		Dim R As Short
		With grdMatrix
			' all non-empty fields  - now update objects
			' ASSUME .FixedCols=1
			For R = .FixedRows To .Rows - 1
				oMet(R - .FixedRows + 1).Wind.Velocity.Value = Val(.get_TextMatrix(R, 1)) * Knots2Ftps
				oMet(R - .FixedRows + 1).Wind.Heading.Value = Val(.get_TextMatrix(R, 2))
				oMet(R - .FixedRows + 1).Wave.SpectrumName.Value = .get_TextMatrix(R, 3)
				oMet(R - .FixedRows + 1).Wave.Height.Value = Val(.get_TextMatrix(R, 4))
				oMet(R - .FixedRows + 1).Wave.Period.Value = Val(.get_TextMatrix(R, 5))
				oMet(R - .FixedRows + 1).Wave.gamma.Value = Val(.get_TextMatrix(R, 6))
				If Collinear Then
					oMet(R - .FixedRows + 1).Wave.Heading.Value = Val(.get_TextMatrix(R, 2))
					oMet(R - .FixedRows + 1).Current.Heading.Value = Val(.get_TextMatrix(R, 2))
					
					.set_TextMatrix(R, 7, Val(.get_TextMatrix(R, 2)))
					.set_TextMatrix(R, 8, Val(.get_TextMatrix(R, 2)))
				Else
					oMet(R - .FixedRows + 1).Wave.Heading.Value = Val(.get_TextMatrix(R, 7))
					oMet(R - .FixedRows + 1).Current.Heading.Value = Val(.get_TextMatrix(R, 8))
				End If
				oMet(R - .FixedRows + 1).Current.SurfaceVel = Val(.get_TextMatrix(R, 9)) * Knots2Ftps
				
				If UDEF Then
					oMet(R - .FixedRows + 1).Wave.SwellHeight.Value = Val(.get_TextMatrix(R, 12))
					oMet(R - .FixedRows + 1).Wave.SwellPeriod.Value = Val(.get_TextMatrix(R, 13))
					oMet(R - .FixedRows + 1).Wave.SwellSpectrumName.Value = .get_TextMatrix(R, 14)
					oMet(R - .FixedRows + 1).Wave.SwellGamma.Value = Val(.get_TextMatrix(R, 15))
					oMet(R - .FixedRows + 1).Wave.SwellHeading.Value = Val(.get_TextMatrix(R, 16))
					
				End If
			Next R
		End With
		
		UpdateMetOceanObjects = True
	End Function
	
	Private Sub UpdateBrokenLines()
		grdMatrix_LeaveCell(grdMatrix, New System.EventArgs())
		
		'UPGRADE_WARNING: Lower bound of array BreakLine was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim BreakLine(NumCases)
		Dim R As Short
		With grdMatrix
			' all non-empty fields  - now update objects
			For R = .FixedRows To .Rows - 1
				BreakLine(R - .FixedRows + 1) = CShort(Val(.get_TextMatrix(R, 11)))
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
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Split_Renamed(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
	
	Private Sub ReadOldFile(ByVal fname As String)
		Dim aline As String
		Dim Fields() As String
		Dim i As Short
		FileOpen(InFile, fname, OpenMode.Input, OpenAccess.Read)
		NumCases = 9
		txtNumCases.Text = CStr(9)
		InitiateObjects()
		InitiateGrid()
		For i = 1 To 9
			aline = LineInput(InFile)
			'UPGRADE_WARNING: Couldn't resolve default property of object Split_Renamed(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Fields = Split_Renamed(aline, " ")
			oMet(i).Wind.Velocity.Value = Fields(1)
			oMet(i).Wind.Heading.Value = Fields(2)
			oMet(i).Current.SurfaceVel = CDbl(Fields(3))
			oMet(i).Current.Heading.Value = Fields(4)
			oMet(i).Wave.Heading.Value = Fields(5)
			oMet(i).Wave.Height.Value = Fields(6)
			oMet(i).Wave.Period.Value = Fields(7)
		Next i
		FileClose(InFile)
		LoadGrid()
	End Sub
	
	Friend Sub LoadGrid()
		
		Dim R, c As Short
		For R = 1 To NumCases
			With grdMatrix
				.Row = R + .FixedRows - 1
				' load data
				.Col = 6
				If Not oMet(R) Is Nothing Then
					If oMet(R).Wind.Velocity.IsEmpty Then
						.set_TextMatrix(.FixedRows + R - 1, 1, "")
					Else
						.set_TextMatrix(.FixedRows + R - 1, 1, VB6.Format(oMet(R).Wind.Velocity.Value * Ftps2Knots, "0.00"))
					End If
					If oMet(R).Wind.Heading.IsEmpty Then
						.set_TextMatrix(.FixedRows + R - 1, 2, "")
					Else
						.set_TextMatrix(.FixedRows + R - 1, 2, oMet(R).Wind.Heading.Value)
					End If
					If oMet(R).Wave.SpectrumName.IsEmpty Then
						.set_TextMatrix(.FixedRows + R - 1, 3, "")
					Else
						.set_TextMatrix(.FixedRows + R - 1, 3, oMet(R).Wave.SpectrumName.Value)
						If InStr(oMet(R).Wave.SpectrumName.Value, "PSMZ") > 0 Then
							.CellBackColor = System.Drawing.SystemColors.Control
						Else
							'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
							.CellBackColor = System.Drawing.ColorTranslator.FromOle(vbNormal)
						End If
					End If
					If oMet(R).Wave.Height.IsEmpty Then
						.set_TextMatrix(.FixedRows + R - 1, 4, "")
					Else
						.set_TextMatrix(.FixedRows + R - 1, 4, oMet(R).Wave.Height.Value)
					End If
					If oMet(R).Wave.Period.IsEmpty Then
						.set_TextMatrix(.FixedRows + R - 1, 5, "")
					Else
						.set_TextMatrix(.FixedRows + R - 1, 5, oMet(R).Wave.Period.Value)
					End If
					If oMet(R).Wave.gamma.IsEmpty Then
						.set_TextMatrix(.FixedRows + R - 1, 6, "")
					Else
						.set_TextMatrix(.FixedRows + R - 1, 6, oMet(R).Wave.gamma.Value)
					End If
					If oMet(R).Wave.Heading.IsEmpty Then
						.set_TextMatrix(.FixedRows + R - 1, 7, "")
					Else
						.set_TextMatrix(.FixedRows + R - 1, 7, oMet(R).Wave.Heading.Value)
					End If
					If oMet(R).Current.Heading.IsEmpty Then
						.set_TextMatrix(.FixedRows + R - 1, 8, "")
					Else
						.set_TextMatrix(.FixedRows + R - 1, 8, oMet(R).Current.Heading.Value)
					End If
					.set_TextMatrix(.FixedRows + R - 1, 9, VB6.Format(oMet(R).Current.SurfaceVel * Ftps2Knots, "0.00"))
					If oMet(R).Current.Profile.Count <= 0 Then
						.set_TextMatrix(.FixedRows + R - 1, 10, "1")
					Else
						.set_TextMatrix(.FixedRows + R - 1, 10, oMet(R).Current.Profile.Count)
					End If
					
					If UDEF Then
						If oMet(R).Wave.Height.IsEmpty Then
							.set_TextMatrix(.FixedRows + R - 1, 12, "")
						Else
							.set_TextMatrix(.FixedRows + R - 1, 12, oMet(R).Wave.SwellHeight.Value)
						End If
						If oMet(R).Wave.SwellPeriod.IsEmpty Then
							.set_TextMatrix(.FixedRows + R - 1, 13, "")
						Else
							.set_TextMatrix(.FixedRows + R - 1, 13, oMet(R).Wave.SwellPeriod.Value)
						End If
						If oMet(R).Wave.SwellSpectrumName.IsEmpty Then
							.set_TextMatrix(.FixedRows + R - 1, 14, "")
						Else
							.set_TextMatrix(.FixedRows + R - 1, 14, oMet(R).Wave.SwellSpectrumName.Value)
							If InStr(oMet(R).Wave.SwellSpectrumName.Value, "PSMZ") > 0 Then
								.CellBackColor = System.Drawing.SystemColors.Control
							Else
								'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
								.CellBackColor = System.Drawing.ColorTranslator.FromOle(vbNormal)
							End If
						End If
						If oMet(R).Wave.SwellGamma.IsEmpty Then
							.set_TextMatrix(.FixedRows + R - 1, 15, "")
						Else
							.set_TextMatrix(.FixedRows + R - 1, 15, oMet(R).Wave.SwellGamma.Value)
						End If
						If oMet(R).Wave.SwellHeading.IsEmpty Then
							.set_TextMatrix(.FixedRows + R - 1, 16, "")
						Else
							.set_TextMatrix(.FixedRows + R - 1, 16, oMet(R).Wave.SwellHeading.Value)
						End If
					End If
					
				End If
			End With
		Next R
		With grdMatrix
			.Col = .FixedCols
			.Row = .FixedRows
		End With
	End Sub
	
	'UPGRADE_WARNING: Event txtNumCases.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
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