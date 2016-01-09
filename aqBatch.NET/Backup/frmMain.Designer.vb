<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmMain
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents chkUDEF As System.Windows.Forms.CheckBox
	Public WithEvents chkVersion As System.Windows.Forms.CheckBox
	Public WithEvents chkCollinear As System.Windows.Forms.CheckBox
	Public WithEvents btnCopyAll As System.Windows.Forms.Button
	Public WithEvents txtAQWAdir As System.Windows.Forms.TextBox
	Public WithEvents btnIntactFile As System.Windows.Forms.Button
	Public WithEvents _optDamageLineType_1 As System.Windows.Forms.RadioButton
	Public WithEvents _optDamageLineType_0 As System.Windows.Forms.RadioButton
	Public WithEvents fraDamageLineType As System.Windows.Forms.GroupBox
	Public WithEvents _optWind_1 As System.Windows.Forms.RadioButton
	Public WithEvents _optWind_0 As System.Windows.Forms.RadioButton
	Public WithEvents fraWind As System.Windows.Forms.GroupBox
	Public WithEvents chkPFLH As System.Windows.Forms.CheckBox
	Public WithEvents btnSilentRun As System.Windows.Forms.Button
	Public WithEvents btnCreateFiles As System.Windows.Forms.Button
	Public WithEvents btnOldAQWA As System.Windows.Forms.Button
	Public WithEvents btnNewAQWA As System.Windows.Forms.Button
	Public WithEvents _optAQWAUnits_1 As System.Windows.Forms.RadioButton
	Public WithEvents _optAQWAUnits_0 As System.Windows.Forms.RadioButton
	Public WithEvents fraUnits As System.Windows.Forms.GroupBox
	Public WithEvents btnPostProcess As System.Windows.Forms.Button
	Public WithEvents _optMoorState_0 As System.Windows.Forms.RadioButton
	Public WithEvents _optMoorState_1 As System.Windows.Forms.RadioButton
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents btnBrowseDir As System.Windows.Forms.Button
	Public WithEvents txtMetoceanName As System.Windows.Forms.TextBox
	Public WithEvents btnSaveMatrix As System.Windows.Forms.Button
	Public WithEvents btnOpenMatrix As System.Windows.Forms.Button
	Public WithEvents btnBrowseBaseDir As System.Windows.Forms.Button
	Public WithEvents txtEdit As System.Windows.Forms.TextBox
	Public WithEvents cboWaveType As System.Windows.Forms.ComboBox
	Public WithEvents btnRun As System.Windows.Forms.Button
	Public WithEvents txtNumCases As System.Windows.Forms.TextBox
	Public dlgFileOpen As System.Windows.Forms.OpenFileDialog
	Public dlgFileSave As System.Windows.Forms.SaveFileDialog
	Public WithEvents grdMatrix As AxMSFlexGridLib.AxMSFlexGrid
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents lblTargetDir As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents lblBaseDir As System.Windows.Forms.Label
	Public WithEvents lblNumCases As System.Windows.Forms.Label
	Public WithEvents optAQWAUnits As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	Public WithEvents optDamageLineType As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	Public WithEvents optMoorState As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	Public WithEvents optWind As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	Public WithEvents mnuOpen As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuSave As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuOpenOldMatrix As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnusep As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuExit As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuFile As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuCreateFiles As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuOldAQWA As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuNewAQWA As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPreProcess As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPostProcess As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuProcess As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuAbout As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuHelp As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMain))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.chkUDEF = New System.Windows.Forms.CheckBox
		Me.chkVersion = New System.Windows.Forms.CheckBox
		Me.chkCollinear = New System.Windows.Forms.CheckBox
		Me.btnCopyAll = New System.Windows.Forms.Button
		Me.txtAQWAdir = New System.Windows.Forms.TextBox
		Me.btnIntactFile = New System.Windows.Forms.Button
		Me.fraDamageLineType = New System.Windows.Forms.GroupBox
		Me._optDamageLineType_1 = New System.Windows.Forms.RadioButton
		Me._optDamageLineType_0 = New System.Windows.Forms.RadioButton
		Me.fraWind = New System.Windows.Forms.GroupBox
		Me._optWind_1 = New System.Windows.Forms.RadioButton
		Me._optWind_0 = New System.Windows.Forms.RadioButton
		Me.chkPFLH = New System.Windows.Forms.CheckBox
		Me.btnSilentRun = New System.Windows.Forms.Button
		Me.btnCreateFiles = New System.Windows.Forms.Button
		Me.btnOldAQWA = New System.Windows.Forms.Button
		Me.btnNewAQWA = New System.Windows.Forms.Button
		Me.fraUnits = New System.Windows.Forms.GroupBox
		Me._optAQWAUnits_1 = New System.Windows.Forms.RadioButton
		Me._optAQWAUnits_0 = New System.Windows.Forms.RadioButton
		Me.btnPostProcess = New System.Windows.Forms.Button
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me._optMoorState_0 = New System.Windows.Forms.RadioButton
		Me._optMoorState_1 = New System.Windows.Forms.RadioButton
		Me.btnBrowseDir = New System.Windows.Forms.Button
		Me.txtMetoceanName = New System.Windows.Forms.TextBox
		Me.btnSaveMatrix = New System.Windows.Forms.Button
		Me.btnOpenMatrix = New System.Windows.Forms.Button
		Me.btnBrowseBaseDir = New System.Windows.Forms.Button
		Me.txtEdit = New System.Windows.Forms.TextBox
		Me.cboWaveType = New System.Windows.Forms.ComboBox
		Me.btnRun = New System.Windows.Forms.Button
		Me.txtNumCases = New System.Windows.Forms.TextBox
		Me.dlgFileOpen = New System.Windows.Forms.OpenFileDialog
		Me.dlgFileSave = New System.Windows.Forms.SaveFileDialog
		Me.grdMatrix = New AxMSFlexGridLib.AxMSFlexGrid
		Me.Label1 = New System.Windows.Forms.Label
		Me.lblTargetDir = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.lblBaseDir = New System.Windows.Forms.Label
		Me.lblNumCases = New System.Windows.Forms.Label
		Me.optAQWAUnits = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(components)
		Me.optDamageLineType = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(components)
		Me.optMoorState = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(components)
		Me.optWind = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(components)
		Me.MainMenu1 = New System.Windows.Forms.MenuStrip
		Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuOpen = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuSave = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuOpenOldMatrix = New System.Windows.Forms.ToolStripMenuItem
		Me.mnusep = New System.Windows.Forms.ToolStripSeparator
		Me.mnuExit = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuProcess = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuCreateFiles = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuOldAQWA = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuNewAQWA = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPreProcess = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPostProcess = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuAbout = New System.Windows.Forms.ToolStripMenuItem
		Me.fraDamageLineType.SuspendLayout()
		Me.fraWind.SuspendLayout()
		Me.fraUnits.SuspendLayout()
		Me.Frame1.SuspendLayout()
		Me.MainMenu1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.grdMatrix, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.optAQWAUnits, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.optDamageLineType, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.optMoorState, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.optWind, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "AQWA Batch Runner - David Tein Consulting Engineers, Ltd."
		Me.ClientSize = New System.Drawing.Size(994, 544)
		Me.Location = New System.Drawing.Point(10, 36)
		Me.Icon = CType(resources.GetObject("frmMain.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmMain"
		Me.chkUDEF.Text = "Swell"
		Me.chkUDEF.Size = New System.Drawing.Size(121, 17)
		Me.chkUDEF.Location = New System.Drawing.Point(280, 512)
		Me.chkUDEF.TabIndex = 40
		Me.chkUDEF.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkUDEF.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkUDEF.BackColor = System.Drawing.SystemColors.Control
		Me.chkUDEF.CausesValidation = True
		Me.chkUDEF.Enabled = True
		Me.chkUDEF.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkUDEF.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkUDEF.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkUDEF.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkUDEF.TabStop = True
		Me.chkUDEF.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkUDEF.Visible = True
		Me.chkUDEF.Name = "chkUDEF"
		Me.chkVersion.Text = "AQWA 15.0"
		Me.chkVersion.Size = New System.Drawing.Size(121, 17)
		Me.chkVersion.Location = New System.Drawing.Point(152, 512)
		Me.chkVersion.TabIndex = 39
		Me.chkVersion.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkVersion.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkVersion.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkVersion.BackColor = System.Drawing.SystemColors.Control
		Me.chkVersion.CausesValidation = True
		Me.chkVersion.Enabled = True
		Me.chkVersion.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkVersion.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkVersion.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkVersion.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkVersion.TabStop = True
		Me.chkVersion.Visible = True
		Me.chkVersion.Name = "chkVersion"
		Me.chkCollinear.Text = "Collinear"
		Me.chkCollinear.Size = New System.Drawing.Size(73, 17)
		Me.chkCollinear.Location = New System.Drawing.Point(352, 232)
		Me.chkCollinear.TabIndex = 38
		Me.chkCollinear.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkCollinear.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkCollinear.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkCollinear.BackColor = System.Drawing.SystemColors.Control
		Me.chkCollinear.CausesValidation = True
		Me.chkCollinear.Enabled = True
		Me.chkCollinear.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkCollinear.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkCollinear.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkCollinear.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkCollinear.TabStop = True
		Me.chkCollinear.Visible = True
		Me.chkCollinear.Name = "chkCollinear"
		Me.btnCopyAll.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnCopyAll.Text = "Copy All"
		Me.btnCopyAll.Size = New System.Drawing.Size(73, 25)
		Me.btnCopyAll.Location = New System.Drawing.Point(440, 232)
		Me.btnCopyAll.TabIndex = 37
		Me.ToolTip1.SetToolTip(Me.btnCopyAll, "Copy Case 1 to All cases except heading data")
		Me.btnCopyAll.BackColor = System.Drawing.SystemColors.Control
		Me.btnCopyAll.CausesValidation = True
		Me.btnCopyAll.Enabled = True
		Me.btnCopyAll.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnCopyAll.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnCopyAll.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnCopyAll.TabStop = True
		Me.btnCopyAll.Name = "btnCopyAll"
		Me.txtAQWAdir.AutoSize = False
		Me.txtAQWAdir.Size = New System.Drawing.Size(161, 19)
		Me.txtAQWAdir.Location = New System.Drawing.Point(104, 232)
		Me.txtAQWAdir.TabIndex = 35
		Me.txtAQWAdir.Text = "aqwa /nowind"
		Me.txtAQWAdir.AcceptsReturn = True
		Me.txtAQWAdir.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtAQWAdir.BackColor = System.Drawing.SystemColors.Window
		Me.txtAQWAdir.CausesValidation = True
		Me.txtAQWAdir.Enabled = True
		Me.txtAQWAdir.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtAQWAdir.HideSelection = True
		Me.txtAQWAdir.ReadOnly = False
		Me.txtAQWAdir.Maxlength = 0
		Me.txtAQWAdir.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtAQWAdir.MultiLine = False
		Me.txtAQWAdir.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtAQWAdir.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtAQWAdir.TabStop = True
		Me.txtAQWAdir.Visible = True
		Me.txtAQWAdir.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtAQWAdir.Name = "txtAQWAdir"
		Me.btnIntactFile.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnIntactFile.Text = "Get Critical Lines from Intact XLS file"
		Me.btnIntactFile.Size = New System.Drawing.Size(189, 26)
		Me.btnIntactFile.Location = New System.Drawing.Point(520, 232)
		Me.btnIntactFile.TabIndex = 34
		Me.btnIntactFile.BackColor = System.Drawing.SystemColors.Control
		Me.btnIntactFile.CausesValidation = True
		Me.btnIntactFile.Enabled = True
		Me.btnIntactFile.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnIntactFile.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnIntactFile.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnIntactFile.TabStop = True
		Me.btnIntactFile.Name = "btnIntactFile"
		Me.fraDamageLineType.Text = "Damage"
		Me.fraDamageLineType.Enabled = False
		Me.fraDamageLineType.Size = New System.Drawing.Size(110, 73)
		Me.fraDamageLineType.Location = New System.Drawing.Point(595, 151)
		Me.fraDamageLineType.TabIndex = 31
		Me.fraDamageLineType.BackColor = System.Drawing.SystemColors.Control
		Me.fraDamageLineType.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraDamageLineType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraDamageLineType.Visible = True
		Me.fraDamageLineType.Name = "fraDamageLineType"
		Me._optDamageLineType_1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optDamageLineType_1.Text = "2nd-Loaded"
		Me._optDamageLineType_1.Size = New System.Drawing.Size(85, 18)
		Me._optDamageLineType_1.Location = New System.Drawing.Point(15, 42)
		Me._optDamageLineType_1.TabIndex = 33
		Me._optDamageLineType_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optDamageLineType_1.BackColor = System.Drawing.SystemColors.Control
		Me._optDamageLineType_1.CausesValidation = True
		Me._optDamageLineType_1.Enabled = True
		Me._optDamageLineType_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optDamageLineType_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._optDamageLineType_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optDamageLineType_1.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optDamageLineType_1.TabStop = True
		Me._optDamageLineType_1.Checked = False
		Me._optDamageLineType_1.Visible = True
		Me._optDamageLineType_1.Name = "_optDamageLineType_1"
		Me._optDamageLineType_0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optDamageLineType_0.Text = "Most-Loaded"
		Me._optDamageLineType_0.Enabled = False
		Me._optDamageLineType_0.Size = New System.Drawing.Size(90, 18)
		Me._optDamageLineType_0.Location = New System.Drawing.Point(15, 20)
		Me._optDamageLineType_0.TabIndex = 32
		Me._optDamageLineType_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optDamageLineType_0.BackColor = System.Drawing.SystemColors.Control
		Me._optDamageLineType_0.CausesValidation = True
		Me._optDamageLineType_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optDamageLineType_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._optDamageLineType_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optDamageLineType_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optDamageLineType_0.TabStop = True
		Me._optDamageLineType_0.Checked = False
		Me._optDamageLineType_0.Visible = True
		Me._optDamageLineType_0.Name = "_optDamageLineType_0"
		Me.fraWind.Text = "Wind"
		Me.fraWind.Size = New System.Drawing.Size(75, 74)
		Me.fraWind.Location = New System.Drawing.Point(415, 152)
		Me.fraWind.TabIndex = 28
		Me.fraWind.BackColor = System.Drawing.SystemColors.Control
		Me.fraWind.Enabled = True
		Me.fraWind.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraWind.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraWind.Visible = True
		Me.fraWind.Name = "fraWind"
		Me._optWind_1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optWind_1.Text = "1- min"
		Me._optWind_1.Size = New System.Drawing.Size(54, 20)
		Me._optWind_1.Location = New System.Drawing.Point(9, 42)
		Me._optWind_1.TabIndex = 30
		Me._optWind_1.Checked = True
		Me._optWind_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optWind_1.BackColor = System.Drawing.SystemColors.Control
		Me._optWind_1.CausesValidation = True
		Me._optWind_1.Enabled = True
		Me._optWind_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optWind_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._optWind_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optWind_1.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optWind_1.TabStop = True
		Me._optWind_1.Visible = True
		Me._optWind_1.Name = "_optWind_1"
		Me._optWind_0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optWind_0.Text = "1-hr"
		Me._optWind_0.Size = New System.Drawing.Size(44, 20)
		Me._optWind_0.Location = New System.Drawing.Point(9, 18)
		Me._optWind_0.TabIndex = 29
		Me._optWind_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optWind_0.BackColor = System.Drawing.SystemColors.Control
		Me._optWind_0.CausesValidation = True
		Me._optWind_0.Enabled = True
		Me._optWind_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optWind_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._optWind_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optWind_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optWind_0.TabStop = True
		Me._optWind_0.Checked = False
		Me._optWind_0.Visible = True
		Me._optWind_0.Name = "_optWind_0"
		Me.chkPFLH.Text = "PFLH"
		Me.chkPFLH.Size = New System.Drawing.Size(65, 19)
		Me.chkPFLH.Location = New System.Drawing.Point(280, 232)
		Me.chkPFLH.TabIndex = 27
		Me.chkPFLH.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkPFLH.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkPFLH.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkPFLH.BackColor = System.Drawing.SystemColors.Control
		Me.chkPFLH.CausesValidation = True
		Me.chkPFLH.Enabled = True
		Me.chkPFLH.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkPFLH.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkPFLH.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkPFLH.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkPFLH.TabStop = True
		Me.chkPFLH.Visible = True
		Me.chkPFLH.Name = "chkPFLH"
		Me.btnSilentRun.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnSilentRun.Text = "&Silent Run"
		Me.btnSilentRun.Size = New System.Drawing.Size(98, 25)
		Me.btnSilentRun.Location = New System.Drawing.Point(514, 508)
		Me.btnSilentRun.TabIndex = 26
		Me.ToolTip1.SetToolTip(Me.btnSilentRun, "Launch DOS window to run all AQWA cases")
		Me.btnSilentRun.BackColor = System.Drawing.SystemColors.Control
		Me.btnSilentRun.CausesValidation = True
		Me.btnSilentRun.Enabled = True
		Me.btnSilentRun.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnSilentRun.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnSilentRun.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnSilentRun.TabStop = True
		Me.btnSilentRun.Name = "btnSilentRun"
		Me.btnCreateFiles.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnCreateFiles.Text = "&Create Case Files"
		Me.btnCreateFiles.Enabled = False
		Me.btnCreateFiles.Size = New System.Drawing.Size(117, 25)
		Me.btnCreateFiles.Location = New System.Drawing.Point(13, 508)
		Me.btnCreateFiles.TabIndex = 25
		Me.ToolTip1.SetToolTip(Me.btnCreateFiles, "Prepare AQWA input files for all environmental headings")
		Me.btnCreateFiles.BackColor = System.Drawing.SystemColors.Control
		Me.btnCreateFiles.CausesValidation = True
		Me.btnCreateFiles.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnCreateFiles.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnCreateFiles.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnCreateFiles.TabStop = True
		Me.btnCreateFiles.Name = "btnCreateFiles"
		Me.btnOldAQWA.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnOldAQWA.Text = "Switch to Old AQWA"
		Me.btnOldAQWA.Enabled = False
		Me.btnOldAQWA.Size = New System.Drawing.Size(140, 25)
		Me.btnOldAQWA.Location = New System.Drawing.Point(133, 508)
		Me.btnOldAQWA.TabIndex = 24
		Me.ToolTip1.SetToolTip(Me.btnOldAQWA, "Set your system ready to run OLD AQWA")
		Me.btnOldAQWA.Visible = False
		Me.btnOldAQWA.BackColor = System.Drawing.SystemColors.Control
		Me.btnOldAQWA.CausesValidation = True
		Me.btnOldAQWA.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnOldAQWA.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnOldAQWA.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnOldAQWA.TabStop = True
		Me.btnOldAQWA.Name = "btnOldAQWA"
		Me.btnNewAQWA.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnNewAQWA.Text = "Restore to New AQWA"
		Me.btnNewAQWA.Enabled = False
		Me.btnNewAQWA.Size = New System.Drawing.Size(141, 25)
		Me.btnNewAQWA.Location = New System.Drawing.Point(275, 508)
		Me.btnNewAQWA.TabIndex = 23
		Me.ToolTip1.SetToolTip(Me.btnNewAQWA, "Set your system ready to run NEW AQWA")
		Me.btnNewAQWA.Visible = False
		Me.btnNewAQWA.BackColor = System.Drawing.SystemColors.Control
		Me.btnNewAQWA.CausesValidation = True
		Me.btnNewAQWA.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnNewAQWA.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnNewAQWA.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnNewAQWA.TabStop = True
		Me.btnNewAQWA.Name = "btnNewAQWA"
		Me.fraUnits.Text = "AQWA Units"
		Me.fraUnits.Size = New System.Drawing.Size(87, 74)
		Me.fraUnits.Location = New System.Drawing.Point(324, 152)
		Me.fraUnits.TabIndex = 20
		Me.fraUnits.BackColor = System.Drawing.SystemColors.Control
		Me.fraUnits.Enabled = True
		Me.fraUnits.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraUnits.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraUnits.Visible = True
		Me.fraUnits.Name = "fraUnits"
		Me._optAQWAUnits_1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optAQWAUnits_1.Text = "Metric"
		Me._optAQWAUnits_1.Size = New System.Drawing.Size(51, 18)
		Me._optAQWAUnits_1.Location = New System.Drawing.Point(23, 43)
		Me._optAQWAUnits_1.TabIndex = 22
		Me._optAQWAUnits_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optAQWAUnits_1.BackColor = System.Drawing.SystemColors.Control
		Me._optAQWAUnits_1.CausesValidation = True
		Me._optAQWAUnits_1.Enabled = True
		Me._optAQWAUnits_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optAQWAUnits_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._optAQWAUnits_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optAQWAUnits_1.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optAQWAUnits_1.TabStop = True
		Me._optAQWAUnits_1.Checked = False
		Me._optAQWAUnits_1.Visible = True
		Me._optAQWAUnits_1.Name = "_optAQWAUnits_1"
		Me._optAQWAUnits_0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optAQWAUnits_0.Text = "US"
		Me._optAQWAUnits_0.Size = New System.Drawing.Size(54, 18)
		Me._optAQWAUnits_0.Location = New System.Drawing.Point(23, 21)
		Me._optAQWAUnits_0.TabIndex = 21
		Me._optAQWAUnits_0.Checked = True
		Me._optAQWAUnits_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optAQWAUnits_0.BackColor = System.Drawing.SystemColors.Control
		Me._optAQWAUnits_0.CausesValidation = True
		Me._optAQWAUnits_0.Enabled = True
		Me._optAQWAUnits_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optAQWAUnits_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._optAQWAUnits_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optAQWAUnits_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optAQWAUnits_0.TabStop = True
		Me._optAQWAUnits_0.Visible = True
		Me._optAQWAUnits_0.Name = "_optAQWAUnits_0"
		Me.btnPostProcess.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnPostProcess.Text = "&Get Results"
		Me.btnPostProcess.Enabled = False
		Me.btnPostProcess.Size = New System.Drawing.Size(98, 25)
		Me.btnPostProcess.Location = New System.Drawing.Point(614, 506)
		Me.btnPostProcess.TabIndex = 19
		Me.ToolTip1.SetToolTip(Me.btnPostProcess, "Launch Excel to display AQWA results.")
		Me.btnPostProcess.BackColor = System.Drawing.SystemColors.Control
		Me.btnPostProcess.CausesValidation = True
		Me.btnPostProcess.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnPostProcess.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnPostProcess.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnPostProcess.TabStop = True
		Me.btnPostProcess.Name = "btnPostProcess"
		Me.Frame1.Text = "Mooring System"
		Me.Frame1.Size = New System.Drawing.Size(93, 73)
		Me.Frame1.Location = New System.Drawing.Point(496, 151)
		Me.Frame1.TabIndex = 16
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Name = "Frame1"
		Me._optMoorState_0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optMoorState_0.Text = "Intact"
		Me._optMoorState_0.Size = New System.Drawing.Size(64, 18)
		Me._optMoorState_0.Location = New System.Drawing.Point(14, 21)
		Me._optMoorState_0.TabIndex = 18
		Me._optMoorState_0.Checked = True
		Me._optMoorState_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optMoorState_0.BackColor = System.Drawing.SystemColors.Control
		Me._optMoorState_0.CausesValidation = True
		Me._optMoorState_0.Enabled = True
		Me._optMoorState_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optMoorState_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._optMoorState_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optMoorState_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optMoorState_0.TabStop = True
		Me._optMoorState_0.Visible = True
		Me._optMoorState_0.Name = "_optMoorState_0"
		Me._optMoorState_1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optMoorState_1.Text = "Damaged"
		Me._optMoorState_1.Size = New System.Drawing.Size(72, 18)
		Me._optMoorState_1.Location = New System.Drawing.Point(14, 43)
		Me._optMoorState_1.TabIndex = 17
		Me._optMoorState_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optMoorState_1.BackColor = System.Drawing.SystemColors.Control
		Me._optMoorState_1.CausesValidation = True
		Me._optMoorState_1.Enabled = True
		Me._optMoorState_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optMoorState_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._optMoorState_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optMoorState_1.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optMoorState_1.TabStop = True
		Me._optMoorState_1.Checked = False
		Me._optMoorState_1.Visible = True
		Me._optMoorState_1.Name = "_optMoorState_1"
		Me.btnBrowseDir.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnBrowseDir.Text = ">>"
		Me.btnBrowseDir.Size = New System.Drawing.Size(24, 24)
		Me.btnBrowseDir.Location = New System.Drawing.Point(695, 81)
		Me.btnBrowseDir.TabIndex = 15
		Me.ToolTip1.SetToolTip(Me.btnBrowseDir, "Browse Directory to run and write your AQWA results")
		Me.btnBrowseDir.BackColor = System.Drawing.SystemColors.Control
		Me.btnBrowseDir.CausesValidation = True
		Me.btnBrowseDir.Enabled = True
		Me.btnBrowseDir.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnBrowseDir.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnBrowseDir.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnBrowseDir.TabStop = True
		Me.btnBrowseDir.Name = "btnBrowseDir"
		Me.txtMetoceanName.AutoSize = False
		Me.txtMetoceanName.Size = New System.Drawing.Size(209, 19)
		Me.txtMetoceanName.Location = New System.Drawing.Point(105, 154)
		Me.txtMetoceanName.TabIndex = 11
		Me.txtMetoceanName.AcceptsReturn = True
		Me.txtMetoceanName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtMetoceanName.BackColor = System.Drawing.SystemColors.Window
		Me.txtMetoceanName.CausesValidation = True
		Me.txtMetoceanName.Enabled = True
		Me.txtMetoceanName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtMetoceanName.HideSelection = True
		Me.txtMetoceanName.ReadOnly = False
		Me.txtMetoceanName.Maxlength = 0
		Me.txtMetoceanName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtMetoceanName.MultiLine = False
		Me.txtMetoceanName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtMetoceanName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtMetoceanName.TabStop = True
		Me.txtMetoceanName.Visible = True
		Me.txtMetoceanName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtMetoceanName.Name = "txtMetoceanName"
		Me.btnSaveMatrix.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnSaveMatrix.Text = "&Save Matrix"
		Me.btnSaveMatrix.Size = New System.Drawing.Size(82, 25)
		Me.btnSaveMatrix.Location = New System.Drawing.Point(229, 184)
		Me.btnSaveMatrix.TabIndex = 10
		Me.btnSaveMatrix.BackColor = System.Drawing.SystemColors.Control
		Me.btnSaveMatrix.CausesValidation = True
		Me.btnSaveMatrix.Enabled = True
		Me.btnSaveMatrix.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnSaveMatrix.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnSaveMatrix.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnSaveMatrix.TabStop = True
		Me.btnSaveMatrix.Name = "btnSaveMatrix"
		Me.btnOpenMatrix.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnOpenMatrix.Text = "&Open Matrix"
		Me.btnOpenMatrix.Size = New System.Drawing.Size(82, 25)
		Me.btnOpenMatrix.Location = New System.Drawing.Point(146, 184)
		Me.btnOpenMatrix.TabIndex = 9
		Me.btnOpenMatrix.BackColor = System.Drawing.SystemColors.Control
		Me.btnOpenMatrix.CausesValidation = True
		Me.btnOpenMatrix.Enabled = True
		Me.btnOpenMatrix.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnOpenMatrix.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnOpenMatrix.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnOpenMatrix.TabStop = True
		Me.btnOpenMatrix.Name = "btnOpenMatrix"
		Me.btnBrowseBaseDir.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnBrowseBaseDir.Text = ">>"
		Me.btnBrowseBaseDir.Size = New System.Drawing.Size(24, 24)
		Me.btnBrowseBaseDir.Location = New System.Drawing.Point(694, 17)
		Me.btnBrowseBaseDir.TabIndex = 6
		Me.ToolTip1.SetToolTip(Me.btnBrowseBaseDir, "Browse for directory of ABRUN without environment")
		Me.btnBrowseBaseDir.BackColor = System.Drawing.SystemColors.Control
		Me.btnBrowseBaseDir.CausesValidation = True
		Me.btnBrowseBaseDir.Enabled = True
		Me.btnBrowseBaseDir.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnBrowseBaseDir.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnBrowseBaseDir.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnBrowseBaseDir.TabStop = True
		Me.btnBrowseBaseDir.Name = "btnBrowseBaseDir"
		Me.txtEdit.AutoSize = False
		Me.txtEdit.Size = New System.Drawing.Size(59, 18)
		Me.txtEdit.Location = New System.Drawing.Point(112, 427)
		Me.txtEdit.TabIndex = 5
		Me.txtEdit.Visible = False
		Me.txtEdit.AcceptsReturn = True
		Me.txtEdit.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtEdit.BackColor = System.Drawing.SystemColors.Window
		Me.txtEdit.CausesValidation = True
		Me.txtEdit.Enabled = True
		Me.txtEdit.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtEdit.HideSelection = True
		Me.txtEdit.ReadOnly = False
		Me.txtEdit.Maxlength = 0
		Me.txtEdit.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtEdit.MultiLine = False
		Me.txtEdit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtEdit.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtEdit.TabStop = True
		Me.txtEdit.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtEdit.Name = "txtEdit"
		Me.cboWaveType.Size = New System.Drawing.Size(142, 21)
		Me.cboWaveType.Location = New System.Drawing.Point(114, 457)
		Me.cboWaveType.Items.AddRange(New Object(){"Pierson-Moskowitz", "Jonswap"})
		Me.cboWaveType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboWaveType.TabIndex = 3
		Me.cboWaveType.Visible = False
		Me.cboWaveType.BackColor = System.Drawing.SystemColors.Window
		Me.cboWaveType.CausesValidation = True
		Me.cboWaveType.Enabled = True
		Me.cboWaveType.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboWaveType.IntegralHeight = True
		Me.cboWaveType.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboWaveType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboWaveType.Sorted = False
		Me.cboWaveType.TabStop = True
		Me.cboWaveType.Name = "cboWaveType"
		Me.btnRun.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnRun.Text = "&Run Cases"
		Me.btnRun.Size = New System.Drawing.Size(89, 25)
		Me.btnRun.Location = New System.Drawing.Point(419, 508)
		Me.btnRun.TabIndex = 2
		Me.ToolTip1.SetToolTip(Me.btnRun, "Launch DOS window to run all AQWA cases")
		Me.btnRun.BackColor = System.Drawing.SystemColors.Control
		Me.btnRun.CausesValidation = True
		Me.btnRun.Enabled = True
		Me.btnRun.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnRun.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnRun.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnRun.TabStop = True
		Me.btnRun.Name = "btnRun"
		Me.txtNumCases.AutoSize = False
		Me.txtNumCases.Size = New System.Drawing.Size(34, 19)
		Me.txtNumCases.Location = New System.Drawing.Point(104, 191)
		Me.txtNumCases.TabIndex = 0
		Me.txtNumCases.Text = "1"
		Me.txtNumCases.AcceptsReturn = True
		Me.txtNumCases.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtNumCases.BackColor = System.Drawing.SystemColors.Window
		Me.txtNumCases.CausesValidation = True
		Me.txtNumCases.Enabled = True
		Me.txtNumCases.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtNumCases.HideSelection = True
		Me.txtNumCases.ReadOnly = False
		Me.txtNumCases.Maxlength = 0
		Me.txtNumCases.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtNumCases.MultiLine = False
		Me.txtNumCases.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtNumCases.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtNumCases.TabStop = True
		Me.txtNumCases.Visible = True
		Me.txtNumCases.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtNumCases.Name = "txtNumCases"
		grdMatrix.OcxState = CType(resources.GetObject("grdMatrix.OcxState"), System.Windows.Forms.AxHost.State)
		Me.grdMatrix.Size = New System.Drawing.Size(948, 236)
		Me.grdMatrix.Location = New System.Drawing.Point(11, 266)
		Me.grdMatrix.TabIndex = 4
		Me.grdMatrix.Name = "grdMatrix"
		Me.Label1.Text = " AQWA path:"
		Me.Label1.Size = New System.Drawing.Size(89, 17)
		Me.Label1.Location = New System.Drawing.Point(8, 232)
		Me.Label1.TabIndex = 36
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.BackColor = System.Drawing.SystemColors.Control
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.lblTargetDir.Size = New System.Drawing.Size(582, 63)
		Me.lblTargetDir.Location = New System.Drawing.Point(104, 81)
		Me.lblTargetDir.TabIndex = 14
		Me.lblTargetDir.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblTargetDir.BackColor = System.Drawing.SystemColors.Control
		Me.lblTargetDir.Enabled = True
		Me.lblTargetDir.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblTargetDir.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblTargetDir.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblTargetDir.UseMnemonic = True
		Me.lblTargetDir.Visible = True
		Me.lblTargetDir.AutoSize = False
		Me.lblTargetDir.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lblTargetDir.Name = "lblTargetDir"
		Me.Label4.Text = "Target Directory:"
		Me.Label4.Size = New System.Drawing.Size(84, 17)
		Me.Label4.Location = New System.Drawing.Point(9, 82)
		Me.Label4.TabIndex = 13
		Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label4.BackColor = System.Drawing.SystemColors.Control
		Me.Label4.Enabled = True
		Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label4.UseMnemonic = True
		Me.Label4.Visible = True
		Me.Label4.AutoSize = False
		Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label4.Name = "Label4"
		Me.Label2.Text = "Metocean Name:"
		Me.Label2.Size = New System.Drawing.Size(92, 17)
		Me.Label2.Location = New System.Drawing.Point(9, 158)
		Me.Label2.TabIndex = 12
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label2.BackColor = System.Drawing.SystemColors.Control
		Me.Label2.Enabled = True
		Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me.Label3.Text = "Base File Directory:"
		Me.Label3.Size = New System.Drawing.Size(93, 17)
		Me.Label3.Location = New System.Drawing.Point(9, 17)
		Me.Label3.TabIndex = 8
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label3.BackColor = System.Drawing.SystemColors.Control
		Me.Label3.Enabled = True
		Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label3.UseMnemonic = True
		Me.Label3.Visible = True
		Me.Label3.AutoSize = False
		Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label3.Name = "Label3"
		Me.lblBaseDir.Size = New System.Drawing.Size(582, 61)
		Me.lblBaseDir.Location = New System.Drawing.Point(105, 13)
		Me.lblBaseDir.TabIndex = 7
		Me.lblBaseDir.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblBaseDir.BackColor = System.Drawing.SystemColors.Control
		Me.lblBaseDir.Enabled = True
		Me.lblBaseDir.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblBaseDir.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblBaseDir.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblBaseDir.UseMnemonic = True
		Me.lblBaseDir.Visible = True
		Me.lblBaseDir.AutoSize = False
		Me.lblBaseDir.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lblBaseDir.Name = "lblBaseDir"
		Me.lblNumCases.Text = "Number of Cases:"
		Me.lblNumCases.Size = New System.Drawing.Size(92, 17)
		Me.lblNumCases.Location = New System.Drawing.Point(9, 192)
		Me.lblNumCases.TabIndex = 1
		Me.lblNumCases.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblNumCases.BackColor = System.Drawing.SystemColors.Control
		Me.lblNumCases.Enabled = True
		Me.lblNumCases.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblNumCases.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblNumCases.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblNumCases.UseMnemonic = True
		Me.lblNumCases.Visible = True
		Me.lblNumCases.AutoSize = False
		Me.lblNumCases.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblNumCases.Name = "lblNumCases"
		Me.mnuFile.Name = "mnuFile"
		Me.mnuFile.Text = "File"
		Me.mnuFile.Checked = False
		Me.mnuFile.Enabled = True
		Me.mnuFile.Visible = True
		Me.mnuOpen.Name = "mnuOpen"
		Me.mnuOpen.Text = "Open Matrix..."
		Me.mnuOpen.Checked = False
		Me.mnuOpen.Enabled = True
		Me.mnuOpen.Visible = True
		Me.mnuSave.Name = "mnuSave"
		Me.mnuSave.Text = "Save Matrix..."
		Me.mnuSave.Checked = False
		Me.mnuSave.Enabled = True
		Me.mnuSave.Visible = True
		Me.mnuOpenOldMatrix.Name = "mnuOpenOldMatrix"
		Me.mnuOpenOldMatrix.Text = "Open Old Matrix"
		Me.mnuOpenOldMatrix.Checked = False
		Me.mnuOpenOldMatrix.Enabled = True
		Me.mnuOpenOldMatrix.Visible = True
		Me.mnusep.Enabled = True
		Me.mnusep.Visible = True
		Me.mnusep.Name = "mnusep"
		Me.mnuExit.Name = "mnuExit"
		Me.mnuExit.Text = "Exit"
		Me.mnuExit.Checked = False
		Me.mnuExit.Enabled = True
		Me.mnuExit.Visible = True
		Me.mnuProcess.Name = "mnuProcess"
		Me.mnuProcess.Text = "&Tools"
		Me.mnuProcess.Checked = False
		Me.mnuProcess.Enabled = True
		Me.mnuProcess.Visible = True
		Me.mnuCreateFiles.Name = "mnuCreateFiles"
		Me.mnuCreateFiles.Text = "Create Files"
		Me.mnuCreateFiles.Checked = False
		Me.mnuCreateFiles.Enabled = True
		Me.mnuCreateFiles.Visible = True
		Me.mnuOldAQWA.Name = "mnuOldAQWA"
		Me.mnuOldAQWA.Text = "Switch To Old AQWA"
		Me.mnuOldAQWA.Checked = False
		Me.mnuOldAQWA.Enabled = True
		Me.mnuOldAQWA.Visible = True
		Me.mnuNewAQWA.Name = "mnuNewAQWA"
		Me.mnuNewAQWA.Text = "Restore To New AQWA"
		Me.mnuNewAQWA.Checked = False
		Me.mnuNewAQWA.Enabled = True
		Me.mnuNewAQWA.Visible = True
		Me.mnuPreProcess.Name = "mnuPreProcess"
		Me.mnuPreProcess.Text = "&Run Cases"
		Me.mnuPreProcess.Checked = False
		Me.mnuPreProcess.Enabled = True
		Me.mnuPreProcess.Visible = True
		Me.mnuPostProcess.Name = "mnuPostProcess"
		Me.mnuPostProcess.Text = "&Get Results..."
		Me.mnuPostProcess.Checked = False
		Me.mnuPostProcess.Enabled = True
		Me.mnuPostProcess.Visible = True
		Me.mnuHelp.Name = "mnuHelp"
		Me.mnuHelp.Text = "Help"
		Me.mnuHelp.Checked = False
		Me.mnuHelp.Enabled = True
		Me.mnuHelp.Visible = True
		Me.mnuAbout.Name = "mnuAbout"
		Me.mnuAbout.Text = "About..."
		Me.mnuAbout.Checked = False
		Me.mnuAbout.Enabled = True
		Me.mnuAbout.Visible = True
		Me.Controls.Add(chkUDEF)
		Me.Controls.Add(chkVersion)
		Me.Controls.Add(chkCollinear)
		Me.Controls.Add(btnCopyAll)
		Me.Controls.Add(txtAQWAdir)
		Me.Controls.Add(btnIntactFile)
		Me.Controls.Add(fraDamageLineType)
		Me.Controls.Add(fraWind)
		Me.Controls.Add(chkPFLH)
		Me.Controls.Add(btnSilentRun)
		Me.Controls.Add(btnCreateFiles)
		Me.Controls.Add(btnOldAQWA)
		Me.Controls.Add(btnNewAQWA)
		Me.Controls.Add(fraUnits)
		Me.Controls.Add(btnPostProcess)
		Me.Controls.Add(Frame1)
		Me.Controls.Add(btnBrowseDir)
		Me.Controls.Add(txtMetoceanName)
		Me.Controls.Add(btnSaveMatrix)
		Me.Controls.Add(btnOpenMatrix)
		Me.Controls.Add(btnBrowseBaseDir)
		Me.Controls.Add(txtEdit)
		Me.Controls.Add(cboWaveType)
		Me.Controls.Add(btnRun)
		Me.Controls.Add(txtNumCases)
		Me.Controls.Add(grdMatrix)
		Me.Controls.Add(Label1)
		Me.Controls.Add(lblTargetDir)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label3)
		Me.Controls.Add(lblBaseDir)
		Me.Controls.Add(lblNumCases)
		Me.fraDamageLineType.Controls.Add(_optDamageLineType_1)
		Me.fraDamageLineType.Controls.Add(_optDamageLineType_0)
		Me.fraWind.Controls.Add(_optWind_1)
		Me.fraWind.Controls.Add(_optWind_0)
		Me.fraUnits.Controls.Add(_optAQWAUnits_1)
		Me.fraUnits.Controls.Add(_optAQWAUnits_0)
		Me.Frame1.Controls.Add(_optMoorState_0)
		Me.Frame1.Controls.Add(_optMoorState_1)
		Me.optAQWAUnits.SetIndex(_optAQWAUnits_1, CType(1, Short))
		Me.optAQWAUnits.SetIndex(_optAQWAUnits_0, CType(0, Short))
		Me.optDamageLineType.SetIndex(_optDamageLineType_1, CType(1, Short))
		Me.optDamageLineType.SetIndex(_optDamageLineType_0, CType(0, Short))
		Me.optMoorState.SetIndex(_optMoorState_0, CType(0, Short))
		Me.optMoorState.SetIndex(_optMoorState_1, CType(1, Short))
		Me.optWind.SetIndex(_optWind_1, CType(1, Short))
		Me.optWind.SetIndex(_optWind_0, CType(0, Short))
		CType(Me.optWind, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.optMoorState, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.optDamageLineType, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.optAQWAUnits, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.grdMatrix, System.ComponentModel.ISupportInitialize).EndInit()
		MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuFile, Me.mnuProcess, Me.mnuHelp})
		mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuOpen, Me.mnuSave, Me.mnuOpenOldMatrix, Me.mnusep, Me.mnuExit})
		mnuProcess.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuCreateFiles, Me.mnuOldAQWA, Me.mnuNewAQWA, Me.mnuPreProcess, Me.mnuPostProcess})
		mnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuAbout})
		Me.Controls.Add(MainMenu1)
		Me.fraDamageLineType.ResumeLayout(False)
		Me.fraWind.ResumeLayout(False)
		Me.fraUnits.ResumeLayout(False)
		Me.Frame1.ResumeLayout(False)
		Me.MainMenu1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class