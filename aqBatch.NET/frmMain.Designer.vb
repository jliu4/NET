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
    Public WithEvents btnRun As System.Windows.Forms.Button
    Public WithEvents txtNumCases As System.Windows.Forms.TextBox
	Public dlgFileOpen As System.Windows.Forms.OpenFileDialog
	Public dlgFileSave As System.Windows.Forms.SaveFileDialog
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
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnCopyAll = New System.Windows.Forms.Button()
        Me.btnSilentRun = New System.Windows.Forms.Button()
        Me.btnCreateFiles = New System.Windows.Forms.Button()
        Me.btnOldAQWA = New System.Windows.Forms.Button()
        Me.btnNewAQWA = New System.Windows.Forms.Button()
        Me.btnPostProcess = New System.Windows.Forms.Button()
        Me.btnBrowseDir = New System.Windows.Forms.Button()
        Me.btnBrowseBaseDir = New System.Windows.Forms.Button()
        Me.btnRun = New System.Windows.Forms.Button()
        Me.chkUDEF = New System.Windows.Forms.CheckBox()
        Me.chkVersion = New System.Windows.Forms.CheckBox()
        Me.chkCollinear = New System.Windows.Forms.CheckBox()
        Me.txtAQWAdir = New System.Windows.Forms.TextBox()
        Me.btnIntactFile = New System.Windows.Forms.Button()
        Me.fraDamageLineType = New System.Windows.Forms.GroupBox()
        Me._optDamageLineType_1 = New System.Windows.Forms.RadioButton()
        Me._optDamageLineType_0 = New System.Windows.Forms.RadioButton()
        Me.fraWind = New System.Windows.Forms.GroupBox()
        Me._optWind_1 = New System.Windows.Forms.RadioButton()
        Me._optWind_0 = New System.Windows.Forms.RadioButton()
        Me.chkPFLH = New System.Windows.Forms.CheckBox()
        Me.fraUnits = New System.Windows.Forms.GroupBox()
        Me._optAQWAUnits_1 = New System.Windows.Forms.RadioButton()
        Me._optAQWAUnits_0 = New System.Windows.Forms.RadioButton()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me._optMoorState_0 = New System.Windows.Forms.RadioButton()
        Me._optMoorState_1 = New System.Windows.Forms.RadioButton()
        Me.txtMetoceanName = New System.Windows.Forms.TextBox()
        Me.btnSaveMatrix = New System.Windows.Forms.Button()
        Me.btnOpenMatrix = New System.Windows.Forms.Button()
        Me.txtNumCases = New System.Windows.Forms.TextBox()
        Me.dlgFileOpen = New System.Windows.Forms.OpenFileDialog()
        Me.dlgFileSave = New System.Windows.Forms.SaveFileDialog()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblTargetDir = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblBaseDir = New System.Windows.Forms.Label()
        Me.lblNumCases = New System.Windows.Forms.Label()
        Me.optAQWAUnits = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.optDamageLineType = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.optMoorState = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.optWind = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.grdMatrix = New System.Windows.Forms.DataGridView()
        Me.fraDamageLineType.SuspendLayout()
        Me.fraWind.SuspendLayout()
        Me.fraUnits.SuspendLayout()
        Me.Frame1.SuspendLayout()
        CType(Me.optAQWAUnits, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optDamageLineType, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optMoorState, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optWind, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdMatrix, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnCopyAll
        '
        Me.btnCopyAll.BackColor = System.Drawing.SystemColors.Control
        Me.btnCopyAll.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCopyAll.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCopyAll.Location = New System.Drawing.Point(440, 232)
        Me.btnCopyAll.Name = "btnCopyAll"
        Me.btnCopyAll.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCopyAll.Size = New System.Drawing.Size(73, 25)
        Me.btnCopyAll.TabIndex = 37
        Me.btnCopyAll.Text = "Copy All"
        Me.ToolTip1.SetToolTip(Me.btnCopyAll, "Copy Case 1 to All cases except heading data")
        Me.btnCopyAll.UseVisualStyleBackColor = False
        '
        'btnSilentRun
        '
        Me.btnSilentRun.BackColor = System.Drawing.SystemColors.Control
        Me.btnSilentRun.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnSilentRun.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnSilentRun.Location = New System.Drawing.Point(514, 508)
        Me.btnSilentRun.Name = "btnSilentRun"
        Me.btnSilentRun.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnSilentRun.Size = New System.Drawing.Size(98, 25)
        Me.btnSilentRun.TabIndex = 26
        Me.btnSilentRun.Text = "&Silent Run"
        Me.ToolTip1.SetToolTip(Me.btnSilentRun, "Launch DOS window to run all AQWA cases")
        Me.btnSilentRun.UseVisualStyleBackColor = False
        '
        'btnCreateFiles
        '
        Me.btnCreateFiles.BackColor = System.Drawing.SystemColors.Control
        Me.btnCreateFiles.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCreateFiles.Enabled = False
        Me.btnCreateFiles.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCreateFiles.Location = New System.Drawing.Point(13, 508)
        Me.btnCreateFiles.Name = "btnCreateFiles"
        Me.btnCreateFiles.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCreateFiles.Size = New System.Drawing.Size(117, 25)
        Me.btnCreateFiles.TabIndex = 25
        Me.btnCreateFiles.Text = "&Create Case Files"
        Me.ToolTip1.SetToolTip(Me.btnCreateFiles, "Prepare AQWA input files for all environmental headings")
        Me.btnCreateFiles.UseVisualStyleBackColor = False
        '
        'btnOldAQWA
        '
        Me.btnOldAQWA.BackColor = System.Drawing.SystemColors.Control
        Me.btnOldAQWA.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnOldAQWA.Enabled = False
        Me.btnOldAQWA.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnOldAQWA.Location = New System.Drawing.Point(133, 508)
        Me.btnOldAQWA.Name = "btnOldAQWA"
        Me.btnOldAQWA.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnOldAQWA.Size = New System.Drawing.Size(140, 25)
        Me.btnOldAQWA.TabIndex = 24
        Me.btnOldAQWA.Text = "Switch to Old AQWA"
        Me.ToolTip1.SetToolTip(Me.btnOldAQWA, "Set your system ready to run OLD AQWA")
        Me.btnOldAQWA.UseVisualStyleBackColor = False
        Me.btnOldAQWA.Visible = False
        '
        'btnNewAQWA
        '
        Me.btnNewAQWA.BackColor = System.Drawing.SystemColors.Control
        Me.btnNewAQWA.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNewAQWA.Enabled = False
        Me.btnNewAQWA.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNewAQWA.Location = New System.Drawing.Point(275, 508)
        Me.btnNewAQWA.Name = "btnNewAQWA"
        Me.btnNewAQWA.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNewAQWA.Size = New System.Drawing.Size(141, 25)
        Me.btnNewAQWA.TabIndex = 23
        Me.btnNewAQWA.Text = "Restore to New AQWA"
        Me.ToolTip1.SetToolTip(Me.btnNewAQWA, "Set your system ready to run NEW AQWA")
        Me.btnNewAQWA.UseVisualStyleBackColor = False
        Me.btnNewAQWA.Visible = False
        '
        'btnPostProcess
        '
        Me.btnPostProcess.BackColor = System.Drawing.SystemColors.Control
        Me.btnPostProcess.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnPostProcess.Enabled = False
        Me.btnPostProcess.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnPostProcess.Location = New System.Drawing.Point(614, 506)
        Me.btnPostProcess.Name = "btnPostProcess"
        Me.btnPostProcess.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnPostProcess.Size = New System.Drawing.Size(98, 25)
        Me.btnPostProcess.TabIndex = 19
        Me.btnPostProcess.Text = "&Get Results"
        Me.ToolTip1.SetToolTip(Me.btnPostProcess, "Launch Excel to display AQWA results.")
        Me.btnPostProcess.UseVisualStyleBackColor = False
        '
        'btnBrowseDir
        '
        Me.btnBrowseDir.BackColor = System.Drawing.SystemColors.Control
        Me.btnBrowseDir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnBrowseDir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnBrowseDir.Location = New System.Drawing.Point(742, 98)
        Me.btnBrowseDir.Name = "btnBrowseDir"
        Me.btnBrowseDir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnBrowseDir.Size = New System.Drawing.Size(38, 24)
        Me.btnBrowseDir.TabIndex = 15
        Me.btnBrowseDir.Text = ">>"
        Me.ToolTip1.SetToolTip(Me.btnBrowseDir, "Browse Directory to run and write your AQWA results")
        Me.btnBrowseDir.UseVisualStyleBackColor = False
        '
        'btnBrowseBaseDir
        '
        Me.btnBrowseBaseDir.BackColor = System.Drawing.SystemColors.Control
        Me.btnBrowseBaseDir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnBrowseBaseDir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnBrowseBaseDir.Location = New System.Drawing.Point(742, 39)
        Me.btnBrowseBaseDir.Name = "btnBrowseBaseDir"
        Me.btnBrowseBaseDir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnBrowseBaseDir.Size = New System.Drawing.Size(38, 24)
        Me.btnBrowseBaseDir.TabIndex = 6
        Me.btnBrowseBaseDir.Text = ">>"
        Me.ToolTip1.SetToolTip(Me.btnBrowseBaseDir, "Browse for directory of ABRUN without environment")
        Me.btnBrowseBaseDir.UseVisualStyleBackColor = False
        '
        'btnRun
        '
        Me.btnRun.BackColor = System.Drawing.SystemColors.Control
        Me.btnRun.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnRun.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnRun.Location = New System.Drawing.Point(419, 508)
        Me.btnRun.Name = "btnRun"
        Me.btnRun.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnRun.Size = New System.Drawing.Size(89, 25)
        Me.btnRun.TabIndex = 2
        Me.btnRun.Text = "&Run Cases"
        Me.ToolTip1.SetToolTip(Me.btnRun, "Launch DOS window to run all AQWA cases")
        Me.btnRun.UseVisualStyleBackColor = False
        '
        'chkUDEF
        '
        Me.chkUDEF.BackColor = System.Drawing.SystemColors.Control
        Me.chkUDEF.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkUDEF.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkUDEF.Location = New System.Drawing.Point(280, 512)
        Me.chkUDEF.Name = "chkUDEF"
        Me.chkUDEF.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkUDEF.Size = New System.Drawing.Size(121, 17)
        Me.chkUDEF.TabIndex = 40
        Me.chkUDEF.Text = "Swell"
        Me.chkUDEF.UseVisualStyleBackColor = False
        '
        'chkVersion
        '
        Me.chkVersion.BackColor = System.Drawing.SystemColors.Control
        Me.chkVersion.Checked = True
        Me.chkVersion.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkVersion.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkVersion.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkVersion.Location = New System.Drawing.Point(152, 512)
        Me.chkVersion.Name = "chkVersion"
        Me.chkVersion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkVersion.Size = New System.Drawing.Size(121, 17)
        Me.chkVersion.TabIndex = 39
        Me.chkVersion.Text = "AQWA 15.0"
        Me.chkVersion.UseVisualStyleBackColor = False
        '
        'chkCollinear
        '
        Me.chkCollinear.BackColor = System.Drawing.SystemColors.Control
        Me.chkCollinear.Checked = True
        Me.chkCollinear.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCollinear.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCollinear.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCollinear.Location = New System.Drawing.Point(352, 232)
        Me.chkCollinear.Name = "chkCollinear"
        Me.chkCollinear.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCollinear.Size = New System.Drawing.Size(73, 17)
        Me.chkCollinear.TabIndex = 38
        Me.chkCollinear.Text = "Collinear"
        Me.chkCollinear.UseVisualStyleBackColor = False
        '
        'txtAQWAdir
        '
        Me.txtAQWAdir.AcceptsReturn = True
        Me.txtAQWAdir.BackColor = System.Drawing.SystemColors.Window
        Me.txtAQWAdir.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAQWAdir.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAQWAdir.Location = New System.Drawing.Point(104, 232)
        Me.txtAQWAdir.Name = "txtAQWAdir"
        Me.txtAQWAdir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAQWAdir.Size = New System.Drawing.Size(161, 20)
        Me.txtAQWAdir.TabIndex = 35
        Me.txtAQWAdir.Text = "aqwa /nowind"
        '
        'btnIntactFile
        '
        Me.btnIntactFile.BackColor = System.Drawing.SystemColors.Control
        Me.btnIntactFile.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnIntactFile.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnIntactFile.Location = New System.Drawing.Point(520, 232)
        Me.btnIntactFile.Name = "btnIntactFile"
        Me.btnIntactFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnIntactFile.Size = New System.Drawing.Size(189, 26)
        Me.btnIntactFile.TabIndex = 34
        Me.btnIntactFile.Text = "Get Critical Lines from Intact XLS file"
        Me.btnIntactFile.UseVisualStyleBackColor = False
        '
        'fraDamageLineType
        '
        Me.fraDamageLineType.BackColor = System.Drawing.SystemColors.Control
        Me.fraDamageLineType.Controls.Add(Me._optDamageLineType_1)
        Me.fraDamageLineType.Controls.Add(Me._optDamageLineType_0)
        Me.fraDamageLineType.Enabled = False
        Me.fraDamageLineType.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraDamageLineType.Location = New System.Drawing.Point(640, 154)
        Me.fraDamageLineType.Name = "fraDamageLineType"
        Me.fraDamageLineType.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraDamageLineType.Size = New System.Drawing.Size(110, 73)
        Me.fraDamageLineType.TabIndex = 31
        Me.fraDamageLineType.TabStop = False
        Me.fraDamageLineType.Text = "Damage"
        '
        '_optDamageLineType_1
        '
        Me._optDamageLineType_1.BackColor = System.Drawing.SystemColors.Control
        Me._optDamageLineType_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optDamageLineType_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optDamageLineType.SetIndex(Me._optDamageLineType_1, CType(1, Short))
        Me._optDamageLineType_1.Location = New System.Drawing.Point(15, 42)
        Me._optDamageLineType_1.Name = "_optDamageLineType_1"
        Me._optDamageLineType_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optDamageLineType_1.Size = New System.Drawing.Size(85, 18)
        Me._optDamageLineType_1.TabIndex = 33
        Me._optDamageLineType_1.TabStop = True
        Me._optDamageLineType_1.Text = "2nd-Loaded"
        Me._optDamageLineType_1.UseVisualStyleBackColor = False
        '
        '_optDamageLineType_0
        '
        Me._optDamageLineType_0.BackColor = System.Drawing.SystemColors.Control
        Me._optDamageLineType_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optDamageLineType_0.Enabled = False
        Me._optDamageLineType_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optDamageLineType.SetIndex(Me._optDamageLineType_0, CType(0, Short))
        Me._optDamageLineType_0.Location = New System.Drawing.Point(15, 20)
        Me._optDamageLineType_0.Name = "_optDamageLineType_0"
        Me._optDamageLineType_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optDamageLineType_0.Size = New System.Drawing.Size(90, 18)
        Me._optDamageLineType_0.TabIndex = 32
        Me._optDamageLineType_0.TabStop = True
        Me._optDamageLineType_0.Text = "Most-Loaded"
        Me._optDamageLineType_0.UseVisualStyleBackColor = False
        '
        'fraWind
        '
        Me.fraWind.BackColor = System.Drawing.SystemColors.Control
        Me.fraWind.Controls.Add(Me._optWind_1)
        Me.fraWind.Controls.Add(Me._optWind_0)
        Me.fraWind.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraWind.Location = New System.Drawing.Point(433, 154)
        Me.fraWind.Name = "fraWind"
        Me.fraWind.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraWind.Size = New System.Drawing.Size(75, 74)
        Me.fraWind.TabIndex = 28
        Me.fraWind.TabStop = False
        Me.fraWind.Text = "Wind"
        '
        '_optWind_1
        '
        Me._optWind_1.BackColor = System.Drawing.SystemColors.Control
        Me._optWind_1.Checked = True
        Me._optWind_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optWind_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optWind.SetIndex(Me._optWind_1, CType(1, Short))
        Me._optWind_1.Location = New System.Drawing.Point(9, 42)
        Me._optWind_1.Name = "_optWind_1"
        Me._optWind_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optWind_1.Size = New System.Drawing.Size(54, 20)
        Me._optWind_1.TabIndex = 30
        Me._optWind_1.TabStop = True
        Me._optWind_1.Text = "1- min"
        Me._optWind_1.UseVisualStyleBackColor = False
        '
        '_optWind_0
        '
        Me._optWind_0.BackColor = System.Drawing.SystemColors.Control
        Me._optWind_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optWind_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optWind.SetIndex(Me._optWind_0, CType(0, Short))
        Me._optWind_0.Location = New System.Drawing.Point(9, 18)
        Me._optWind_0.Name = "_optWind_0"
        Me._optWind_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optWind_0.Size = New System.Drawing.Size(44, 20)
        Me._optWind_0.TabIndex = 29
        Me._optWind_0.TabStop = True
        Me._optWind_0.Text = "1-hr"
        Me._optWind_0.UseVisualStyleBackColor = False
        '
        'chkPFLH
        '
        Me.chkPFLH.BackColor = System.Drawing.SystemColors.Control
        Me.chkPFLH.Checked = True
        Me.chkPFLH.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkPFLH.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkPFLH.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkPFLH.Location = New System.Drawing.Point(280, 232)
        Me.chkPFLH.Name = "chkPFLH"
        Me.chkPFLH.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkPFLH.Size = New System.Drawing.Size(65, 19)
        Me.chkPFLH.TabIndex = 27
        Me.chkPFLH.Text = "PFLH"
        Me.chkPFLH.UseVisualStyleBackColor = False
        '
        'fraUnits
        '
        Me.fraUnits.BackColor = System.Drawing.SystemColors.Control
        Me.fraUnits.Controls.Add(Me._optAQWAUnits_1)
        Me.fraUnits.Controls.Add(Me._optAQWAUnits_0)
        Me.fraUnits.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraUnits.Location = New System.Drawing.Point(338, 154)
        Me.fraUnits.Name = "fraUnits"
        Me.fraUnits.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraUnits.Size = New System.Drawing.Size(87, 74)
        Me.fraUnits.TabIndex = 20
        Me.fraUnits.TabStop = False
        Me.fraUnits.Text = "AQWA Units"
        '
        '_optAQWAUnits_1
        '
        Me._optAQWAUnits_1.BackColor = System.Drawing.SystemColors.Control
        Me._optAQWAUnits_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optAQWAUnits_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optAQWAUnits.SetIndex(Me._optAQWAUnits_1, CType(1, Short))
        Me._optAQWAUnits_1.Location = New System.Drawing.Point(23, 43)
        Me._optAQWAUnits_1.Name = "_optAQWAUnits_1"
        Me._optAQWAUnits_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optAQWAUnits_1.Size = New System.Drawing.Size(51, 18)
        Me._optAQWAUnits_1.TabIndex = 22
        Me._optAQWAUnits_1.TabStop = True
        Me._optAQWAUnits_1.Text = "Metric"
        Me._optAQWAUnits_1.UseVisualStyleBackColor = False
        '
        '_optAQWAUnits_0
        '
        Me._optAQWAUnits_0.BackColor = System.Drawing.SystemColors.Control
        Me._optAQWAUnits_0.Checked = True
        Me._optAQWAUnits_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optAQWAUnits_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optAQWAUnits.SetIndex(Me._optAQWAUnits_0, CType(0, Short))
        Me._optAQWAUnits_0.Location = New System.Drawing.Point(23, 21)
        Me._optAQWAUnits_0.Name = "_optAQWAUnits_0"
        Me._optAQWAUnits_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optAQWAUnits_0.Size = New System.Drawing.Size(54, 18)
        Me._optAQWAUnits_0.TabIndex = 21
        Me._optAQWAUnits_0.TabStop = True
        Me._optAQWAUnits_0.Text = "US"
        Me._optAQWAUnits_0.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me._optMoorState_0)
        Me.Frame1.Controls.Add(Me._optMoorState_1)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(514, 154)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(116, 73)
        Me.Frame1.TabIndex = 16
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Mooring System"
        '
        '_optMoorState_0
        '
        Me._optMoorState_0.BackColor = System.Drawing.SystemColors.Control
        Me._optMoorState_0.Checked = True
        Me._optMoorState_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoorState_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMoorState.SetIndex(Me._optMoorState_0, CType(0, Short))
        Me._optMoorState_0.Location = New System.Drawing.Point(14, 21)
        Me._optMoorState_0.Name = "_optMoorState_0"
        Me._optMoorState_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoorState_0.Size = New System.Drawing.Size(64, 18)
        Me._optMoorState_0.TabIndex = 18
        Me._optMoorState_0.TabStop = True
        Me._optMoorState_0.Text = "Intact"
        Me._optMoorState_0.UseVisualStyleBackColor = False
        '
        '_optMoorState_1
        '
        Me._optMoorState_1.BackColor = System.Drawing.SystemColors.Control
        Me._optMoorState_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoorState_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMoorState.SetIndex(Me._optMoorState_1, CType(1, Short))
        Me._optMoorState_1.Location = New System.Drawing.Point(14, 43)
        Me._optMoorState_1.Name = "_optMoorState_1"
        Me._optMoorState_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoorState_1.Size = New System.Drawing.Size(72, 18)
        Me._optMoorState_1.TabIndex = 17
        Me._optMoorState_1.TabStop = True
        Me._optMoorState_1.Text = "Damaged"
        Me._optMoorState_1.UseVisualStyleBackColor = False
        '
        'txtMetoceanName
        '
        Me.txtMetoceanName.AcceptsReturn = True
        Me.txtMetoceanName.BackColor = System.Drawing.SystemColors.Window
        Me.txtMetoceanName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMetoceanName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMetoceanName.Location = New System.Drawing.Point(122, 154)
        Me.txtMetoceanName.Name = "txtMetoceanName"
        Me.txtMetoceanName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMetoceanName.Size = New System.Drawing.Size(210, 20)
        Me.txtMetoceanName.TabIndex = 11
        '
        'btnSaveMatrix
        '
        Me.btnSaveMatrix.BackColor = System.Drawing.SystemColors.Control
        Me.btnSaveMatrix.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnSaveMatrix.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnSaveMatrix.Location = New System.Drawing.Point(250, 186)
        Me.btnSaveMatrix.Name = "btnSaveMatrix"
        Me.btnSaveMatrix.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnSaveMatrix.Size = New System.Drawing.Size(82, 25)
        Me.btnSaveMatrix.TabIndex = 10
        Me.btnSaveMatrix.Text = "&Save Matrix"
        Me.btnSaveMatrix.UseVisualStyleBackColor = False
        '
        'btnOpenMatrix
        '
        Me.btnOpenMatrix.BackColor = System.Drawing.SystemColors.Control
        Me.btnOpenMatrix.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnOpenMatrix.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnOpenMatrix.Location = New System.Drawing.Point(162, 186)
        Me.btnOpenMatrix.Name = "btnOpenMatrix"
        Me.btnOpenMatrix.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnOpenMatrix.Size = New System.Drawing.Size(82, 25)
        Me.btnOpenMatrix.TabIndex = 9
        Me.btnOpenMatrix.Text = "&Open Matrix"
        Me.btnOpenMatrix.UseVisualStyleBackColor = False
        '
        'txtNumCases
        '
        Me.txtNumCases.AcceptsReturn = True
        Me.txtNumCases.BackColor = System.Drawing.SystemColors.Window
        Me.txtNumCases.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNumCases.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNumCases.Location = New System.Drawing.Point(122, 189)
        Me.txtNumCases.Name = "txtNumCases"
        Me.txtNumCases.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNumCases.Size = New System.Drawing.Size(34, 20)
        Me.txtNumCases.TabIndex = 0
        Me.txtNumCases.Text = "1"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(8, 232)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(89, 17)
        Me.Label1.TabIndex = 36
        Me.Label1.Text = " AQWA path:"
        '
        'lblTargetDir
        '
        Me.lblTargetDir.BackColor = System.Drawing.SystemColors.Control
        Me.lblTargetDir.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTargetDir.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTargetDir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTargetDir.Location = New System.Drawing.Point(122, 98)
        Me.lblTargetDir.Name = "lblTargetDir"
        Me.lblTargetDir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTargetDir.Size = New System.Drawing.Size(596, 34)
        Me.lblTargetDir.TabIndex = 14
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(10, 99)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(120, 23)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "Target Directory:"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(9, 158)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(92, 17)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Metocean Name:"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(6, 39)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(110, 24)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Base File Directory:"
        '
        'lblBaseDir
        '
        Me.lblBaseDir.BackColor = System.Drawing.SystemColors.Control
        Me.lblBaseDir.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblBaseDir.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblBaseDir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblBaseDir.Location = New System.Drawing.Point(122, 39)
        Me.lblBaseDir.Name = "lblBaseDir"
        Me.lblBaseDir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblBaseDir.Size = New System.Drawing.Size(596, 35)
        Me.lblBaseDir.TabIndex = 7
        '
        'lblNumCases
        '
        Me.lblNumCases.BackColor = System.Drawing.SystemColors.Control
        Me.lblNumCases.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblNumCases.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblNumCases.Location = New System.Drawing.Point(9, 192)
        Me.lblNumCases.Name = "lblNumCases"
        Me.lblNumCases.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblNumCases.Size = New System.Drawing.Size(92, 17)
        Me.lblNumCases.TabIndex = 1
        Me.lblNumCases.Text = "Number of Cases:"
        '
        'optAQWAUnits
        '
        '
        'optDamageLineType
        '
        '
        'optMoorState
        '
        '
        'grdMatrix
        '
        Me.grdMatrix.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdMatrix.Location = New System.Drawing.Point(18, 271)
        Me.grdMatrix.Name = "grdMatrix"
        Me.grdMatrix.Size = New System.Drawing.Size(700, 229)
        Me.grdMatrix.TabIndex = 42
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(994, 544)
        Me.Controls.Add(Me.grdMatrix)
        Me.Controls.Add(Me.chkUDEF)
        Me.Controls.Add(Me.chkVersion)
        Me.Controls.Add(Me.chkCollinear)
        Me.Controls.Add(Me.btnCopyAll)
        Me.Controls.Add(Me.txtAQWAdir)
        Me.Controls.Add(Me.btnIntactFile)
        Me.Controls.Add(Me.fraDamageLineType)
        Me.Controls.Add(Me.fraWind)
        Me.Controls.Add(Me.chkPFLH)
        Me.Controls.Add(Me.btnSilentRun)
        Me.Controls.Add(Me.btnCreateFiles)
        Me.Controls.Add(Me.btnOldAQWA)
        Me.Controls.Add(Me.btnNewAQWA)
        Me.Controls.Add(Me.fraUnits)
        Me.Controls.Add(Me.btnPostProcess)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.btnBrowseDir)
        Me.Controls.Add(Me.txtMetoceanName)
        Me.Controls.Add(Me.btnSaveMatrix)
        Me.Controls.Add(Me.btnOpenMatrix)
        Me.Controls.Add(Me.btnBrowseBaseDir)
        Me.Controls.Add(Me.btnRun)
        Me.Controls.Add(Me.txtNumCases)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblTargetDir)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.lblBaseDir)
        Me.Controls.Add(Me.lblNumCases)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(10, 36)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmMain"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "AQWA Batch Runner - David Tein Consulting Engineers, Ltd."
        Me.fraDamageLineType.ResumeLayout(False)
        Me.fraWind.ResumeLayout(False)
        Me.fraUnits.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        CType(Me.optAQWAUnits, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optDamageLineType, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optMoorState, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optWind, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdMatrix, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents grdMatrix As DataGridView
#End Region
End Class