<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmAnalysesB
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
	Public WithEvents txtNumLinesPerGroup As System.Windows.Forms.TextBox
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents fraMoorGroup As System.Windows.Forms.GroupBox
	Public WithEvents txtSubTitle As System.Windows.Forms.TextBox
	Public WithEvents txtReportTitle As System.Windows.Forms.TextBox
	Public WithEvents txtInitYaw As System.Windows.Forms.TextBox
	Public WithEvents txtInitSway As System.Windows.Forms.TextBox
	Public WithEvents txtInitSurge As System.Windows.Forms.TextBox
	Public WithEvents _lblInitPos_2 As System.Windows.Forms.Label
	Public WithEvents _lblInitPos_1 As System.Windows.Forms.Label
	Public WithEvents _lblInitPos_0 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_4 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_3 As System.Windows.Forms.Label
	Public WithEvents _lblAngleUnit_3 As System.Windows.Forms.Label
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents btnReport As System.Windows.Forms.Button
	Public WithEvents lblLocation As System.Windows.Forms.Label
	Public WithEvents lblClientName As System.Windows.Forms.Label
	Public WithEvents fraReport As System.Windows.Forms.GroupBox
	Public dlgFileOpen As System.Windows.Forms.OpenFileDialog
	Public dlgFileSave As System.Windows.Forms.SaveFileDialog
	Public WithEvents btnBrowse As System.Windows.Forms.Button
	Public WithEvents txtFile As System.Windows.Forms.TextBox
	Public WithEvents fraFile As System.Windows.Forms.GroupBox
	Public WithEvents _txtHeadings_2 As System.Windows.Forms.TextBox
	Public WithEvents _txtHeadings_1 As System.Windows.Forms.TextBox
	Public WithEvents _txtHeadings_0 As System.Windows.Forms.TextBox
	Public WithEvents _lblAngleUnit_2 As System.Windows.Forms.Label
	Public WithEvents _lblAngleUnit_1 As System.Windows.Forms.Label
	Public WithEvents _lblHeadings_3 As System.Windows.Forms.Label
	Public WithEvents _lblHeadings_2 As System.Windows.Forms.Label
	Public WithEvents _lblHeadings_1 As System.Windows.Forms.Label
	Public WithEvents _lblHeadings_0 As System.Windows.Forms.Label
	Public WithEvents fraHeadings As System.Windows.Forms.GroupBox
	Public WithEvents btnAnalysis As System.Windows.Forms.Button
	Public WithEvents btnCancel As System.Windows.Forms.Button
	Public WithEvents _txtVslSt_0 As System.Windows.Forms.TextBox
	Public WithEvents _txtVslSt_1 As System.Windows.Forms.TextBox
	Public WithEvents _txtVslSt_2 As System.Windows.Forms.TextBox
	Public WithEvents _txtVslSt_3 As System.Windows.Forms.TextBox
	Public WithEvents _lblLengthUnit_2 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_1 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_4 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_0 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_0 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_1 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_3 As System.Windows.Forms.Label
	Public WithEvents _lblAngleUnit_0 As System.Windows.Forms.Label
	Public WithEvents fraVesselLoc As System.Windows.Forms.GroupBox
	Public WithEvents txtEnvironment As System.Windows.Forms.TextBox
	Public WithEvents btnEnvironment As System.Windows.Forms.Button
	Public WithEvents fraEnvironment As System.Windows.Forms.GroupBox
	Public WithEvents _lblForceUnit_0 As System.Windows.Forms.Label
	Public WithEvents _lblVelUnit_0 As System.Windows.Forms.Label
	Public WithEvents lblAngleUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblForceUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblHeadings As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblInitPos As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblLengthUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblVelUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblVslSt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents txtHeadings As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	Public WithEvents txtVslSt As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAnalysesB))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.fraMoorGroup = New System.Windows.Forms.GroupBox
		Me.txtNumLinesPerGroup = New System.Windows.Forms.TextBox
		Me.Label1 = New System.Windows.Forms.Label
		Me.fraReport = New System.Windows.Forms.GroupBox
		Me.txtSubTitle = New System.Windows.Forms.TextBox
		Me.txtReportTitle = New System.Windows.Forms.TextBox
		Me.Frame2 = New System.Windows.Forms.GroupBox
		Me.txtInitYaw = New System.Windows.Forms.TextBox
		Me.txtInitSway = New System.Windows.Forms.TextBox
		Me.txtInitSurge = New System.Windows.Forms.TextBox
		Me._lblInitPos_2 = New System.Windows.Forms.Label
		Me._lblInitPos_1 = New System.Windows.Forms.Label
		Me._lblInitPos_0 = New System.Windows.Forms.Label
		Me._lblLengthUnit_4 = New System.Windows.Forms.Label
		Me._lblLengthUnit_3 = New System.Windows.Forms.Label
		Me._lblAngleUnit_3 = New System.Windows.Forms.Label
		Me.btnReport = New System.Windows.Forms.Button
		Me.lblLocation = New System.Windows.Forms.Label
		Me.lblClientName = New System.Windows.Forms.Label
		Me.dlgFileOpen = New System.Windows.Forms.OpenFileDialog
		Me.dlgFileSave = New System.Windows.Forms.SaveFileDialog
		Me.fraFile = New System.Windows.Forms.GroupBox
		Me.btnBrowse = New System.Windows.Forms.Button
		Me.txtFile = New System.Windows.Forms.TextBox
		Me.fraHeadings = New System.Windows.Forms.GroupBox
		Me._txtHeadings_2 = New System.Windows.Forms.TextBox
		Me._txtHeadings_1 = New System.Windows.Forms.TextBox
		Me._txtHeadings_0 = New System.Windows.Forms.TextBox
		Me._lblAngleUnit_2 = New System.Windows.Forms.Label
		Me._lblAngleUnit_1 = New System.Windows.Forms.Label
		Me._lblHeadings_3 = New System.Windows.Forms.Label
		Me._lblHeadings_2 = New System.Windows.Forms.Label
		Me._lblHeadings_1 = New System.Windows.Forms.Label
		Me._lblHeadings_0 = New System.Windows.Forms.Label
		Me.btnAnalysis = New System.Windows.Forms.Button
		Me.btnCancel = New System.Windows.Forms.Button
		Me.fraVesselLoc = New System.Windows.Forms.GroupBox
		Me._txtVslSt_0 = New System.Windows.Forms.TextBox
		Me._txtVslSt_1 = New System.Windows.Forms.TextBox
		Me._txtVslSt_2 = New System.Windows.Forms.TextBox
		Me._txtVslSt_3 = New System.Windows.Forms.TextBox
		Me._lblLengthUnit_2 = New System.Windows.Forms.Label
		Me._lblLengthUnit_1 = New System.Windows.Forms.Label
		Me._lblVslSt_4 = New System.Windows.Forms.Label
		Me._lblLengthUnit_0 = New System.Windows.Forms.Label
		Me._lblVslSt_0 = New System.Windows.Forms.Label
		Me._lblVslSt_1 = New System.Windows.Forms.Label
		Me._lblVslSt_3 = New System.Windows.Forms.Label
		Me._lblAngleUnit_0 = New System.Windows.Forms.Label
		Me.fraEnvironment = New System.Windows.Forms.GroupBox
		Me.txtEnvironment = New System.Windows.Forms.TextBox
		Me.btnEnvironment = New System.Windows.Forms.Button
		Me._lblForceUnit_0 = New System.Windows.Forms.Label
		Me._lblVelUnit_0 = New System.Windows.Forms.Label
		Me.lblAngleUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.lblForceUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.lblHeadings = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.lblInitPos = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.lblLengthUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.lblVelUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.lblVslSt = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.txtHeadings = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(components)
		Me.txtVslSt = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(components)
		Me.fraMoorGroup.SuspendLayout()
		Me.fraReport.SuspendLayout()
		Me.Frame2.SuspendLayout()
		Me.fraFile.SuspendLayout()
		Me.fraHeadings.SuspendLayout()
		Me.fraVesselLoc.SuspendLayout()
		Me.fraEnvironment.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.lblAngleUnit, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lblForceUnit, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lblHeadings, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lblInitPos, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lblLengthUnit, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lblVelUnit, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lblVslSt, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.txtHeadings, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.txtVslSt, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.Text = "Mooring Analyses with Multiple Environmental Headings"
		Me.ClientSize = New System.Drawing.Size(609, 371)
		Me.Location = New System.Drawing.Point(4, 23)
		Me.Icon = CType(resources.GetObject("frmAnalysesB.Icon"), System.Drawing.Icon)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmAnalysesB"
		Me.fraMoorGroup.Text = "Grouping Pattern:"
		Me.fraMoorGroup.Size = New System.Drawing.Size(149, 59)
		Me.fraMoorGroup.Location = New System.Drawing.Point(376, 299)
		Me.fraMoorGroup.TabIndex = 47
		Me.fraMoorGroup.BackColor = System.Drawing.SystemColors.Control
		Me.fraMoorGroup.Enabled = True
		Me.fraMoorGroup.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraMoorGroup.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraMoorGroup.Visible = True
		Me.fraMoorGroup.Name = "fraMoorGroup"
		Me.txtNumLinesPerGroup.AutoSize = False
		Me.txtNumLinesPerGroup.Size = New System.Drawing.Size(28, 20)
		Me.txtNumLinesPerGroup.Location = New System.Drawing.Point(106, 23)
		Me.txtNumLinesPerGroup.TabIndex = 48
		Me.txtNumLinesPerGroup.AcceptsReturn = True
		Me.txtNumLinesPerGroup.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtNumLinesPerGroup.BackColor = System.Drawing.SystemColors.Window
		Me.txtNumLinesPerGroup.CausesValidation = True
		Me.txtNumLinesPerGroup.Enabled = True
		Me.txtNumLinesPerGroup.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtNumLinesPerGroup.HideSelection = True
		Me.txtNumLinesPerGroup.ReadOnly = False
		Me.txtNumLinesPerGroup.Maxlength = 0
		Me.txtNumLinesPerGroup.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtNumLinesPerGroup.MultiLine = False
		Me.txtNumLinesPerGroup.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtNumLinesPerGroup.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtNumLinesPerGroup.TabStop = True
		Me.txtNumLinesPerGroup.Visible = True
		Me.txtNumLinesPerGroup.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtNumLinesPerGroup.Name = "txtNumLinesPerGroup"
		Me.Label1.Text = "Number of Lines per Group:"
		Me.Label1.Size = New System.Drawing.Size(78, 33)
		Me.Label1.Location = New System.Drawing.Point(13, 19)
		Me.Label1.TabIndex = 49
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
		Me.fraReport.Text = "Reporting"
		Me.fraReport.Size = New System.Drawing.Size(289, 209)
		Me.fraReport.Location = New System.Drawing.Point(235, 78)
		Me.fraReport.TabIndex = 31
		Me.fraReport.BackColor = System.Drawing.SystemColors.Control
		Me.fraReport.Enabled = True
		Me.fraReport.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraReport.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraReport.Visible = True
		Me.fraReport.Name = "fraReport"
		Me.txtSubTitle.AutoSize = False
		Me.txtSubTitle.Size = New System.Drawing.Size(199, 19)
		Me.txtSubTitle.Location = New System.Drawing.Point(73, 44)
		Me.txtSubTitle.TabIndex = 44
		Me.txtSubTitle.TabStop = False
		Me.txtSubTitle.AcceptsReturn = True
		Me.txtSubTitle.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtSubTitle.BackColor = System.Drawing.SystemColors.Window
		Me.txtSubTitle.CausesValidation = True
		Me.txtSubTitle.Enabled = True
		Me.txtSubTitle.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtSubTitle.HideSelection = True
		Me.txtSubTitle.ReadOnly = False
		Me.txtSubTitle.Maxlength = 0
		Me.txtSubTitle.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtSubTitle.MultiLine = False
		Me.txtSubTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtSubTitle.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtSubTitle.Visible = True
		Me.txtSubTitle.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtSubTitle.Name = "txtSubTitle"
		Me.txtReportTitle.AutoSize = False
		Me.txtReportTitle.Size = New System.Drawing.Size(199, 19)
		Me.txtReportTitle.Location = New System.Drawing.Point(73, 19)
		Me.txtReportTitle.TabIndex = 43
		Me.txtReportTitle.TabStop = False
		Me.txtReportTitle.Text = "Mooring Analysis Results"
		Me.txtReportTitle.AcceptsReturn = True
		Me.txtReportTitle.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtReportTitle.BackColor = System.Drawing.SystemColors.Window
		Me.txtReportTitle.CausesValidation = True
		Me.txtReportTitle.Enabled = True
		Me.txtReportTitle.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtReportTitle.HideSelection = True
		Me.txtReportTitle.ReadOnly = False
		Me.txtReportTitle.Maxlength = 0
		Me.txtReportTitle.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtReportTitle.MultiLine = False
		Me.txtReportTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtReportTitle.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtReportTitle.Visible = True
		Me.txtReportTitle.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtReportTitle.Name = "txtReportTitle"
		Me.Frame2.Text = "Base Vessel Motion:"
		Me.Frame2.Size = New System.Drawing.Size(174, 108)
		Me.Frame2.Location = New System.Drawing.Point(13, 79)
		Me.Frame2.TabIndex = 33
		Me.Frame2.BackColor = System.Drawing.SystemColors.Control
		Me.Frame2.Enabled = True
		Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame2.Visible = True
		Me.Frame2.Name = "Frame2"
		Me.txtInitYaw.AutoSize = False
		Me.txtInitYaw.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me.txtInitYaw.Size = New System.Drawing.Size(67, 19)
		Me.txtInitYaw.Location = New System.Drawing.Point(77, 74)
		Me.txtInitYaw.TabIndex = 36
		Me.txtInitYaw.TabStop = False
		Me.txtInitYaw.Text = "0"
		Me.txtInitYaw.AcceptsReturn = True
		Me.txtInitYaw.BackColor = System.Drawing.SystemColors.Window
		Me.txtInitYaw.CausesValidation = True
		Me.txtInitYaw.Enabled = True
		Me.txtInitYaw.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtInitYaw.HideSelection = True
		Me.txtInitYaw.ReadOnly = False
		Me.txtInitYaw.Maxlength = 0
		Me.txtInitYaw.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtInitYaw.MultiLine = False
		Me.txtInitYaw.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtInitYaw.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtInitYaw.Visible = True
		Me.txtInitYaw.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtInitYaw.Name = "txtInitYaw"
		Me.txtInitSway.AutoSize = False
		Me.txtInitSway.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me.txtInitSway.Size = New System.Drawing.Size(67, 19)
		Me.txtInitSway.Location = New System.Drawing.Point(77, 49)
		Me.txtInitSway.TabIndex = 35
		Me.txtInitSway.TabStop = False
		Me.txtInitSway.Text = "0"
		Me.txtInitSway.AcceptsReturn = True
		Me.txtInitSway.BackColor = System.Drawing.SystemColors.Window
		Me.txtInitSway.CausesValidation = True
		Me.txtInitSway.Enabled = True
		Me.txtInitSway.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtInitSway.HideSelection = True
		Me.txtInitSway.ReadOnly = False
		Me.txtInitSway.Maxlength = 0
		Me.txtInitSway.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtInitSway.MultiLine = False
		Me.txtInitSway.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtInitSway.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtInitSway.Visible = True
		Me.txtInitSway.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtInitSway.Name = "txtInitSway"
		Me.txtInitSurge.AutoSize = False
		Me.txtInitSurge.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me.txtInitSurge.Size = New System.Drawing.Size(67, 19)
		Me.txtInitSurge.Location = New System.Drawing.Point(77, 26)
		Me.txtInitSurge.TabIndex = 34
		Me.txtInitSurge.TabStop = False
		Me.txtInitSurge.Text = "0"
		Me.txtInitSurge.AcceptsReturn = True
		Me.txtInitSurge.BackColor = System.Drawing.SystemColors.Window
		Me.txtInitSurge.CausesValidation = True
		Me.txtInitSurge.Enabled = True
		Me.txtInitSurge.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtInitSurge.HideSelection = True
		Me.txtInitSurge.ReadOnly = False
		Me.txtInitSurge.Maxlength = 0
		Me.txtInitSurge.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtInitSurge.MultiLine = False
		Me.txtInitSurge.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtInitSurge.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtInitSurge.Visible = True
		Me.txtInitSurge.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtInitSurge.Name = "txtInitSurge"
		Me._lblInitPos_2.Text = "Yaw:"
		Me._lblInitPos_2.Size = New System.Drawing.Size(37, 17)
		Me._lblInitPos_2.Location = New System.Drawing.Point(35, 75)
		Me._lblInitPos_2.TabIndex = 42
		Me._lblInitPos_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblInitPos_2.BackColor = System.Drawing.SystemColors.Control
		Me._lblInitPos_2.Enabled = True
		Me._lblInitPos_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblInitPos_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblInitPos_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblInitPos_2.UseMnemonic = True
		Me._lblInitPos_2.Visible = True
		Me._lblInitPos_2.AutoSize = False
		Me._lblInitPos_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblInitPos_2.Name = "_lblInitPos_2"
		Me._lblInitPos_1.Text = "Sway:"
		Me._lblInitPos_1.Size = New System.Drawing.Size(37, 17)
		Me._lblInitPos_1.Location = New System.Drawing.Point(34, 51)
		Me._lblInitPos_1.TabIndex = 41
		Me._lblInitPos_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblInitPos_1.BackColor = System.Drawing.SystemColors.Control
		Me._lblInitPos_1.Enabled = True
		Me._lblInitPos_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblInitPos_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblInitPos_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblInitPos_1.UseMnemonic = True
		Me._lblInitPos_1.Visible = True
		Me._lblInitPos_1.AutoSize = False
		Me._lblInitPos_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblInitPos_1.Name = "_lblInitPos_1"
		Me._lblInitPos_0.Text = "Surge:"
		Me._lblInitPos_0.Size = New System.Drawing.Size(37, 17)
		Me._lblInitPos_0.Location = New System.Drawing.Point(35, 27)
		Me._lblInitPos_0.TabIndex = 40
		Me._lblInitPos_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblInitPos_0.BackColor = System.Drawing.SystemColors.Control
		Me._lblInitPos_0.Enabled = True
		Me._lblInitPos_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblInitPos_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblInitPos_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblInitPos_0.UseMnemonic = True
		Me._lblInitPos_0.Visible = True
		Me._lblInitPos_0.AutoSize = False
		Me._lblInitPos_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblInitPos_0.Name = "_lblInitPos_0"
		Me._lblLengthUnit_4.Text = "ft"
		Me._lblLengthUnit_4.Size = New System.Drawing.Size(17, 17)
		Me._lblLengthUnit_4.Location = New System.Drawing.Point(149, 48)
		Me._lblLengthUnit_4.TabIndex = 39
		Me._lblLengthUnit_4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblLengthUnit_4.BackColor = System.Drawing.SystemColors.Control
		Me._lblLengthUnit_4.Enabled = True
		Me._lblLengthUnit_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblLengthUnit_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblLengthUnit_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblLengthUnit_4.UseMnemonic = True
		Me._lblLengthUnit_4.Visible = True
		Me._lblLengthUnit_4.AutoSize = False
		Me._lblLengthUnit_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblLengthUnit_4.Name = "_lblLengthUnit_4"
		Me._lblLengthUnit_3.Text = "ft"
		Me._lblLengthUnit_3.Size = New System.Drawing.Size(17, 17)
		Me._lblLengthUnit_3.Location = New System.Drawing.Point(149, 28)
		Me._lblLengthUnit_3.TabIndex = 38
		Me._lblLengthUnit_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblLengthUnit_3.BackColor = System.Drawing.SystemColors.Control
		Me._lblLengthUnit_3.Enabled = True
		Me._lblLengthUnit_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblLengthUnit_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblLengthUnit_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblLengthUnit_3.UseMnemonic = True
		Me._lblLengthUnit_3.Visible = True
		Me._lblLengthUnit_3.AutoSize = False
		Me._lblLengthUnit_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblLengthUnit_3.Name = "_lblLengthUnit_3"
		Me._lblAngleUnit_3.Text = "Invalid_string_refer_to_original_code"
		Me._lblAngleUnit_3.Size = New System.Drawing.Size(17, 17)
		Me._lblAngleUnit_3.Location = New System.Drawing.Point(148, 72)
		Me._lblAngleUnit_3.TabIndex = 37
		Me._lblAngleUnit_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblAngleUnit_3.BackColor = System.Drawing.SystemColors.Control
		Me._lblAngleUnit_3.Enabled = True
		Me._lblAngleUnit_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblAngleUnit_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblAngleUnit_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblAngleUnit_3.UseMnemonic = True
		Me._lblAngleUnit_3.Visible = True
		Me._lblAngleUnit_3.AutoSize = False
		Me._lblAngleUnit_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblAngleUnit_3.Name = "_lblAngleUnit_3"
		Me.btnReport.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnReport.Text = "&Report"
		Me.btnReport.Size = New System.Drawing.Size(70, 25)
		Me.btnReport.Location = New System.Drawing.Point(197, 84)
		Me.btnReport.TabIndex = 32
		Me.btnReport.BackColor = System.Drawing.SystemColors.Control
		Me.btnReport.CausesValidation = True
		Me.btnReport.Enabled = True
		Me.btnReport.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnReport.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnReport.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnReport.TabStop = True
		Me.btnReport.Name = "btnReport"
		Me.lblLocation.Text = "Metocean:"
		Me.lblLocation.Size = New System.Drawing.Size(62, 17)
		Me.lblLocation.Location = New System.Drawing.Point(13, 46)
		Me.lblLocation.TabIndex = 46
		Me.lblLocation.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblLocation.BackColor = System.Drawing.SystemColors.Control
		Me.lblLocation.Enabled = True
		Me.lblLocation.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblLocation.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblLocation.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblLocation.UseMnemonic = True
		Me.lblLocation.Visible = True
		Me.lblLocation.AutoSize = False
		Me.lblLocation.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblLocation.Name = "lblLocation"
		Me.lblClientName.Text = "Main Title:"
		Me.lblClientName.Size = New System.Drawing.Size(73, 17)
		Me.lblClientName.Location = New System.Drawing.Point(13, 22)
		Me.lblClientName.TabIndex = 45
		Me.lblClientName.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblClientName.BackColor = System.Drawing.SystemColors.Control
		Me.lblClientName.Enabled = True
		Me.lblClientName.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblClientName.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblClientName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblClientName.UseMnemonic = True
		Me.lblClientName.Visible = True
		Me.lblClientName.AutoSize = False
		Me.lblClientName.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblClientName.Name = "lblClientName"
		Me.fraFile.Text = "Result File Name"
		Me.fraFile.ForeColor = System.Drawing.Color.Black
		Me.fraFile.Size = New System.Drawing.Size(514, 57)
		Me.fraFile.Location = New System.Drawing.Point(9, 10)
		Me.fraFile.TabIndex = 24
		Me.fraFile.BackColor = System.Drawing.SystemColors.Control
		Me.fraFile.Enabled = True
		Me.fraFile.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraFile.Visible = True
		Me.fraFile.Name = "fraFile"
		Me.btnBrowse.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnBrowse.Text = "Browse..."
		Me.btnBrowse.Size = New System.Drawing.Size(65, 25)
		Me.btnBrowse.Location = New System.Drawing.Point(432, 19)
		Me.btnBrowse.TabIndex = 26
		Me.btnBrowse.BackColor = System.Drawing.SystemColors.Control
		Me.btnBrowse.CausesValidation = True
		Me.btnBrowse.Enabled = True
		Me.btnBrowse.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnBrowse.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnBrowse.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnBrowse.TabStop = True
		Me.btnBrowse.Name = "btnBrowse"
		Me.txtFile.AutoSize = False
		Me.txtFile.Size = New System.Drawing.Size(409, 19)
		Me.txtFile.Location = New System.Drawing.Point(16, 24)
		Me.txtFile.TabIndex = 25
		Me.txtFile.Text = "vessel.sta"
		Me.txtFile.AcceptsReturn = True
		Me.txtFile.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtFile.BackColor = System.Drawing.SystemColors.Window
		Me.txtFile.CausesValidation = True
		Me.txtFile.Enabled = True
		Me.txtFile.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtFile.HideSelection = True
		Me.txtFile.ReadOnly = False
		Me.txtFile.Maxlength = 0
		Me.txtFile.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtFile.MultiLine = False
		Me.txtFile.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtFile.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtFile.TabStop = True
		Me.txtFile.Visible = True
		Me.txtFile.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtFile.Name = "txtFile"
		Me.fraHeadings.Text = "Colinear Environmental Headings (TN)"
		Me.fraHeadings.ForeColor = System.Drawing.Color.Black
		Me.fraHeadings.Size = New System.Drawing.Size(355, 60)
		Me.fraHeadings.Location = New System.Drawing.Point(10, 298)
		Me.fraHeadings.TabIndex = 16
		Me.fraHeadings.BackColor = System.Drawing.SystemColors.Control
		Me.fraHeadings.Enabled = True
		Me.fraHeadings.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraHeadings.Visible = True
		Me.fraHeadings.Name = "fraHeadings"
		Me._txtHeadings_2.AutoSize = False
		Me._txtHeadings_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me._txtHeadings_2.Size = New System.Drawing.Size(35, 19)
		Me._txtHeadings_2.Location = New System.Drawing.Point(253, 23)
		Me._txtHeadings_2.TabIndex = 21
		Me._txtHeadings_2.TabStop = False
		Me._txtHeadings_2.Text = "9"
		Me._txtHeadings_2.AcceptsReturn = True
		Me._txtHeadings_2.BackColor = System.Drawing.SystemColors.Window
		Me._txtHeadings_2.CausesValidation = True
		Me._txtHeadings_2.Enabled = True
		Me._txtHeadings_2.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtHeadings_2.HideSelection = True
		Me._txtHeadings_2.ReadOnly = False
		Me._txtHeadings_2.Maxlength = 0
		Me._txtHeadings_2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtHeadings_2.MultiLine = False
		Me._txtHeadings_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtHeadings_2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtHeadings_2.Visible = True
		Me._txtHeadings_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtHeadings_2.Name = "_txtHeadings_2"
		Me._txtHeadings_1.AutoSize = False
		Me._txtHeadings_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me._txtHeadings_1.Size = New System.Drawing.Size(43, 19)
		Me._txtHeadings_1.Location = New System.Drawing.Point(138, 24)
		Me._txtHeadings_1.TabIndex = 19
		Me._txtHeadings_1.TabStop = False
		Me._txtHeadings_1.Text = "180"
		Me._txtHeadings_1.AcceptsReturn = True
		Me._txtHeadings_1.BackColor = System.Drawing.SystemColors.Window
		Me._txtHeadings_1.CausesValidation = True
		Me._txtHeadings_1.Enabled = True
		Me._txtHeadings_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtHeadings_1.HideSelection = True
		Me._txtHeadings_1.ReadOnly = False
		Me._txtHeadings_1.Maxlength = 0
		Me._txtHeadings_1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtHeadings_1.MultiLine = False
		Me._txtHeadings_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtHeadings_1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtHeadings_1.Visible = True
		Me._txtHeadings_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtHeadings_1.Name = "_txtHeadings_1"
		Me._txtHeadings_0.AutoSize = False
		Me._txtHeadings_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me._txtHeadings_0.Size = New System.Drawing.Size(45, 19)
		Me._txtHeadings_0.Location = New System.Drawing.Point(49, 24)
		Me._txtHeadings_0.TabIndex = 17
		Me._txtHeadings_0.TabStop = False
		Me._txtHeadings_0.Text = "0"
		Me._txtHeadings_0.AcceptsReturn = True
		Me._txtHeadings_0.BackColor = System.Drawing.SystemColors.Window
		Me._txtHeadings_0.CausesValidation = True
		Me._txtHeadings_0.Enabled = True
		Me._txtHeadings_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtHeadings_0.HideSelection = True
		Me._txtHeadings_0.ReadOnly = False
		Me._txtHeadings_0.Maxlength = 0
		Me._txtHeadings_0.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtHeadings_0.MultiLine = False
		Me._txtHeadings_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtHeadings_0.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtHeadings_0.Visible = True
		Me._txtHeadings_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtHeadings_0.Name = "_txtHeadings_0"
		Me._lblAngleUnit_2.Text = "Invalid_string_refer_to_original_code"
		Me._lblAngleUnit_2.Size = New System.Drawing.Size(17, 17)
		Me._lblAngleUnit_2.Location = New System.Drawing.Point(182, 22)
		Me._lblAngleUnit_2.TabIndex = 30
		Me._lblAngleUnit_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblAngleUnit_2.BackColor = System.Drawing.SystemColors.Control
		Me._lblAngleUnit_2.Enabled = True
		Me._lblAngleUnit_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblAngleUnit_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblAngleUnit_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblAngleUnit_2.UseMnemonic = True
		Me._lblAngleUnit_2.Visible = True
		Me._lblAngleUnit_2.AutoSize = False
		Me._lblAngleUnit_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblAngleUnit_2.Name = "_lblAngleUnit_2"
		Me._lblAngleUnit_1.Text = "Invalid_string_refer_to_original_code"
		Me._lblAngleUnit_1.Size = New System.Drawing.Size(17, 17)
		Me._lblAngleUnit_1.Location = New System.Drawing.Point(95, 23)
		Me._lblAngleUnit_1.TabIndex = 29
		Me._lblAngleUnit_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblAngleUnit_1.BackColor = System.Drawing.SystemColors.Control
		Me._lblAngleUnit_1.Enabled = True
		Me._lblAngleUnit_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblAngleUnit_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblAngleUnit_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblAngleUnit_1.UseMnemonic = True
		Me._lblAngleUnit_1.Visible = True
		Me._lblAngleUnit_1.AutoSize = False
		Me._lblAngleUnit_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblAngleUnit_1.Name = "_lblAngleUnit_1"
		Me._lblHeadings_3.Text = "headings"
		Me._lblHeadings_3.Size = New System.Drawing.Size(49, 17)
		Me._lblHeadings_3.Location = New System.Drawing.Point(295, 26)
		Me._lblHeadings_3.TabIndex = 23
		Me._lblHeadings_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblHeadings_3.BackColor = System.Drawing.SystemColors.Control
		Me._lblHeadings_3.Enabled = True
		Me._lblHeadings_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblHeadings_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblHeadings_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblHeadings_3.UseMnemonic = True
		Me._lblHeadings_3.Visible = True
		Me._lblHeadings_3.AutoSize = False
		Me._lblHeadings_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblHeadings_3.Name = "_lblHeadings_3"
		Me._lblHeadings_2.Text = "with total"
		Me._lblHeadings_2.Size = New System.Drawing.Size(49, 17)
		Me._lblHeadings_2.Location = New System.Drawing.Point(201, 25)
		Me._lblHeadings_2.TabIndex = 22
		Me._lblHeadings_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblHeadings_2.BackColor = System.Drawing.SystemColors.Control
		Me._lblHeadings_2.Enabled = True
		Me._lblHeadings_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblHeadings_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblHeadings_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblHeadings_2.UseMnemonic = True
		Me._lblHeadings_2.Visible = True
		Me._lblHeadings_2.AutoSize = False
		Me._lblHeadings_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblHeadings_2.Name = "_lblHeadings_2"
		Me._lblHeadings_1.Text = "to"
		Me._lblHeadings_1.Size = New System.Drawing.Size(21, 17)
		Me._lblHeadings_1.Location = New System.Drawing.Point(115, 24)
		Me._lblHeadings_1.TabIndex = 20
		Me._lblHeadings_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblHeadings_1.BackColor = System.Drawing.SystemColors.Control
		Me._lblHeadings_1.Enabled = True
		Me._lblHeadings_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblHeadings_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblHeadings_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblHeadings_1.UseMnemonic = True
		Me._lblHeadings_1.Visible = True
		Me._lblHeadings_1.AutoSize = False
		Me._lblHeadings_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblHeadings_1.Name = "_lblHeadings_1"
		Me._lblHeadings_0.Text = "From"
		Me._lblHeadings_0.Size = New System.Drawing.Size(38, 17)
		Me._lblHeadings_0.Location = New System.Drawing.Point(17, 25)
		Me._lblHeadings_0.TabIndex = 18
		Me._lblHeadings_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblHeadings_0.BackColor = System.Drawing.SystemColors.Control
		Me._lblHeadings_0.Enabled = True
		Me._lblHeadings_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblHeadings_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblHeadings_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblHeadings_0.UseMnemonic = True
		Me._lblHeadings_0.Visible = True
		Me._lblHeadings_0.AutoSize = False
		Me._lblHeadings_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblHeadings_0.Name = "_lblHeadings_0"
		Me.btnAnalysis.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnAnalysis.Text = "St&art"
		Me.btnAnalysis.Size = New System.Drawing.Size(65, 25)
		Me.btnAnalysis.Location = New System.Drawing.Point(532, 45)
		Me.btnAnalysis.TabIndex = 15
		Me.btnAnalysis.BackColor = System.Drawing.SystemColors.Control
		Me.btnAnalysis.CausesValidation = True
		Me.btnAnalysis.Enabled = True
		Me.btnAnalysis.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnAnalysis.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnAnalysis.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnAnalysis.TabStop = True
		Me.btnAnalysis.Name = "btnAnalysis"
		Me.btnCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnCancel.Text = "Cancel"
		Me.btnCancel.Size = New System.Drawing.Size(65, 25)
		Me.btnCancel.Location = New System.Drawing.Point(532, 13)
		Me.btnCancel.TabIndex = 14
		Me.btnCancel.BackColor = System.Drawing.SystemColors.Control
		Me.btnCancel.CausesValidation = True
		Me.btnCancel.Enabled = True
		Me.btnCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnCancel.TabStop = True
		Me.btnCancel.Name = "btnCancel"
		Me.fraVesselLoc.Text = "Current Vessel Station :"
		Me.fraVesselLoc.ForeColor = System.Drawing.Color.Black
		Me.fraVesselLoc.Size = New System.Drawing.Size(214, 141)
		Me.fraVesselLoc.Location = New System.Drawing.Point(10, 146)
		Me.fraVesselLoc.TabIndex = 3
		Me.fraVesselLoc.BackColor = System.Drawing.SystemColors.Control
		Me.fraVesselLoc.Enabled = True
		Me.fraVesselLoc.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraVesselLoc.Visible = True
		Me.fraVesselLoc.Name = "fraVesselLoc"
		Me._txtVslSt_0.AutoSize = False
		Me._txtVslSt_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me._txtVslSt_0.BackColor = System.Drawing.SystemColors.Control
		Me._txtVslSt_0.Size = New System.Drawing.Size(104, 19)
		Me._txtVslSt_0.Location = New System.Drawing.Point(71, 26)
		Me._txtVslSt_0.ReadOnly = True
		Me._txtVslSt_0.TabIndex = 7
		Me._txtVslSt_0.TabStop = False
		Me._txtVslSt_0.Text = "0"
		Me._txtVslSt_0.AcceptsReturn = True
		Me._txtVslSt_0.CausesValidation = True
		Me._txtVslSt_0.Enabled = True
		Me._txtVslSt_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtVslSt_0.HideSelection = True
		Me._txtVslSt_0.Maxlength = 0
		Me._txtVslSt_0.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtVslSt_0.MultiLine = False
		Me._txtVslSt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtVslSt_0.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtVslSt_0.Visible = True
		Me._txtVslSt_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtVslSt_0.Name = "_txtVslSt_0"
		Me._txtVslSt_1.AutoSize = False
		Me._txtVslSt_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me._txtVslSt_1.BackColor = System.Drawing.SystemColors.Control
		Me._txtVslSt_1.Size = New System.Drawing.Size(104, 19)
		Me._txtVslSt_1.Location = New System.Drawing.Point(71, 53)
		Me._txtVslSt_1.ReadOnly = True
		Me._txtVslSt_1.TabIndex = 6
		Me._txtVslSt_1.TabStop = False
		Me._txtVslSt_1.Text = "0"
		Me._txtVslSt_1.AcceptsReturn = True
		Me._txtVslSt_1.CausesValidation = True
		Me._txtVslSt_1.Enabled = True
		Me._txtVslSt_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtVslSt_1.HideSelection = True
		Me._txtVslSt_1.Maxlength = 0
		Me._txtVslSt_1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtVslSt_1.MultiLine = False
		Me._txtVslSt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtVslSt_1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtVslSt_1.Visible = True
		Me._txtVslSt_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtVslSt_1.Name = "_txtVslSt_1"
		Me._txtVslSt_2.AutoSize = False
		Me._txtVslSt_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me._txtVslSt_2.BackColor = System.Drawing.SystemColors.Control
		Me._txtVslSt_2.Size = New System.Drawing.Size(57, 19)
		Me._txtVslSt_2.Location = New System.Drawing.Point(118, 79)
		Me._txtVslSt_2.ReadOnly = True
		Me._txtVslSt_2.TabIndex = 5
		Me._txtVslSt_2.TabStop = False
		Me._txtVslSt_2.Text = "0"
		Me._txtVslSt_2.AcceptsReturn = True
		Me._txtVslSt_2.CausesValidation = True
		Me._txtVslSt_2.Enabled = True
		Me._txtVslSt_2.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtVslSt_2.HideSelection = True
		Me._txtVslSt_2.Maxlength = 0
		Me._txtVslSt_2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtVslSt_2.MultiLine = False
		Me._txtVslSt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtVslSt_2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtVslSt_2.Visible = True
		Me._txtVslSt_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtVslSt_2.Name = "_txtVslSt_2"
		Me._txtVslSt_3.AutoSize = False
		Me._txtVslSt_3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me._txtVslSt_3.BackColor = System.Drawing.SystemColors.Control
		Me._txtVslSt_3.Size = New System.Drawing.Size(57, 19)
		Me._txtVslSt_3.Location = New System.Drawing.Point(118, 106)
		Me._txtVslSt_3.ReadOnly = True
		Me._txtVslSt_3.TabIndex = 4
		Me._txtVslSt_3.TabStop = False
		Me._txtVslSt_3.Text = "0"
		Me._txtVslSt_3.AcceptsReturn = True
		Me._txtVslSt_3.CausesValidation = True
		Me._txtVslSt_3.Enabled = True
		Me._txtVslSt_3.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtVslSt_3.HideSelection = True
		Me._txtVslSt_3.Maxlength = 0
		Me._txtVslSt_3.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtVslSt_3.MultiLine = False
		Me._txtVslSt_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtVslSt_3.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtVslSt_3.Visible = True
		Me._txtVslSt_3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtVslSt_3.Name = "_txtVslSt_3"
		Me._lblLengthUnit_2.Text = "ft"
		Me._lblLengthUnit_2.Size = New System.Drawing.Size(17, 17)
		Me._lblLengthUnit_2.Location = New System.Drawing.Point(183, 108)
		Me._lblLengthUnit_2.TabIndex = 28
		Me._lblLengthUnit_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblLengthUnit_2.BackColor = System.Drawing.SystemColors.Control
		Me._lblLengthUnit_2.Enabled = True
		Me._lblLengthUnit_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblLengthUnit_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblLengthUnit_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblLengthUnit_2.UseMnemonic = True
		Me._lblLengthUnit_2.Visible = True
		Me._lblLengthUnit_2.AutoSize = False
		Me._lblLengthUnit_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblLengthUnit_2.Name = "_lblLengthUnit_2"
		Me._lblLengthUnit_1.Text = "ft"
		Me._lblLengthUnit_1.Size = New System.Drawing.Size(17, 17)
		Me._lblLengthUnit_1.Location = New System.Drawing.Point(183, 53)
		Me._lblLengthUnit_1.TabIndex = 27
		Me._lblLengthUnit_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblLengthUnit_1.BackColor = System.Drawing.SystemColors.Control
		Me._lblLengthUnit_1.Enabled = True
		Me._lblLengthUnit_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblLengthUnit_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblLengthUnit_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblLengthUnit_1.UseMnemonic = True
		Me._lblLengthUnit_1.Visible = True
		Me._lblLengthUnit_1.AutoSize = False
		Me._lblLengthUnit_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblLengthUnit_1.Name = "_lblLengthUnit_1"
		Me._lblVslSt_4.Text = "Draft"
		Me._lblVslSt_4.Size = New System.Drawing.Size(25, 17)
		Me._lblVslSt_4.Location = New System.Drawing.Point(29, 107)
		Me._lblVslSt_4.TabIndex = 13
		Me._lblVslSt_4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblVslSt_4.BackColor = System.Drawing.SystemColors.Control
		Me._lblVslSt_4.Enabled = True
		Me._lblVslSt_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblVslSt_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblVslSt_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblVslSt_4.UseMnemonic = True
		Me._lblVslSt_4.Visible = True
		Me._lblVslSt_4.AutoSize = False
		Me._lblVslSt_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblVslSt_4.Name = "_lblVslSt_4"
		Me._lblLengthUnit_0.Text = "ft"
		Me._lblLengthUnit_0.Size = New System.Drawing.Size(17, 17)
		Me._lblLengthUnit_0.Location = New System.Drawing.Point(183, 26)
		Me._lblLengthUnit_0.TabIndex = 12
		Me._lblLengthUnit_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblLengthUnit_0.BackColor = System.Drawing.SystemColors.Control
		Me._lblLengthUnit_0.Enabled = True
		Me._lblLengthUnit_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblLengthUnit_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblLengthUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblLengthUnit_0.UseMnemonic = True
		Me._lblLengthUnit_0.Visible = True
		Me._lblLengthUnit_0.AutoSize = False
		Me._lblLengthUnit_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblLengthUnit_0.Name = "_lblLengthUnit_0"
		Me._lblVslSt_0.Text = "x (E)"
		Me._lblVslSt_0.Size = New System.Drawing.Size(33, 17)
		Me._lblVslSt_0.Location = New System.Drawing.Point(29, 26)
		Me._lblVslSt_0.TabIndex = 11
		Me._lblVslSt_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblVslSt_0.BackColor = System.Drawing.SystemColors.Control
		Me._lblVslSt_0.Enabled = True
		Me._lblVslSt_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblVslSt_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblVslSt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblVslSt_0.UseMnemonic = True
		Me._lblVslSt_0.Visible = True
		Me._lblVslSt_0.AutoSize = False
		Me._lblVslSt_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblVslSt_0.Name = "_lblVslSt_0"
		Me._lblVslSt_1.Text = "y (N)"
		Me._lblVslSt_1.Size = New System.Drawing.Size(33, 17)
		Me._lblVslSt_1.Location = New System.Drawing.Point(29, 53)
		Me._lblVslSt_1.TabIndex = 10
		Me._lblVslSt_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblVslSt_1.BackColor = System.Drawing.SystemColors.Control
		Me._lblVslSt_1.Enabled = True
		Me._lblVslSt_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblVslSt_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblVslSt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblVslSt_1.UseMnemonic = True
		Me._lblVslSt_1.Visible = True
		Me._lblVslSt_1.AutoSize = False
		Me._lblVslSt_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblVslSt_1.Name = "_lblVslSt_1"
		Me._lblVslSt_3.Text = "Heading"
		Me._lblVslSt_3.Size = New System.Drawing.Size(49, 17)
		Me._lblVslSt_3.Location = New System.Drawing.Point(29, 80)
		Me._lblVslSt_3.TabIndex = 9
		Me._lblVslSt_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblVslSt_3.BackColor = System.Drawing.SystemColors.Control
		Me._lblVslSt_3.Enabled = True
		Me._lblVslSt_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblVslSt_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblVslSt_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblVslSt_3.UseMnemonic = True
		Me._lblVslSt_3.Visible = True
		Me._lblVslSt_3.AutoSize = False
		Me._lblVslSt_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblVslSt_3.Name = "_lblVslSt_3"
		Me._lblAngleUnit_0.Text = "Invalid_string_refer_to_original_code"
		Me._lblAngleUnit_0.Size = New System.Drawing.Size(17, 17)
		Me._lblAngleUnit_0.Location = New System.Drawing.Point(183, 81)
		Me._lblAngleUnit_0.TabIndex = 8
		Me._lblAngleUnit_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblAngleUnit_0.BackColor = System.Drawing.SystemColors.Control
		Me._lblAngleUnit_0.Enabled = True
		Me._lblAngleUnit_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblAngleUnit_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblAngleUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblAngleUnit_0.UseMnemonic = True
		Me._lblAngleUnit_0.Visible = True
		Me._lblAngleUnit_0.AutoSize = False
		Me._lblAngleUnit_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblAngleUnit_0.Name = "_lblAngleUnit_0"
		Me.fraEnvironment.Text = "Environment Condition"
		Me.fraEnvironment.ForeColor = System.Drawing.Color.Black
		Me.fraEnvironment.Size = New System.Drawing.Size(216, 58)
		Me.fraEnvironment.Location = New System.Drawing.Point(9, 79)
		Me.fraEnvironment.TabIndex = 0
		Me.fraEnvironment.BackColor = System.Drawing.SystemColors.Control
		Me.fraEnvironment.Enabled = True
		Me.fraEnvironment.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraEnvironment.Visible = True
		Me.fraEnvironment.Name = "fraEnvironment"
		Me.txtEnvironment.AutoSize = False
		Me.txtEnvironment.BackColor = System.Drawing.SystemColors.Control
		Me.txtEnvironment.Size = New System.Drawing.Size(131, 19)
		Me.txtEnvironment.Location = New System.Drawing.Point(16, 24)
		Me.txtEnvironment.ReadOnly = True
		Me.txtEnvironment.TabIndex = 2
		Me.txtEnvironment.Text = "Current Environment"
		Me.txtEnvironment.AcceptsReturn = True
		Me.txtEnvironment.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtEnvironment.CausesValidation = True
		Me.txtEnvironment.Enabled = True
		Me.txtEnvironment.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtEnvironment.HideSelection = True
		Me.txtEnvironment.Maxlength = 0
		Me.txtEnvironment.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtEnvironment.MultiLine = False
		Me.txtEnvironment.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtEnvironment.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtEnvironment.TabStop = True
		Me.txtEnvironment.Visible = True
		Me.txtEnvironment.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtEnvironment.Name = "txtEnvironment"
		Me.btnEnvironment.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnEnvironment.Text = "Edit..."
		Me.btnEnvironment.Size = New System.Drawing.Size(50, 25)
		Me.btnEnvironment.Location = New System.Drawing.Point(154, 22)
		Me.btnEnvironment.TabIndex = 1
		Me.btnEnvironment.BackColor = System.Drawing.SystemColors.Control
		Me.btnEnvironment.CausesValidation = True
		Me.btnEnvironment.Enabled = True
		Me.btnEnvironment.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnEnvironment.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnEnvironment.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnEnvironment.TabStop = True
		Me.btnEnvironment.Name = "btnEnvironment"
		Me._lblForceUnit_0.Text = "kips"
		Me._lblForceUnit_0.Size = New System.Drawing.Size(25, 17)
		Me._lblForceUnit_0.Location = New System.Drawing.Point(1, 16)
		Me._lblForceUnit_0.TabIndex = 51
		Me._lblForceUnit_0.Visible = False
		Me._lblForceUnit_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblForceUnit_0.BackColor = System.Drawing.SystemColors.Control
		Me._lblForceUnit_0.Enabled = True
		Me._lblForceUnit_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblForceUnit_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblForceUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblForceUnit_0.UseMnemonic = True
		Me._lblForceUnit_0.AutoSize = False
		Me._lblForceUnit_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblForceUnit_0.Name = "_lblForceUnit_0"
		Me._lblVelUnit_0.Text = "kn"
		Me._lblVelUnit_0.Size = New System.Drawing.Size(25, 17)
		Me._lblVelUnit_0.Location = New System.Drawing.Point(0, 0)
		Me._lblVelUnit_0.TabIndex = 50
		Me._lblVelUnit_0.Visible = False
		Me._lblVelUnit_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblVelUnit_0.BackColor = System.Drawing.SystemColors.Control
		Me._lblVelUnit_0.Enabled = True
		Me._lblVelUnit_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblVelUnit_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblVelUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblVelUnit_0.UseMnemonic = True
		Me._lblVelUnit_0.AutoSize = False
		Me._lblVelUnit_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblVelUnit_0.Name = "_lblVelUnit_0"
		Me.Controls.Add(fraMoorGroup)
		Me.Controls.Add(fraReport)
		Me.Controls.Add(fraFile)
		Me.Controls.Add(fraHeadings)
		Me.Controls.Add(btnAnalysis)
		Me.Controls.Add(btnCancel)
		Me.Controls.Add(fraVesselLoc)
		Me.Controls.Add(fraEnvironment)
		Me.Controls.Add(_lblForceUnit_0)
		Me.Controls.Add(_lblVelUnit_0)
		Me.fraMoorGroup.Controls.Add(txtNumLinesPerGroup)
		Me.fraMoorGroup.Controls.Add(Label1)
		Me.fraReport.Controls.Add(txtSubTitle)
		Me.fraReport.Controls.Add(txtReportTitle)
		Me.fraReport.Controls.Add(Frame2)
		Me.fraReport.Controls.Add(btnReport)
		Me.fraReport.Controls.Add(lblLocation)
		Me.fraReport.Controls.Add(lblClientName)
		Me.Frame2.Controls.Add(txtInitYaw)
		Me.Frame2.Controls.Add(txtInitSway)
		Me.Frame2.Controls.Add(txtInitSurge)
		Me.Frame2.Controls.Add(_lblInitPos_2)
		Me.Frame2.Controls.Add(_lblInitPos_1)
		Me.Frame2.Controls.Add(_lblInitPos_0)
		Me.Frame2.Controls.Add(_lblLengthUnit_4)
		Me.Frame2.Controls.Add(_lblLengthUnit_3)
		Me.Frame2.Controls.Add(_lblAngleUnit_3)
		Me.fraFile.Controls.Add(btnBrowse)
		Me.fraFile.Controls.Add(txtFile)
		Me.fraHeadings.Controls.Add(_txtHeadings_2)
		Me.fraHeadings.Controls.Add(_txtHeadings_1)
		Me.fraHeadings.Controls.Add(_txtHeadings_0)
		Me.fraHeadings.Controls.Add(_lblAngleUnit_2)
		Me.fraHeadings.Controls.Add(_lblAngleUnit_1)
		Me.fraHeadings.Controls.Add(_lblHeadings_3)
		Me.fraHeadings.Controls.Add(_lblHeadings_2)
		Me.fraHeadings.Controls.Add(_lblHeadings_1)
		Me.fraHeadings.Controls.Add(_lblHeadings_0)
		Me.fraVesselLoc.Controls.Add(_txtVslSt_0)
		Me.fraVesselLoc.Controls.Add(_txtVslSt_1)
		Me.fraVesselLoc.Controls.Add(_txtVslSt_2)
		Me.fraVesselLoc.Controls.Add(_txtVslSt_3)
		Me.fraVesselLoc.Controls.Add(_lblLengthUnit_2)
		Me.fraVesselLoc.Controls.Add(_lblLengthUnit_1)
		Me.fraVesselLoc.Controls.Add(_lblVslSt_4)
		Me.fraVesselLoc.Controls.Add(_lblLengthUnit_0)
		Me.fraVesselLoc.Controls.Add(_lblVslSt_0)
		Me.fraVesselLoc.Controls.Add(_lblVslSt_1)
		Me.fraVesselLoc.Controls.Add(_lblVslSt_3)
		Me.fraVesselLoc.Controls.Add(_lblAngleUnit_0)
		Me.fraEnvironment.Controls.Add(txtEnvironment)
		Me.fraEnvironment.Controls.Add(btnEnvironment)
		Me.lblAngleUnit.SetIndex(_lblAngleUnit_3, CType(3, Short))
		Me.lblAngleUnit.SetIndex(_lblAngleUnit_2, CType(2, Short))
		Me.lblAngleUnit.SetIndex(_lblAngleUnit_1, CType(1, Short))
		Me.lblAngleUnit.SetIndex(_lblAngleUnit_0, CType(0, Short))
		Me.lblForceUnit.SetIndex(_lblForceUnit_0, CType(0, Short))
		Me.lblHeadings.SetIndex(_lblHeadings_3, CType(3, Short))
		Me.lblHeadings.SetIndex(_lblHeadings_2, CType(2, Short))
		Me.lblHeadings.SetIndex(_lblHeadings_1, CType(1, Short))
		Me.lblHeadings.SetIndex(_lblHeadings_0, CType(0, Short))
		Me.lblInitPos.SetIndex(_lblInitPos_2, CType(2, Short))
		Me.lblInitPos.SetIndex(_lblInitPos_1, CType(1, Short))
		Me.lblInitPos.SetIndex(_lblInitPos_0, CType(0, Short))
		Me.lblLengthUnit.SetIndex(_lblLengthUnit_4, CType(4, Short))
		Me.lblLengthUnit.SetIndex(_lblLengthUnit_3, CType(3, Short))
		Me.lblLengthUnit.SetIndex(_lblLengthUnit_2, CType(2, Short))
		Me.lblLengthUnit.SetIndex(_lblLengthUnit_1, CType(1, Short))
		Me.lblLengthUnit.SetIndex(_lblLengthUnit_0, CType(0, Short))
		Me.lblVelUnit.SetIndex(_lblVelUnit_0, CType(0, Short))
		Me.lblVslSt.SetIndex(_lblVslSt_4, CType(4, Short))
		Me.lblVslSt.SetIndex(_lblVslSt_0, CType(0, Short))
		Me.lblVslSt.SetIndex(_lblVslSt_1, CType(1, Short))
		Me.lblVslSt.SetIndex(_lblVslSt_3, CType(3, Short))
		Me.txtHeadings.SetIndex(_txtHeadings_2, CType(2, Short))
		Me.txtHeadings.SetIndex(_txtHeadings_1, CType(1, Short))
		Me.txtHeadings.SetIndex(_txtHeadings_0, CType(0, Short))
		Me.txtVslSt.SetIndex(_txtVslSt_0, CType(0, Short))
		Me.txtVslSt.SetIndex(_txtVslSt_1, CType(1, Short))
		Me.txtVslSt.SetIndex(_txtVslSt_2, CType(2, Short))
		Me.txtVslSt.SetIndex(_txtVslSt_3, CType(3, Short))
		CType(Me.txtVslSt, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.txtHeadings, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lblVslSt, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lblVelUnit, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lblLengthUnit, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lblInitPos, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lblHeadings, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lblForceUnit, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lblAngleUnit, System.ComponentModel.ISupportInitialize).EndInit()
		Me.fraMoorGroup.ResumeLayout(False)
		Me.fraReport.ResumeLayout(False)
		Me.Frame2.ResumeLayout(False)
		Me.fraFile.ResumeLayout(False)
		Me.fraHeadings.ResumeLayout(False)
		Me.fraVesselLoc.ResumeLayout(False)
		Me.fraEnvironment.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class