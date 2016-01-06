<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmOptions
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
	Public WithEvents Text2 As System.Windows.Forms.TextBox
	Public WithEvents txtNumLineSegments As System.Windows.Forms.TextBox
	Public WithEvents txtMinSegLength As System.Windows.Forms.TextBox
	Public WithEvents txtMaxIter As System.Windows.Forms.TextBox
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_3 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents txtStartFreq As System.Windows.Forms.TextBox
	Public WithEvents Text1 As System.Windows.Forms.TextBox
	Public WithEvents Label10 As System.Windows.Forms.Label
	Public WithEvents Label9 As System.Windows.Forms.Label
	Public WithEvents Label8 As System.Windows.Forms.Label
	Public WithEvents Label7 As System.Windows.Forms.Label
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents fraIntegration As System.Windows.Forms.GroupBox
	Public WithEvents _optScopeRef_1 As System.Windows.Forms.RadioButton
	Public WithEvents _optScopeRef_0 As System.Windows.Forms.RadioButton
	Public WithEvents fraScopeRef As System.Windows.Forms.GroupBox
	Public WithEvents txtVesselHdgTol As System.Windows.Forms.TextBox
	Public WithEvents txtVesselOffsetTol As System.Windows.Forms.TextBox
	Public WithEvents txtLineTensionTol As System.Windows.Forms.TextBox
	Public WithEvents txtAnchorDepthTol As System.Windows.Forms.TextBox
	Public WithEvents txtSpreadAngleTol As System.Windows.Forms.TextBox
	Public WithEvents txtScopeTol As System.Windows.Forms.TextBox
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_2 As System.Windows.Forms.Label
	Public WithEvents lblVesselHdgTol As System.Windows.Forms.Label
	Public WithEvents lblVesselOffsetTol As System.Windows.Forms.Label
	Public WithEvents lblForceUnit As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_1 As System.Windows.Forms.Label
	Public WithEvents lblAngleUnit As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_0 As System.Windows.Forms.Label
	Public WithEvents lblForceTol As System.Windows.Forms.Label
	Public WithEvents lblDepthTol As System.Windows.Forms.Label
	Public WithEvents lblAngleTol As System.Windows.Forms.Label
	Public WithEvents lblDistTol As System.Windows.Forms.Label
	Public WithEvents fraTol As System.Windows.Forms.GroupBox
	Public WithEvents btnOK As System.Windows.Forms.Button
	Public WithEvents btnCancel As System.Windows.Forms.Button
	Public WithEvents _optUnit_1 As System.Windows.Forms.RadioButton
	Public WithEvents _optUnit_0 As System.Windows.Forms.RadioButton
	Public WithEvents fraUnit As System.Windows.Forms.GroupBox
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_4 As System.Windows.Forms.Label
	Public WithEvents lblLengthUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents optScopeRef As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	Public WithEvents optUnit As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmOptions))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Text2 = New System.Windows.Forms.TextBox
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.txtNumLineSegments = New System.Windows.Forms.TextBox
		Me.txtMinSegLength = New System.Windows.Forms.TextBox
		Me.txtMaxIter = New System.Windows.Forms.TextBox
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me._lblLengthUnit_3 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.fraIntegration = New System.Windows.Forms.GroupBox
		Me.txtStartFreq = New System.Windows.Forms.TextBox
		Me.Text1 = New System.Windows.Forms.TextBox
		Me.Label10 = New System.Windows.Forms.Label
		Me.Label9 = New System.Windows.Forms.Label
		Me.Label8 = New System.Windows.Forms.Label
		Me.Label7 = New System.Windows.Forms.Label
		Me.Label6 = New System.Windows.Forms.Label
		Me.fraScopeRef = New System.Windows.Forms.GroupBox
		Me._optScopeRef_1 = New System.Windows.Forms.RadioButton
		Me._optScopeRef_0 = New System.Windows.Forms.RadioButton
		Me.fraTol = New System.Windows.Forms.GroupBox
		Me.txtVesselHdgTol = New System.Windows.Forms.TextBox
		Me.txtVesselOffsetTol = New System.Windows.Forms.TextBox
		Me.txtLineTensionTol = New System.Windows.Forms.TextBox
		Me.txtAnchorDepthTol = New System.Windows.Forms.TextBox
		Me.txtSpreadAngleTol = New System.Windows.Forms.TextBox
		Me.txtScopeTol = New System.Windows.Forms.TextBox
		Me.Label1 = New System.Windows.Forms.Label
		Me._lblLengthUnit_2 = New System.Windows.Forms.Label
		Me.lblVesselHdgTol = New System.Windows.Forms.Label
		Me.lblVesselOffsetTol = New System.Windows.Forms.Label
		Me.lblForceUnit = New System.Windows.Forms.Label
		Me._lblLengthUnit_1 = New System.Windows.Forms.Label
		Me.lblAngleUnit = New System.Windows.Forms.Label
		Me._lblLengthUnit_0 = New System.Windows.Forms.Label
		Me.lblForceTol = New System.Windows.Forms.Label
		Me.lblDepthTol = New System.Windows.Forms.Label
		Me.lblAngleTol = New System.Windows.Forms.Label
		Me.lblDistTol = New System.Windows.Forms.Label
		Me.btnOK = New System.Windows.Forms.Button
		Me.btnCancel = New System.Windows.Forms.Button
		Me.fraUnit = New System.Windows.Forms.GroupBox
		Me._optUnit_1 = New System.Windows.Forms.RadioButton
		Me._optUnit_0 = New System.Windows.Forms.RadioButton
		Me.Label5 = New System.Windows.Forms.Label
		Me._lblLengthUnit_4 = New System.Windows.Forms.Label
		Me.lblLengthUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.optScopeRef = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(components)
		Me.optUnit = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(components)
		Me.Frame1.SuspendLayout()
		Me.fraIntegration.SuspendLayout()
		Me.fraScopeRef.SuspendLayout()
		Me.fraTol.SuspendLayout()
		Me.fraUnit.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.lblLengthUnit, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.optScopeRef, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.optUnit, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.Text = "Options"
		Me.ClientSize = New System.Drawing.Size(231, 91)
		Me.Location = New System.Drawing.Point(4, 30)
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
		Me.Name = "frmOptions"
		Me.Text2.AutoSize = False
		Me.Text2.Size = New System.Drawing.Size(45, 21)
		Me.Text2.Location = New System.Drawing.Point(204, 455)
		Me.Text2.TabIndex = 36
		Me.Text2.Text = "2"
		Me.Text2.AcceptsReturn = True
		Me.Text2.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text2.BackColor = System.Drawing.SystemColors.Window
		Me.Text2.CausesValidation = True
		Me.Text2.Enabled = True
		Me.Text2.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text2.HideSelection = True
		Me.Text2.ReadOnly = False
		Me.Text2.Maxlength = 0
		Me.Text2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text2.MultiLine = False
		Me.Text2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text2.TabStop = True
		Me.Text2.Visible = True
		Me.Text2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text2.Name = "Text2"
		Me.Frame1.Text = "Capacity Limits"
		Me.Frame1.Size = New System.Drawing.Size(285, 146)
		Me.Frame1.Location = New System.Drawing.Point(8, 295)
		Me.Frame1.TabIndex = 28
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Name = "Frame1"
		Me.txtNumLineSegments.AutoSize = False
		Me.txtNumLineSegments.Size = New System.Drawing.Size(45, 21)
		Me.txtNumLineSegments.Location = New System.Drawing.Point(210, 71)
		Me.txtNumLineSegments.TabIndex = 34
		Me.txtNumLineSegments.Text = "20"
		Me.txtNumLineSegments.AcceptsReturn = True
		Me.txtNumLineSegments.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtNumLineSegments.BackColor = System.Drawing.SystemColors.Window
		Me.txtNumLineSegments.CausesValidation = True
		Me.txtNumLineSegments.Enabled = True
		Me.txtNumLineSegments.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtNumLineSegments.HideSelection = True
		Me.txtNumLineSegments.ReadOnly = False
		Me.txtNumLineSegments.Maxlength = 0
		Me.txtNumLineSegments.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtNumLineSegments.MultiLine = False
		Me.txtNumLineSegments.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtNumLineSegments.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtNumLineSegments.TabStop = True
		Me.txtNumLineSegments.Visible = True
		Me.txtNumLineSegments.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtNumLineSegments.Name = "txtNumLineSegments"
		Me.txtMinSegLength.AutoSize = False
		Me.txtMinSegLength.Size = New System.Drawing.Size(45, 21)
		Me.txtMinSegLength.Location = New System.Drawing.Point(211, 21)
		Me.txtMinSegLength.TabIndex = 30
		Me.txtMinSegLength.Text = "50"
		Me.txtMinSegLength.AcceptsReturn = True
		Me.txtMinSegLength.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtMinSegLength.BackColor = System.Drawing.SystemColors.Window
		Me.txtMinSegLength.CausesValidation = True
		Me.txtMinSegLength.Enabled = True
		Me.txtMinSegLength.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtMinSegLength.HideSelection = True
		Me.txtMinSegLength.ReadOnly = False
		Me.txtMinSegLength.Maxlength = 0
		Me.txtMinSegLength.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtMinSegLength.MultiLine = False
		Me.txtMinSegLength.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtMinSegLength.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtMinSegLength.TabStop = True
		Me.txtMinSegLength.Visible = True
		Me.txtMinSegLength.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtMinSegLength.Name = "txtMinSegLength"
		Me.txtMaxIter.AutoSize = False
		Me.txtMaxIter.Size = New System.Drawing.Size(45, 21)
		Me.txtMaxIter.Location = New System.Drawing.Point(211, 44)
		Me.txtMaxIter.TabIndex = 29
		Me.txtMaxIter.Text = "2000"
		Me.txtMaxIter.AcceptsReturn = True
		Me.txtMaxIter.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtMaxIter.BackColor = System.Drawing.SystemColors.Window
		Me.txtMaxIter.CausesValidation = True
		Me.txtMaxIter.Enabled = True
		Me.txtMaxIter.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtMaxIter.HideSelection = True
		Me.txtMaxIter.ReadOnly = False
		Me.txtMaxIter.Maxlength = 0
		Me.txtMaxIter.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtMaxIter.MultiLine = False
		Me.txtMaxIter.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtMaxIter.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtMaxIter.TabStop = True
		Me.txtMaxIter.Visible = True
		Me.txtMaxIter.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtMaxIter.Name = "txtMaxIter"
		Me.Label4.Text = "Number of Line Segments:"
		Me.Label4.Size = New System.Drawing.Size(167, 16)
		Me.Label4.Location = New System.Drawing.Point(23, 73)
		Me.Label4.TabIndex = 35
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
		Me.Label2.Text = "Min. Line Segment Length:"
		Me.Label2.Size = New System.Drawing.Size(167, 16)
		Me.Label2.Location = New System.Drawing.Point(23, 24)
		Me.Label2.TabIndex = 33
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
		Me._lblLengthUnit_3.Text = "ft"
		Me._lblLengthUnit_3.Size = New System.Drawing.Size(17, 16)
		Me._lblLengthUnit_3.Location = New System.Drawing.Point(263, 24)
		Me._lblLengthUnit_3.TabIndex = 32
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
		Me.Label3.Text = "Max Iteration:"
		Me.Label3.Size = New System.Drawing.Size(100, 16)
		Me.Label3.Location = New System.Drawing.Point(23, 47)
		Me.Label3.TabIndex = 31
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
		Me.fraIntegration.Text = "Energy Spectrum Integration"
		Me.fraIntegration.Size = New System.Drawing.Size(285, 273)
		Me.fraIntegration.Location = New System.Drawing.Point(302, 99)
		Me.fraIntegration.TabIndex = 27
		Me.fraIntegration.BackColor = System.Drawing.SystemColors.Control
		Me.fraIntegration.Enabled = True
		Me.fraIntegration.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraIntegration.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraIntegration.Visible = True
		Me.fraIntegration.Name = "fraIntegration"
		Me.txtStartFreq.AutoSize = False
		Me.txtStartFreq.Size = New System.Drawing.Size(45, 21)
		Me.txtStartFreq.Location = New System.Drawing.Point(198, 22)
		Me.txtStartFreq.TabIndex = 40
		Me.txtStartFreq.Text = "50"
		Me.txtStartFreq.AcceptsReturn = True
		Me.txtStartFreq.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtStartFreq.BackColor = System.Drawing.SystemColors.Window
		Me.txtStartFreq.CausesValidation = True
		Me.txtStartFreq.Enabled = True
		Me.txtStartFreq.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtStartFreq.HideSelection = True
		Me.txtStartFreq.ReadOnly = False
		Me.txtStartFreq.Maxlength = 0
		Me.txtStartFreq.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtStartFreq.MultiLine = False
		Me.txtStartFreq.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtStartFreq.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtStartFreq.TabStop = True
		Me.txtStartFreq.Visible = True
		Me.txtStartFreq.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtStartFreq.Name = "txtStartFreq"
		Me.Text1.AutoSize = False
		Me.Text1.Size = New System.Drawing.Size(45, 21)
		Me.Text1.Location = New System.Drawing.Point(198, 45)
		Me.Text1.TabIndex = 39
		Me.Text1.Text = "2000"
		Me.Text1.AcceptsReturn = True
		Me.Text1.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text1.BackColor = System.Drawing.SystemColors.Window
		Me.Text1.CausesValidation = True
		Me.Text1.Enabled = True
		Me.Text1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text1.HideSelection = True
		Me.Text1.ReadOnly = False
		Me.Text1.Maxlength = 0
		Me.Text1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text1.MultiLine = False
		Me.Text1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text1.TabStop = True
		Me.Text1.Visible = True
		Me.Text1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text1.Name = "Text1"
		Me.Label10.Text = "Wave Frequency Increment:"
		Me.Label10.Size = New System.Drawing.Size(167, 16)
		Me.Label10.Location = New System.Drawing.Point(13, 72)
		Me.Label10.TabIndex = 45
		Me.Label10.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label10.BackColor = System.Drawing.SystemColors.Control
		Me.Label10.Enabled = True
		Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label10.UseMnemonic = True
		Me.Label10.Visible = True
		Me.Label10.AutoSize = False
		Me.Label10.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label10.Name = "Label10"
		Me.Label9.Text = "Start Low Frequency:"
		Me.Label9.Size = New System.Drawing.Size(167, 16)
		Me.Label9.Location = New System.Drawing.Point(16, 104)
		Me.Label9.TabIndex = 44
		Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label9.BackColor = System.Drawing.SystemColors.Control
		Me.Label9.Enabled = True
		Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label9.UseMnemonic = True
		Me.Label9.Visible = True
		Me.Label9.AutoSize = False
		Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label9.Name = "Label9"
		Me.Label8.Text = "End Low Frequency:"
		Me.Label8.Size = New System.Drawing.Size(167, 16)
		Me.Label8.Location = New System.Drawing.Point(19, 135)
		Me.Label8.TabIndex = 43
		Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label8.BackColor = System.Drawing.SystemColors.Control
		Me.Label8.Enabled = True
		Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label8.UseMnemonic = True
		Me.Label8.Visible = True
		Me.Label8.AutoSize = False
		Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label8.Name = "Label8"
		Me.Label7.Text = "Start Wave Frequency:"
		Me.Label7.Size = New System.Drawing.Size(167, 16)
		Me.Label7.Location = New System.Drawing.Point(9, 26)
		Me.Label7.TabIndex = 42
		Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label7.BackColor = System.Drawing.SystemColors.Control
		Me.Label7.Enabled = True
		Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label7.UseMnemonic = True
		Me.Label7.Visible = True
		Me.Label7.AutoSize = False
		Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label7.Name = "Label7"
		Me.Label6.Text = "End Wave Frequency:"
		Me.Label6.Size = New System.Drawing.Size(167, 16)
		Me.Label6.Location = New System.Drawing.Point(12, 46)
		Me.Label6.TabIndex = 41
		Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label6.BackColor = System.Drawing.SystemColors.Control
		Me.Label6.Enabled = True
		Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label6.UseMnemonic = True
		Me.Label6.Visible = True
		Me.Label6.AutoSize = False
		Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label6.Name = "Label6"
		Me.fraScopeRef.Text = "Scope Reference"
		Me.fraScopeRef.ForeColor = System.Drawing.Color.Black
		Me.fraScopeRef.Size = New System.Drawing.Size(138, 79)
		Me.fraScopeRef.Location = New System.Drawing.Point(274, 9)
		Me.fraScopeRef.TabIndex = 6
		Me.fraScopeRef.Visible = False
		Me.fraScopeRef.BackColor = System.Drawing.SystemColors.Control
		Me.fraScopeRef.Enabled = True
		Me.fraScopeRef.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraScopeRef.Name = "fraScopeRef"
		Me._optScopeRef_1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optScopeRef_1.Text = "Vessel Center"
		Me._optScopeRef_1.Size = New System.Drawing.Size(98, 22)
		Me._optScopeRef_1.Location = New System.Drawing.Point(20, 44)
		Me._optScopeRef_1.TabIndex = 8
		Me._optScopeRef_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optScopeRef_1.BackColor = System.Drawing.SystemColors.Control
		Me._optScopeRef_1.CausesValidation = True
		Me._optScopeRef_1.Enabled = True
		Me._optScopeRef_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optScopeRef_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._optScopeRef_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optScopeRef_1.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optScopeRef_1.TabStop = True
		Me._optScopeRef_1.Checked = False
		Me._optScopeRef_1.Visible = True
		Me._optScopeRef_1.Name = "_optScopeRef_1"
		Me._optScopeRef_0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optScopeRef_0.Text = "Fairlead"
		Me._optScopeRef_0.Size = New System.Drawing.Size(83, 22)
		Me._optScopeRef_0.Location = New System.Drawing.Point(20, 20)
		Me._optScopeRef_0.TabIndex = 7
		Me._optScopeRef_0.Checked = True
		Me._optScopeRef_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optScopeRef_0.BackColor = System.Drawing.SystemColors.Control
		Me._optScopeRef_0.CausesValidation = True
		Me._optScopeRef_0.Enabled = True
		Me._optScopeRef_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optScopeRef_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._optScopeRef_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optScopeRef_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optScopeRef_0.TabStop = True
		Me._optScopeRef_0.Visible = True
		Me._optScopeRef_0.Name = "_optScopeRef_0"
		Me.fraTol.Text = "Precision/Tolerance"
		Me.fraTol.Size = New System.Drawing.Size(287, 191)
		Me.fraTol.Location = New System.Drawing.Point(7, 95)
		Me.fraTol.TabIndex = 5
		Me.fraTol.BackColor = System.Drawing.SystemColors.Control
		Me.fraTol.Enabled = True
		Me.fraTol.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraTol.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraTol.Visible = True
		Me.fraTol.Name = "fraTol"
		Me.txtVesselHdgTol.AutoSize = False
		Me.txtVesselHdgTol.Size = New System.Drawing.Size(60, 21)
		Me.txtVesselHdgTol.Location = New System.Drawing.Point(127, 155)
		Me.txtVesselHdgTol.TabIndex = 25
		Me.txtVesselHdgTol.Text = "0.0001"
		Me.txtVesselHdgTol.AcceptsReturn = True
		Me.txtVesselHdgTol.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtVesselHdgTol.BackColor = System.Drawing.SystemColors.Window
		Me.txtVesselHdgTol.CausesValidation = True
		Me.txtVesselHdgTol.Enabled = True
		Me.txtVesselHdgTol.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtVesselHdgTol.HideSelection = True
		Me.txtVesselHdgTol.ReadOnly = False
		Me.txtVesselHdgTol.Maxlength = 0
		Me.txtVesselHdgTol.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtVesselHdgTol.MultiLine = False
		Me.txtVesselHdgTol.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtVesselHdgTol.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtVesselHdgTol.TabStop = True
		Me.txtVesselHdgTol.Visible = True
		Me.txtVesselHdgTol.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtVesselHdgTol.Name = "txtVesselHdgTol"
		Me.txtVesselOffsetTol.AutoSize = False
		Me.txtVesselOffsetTol.Size = New System.Drawing.Size(60, 21)
		Me.txtVesselOffsetTol.Location = New System.Drawing.Point(127, 129)
		Me.txtVesselOffsetTol.TabIndex = 23
		Me.txtVesselOffsetTol.Text = "0.0001"
		Me.txtVesselOffsetTol.AcceptsReturn = True
		Me.txtVesselOffsetTol.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtVesselOffsetTol.BackColor = System.Drawing.SystemColors.Window
		Me.txtVesselOffsetTol.CausesValidation = True
		Me.txtVesselOffsetTol.Enabled = True
		Me.txtVesselOffsetTol.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtVesselOffsetTol.HideSelection = True
		Me.txtVesselOffsetTol.ReadOnly = False
		Me.txtVesselOffsetTol.Maxlength = 0
		Me.txtVesselOffsetTol.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtVesselOffsetTol.MultiLine = False
		Me.txtVesselOffsetTol.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtVesselOffsetTol.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtVesselOffsetTol.TabStop = True
		Me.txtVesselOffsetTol.Visible = True
		Me.txtVesselOffsetTol.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtVesselOffsetTol.Name = "txtVesselOffsetTol"
		Me.txtLineTensionTol.AutoSize = False
		Me.txtLineTensionTol.Size = New System.Drawing.Size(60, 21)
		Me.txtLineTensionTol.Location = New System.Drawing.Point(127, 103)
		Me.txtLineTensionTol.TabIndex = 19
		Me.txtLineTensionTol.Text = "0.5"
		Me.txtLineTensionTol.AcceptsReturn = True
		Me.txtLineTensionTol.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtLineTensionTol.BackColor = System.Drawing.SystemColors.Window
		Me.txtLineTensionTol.CausesValidation = True
		Me.txtLineTensionTol.Enabled = True
		Me.txtLineTensionTol.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtLineTensionTol.HideSelection = True
		Me.txtLineTensionTol.ReadOnly = False
		Me.txtLineTensionTol.Maxlength = 0
		Me.txtLineTensionTol.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtLineTensionTol.MultiLine = False
		Me.txtLineTensionTol.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtLineTensionTol.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtLineTensionTol.TabStop = True
		Me.txtLineTensionTol.Visible = True
		Me.txtLineTensionTol.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtLineTensionTol.Name = "txtLineTensionTol"
		Me.txtAnchorDepthTol.AutoSize = False
		Me.txtAnchorDepthTol.Size = New System.Drawing.Size(60, 21)
		Me.txtAnchorDepthTol.Location = New System.Drawing.Point(127, 78)
		Me.txtAnchorDepthTol.TabIndex = 18
		Me.txtAnchorDepthTol.Text = "0.0005"
		Me.txtAnchorDepthTol.AcceptsReturn = True
		Me.txtAnchorDepthTol.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtAnchorDepthTol.BackColor = System.Drawing.SystemColors.Window
		Me.txtAnchorDepthTol.CausesValidation = True
		Me.txtAnchorDepthTol.Enabled = True
		Me.txtAnchorDepthTol.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtAnchorDepthTol.HideSelection = True
		Me.txtAnchorDepthTol.ReadOnly = False
		Me.txtAnchorDepthTol.Maxlength = 0
		Me.txtAnchorDepthTol.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtAnchorDepthTol.MultiLine = False
		Me.txtAnchorDepthTol.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtAnchorDepthTol.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtAnchorDepthTol.TabStop = True
		Me.txtAnchorDepthTol.Visible = True
		Me.txtAnchorDepthTol.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtAnchorDepthTol.Name = "txtAnchorDepthTol"
		Me.txtSpreadAngleTol.AutoSize = False
		Me.txtSpreadAngleTol.Size = New System.Drawing.Size(60, 21)
		Me.txtSpreadAngleTol.Location = New System.Drawing.Point(127, 52)
		Me.txtSpreadAngleTol.TabIndex = 15
		Me.txtSpreadAngleTol.Text = "0.001"
		Me.txtSpreadAngleTol.AcceptsReturn = True
		Me.txtSpreadAngleTol.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtSpreadAngleTol.BackColor = System.Drawing.SystemColors.Window
		Me.txtSpreadAngleTol.CausesValidation = True
		Me.txtSpreadAngleTol.Enabled = True
		Me.txtSpreadAngleTol.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtSpreadAngleTol.HideSelection = True
		Me.txtSpreadAngleTol.ReadOnly = False
		Me.txtSpreadAngleTol.Maxlength = 0
		Me.txtSpreadAngleTol.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtSpreadAngleTol.MultiLine = False
		Me.txtSpreadAngleTol.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtSpreadAngleTol.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtSpreadAngleTol.TabStop = True
		Me.txtSpreadAngleTol.Visible = True
		Me.txtSpreadAngleTol.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtSpreadAngleTol.Name = "txtSpreadAngleTol"
		Me.txtScopeTol.AutoSize = False
		Me.txtScopeTol.Size = New System.Drawing.Size(60, 21)
		Me.txtScopeTol.Location = New System.Drawing.Point(127, 26)
		Me.txtScopeTol.TabIndex = 13
		Me.txtScopeTol.Text = "0.00005"
		Me.txtScopeTol.AcceptsReturn = True
		Me.txtScopeTol.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtScopeTol.BackColor = System.Drawing.SystemColors.Window
		Me.txtScopeTol.CausesValidation = True
		Me.txtScopeTol.Enabled = True
		Me.txtScopeTol.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtScopeTol.HideSelection = True
		Me.txtScopeTol.ReadOnly = False
		Me.txtScopeTol.Maxlength = 0
		Me.txtScopeTol.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtScopeTol.MultiLine = False
		Me.txtScopeTol.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtScopeTol.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtScopeTol.TabStop = True
		Me.txtScopeTol.Visible = True
		Me.txtScopeTol.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtScopeTol.Name = "txtScopeTol"
		Me.Label1.Text = "deg"
		Me.Label1.Size = New System.Drawing.Size(37, 16)
		Me.Label1.Location = New System.Drawing.Point(196, 157)
		Me.Label1.TabIndex = 26
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
		Me._lblLengthUnit_2.Text = "ft"
		Me._lblLengthUnit_2.Size = New System.Drawing.Size(17, 16)
		Me._lblLengthUnit_2.Location = New System.Drawing.Point(199, 134)
		Me._lblLengthUnit_2.TabIndex = 24
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
		Me.lblVesselHdgTol.Text = "Vessel Offset:"
		Me.lblVesselHdgTol.Size = New System.Drawing.Size(91, 16)
		Me.lblVesselHdgTol.Location = New System.Drawing.Point(34, 160)
		Me.lblVesselHdgTol.TabIndex = 22
		Me.lblVesselHdgTol.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblVesselHdgTol.BackColor = System.Drawing.SystemColors.Control
		Me.lblVesselHdgTol.Enabled = True
		Me.lblVesselHdgTol.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblVesselHdgTol.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblVesselHdgTol.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblVesselHdgTol.UseMnemonic = True
		Me.lblVesselHdgTol.Visible = True
		Me.lblVesselHdgTol.AutoSize = False
		Me.lblVesselHdgTol.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblVesselHdgTol.Name = "lblVesselHdgTol"
		Me.lblVesselOffsetTol.Text = "Vessel Offset:"
		Me.lblVesselOffsetTol.Size = New System.Drawing.Size(91, 16)
		Me.lblVesselOffsetTol.Location = New System.Drawing.Point(34, 134)
		Me.lblVesselOffsetTol.TabIndex = 21
		Me.lblVesselOffsetTol.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblVesselOffsetTol.BackColor = System.Drawing.SystemColors.Control
		Me.lblVesselOffsetTol.Enabled = True
		Me.lblVesselOffsetTol.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblVesselOffsetTol.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblVesselOffsetTol.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblVesselOffsetTol.UseMnemonic = True
		Me.lblVesselOffsetTol.Visible = True
		Me.lblVesselOffsetTol.AutoSize = False
		Me.lblVesselOffsetTol.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblVesselOffsetTol.Name = "lblVesselOffsetTol"
		Me.lblForceUnit.Text = "kips"
		Me.lblForceUnit.Size = New System.Drawing.Size(37, 16)
		Me.lblForceUnit.Location = New System.Drawing.Point(200, 101)
		Me.lblForceUnit.TabIndex = 20
		Me.lblForceUnit.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblForceUnit.BackColor = System.Drawing.SystemColors.Control
		Me.lblForceUnit.Enabled = True
		Me.lblForceUnit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblForceUnit.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblForceUnit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblForceUnit.UseMnemonic = True
		Me.lblForceUnit.Visible = True
		Me.lblForceUnit.AutoSize = False
		Me.lblForceUnit.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblForceUnit.Name = "lblForceUnit"
		Me._lblLengthUnit_1.Text = "ft"
		Me._lblLengthUnit_1.Size = New System.Drawing.Size(17, 16)
		Me._lblLengthUnit_1.Location = New System.Drawing.Point(200, 77)
		Me._lblLengthUnit_1.TabIndex = 17
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
		Me.lblAngleUnit.Text = "deg"
		Me.lblAngleUnit.Size = New System.Drawing.Size(37, 16)
		Me.lblAngleUnit.Location = New System.Drawing.Point(198, 54)
		Me.lblAngleUnit.TabIndex = 16
		Me.lblAngleUnit.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblAngleUnit.BackColor = System.Drawing.SystemColors.Control
		Me.lblAngleUnit.Enabled = True
		Me.lblAngleUnit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblAngleUnit.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblAngleUnit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblAngleUnit.UseMnemonic = True
		Me.lblAngleUnit.Visible = True
		Me.lblAngleUnit.AutoSize = False
		Me.lblAngleUnit.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblAngleUnit.Name = "lblAngleUnit"
		Me._lblLengthUnit_0.Text = "ft"
		Me._lblLengthUnit_0.Size = New System.Drawing.Size(17, 16)
		Me._lblLengthUnit_0.Location = New System.Drawing.Point(201, 31)
		Me._lblLengthUnit_0.TabIndex = 14
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
		Me.lblForceTol.Text = "Line Tension:"
		Me.lblForceTol.Size = New System.Drawing.Size(112, 16)
		Me.lblForceTol.Location = New System.Drawing.Point(34, 108)
		Me.lblForceTol.TabIndex = 12
		Me.lblForceTol.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblForceTol.BackColor = System.Drawing.SystemColors.Control
		Me.lblForceTol.Enabled = True
		Me.lblForceTol.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblForceTol.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblForceTol.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblForceTol.UseMnemonic = True
		Me.lblForceTol.Visible = True
		Me.lblForceTol.AutoSize = False
		Me.lblForceTol.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblForceTol.Name = "lblForceTol"
		Me.lblDepthTol.Text = "Anchor Depth:"
		Me.lblDepthTol.Size = New System.Drawing.Size(109, 16)
		Me.lblDepthTol.Location = New System.Drawing.Point(34, 81)
		Me.lblDepthTol.TabIndex = 11
		Me.lblDepthTol.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblDepthTol.BackColor = System.Drawing.SystemColors.Control
		Me.lblDepthTol.Enabled = True
		Me.lblDepthTol.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblDepthTol.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblDepthTol.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblDepthTol.UseMnemonic = True
		Me.lblDepthTol.Visible = True
		Me.lblDepthTol.AutoSize = False
		Me.lblDepthTol.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblDepthTol.Name = "lblDepthTol"
		Me.lblAngleTol.Text = "Spread Angle:"
		Me.lblAngleTol.Size = New System.Drawing.Size(94, 16)
		Me.lblAngleTol.Location = New System.Drawing.Point(34, 55)
		Me.lblAngleTol.TabIndex = 10
		Me.lblAngleTol.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblAngleTol.BackColor = System.Drawing.SystemColors.Control
		Me.lblAngleTol.Enabled = True
		Me.lblAngleTol.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblAngleTol.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblAngleTol.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblAngleTol.UseMnemonic = True
		Me.lblAngleTol.Visible = True
		Me.lblAngleTol.AutoSize = False
		Me.lblAngleTol.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblAngleTol.Name = "lblAngleTol"
		Me.lblDistTol.Text = "Scope:"
		Me.lblDistTol.Size = New System.Drawing.Size(65, 16)
		Me.lblDistTol.Location = New System.Drawing.Point(34, 29)
		Me.lblDistTol.TabIndex = 9
		Me.lblDistTol.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblDistTol.BackColor = System.Drawing.SystemColors.Control
		Me.lblDistTol.Enabled = True
		Me.lblDistTol.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblDistTol.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblDistTol.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblDistTol.UseMnemonic = True
		Me.lblDistTol.Visible = True
		Me.lblDistTol.AutoSize = False
		Me.lblDistTol.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblDistTol.Name = "lblDistTol"
		Me.btnOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnOK.Text = "OK"
		Me.btnOK.Size = New System.Drawing.Size(70, 25)
		Me.btnOK.Location = New System.Drawing.Point(155, 14)
		Me.btnOK.TabIndex = 4
		Me.ToolTip1.SetToolTip(Me.btnOK, "Save Final Position")
		Me.btnOK.BackColor = System.Drawing.SystemColors.Control
		Me.btnOK.CausesValidation = True
		Me.btnOK.Enabled = True
		Me.btnOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnOK.TabStop = True
		Me.btnOK.Name = "btnOK"
		Me.btnCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.btnCancel
		Me.btnCancel.Text = "Cancel"
		Me.btnCancel.Size = New System.Drawing.Size(70, 25)
		Me.btnCancel.Location = New System.Drawing.Point(155, 43)
		Me.btnCancel.TabIndex = 3
		Me.btnCancel.BackColor = System.Drawing.SystemColors.Control
		Me.btnCancel.CausesValidation = True
		Me.btnCancel.Enabled = True
		Me.btnCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnCancel.TabStop = True
		Me.btnCancel.Name = "btnCancel"
		Me.fraUnit.Text = "Unit System"
		Me.fraUnit.ForeColor = System.Drawing.Color.Black
		Me.fraUnit.Size = New System.Drawing.Size(142, 79)
		Me.fraUnit.Location = New System.Drawing.Point(7, 8)
		Me.fraUnit.TabIndex = 0
		Me.fraUnit.BackColor = System.Drawing.SystemColors.Control
		Me.fraUnit.Enabled = True
		Me.fraUnit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraUnit.Visible = True
		Me.fraUnit.Name = "fraUnit"
		Me._optUnit_1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optUnit_1.Text = "Imperial (English)"
		Me._optUnit_1.Size = New System.Drawing.Size(107, 22)
		Me._optUnit_1.Location = New System.Drawing.Point(20, 44)
		Me._optUnit_1.TabIndex = 2
		Me._optUnit_1.Checked = True
		Me._optUnit_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optUnit_1.BackColor = System.Drawing.SystemColors.Control
		Me._optUnit_1.CausesValidation = True
		Me._optUnit_1.Enabled = True
		Me._optUnit_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optUnit_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._optUnit_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optUnit_1.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optUnit_1.TabStop = True
		Me._optUnit_1.Visible = True
		Me._optUnit_1.Name = "_optUnit_1"
		Me._optUnit_0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optUnit_0.Text = "Metric (S.I.)"
		Me._optUnit_0.Size = New System.Drawing.Size(83, 22)
		Me._optUnit_0.Location = New System.Drawing.Point(20, 21)
		Me._optUnit_0.TabIndex = 1
		Me._optUnit_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optUnit_0.BackColor = System.Drawing.SystemColors.Control
		Me._optUnit_0.CausesValidation = True
		Me._optUnit_0.Enabled = True
		Me._optUnit_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optUnit_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._optUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optUnit_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optUnit_0.TabStop = True
		Me._optUnit_0.Checked = False
		Me._optUnit_0.Visible = True
		Me._optUnit_0.Name = "_optUnit_0"
		Me.Label5.Text = "Buoy Height:"
		Me.Label5.Size = New System.Drawing.Size(146, 16)
		Me.Label5.Location = New System.Drawing.Point(16, 457)
		Me.Label5.TabIndex = 38
		Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label5.BackColor = System.Drawing.SystemColors.Control
		Me.Label5.Enabled = True
		Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label5.UseMnemonic = True
		Me.Label5.Visible = True
		Me.Label5.AutoSize = False
		Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label5.Name = "Label5"
		Me._lblLengthUnit_4.Text = "ft"
		Me._lblLengthUnit_4.Size = New System.Drawing.Size(17, 16)
		Me._lblLengthUnit_4.Location = New System.Drawing.Point(256, 458)
		Me._lblLengthUnit_4.TabIndex = 37
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
		Me.Controls.Add(Text2)
		Me.Controls.Add(Frame1)
		Me.Controls.Add(fraIntegration)
		Me.Controls.Add(fraScopeRef)
		Me.Controls.Add(fraTol)
		Me.Controls.Add(btnOK)
		Me.Controls.Add(btnCancel)
		Me.Controls.Add(fraUnit)
		Me.Controls.Add(Label5)
		Me.Controls.Add(_lblLengthUnit_4)
		Me.Frame1.Controls.Add(txtNumLineSegments)
		Me.Frame1.Controls.Add(txtMinSegLength)
		Me.Frame1.Controls.Add(txtMaxIter)
		Me.Frame1.Controls.Add(Label4)
		Me.Frame1.Controls.Add(Label2)
		Me.Frame1.Controls.Add(_lblLengthUnit_3)
		Me.Frame1.Controls.Add(Label3)
		Me.fraIntegration.Controls.Add(txtStartFreq)
		Me.fraIntegration.Controls.Add(Text1)
		Me.fraIntegration.Controls.Add(Label10)
		Me.fraIntegration.Controls.Add(Label9)
		Me.fraIntegration.Controls.Add(Label8)
		Me.fraIntegration.Controls.Add(Label7)
		Me.fraIntegration.Controls.Add(Label6)
		Me.fraScopeRef.Controls.Add(_optScopeRef_1)
		Me.fraScopeRef.Controls.Add(_optScopeRef_0)
		Me.fraTol.Controls.Add(txtVesselHdgTol)
		Me.fraTol.Controls.Add(txtVesselOffsetTol)
		Me.fraTol.Controls.Add(txtLineTensionTol)
		Me.fraTol.Controls.Add(txtAnchorDepthTol)
		Me.fraTol.Controls.Add(txtSpreadAngleTol)
		Me.fraTol.Controls.Add(txtScopeTol)
		Me.fraTol.Controls.Add(Label1)
		Me.fraTol.Controls.Add(_lblLengthUnit_2)
		Me.fraTol.Controls.Add(lblVesselHdgTol)
		Me.fraTol.Controls.Add(lblVesselOffsetTol)
		Me.fraTol.Controls.Add(lblForceUnit)
		Me.fraTol.Controls.Add(_lblLengthUnit_1)
		Me.fraTol.Controls.Add(lblAngleUnit)
		Me.fraTol.Controls.Add(_lblLengthUnit_0)
		Me.fraTol.Controls.Add(lblForceTol)
		Me.fraTol.Controls.Add(lblDepthTol)
		Me.fraTol.Controls.Add(lblAngleTol)
		Me.fraTol.Controls.Add(lblDistTol)
		Me.fraUnit.Controls.Add(_optUnit_1)
		Me.fraUnit.Controls.Add(_optUnit_0)
		Me.lblLengthUnit.SetIndex(_lblLengthUnit_3, CType(3, Short))
		Me.lblLengthUnit.SetIndex(_lblLengthUnit_2, CType(2, Short))
		Me.lblLengthUnit.SetIndex(_lblLengthUnit_1, CType(1, Short))
		Me.lblLengthUnit.SetIndex(_lblLengthUnit_0, CType(0, Short))
		Me.lblLengthUnit.SetIndex(_lblLengthUnit_4, CType(4, Short))
		Me.optScopeRef.SetIndex(_optScopeRef_1, CType(1, Short))
		Me.optScopeRef.SetIndex(_optScopeRef_0, CType(0, Short))
		Me.optUnit.SetIndex(_optUnit_1, CType(1, Short))
		Me.optUnit.SetIndex(_optUnit_0, CType(0, Short))
		CType(Me.optUnit, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.optScopeRef, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lblLengthUnit, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Frame1.ResumeLayout(False)
		Me.fraIntegration.ResumeLayout(False)
		Me.fraScopeRef.ResumeLayout(False)
		Me.fraTol.ResumeLayout(False)
		Me.fraUnit.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class