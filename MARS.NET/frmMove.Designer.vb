<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmMove
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
	Public WithEvents btnReport As System.Windows.Forms.Button
	Public WithEvents txtClientName As System.Windows.Forms.TextBox
	Public WithEvents txtLocationName As System.Windows.Forms.TextBox
	Public WithEvents lblClientName As System.Windows.Forms.Label
	Public WithEvents lblLocation As System.Windows.Forms.Label
	Public WithEvents fraReport As System.Windows.Forms.GroupBox
	Public WithEvents _txtExtLoad_5 As System.Windows.Forms.TextBox
	Public WithEvents _txtExtLoad_4 As System.Windows.Forms.TextBox
	Public WithEvents _txtExtLoad_3 As System.Windows.Forms.TextBox
	Public WithEvents _lblForceUnit_6 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_8 As System.Windows.Forms.Label
	Public WithEvents _lblForceUnit_4 As System.Windows.Forms.Label
	Public WithEvents _lblForceUnit_3 As System.Windows.Forms.Label
	Public WithEvents _lblExtLoad_5 As System.Windows.Forms.Label
	Public WithEvents _lblExtLoad_4 As System.Windows.Forms.Label
	Public WithEvents _lblExtLoad_3 As System.Windows.Forms.Label
	Public WithEvents _lblExtUnits_3 As System.Windows.Forms.Label
	Public WithEvents _fraExternalLoad_1 As System.Windows.Forms.GroupBox
	Public WithEvents btnSave As System.Windows.Forms.Button
	Public WithEvents SysInfo1 As AxSysInfoLib.AxSysInfo
	Public WithEvents btnCancel As System.Windows.Forms.Button
	Public WithEvents _txtExtLoad_0 As System.Windows.Forms.TextBox
	Public WithEvents _txtExtLoad_1 As System.Windows.Forms.TextBox
	Public WithEvents _txtExtLoad_2 As System.Windows.Forms.TextBox
	Public WithEvents _lblLengthUnit_7 As System.Windows.Forms.Label
	Public WithEvents _lblForceUnit_2 As System.Windows.Forms.Label
	Public WithEvents _lblForceUnit_1 As System.Windows.Forms.Label
	Public WithEvents _lblExtUnits_2 As System.Windows.Forms.Label
	Public WithEvents _lblForceUnit_0 As System.Windows.Forms.Label
	Public WithEvents _lblExtLoad_0 As System.Windows.Forms.Label
	Public WithEvents _lblExtLoad_1 As System.Windows.Forms.Label
	Public WithEvents _lblExtLoad_2 As System.Windows.Forms.Label
	Public WithEvents _fraExternalLoad_0 As System.Windows.Forms.GroupBox
	Public WithEvents _btnPayout_0 As System.Windows.Forms.Button
	Public WithEvents _btnPayout_1 As System.Windows.Forms.Button
	Public WithEvents grdPayout As AxMSFlexGridLib.AxMSFlexGrid
	Public WithEvents fraPayout As System.Windows.Forms.GroupBox
	Public WithEvents txtConditions As System.Windows.Forms.TextBox
	Public WithEvents grdLC As AxMSFlexGridLib.AxMSFlexGrid
	Public WithEvents _lblForceUnit_5 As System.Windows.Forms.Label
	Public WithEvents lblConditions As System.Windows.Forms.Label
	Public WithEvents fraConditions As System.Windows.Forms.GroupBox
	Public WithEvents btnPosition As System.Windows.Forms.Button
	Public WithEvents _txtVslSt_9 As System.Windows.Forms.TextBox
	Public WithEvents _txtVslSt_8 As System.Windows.Forms.TextBox
	Public WithEvents _cboWells_1 As System.Windows.Forms.ComboBox
	Public WithEvents _optInputSystem_3 As System.Windows.Forms.RadioButton
	Public WithEvents _optInputSystem_2 As System.Windows.Forms.RadioButton
	Public WithEvents _txtVslSt_10 As System.Windows.Forms.TextBox
	Public WithEvents _txtVslSt_7 As System.Windows.Forms.TextBox
	Public WithEvents _txtVslSt_6 As System.Windows.Forms.TextBox
	Public WithEvents _lblLengthUnit_6 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_5 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_4 As System.Windows.Forms.Label
	Public WithEvents _lblVslStUnit_9 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_7 As System.Windows.Forms.Label
	Public WithEvents _lblVslStUnit_10 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_8 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_6 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_5 As System.Windows.Forms.Label
	Public WithEvents fraVesselLoc1 As System.Windows.Forms.GroupBox
	Public WithEvents _optInputSystem_0 As System.Windows.Forms.RadioButton
	Public WithEvents _cboWells_0 As System.Windows.Forms.ComboBox
	Public WithEvents _txtVslSt_2 As System.Windows.Forms.TextBox
	Public WithEvents _txtVslSt_3 As System.Windows.Forms.TextBox
	Public WithEvents _txtVslSt_0 As System.Windows.Forms.TextBox
	Public WithEvents _txtVslSt_1 As System.Windows.Forms.TextBox
	Public WithEvents _txtVslSt_4 As System.Windows.Forms.TextBox
	Public WithEvents _txtVslSt_5 As System.Windows.Forms.TextBox
	Public WithEvents _optInputSystem_1 As System.Windows.Forms.RadioButton
	Public WithEvents _lblLengthUnit_3 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_2 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_1 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_2 As System.Windows.Forms.Label
	Public WithEvents _lblVslStUnit_3 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_4 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_0 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_0 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_1 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_3 As System.Windows.Forms.Label
	Public WithEvents _lblVslStUnit_4 As System.Windows.Forms.Label
	Public WithEvents fraVesselLoc As System.Windows.Forms.GroupBox
	Public WithEvents _lblForceUnit_7 As System.Windows.Forms.Label
	Public WithEvents _lblVelUnit_0 As System.Windows.Forms.Label
	Public WithEvents _lblGenCmt_3 As System.Windows.Forms.Label
	Public WithEvents _lblGenCmt_2 As System.Windows.Forms.Label
	Public WithEvents _lblGenCmt_1 As System.Windows.Forms.Label
	Public WithEvents _lblGenCmt_0 As System.Windows.Forms.Label
	Public WithEvents btnPayout As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
	Public WithEvents cboWells As Microsoft.VisualBasic.Compatibility.VB6.ComboBoxArray
	Public WithEvents fraExternalLoad As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
	Public WithEvents lblExtLoad As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblExtUnits As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblForceUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblGenCmt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblLengthUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblVelUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblVslSt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblVslStUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents optInputSystem As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	Public WithEvents txtExtLoad As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	Public WithEvents txtVslSt As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMove))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.fraReport = New System.Windows.Forms.GroupBox
		Me.btnReport = New System.Windows.Forms.Button
		Me.txtClientName = New System.Windows.Forms.TextBox
		Me.txtLocationName = New System.Windows.Forms.TextBox
		Me.lblClientName = New System.Windows.Forms.Label
		Me.lblLocation = New System.Windows.Forms.Label
		Me._fraExternalLoad_1 = New System.Windows.Forms.GroupBox
		Me._txtExtLoad_5 = New System.Windows.Forms.TextBox
		Me._txtExtLoad_4 = New System.Windows.Forms.TextBox
		Me._txtExtLoad_3 = New System.Windows.Forms.TextBox
		Me._lblForceUnit_6 = New System.Windows.Forms.Label
		Me._lblLengthUnit_8 = New System.Windows.Forms.Label
		Me._lblForceUnit_4 = New System.Windows.Forms.Label
		Me._lblForceUnit_3 = New System.Windows.Forms.Label
		Me._lblExtLoad_5 = New System.Windows.Forms.Label
		Me._lblExtLoad_4 = New System.Windows.Forms.Label
		Me._lblExtLoad_3 = New System.Windows.Forms.Label
		Me._lblExtUnits_3 = New System.Windows.Forms.Label
		Me.btnSave = New System.Windows.Forms.Button
		Me.SysInfo1 = New AxSysInfoLib.AxSysInfo
		Me.btnCancel = New System.Windows.Forms.Button
		Me._fraExternalLoad_0 = New System.Windows.Forms.GroupBox
		Me._txtExtLoad_0 = New System.Windows.Forms.TextBox
		Me._txtExtLoad_1 = New System.Windows.Forms.TextBox
		Me._txtExtLoad_2 = New System.Windows.Forms.TextBox
		Me._lblLengthUnit_7 = New System.Windows.Forms.Label
		Me._lblForceUnit_2 = New System.Windows.Forms.Label
		Me._lblForceUnit_1 = New System.Windows.Forms.Label
		Me._lblExtUnits_2 = New System.Windows.Forms.Label
		Me._lblForceUnit_0 = New System.Windows.Forms.Label
		Me._lblExtLoad_0 = New System.Windows.Forms.Label
		Me._lblExtLoad_1 = New System.Windows.Forms.Label
		Me._lblExtLoad_2 = New System.Windows.Forms.Label
		Me.fraPayout = New System.Windows.Forms.GroupBox
		Me._btnPayout_0 = New System.Windows.Forms.Button
		Me._btnPayout_1 = New System.Windows.Forms.Button
		Me.grdPayout = New AxMSFlexGridLib.AxMSFlexGrid
		Me.fraConditions = New System.Windows.Forms.GroupBox
		Me.txtConditions = New System.Windows.Forms.TextBox
		Me.grdLC = New AxMSFlexGridLib.AxMSFlexGrid
		Me._lblForceUnit_5 = New System.Windows.Forms.Label
		Me.lblConditions = New System.Windows.Forms.Label
		Me.fraVesselLoc1 = New System.Windows.Forms.GroupBox
		Me.btnPosition = New System.Windows.Forms.Button
		Me._txtVslSt_9 = New System.Windows.Forms.TextBox
		Me._txtVslSt_8 = New System.Windows.Forms.TextBox
		Me._cboWells_1 = New System.Windows.Forms.ComboBox
		Me._optInputSystem_3 = New System.Windows.Forms.RadioButton
		Me._optInputSystem_2 = New System.Windows.Forms.RadioButton
		Me._txtVslSt_10 = New System.Windows.Forms.TextBox
		Me._txtVslSt_7 = New System.Windows.Forms.TextBox
		Me._txtVslSt_6 = New System.Windows.Forms.TextBox
		Me._lblLengthUnit_6 = New System.Windows.Forms.Label
		Me._lblLengthUnit_5 = New System.Windows.Forms.Label
		Me._lblLengthUnit_4 = New System.Windows.Forms.Label
		Me._lblVslStUnit_9 = New System.Windows.Forms.Label
		Me._lblVslSt_7 = New System.Windows.Forms.Label
		Me._lblVslStUnit_10 = New System.Windows.Forms.Label
		Me._lblVslSt_8 = New System.Windows.Forms.Label
		Me._lblVslSt_6 = New System.Windows.Forms.Label
		Me._lblVslSt_5 = New System.Windows.Forms.Label
		Me.fraVesselLoc = New System.Windows.Forms.GroupBox
		Me._optInputSystem_0 = New System.Windows.Forms.RadioButton
		Me._cboWells_0 = New System.Windows.Forms.ComboBox
		Me._txtVslSt_2 = New System.Windows.Forms.TextBox
		Me._txtVslSt_3 = New System.Windows.Forms.TextBox
		Me._txtVslSt_0 = New System.Windows.Forms.TextBox
		Me._txtVslSt_1 = New System.Windows.Forms.TextBox
		Me._txtVslSt_4 = New System.Windows.Forms.TextBox
		Me._txtVslSt_5 = New System.Windows.Forms.TextBox
		Me._optInputSystem_1 = New System.Windows.Forms.RadioButton
		Me._lblLengthUnit_3 = New System.Windows.Forms.Label
		Me._lblLengthUnit_2 = New System.Windows.Forms.Label
		Me._lblLengthUnit_1 = New System.Windows.Forms.Label
		Me._lblVslSt_2 = New System.Windows.Forms.Label
		Me._lblVslStUnit_3 = New System.Windows.Forms.Label
		Me._lblVslSt_4 = New System.Windows.Forms.Label
		Me._lblLengthUnit_0 = New System.Windows.Forms.Label
		Me._lblVslSt_0 = New System.Windows.Forms.Label
		Me._lblVslSt_1 = New System.Windows.Forms.Label
		Me._lblVslSt_3 = New System.Windows.Forms.Label
		Me._lblVslStUnit_4 = New System.Windows.Forms.Label
		Me._lblForceUnit_7 = New System.Windows.Forms.Label
		Me._lblVelUnit_0 = New System.Windows.Forms.Label
		Me._lblGenCmt_3 = New System.Windows.Forms.Label
		Me._lblGenCmt_2 = New System.Windows.Forms.Label
		Me._lblGenCmt_1 = New System.Windows.Forms.Label
		Me._lblGenCmt_0 = New System.Windows.Forms.Label
		Me.btnPayout = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(components)
		Me.cboWells = New Microsoft.VisualBasic.Compatibility.VB6.ComboBoxArray(components)
		Me.fraExternalLoad = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(components)
		Me.lblExtLoad = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.lblExtUnits = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.lblForceUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.lblGenCmt = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.lblLengthUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.lblVelUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.lblVslSt = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.lblVslStUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.optInputSystem = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(components)
		Me.txtExtLoad = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(components)
		Me.txtVslSt = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(components)
		Me.fraReport.SuspendLayout()
		Me._fraExternalLoad_1.SuspendLayout()
		Me._fraExternalLoad_0.SuspendLayout()
		Me.fraPayout.SuspendLayout()
		Me.fraConditions.SuspendLayout()
		Me.fraVesselLoc1.SuspendLayout()
		Me.fraVesselLoc.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.SysInfo1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.grdPayout, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.grdLC, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.btnPayout, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.cboWells, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.fraExternalLoad, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lblExtLoad, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lblExtUnits, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lblForceUnit, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lblGenCmt, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lblLengthUnit, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lblVelUnit, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lblVslSt, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lblVslStUnit, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.optInputSystem, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.txtExtLoad, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.txtVslSt, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.Text = " Move Around"
		Me.ClientSize = New System.Drawing.Size(816, 587)
		Me.Location = New System.Drawing.Point(4, 23)
		Me.Icon = CType(resources.GetObject("frmMove.Icon"), System.Drawing.Icon)
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
		Me.Name = "frmMove"
		Me.fraReport.Text = "Report Titles"
		Me.fraReport.Size = New System.Drawing.Size(429, 78)
		Me.fraReport.Location = New System.Drawing.Point(308, 495)
		Me.fraReport.TabIndex = 79
		Me.fraReport.BackColor = System.Drawing.SystemColors.Control
		Me.fraReport.Enabled = True
		Me.fraReport.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraReport.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraReport.Visible = True
		Me.fraReport.Name = "fraReport"
		Me.btnReport.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnReport.Text = "&Report"
		Me.btnReport.Size = New System.Drawing.Size(65, 28)
		Me.btnReport.Location = New System.Drawing.Point(353, 18)
		Me.btnReport.TabIndex = 84
		Me.btnReport.BackColor = System.Drawing.SystemColors.Control
		Me.btnReport.CausesValidation = True
		Me.btnReport.Enabled = True
		Me.btnReport.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnReport.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnReport.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnReport.TabStop = True
		Me.btnReport.Name = "btnReport"
		Me.txtClientName.AutoSize = False
		Me.txtClientName.Size = New System.Drawing.Size(240, 19)
		Me.txtClientName.Location = New System.Drawing.Point(98, 20)
		Me.txtClientName.TabIndex = 81
		Me.txtClientName.TabStop = False
		Me.txtClientName.Text = "Test Client "
		Me.txtClientName.AcceptsReturn = True
		Me.txtClientName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtClientName.BackColor = System.Drawing.SystemColors.Window
		Me.txtClientName.CausesValidation = True
		Me.txtClientName.Enabled = True
		Me.txtClientName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtClientName.HideSelection = True
		Me.txtClientName.ReadOnly = False
		Me.txtClientName.Maxlength = 0
		Me.txtClientName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtClientName.MultiLine = False
		Me.txtClientName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtClientName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtClientName.Visible = True
		Me.txtClientName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtClientName.Name = "txtClientName"
		Me.txtLocationName.AutoSize = False
		Me.txtLocationName.Size = New System.Drawing.Size(240, 19)
		Me.txtLocationName.Location = New System.Drawing.Point(98, 44)
		Me.txtLocationName.TabIndex = 80
		Me.txtLocationName.TabStop = False
		Me.txtLocationName.Text = "Test Location "
		Me.txtLocationName.AcceptsReturn = True
		Me.txtLocationName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtLocationName.BackColor = System.Drawing.SystemColors.Window
		Me.txtLocationName.CausesValidation = True
		Me.txtLocationName.Enabled = True
		Me.txtLocationName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtLocationName.HideSelection = True
		Me.txtLocationName.ReadOnly = False
		Me.txtLocationName.Maxlength = 0
		Me.txtLocationName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtLocationName.MultiLine = False
		Me.txtLocationName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtLocationName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtLocationName.Visible = True
		Me.txtLocationName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtLocationName.Name = "txtLocationName"
		Me.lblClientName.Text = "Prepared for"
		Me.lblClientName.Size = New System.Drawing.Size(73, 17)
		Me.lblClientName.Location = New System.Drawing.Point(15, 24)
		Me.lblClientName.TabIndex = 83
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
		Me.lblLocation.Text = "Location Name:"
		Me.lblLocation.Size = New System.Drawing.Size(112, 17)
		Me.lblLocation.Location = New System.Drawing.Point(14, 48)
		Me.lblLocation.TabIndex = 82
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
		Me._fraExternalLoad_1.Text = "Mooring System Forces after Payout Adjustment"
		Me._fraExternalLoad_1.Size = New System.Drawing.Size(361, 89)
		Me._fraExternalLoad_1.Location = New System.Drawing.Point(375, 134)
		Me._fraExternalLoad_1.TabIndex = 45
		Me._fraExternalLoad_1.BackColor = System.Drawing.SystemColors.Control
		Me._fraExternalLoad_1.Enabled = True
		Me._fraExternalLoad_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._fraExternalLoad_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._fraExternalLoad_1.Visible = True
		Me._fraExternalLoad_1.Name = "_fraExternalLoad_1"
		Me._txtExtLoad_5.AutoSize = False
		Me._txtExtLoad_5.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me._txtExtLoad_5.BackColor = System.Drawing.SystemColors.Control
		Me._txtExtLoad_5.Size = New System.Drawing.Size(65, 19)
		Me._txtExtLoad_5.Location = New System.Drawing.Point(56, 56)
		Me._txtExtLoad_5.ReadOnly = True
		Me._txtExtLoad_5.TabIndex = 51
		Me._txtExtLoad_5.TabStop = False
		Me._txtExtLoad_5.Text = "0"
		Me._txtExtLoad_5.AcceptsReturn = True
		Me._txtExtLoad_5.CausesValidation = True
		Me._txtExtLoad_5.Enabled = True
		Me._txtExtLoad_5.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtExtLoad_5.HideSelection = True
		Me._txtExtLoad_5.Maxlength = 0
		Me._txtExtLoad_5.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtExtLoad_5.MultiLine = False
		Me._txtExtLoad_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtExtLoad_5.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtExtLoad_5.Visible = True
		Me._txtExtLoad_5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtExtLoad_5.Name = "_txtExtLoad_5"
		Me._txtExtLoad_4.AutoSize = False
		Me._txtExtLoad_4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me._txtExtLoad_4.BackColor = System.Drawing.SystemColors.Control
		Me._txtExtLoad_4.Size = New System.Drawing.Size(65, 19)
		Me._txtExtLoad_4.Location = New System.Drawing.Point(216, 24)
		Me._txtExtLoad_4.ReadOnly = True
		Me._txtExtLoad_4.TabIndex = 49
		Me._txtExtLoad_4.TabStop = False
		Me._txtExtLoad_4.Text = "0"
		Me._txtExtLoad_4.AcceptsReturn = True
		Me._txtExtLoad_4.CausesValidation = True
		Me._txtExtLoad_4.Enabled = True
		Me._txtExtLoad_4.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtExtLoad_4.HideSelection = True
		Me._txtExtLoad_4.Maxlength = 0
		Me._txtExtLoad_4.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtExtLoad_4.MultiLine = False
		Me._txtExtLoad_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtExtLoad_4.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtExtLoad_4.Visible = True
		Me._txtExtLoad_4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtExtLoad_4.Name = "_txtExtLoad_4"
		Me._txtExtLoad_3.AutoSize = False
		Me._txtExtLoad_3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me._txtExtLoad_3.BackColor = System.Drawing.SystemColors.Control
		Me._txtExtLoad_3.Size = New System.Drawing.Size(65, 19)
		Me._txtExtLoad_3.Location = New System.Drawing.Point(56, 24)
		Me._txtExtLoad_3.ReadOnly = True
		Me._txtExtLoad_3.TabIndex = 47
		Me._txtExtLoad_3.TabStop = False
		Me._txtExtLoad_3.Text = "0"
		Me._txtExtLoad_3.AcceptsReturn = True
		Me._txtExtLoad_3.CausesValidation = True
		Me._txtExtLoad_3.Enabled = True
		Me._txtExtLoad_3.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtExtLoad_3.HideSelection = True
		Me._txtExtLoad_3.Maxlength = 0
		Me._txtExtLoad_3.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtExtLoad_3.MultiLine = False
		Me._txtExtLoad_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtExtLoad_3.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtExtLoad_3.Visible = True
		Me._txtExtLoad_3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtExtLoad_3.Name = "_txtExtLoad_3"
		Me._lblForceUnit_6.Text = "kips"
		Me._lblForceUnit_6.Size = New System.Drawing.Size(22, 17)
		Me._lblForceUnit_6.Location = New System.Drawing.Point(127, 56)
		Me._lblForceUnit_6.TabIndex = 78
		Me._lblForceUnit_6.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblForceUnit_6.BackColor = System.Drawing.SystemColors.Control
		Me._lblForceUnit_6.Enabled = True
		Me._lblForceUnit_6.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblForceUnit_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblForceUnit_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblForceUnit_6.UseMnemonic = True
		Me._lblForceUnit_6.Visible = True
		Me._lblForceUnit_6.AutoSize = False
		Me._lblForceUnit_6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblForceUnit_6.Name = "_lblForceUnit_6"
		Me._lblLengthUnit_8.Text = "ft"
		Me._lblLengthUnit_8.Size = New System.Drawing.Size(17, 17)
		Me._lblLengthUnit_8.Location = New System.Drawing.Point(161, 57)
		Me._lblLengthUnit_8.TabIndex = 77
		Me._lblLengthUnit_8.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblLengthUnit_8.BackColor = System.Drawing.SystemColors.Control
		Me._lblLengthUnit_8.Enabled = True
		Me._lblLengthUnit_8.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblLengthUnit_8.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblLengthUnit_8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblLengthUnit_8.UseMnemonic = True
		Me._lblLengthUnit_8.Visible = True
		Me._lblLengthUnit_8.AutoSize = False
		Me._lblLengthUnit_8.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblLengthUnit_8.Name = "_lblLengthUnit_8"
		Me._lblForceUnit_4.Text = "kips"
		Me._lblForceUnit_4.Size = New System.Drawing.Size(25, 17)
		Me._lblForceUnit_4.Location = New System.Drawing.Point(285, 26)
		Me._lblForceUnit_4.TabIndex = 74
		Me._lblForceUnit_4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblForceUnit_4.BackColor = System.Drawing.SystemColors.Control
		Me._lblForceUnit_4.Enabled = True
		Me._lblForceUnit_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblForceUnit_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblForceUnit_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblForceUnit_4.UseMnemonic = True
		Me._lblForceUnit_4.Visible = True
		Me._lblForceUnit_4.AutoSize = False
		Me._lblForceUnit_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblForceUnit_4.Name = "_lblForceUnit_4"
		Me._lblForceUnit_3.Text = "kips"
		Me._lblForceUnit_3.Size = New System.Drawing.Size(25, 17)
		Me._lblForceUnit_3.Location = New System.Drawing.Point(125, 25)
		Me._lblForceUnit_3.TabIndex = 73
		Me._lblForceUnit_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblForceUnit_3.BackColor = System.Drawing.SystemColors.Control
		Me._lblForceUnit_3.Enabled = True
		Me._lblForceUnit_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblForceUnit_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblForceUnit_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblForceUnit_3.UseMnemonic = True
		Me._lblForceUnit_3.Visible = True
		Me._lblForceUnit_3.AutoSize = False
		Me._lblForceUnit_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblForceUnit_3.Name = "_lblForceUnit_3"
		Me._lblExtLoad_5.Text = "Yaw"
		Me._lblExtLoad_5.Size = New System.Drawing.Size(33, 17)
		Me._lblExtLoad_5.Location = New System.Drawing.Point(16, 56)
		Me._lblExtLoad_5.TabIndex = 50
		Me._lblExtLoad_5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblExtLoad_5.BackColor = System.Drawing.SystemColors.Control
		Me._lblExtLoad_5.Enabled = True
		Me._lblExtLoad_5.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblExtLoad_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblExtLoad_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblExtLoad_5.UseMnemonic = True
		Me._lblExtLoad_5.Visible = True
		Me._lblExtLoad_5.AutoSize = False
		Me._lblExtLoad_5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblExtLoad_5.Name = "_lblExtLoad_5"
		Me._lblExtLoad_4.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._lblExtLoad_4.Text = "y (N) "
		Me._lblExtLoad_4.Size = New System.Drawing.Size(33, 17)
		Me._lblExtLoad_4.Location = New System.Drawing.Point(176, 24)
		Me._lblExtLoad_4.TabIndex = 48
		Me._lblExtLoad_4.BackColor = System.Drawing.SystemColors.Control
		Me._lblExtLoad_4.Enabled = True
		Me._lblExtLoad_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblExtLoad_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblExtLoad_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblExtLoad_4.UseMnemonic = True
		Me._lblExtLoad_4.Visible = True
		Me._lblExtLoad_4.AutoSize = False
		Me._lblExtLoad_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblExtLoad_4.Name = "_lblExtLoad_4"
		Me._lblExtLoad_3.Text = "x (E)"
		Me._lblExtLoad_3.Size = New System.Drawing.Size(33, 17)
		Me._lblExtLoad_3.Location = New System.Drawing.Point(16, 24)
		Me._lblExtLoad_3.TabIndex = 46
		Me._lblExtLoad_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblExtLoad_3.BackColor = System.Drawing.SystemColors.Control
		Me._lblExtLoad_3.Enabled = True
		Me._lblExtLoad_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblExtLoad_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblExtLoad_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblExtLoad_3.UseMnemonic = True
		Me._lblExtLoad_3.Visible = True
		Me._lblExtLoad_3.AutoSize = False
		Me._lblExtLoad_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblExtLoad_3.Name = "_lblExtLoad_3"
		Me._lblExtUnits_3.Text = "-"
		Me._lblExtUnits_3.Size = New System.Drawing.Size(3, 13)
		Me._lblExtUnits_3.Location = New System.Drawing.Point(152, 56)
		Me._lblExtUnits_3.TabIndex = 52
		Me._lblExtUnits_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblExtUnits_3.BackColor = System.Drawing.SystemColors.Control
		Me._lblExtUnits_3.Enabled = True
		Me._lblExtUnits_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblExtUnits_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblExtUnits_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblExtUnits_3.UseMnemonic = True
		Me._lblExtUnits_3.Visible = True
		Me._lblExtUnits_3.AutoSize = False
		Me._lblExtUnits_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblExtUnits_3.Name = "_lblExtUnits_3"
		Me.btnSave.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnSave.Text = "&Save"
		Me.btnSave.Size = New System.Drawing.Size(65, 28)
		Me.btnSave.Location = New System.Drawing.Point(744, 46)
		Me.btnSave.TabIndex = 0
		Me.ToolTip1.SetToolTip(Me.btnSave, "Save Payout and Station")
		Me.btnSave.BackColor = System.Drawing.SystemColors.Control
		Me.btnSave.CausesValidation = True
		Me.btnSave.Enabled = True
		Me.btnSave.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnSave.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnSave.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnSave.TabStop = True
		Me.btnSave.Name = "btnSave"
		SysInfo1.OcxState = CType(resources.GetObject("SysInfo1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.SysInfo1.Location = New System.Drawing.Point(440, 176)
		Me.SysInfo1.Name = "SysInfo1"
		Me.btnCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnCancel.Text = "Cancel"
		Me.btnCancel.Size = New System.Drawing.Size(65, 28)
		Me.btnCancel.Location = New System.Drawing.Point(744, 14)
		Me.btnCancel.TabIndex = 1
		Me.btnCancel.BackColor = System.Drawing.SystemColors.Control
		Me.btnCancel.CausesValidation = True
		Me.btnCancel.Enabled = True
		Me.btnCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnCancel.TabStop = True
		Me.btnCancel.Name = "btnCancel"
		Me._fraExternalLoad_0.Text = "Current Mooring System Forces"
		Me._fraExternalLoad_0.Size = New System.Drawing.Size(359, 89)
		Me._fraExternalLoad_0.Location = New System.Drawing.Point(8, 135)
		Me._fraExternalLoad_0.TabIndex = 36
		Me._fraExternalLoad_0.BackColor = System.Drawing.SystemColors.Control
		Me._fraExternalLoad_0.Enabled = True
		Me._fraExternalLoad_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._fraExternalLoad_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._fraExternalLoad_0.Visible = True
		Me._fraExternalLoad_0.Name = "_fraExternalLoad_0"
		Me._txtExtLoad_0.AutoSize = False
		Me._txtExtLoad_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me._txtExtLoad_0.BackColor = System.Drawing.SystemColors.Control
		Me._txtExtLoad_0.Size = New System.Drawing.Size(65, 19)
		Me._txtExtLoad_0.Location = New System.Drawing.Point(56, 24)
		Me._txtExtLoad_0.ReadOnly = True
		Me._txtExtLoad_0.TabIndex = 38
		Me._txtExtLoad_0.TabStop = False
		Me._txtExtLoad_0.Text = "0"
		Me._txtExtLoad_0.AcceptsReturn = True
		Me._txtExtLoad_0.CausesValidation = True
		Me._txtExtLoad_0.Enabled = True
		Me._txtExtLoad_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtExtLoad_0.HideSelection = True
		Me._txtExtLoad_0.Maxlength = 0
		Me._txtExtLoad_0.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtExtLoad_0.MultiLine = False
		Me._txtExtLoad_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtExtLoad_0.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtExtLoad_0.Visible = True
		Me._txtExtLoad_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtExtLoad_0.Name = "_txtExtLoad_0"
		Me._txtExtLoad_1.AutoSize = False
		Me._txtExtLoad_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me._txtExtLoad_1.BackColor = System.Drawing.SystemColors.Control
		Me._txtExtLoad_1.Size = New System.Drawing.Size(65, 19)
		Me._txtExtLoad_1.Location = New System.Drawing.Point(216, 24)
		Me._txtExtLoad_1.ReadOnly = True
		Me._txtExtLoad_1.TabIndex = 41
		Me._txtExtLoad_1.TabStop = False
		Me._txtExtLoad_1.Text = "0"
		Me._txtExtLoad_1.AcceptsReturn = True
		Me._txtExtLoad_1.CausesValidation = True
		Me._txtExtLoad_1.Enabled = True
		Me._txtExtLoad_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtExtLoad_1.HideSelection = True
		Me._txtExtLoad_1.Maxlength = 0
		Me._txtExtLoad_1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtExtLoad_1.MultiLine = False
		Me._txtExtLoad_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtExtLoad_1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtExtLoad_1.Visible = True
		Me._txtExtLoad_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtExtLoad_1.Name = "_txtExtLoad_1"
		Me._txtExtLoad_2.AutoSize = False
		Me._txtExtLoad_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me._txtExtLoad_2.BackColor = System.Drawing.SystemColors.Control
		Me._txtExtLoad_2.Size = New System.Drawing.Size(65, 19)
		Me._txtExtLoad_2.Location = New System.Drawing.Point(56, 56)
		Me._txtExtLoad_2.ReadOnly = True
		Me._txtExtLoad_2.TabIndex = 43
		Me._txtExtLoad_2.TabStop = False
		Me._txtExtLoad_2.Text = "0"
		Me._txtExtLoad_2.AcceptsReturn = True
		Me._txtExtLoad_2.CausesValidation = True
		Me._txtExtLoad_2.Enabled = True
		Me._txtExtLoad_2.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtExtLoad_2.HideSelection = True
		Me._txtExtLoad_2.Maxlength = 0
		Me._txtExtLoad_2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtExtLoad_2.MultiLine = False
		Me._txtExtLoad_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtExtLoad_2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtExtLoad_2.Visible = True
		Me._txtExtLoad_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtExtLoad_2.Name = "_txtExtLoad_2"
		Me._lblLengthUnit_7.Text = "ft"
		Me._lblLengthUnit_7.Size = New System.Drawing.Size(17, 17)
		Me._lblLengthUnit_7.Location = New System.Drawing.Point(159, 58)
		Me._lblLengthUnit_7.TabIndex = 76
		Me._lblLengthUnit_7.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblLengthUnit_7.BackColor = System.Drawing.SystemColors.Control
		Me._lblLengthUnit_7.Enabled = True
		Me._lblLengthUnit_7.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblLengthUnit_7.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblLengthUnit_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblLengthUnit_7.UseMnemonic = True
		Me._lblLengthUnit_7.Visible = True
		Me._lblLengthUnit_7.AutoSize = False
		Me._lblLengthUnit_7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblLengthUnit_7.Name = "_lblLengthUnit_7"
		Me._lblForceUnit_2.Text = "kips"
		Me._lblForceUnit_2.Size = New System.Drawing.Size(25, 17)
		Me._lblForceUnit_2.Location = New System.Drawing.Point(128, 57)
		Me._lblForceUnit_2.TabIndex = 72
		Me._lblForceUnit_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblForceUnit_2.BackColor = System.Drawing.SystemColors.Control
		Me._lblForceUnit_2.Enabled = True
		Me._lblForceUnit_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblForceUnit_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblForceUnit_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblForceUnit_2.UseMnemonic = True
		Me._lblForceUnit_2.Visible = True
		Me._lblForceUnit_2.AutoSize = False
		Me._lblForceUnit_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblForceUnit_2.Name = "_lblForceUnit_2"
		Me._lblForceUnit_1.Text = "kips"
		Me._lblForceUnit_1.Size = New System.Drawing.Size(25, 17)
		Me._lblForceUnit_1.Location = New System.Drawing.Point(284, 25)
		Me._lblForceUnit_1.TabIndex = 71
		Me._lblForceUnit_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblForceUnit_1.BackColor = System.Drawing.SystemColors.Control
		Me._lblForceUnit_1.Enabled = True
		Me._lblForceUnit_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblForceUnit_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblForceUnit_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblForceUnit_1.UseMnemonic = True
		Me._lblForceUnit_1.Visible = True
		Me._lblForceUnit_1.AutoSize = False
		Me._lblForceUnit_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblForceUnit_1.Name = "_lblForceUnit_1"
		Me._lblExtUnits_2.Text = "-"
		Me._lblExtUnits_2.Size = New System.Drawing.Size(3, 13)
		Me._lblExtUnits_2.Location = New System.Drawing.Point(152, 58)
		Me._lblExtUnits_2.TabIndex = 44
		Me._lblExtUnits_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblExtUnits_2.BackColor = System.Drawing.SystemColors.Control
		Me._lblExtUnits_2.Enabled = True
		Me._lblExtUnits_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblExtUnits_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblExtUnits_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblExtUnits_2.UseMnemonic = True
		Me._lblExtUnits_2.Visible = True
		Me._lblExtUnits_2.AutoSize = False
		Me._lblExtUnits_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblExtUnits_2.Name = "_lblExtUnits_2"
		Me._lblForceUnit_0.Text = "kips"
		Me._lblForceUnit_0.Size = New System.Drawing.Size(25, 17)
		Me._lblForceUnit_0.Location = New System.Drawing.Point(128, 24)
		Me._lblForceUnit_0.TabIndex = 39
		Me._lblForceUnit_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblForceUnit_0.BackColor = System.Drawing.SystemColors.Control
		Me._lblForceUnit_0.Enabled = True
		Me._lblForceUnit_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblForceUnit_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblForceUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblForceUnit_0.UseMnemonic = True
		Me._lblForceUnit_0.Visible = True
		Me._lblForceUnit_0.AutoSize = False
		Me._lblForceUnit_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblForceUnit_0.Name = "_lblForceUnit_0"
		Me._lblExtLoad_0.Text = "x (E)"
		Me._lblExtLoad_0.Size = New System.Drawing.Size(33, 17)
		Me._lblExtLoad_0.Location = New System.Drawing.Point(16, 24)
		Me._lblExtLoad_0.TabIndex = 37
		Me._lblExtLoad_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblExtLoad_0.BackColor = System.Drawing.SystemColors.Control
		Me._lblExtLoad_0.Enabled = True
		Me._lblExtLoad_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblExtLoad_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblExtLoad_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblExtLoad_0.UseMnemonic = True
		Me._lblExtLoad_0.Visible = True
		Me._lblExtLoad_0.AutoSize = False
		Me._lblExtLoad_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblExtLoad_0.Name = "_lblExtLoad_0"
		Me._lblExtLoad_1.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._lblExtLoad_1.Text = "y (N) "
		Me._lblExtLoad_1.Size = New System.Drawing.Size(33, 17)
		Me._lblExtLoad_1.Location = New System.Drawing.Point(176, 24)
		Me._lblExtLoad_1.TabIndex = 40
		Me._lblExtLoad_1.BackColor = System.Drawing.SystemColors.Control
		Me._lblExtLoad_1.Enabled = True
		Me._lblExtLoad_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblExtLoad_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblExtLoad_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblExtLoad_1.UseMnemonic = True
		Me._lblExtLoad_1.Visible = True
		Me._lblExtLoad_1.AutoSize = False
		Me._lblExtLoad_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblExtLoad_1.Name = "_lblExtLoad_1"
		Me._lblExtLoad_2.Text = "Yaw"
		Me._lblExtLoad_2.Size = New System.Drawing.Size(33, 17)
		Me._lblExtLoad_2.Location = New System.Drawing.Point(16, 56)
		Me._lblExtLoad_2.TabIndex = 42
		Me._lblExtLoad_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblExtLoad_2.BackColor = System.Drawing.SystemColors.Control
		Me._lblExtLoad_2.Enabled = True
		Me._lblExtLoad_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblExtLoad_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblExtLoad_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblExtLoad_2.UseMnemonic = True
		Me._lblExtLoad_2.Visible = True
		Me._lblExtLoad_2.AutoSize = False
		Me._lblExtLoad_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblExtLoad_2.Name = "_lblExtLoad_2"
		Me.fraPayout.Text = "Mooring Payout Adjustment"
		Me.fraPayout.Size = New System.Drawing.Size(487, 257)
		Me.fraPayout.Location = New System.Drawing.Point(250, 231)
		Me.fraPayout.TabIndex = 57
		Me.fraPayout.BackColor = System.Drawing.SystemColors.Control
		Me.fraPayout.Enabled = True
		Me.fraPayout.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraPayout.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraPayout.Visible = True
		Me.fraPayout.Name = "fraPayout"
		Me._btnPayout_0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._btnPayout_0.Text = "Payout by External Load"
		Me._btnPayout_0.Size = New System.Drawing.Size(150, 25)
		Me._btnPayout_0.Location = New System.Drawing.Point(160, 23)
		Me._btnPayout_0.TabIndex = 64
		Me.ToolTip1.SetToolTip(Me._btnPayout_0, "Compute Payout by Position and External Load")
		Me._btnPayout_0.BackColor = System.Drawing.SystemColors.Control
		Me._btnPayout_0.CausesValidation = True
		Me._btnPayout_0.Enabled = True
		Me._btnPayout_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._btnPayout_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._btnPayout_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._btnPayout_0.TabStop = True
		Me._btnPayout_0.Name = "_btnPayout_0"
		Me._btnPayout_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._btnPayout_1.Text = "Payout by Preset Tension"
		Me._btnPayout_1.Size = New System.Drawing.Size(150, 25)
		Me._btnPayout_1.Location = New System.Drawing.Point(317, 23)
		Me._btnPayout_1.TabIndex = 58
		Me.ToolTip1.SetToolTip(Me._btnPayout_1, "Compute Payout by Position and Preset Tension")
		Me._btnPayout_1.BackColor = System.Drawing.SystemColors.Control
		Me._btnPayout_1.CausesValidation = True
		Me._btnPayout_1.Enabled = True
		Me._btnPayout_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._btnPayout_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._btnPayout_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._btnPayout_1.TabStop = True
		Me._btnPayout_1.Name = "_btnPayout_1"
		grdPayout.OcxState = CType(resources.GetObject("grdPayout.OcxState"), System.Windows.Forms.AxHost.State)
		Me.grdPayout.Size = New System.Drawing.Size(454, 177)
		Me.grdPayout.Location = New System.Drawing.Point(16, 64)
		Me.grdPayout.TabIndex = 59
		Me.grdPayout.Name = "grdPayout"
		Me.fraConditions.Text = "Mooring System Conditions"
		Me.fraConditions.Size = New System.Drawing.Size(233, 257)
		Me.fraConditions.Location = New System.Drawing.Point(8, 232)
		Me.fraConditions.TabIndex = 53
		Me.fraConditions.BackColor = System.Drawing.SystemColors.Control
		Me.fraConditions.Enabled = True
		Me.fraConditions.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraConditions.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraConditions.Visible = True
		Me.fraConditions.Name = "fraConditions"
		Me.txtConditions.AutoSize = False
		Me.txtConditions.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me.txtConditions.Size = New System.Drawing.Size(56, 19)
		Me.txtConditions.Location = New System.Drawing.Point(144, 24)
		Me.txtConditions.TabIndex = 55
		Me.txtConditions.Text = "0"
		Me.txtConditions.AcceptsReturn = True
		Me.txtConditions.BackColor = System.Drawing.SystemColors.Window
		Me.txtConditions.CausesValidation = True
		Me.txtConditions.Enabled = True
		Me.txtConditions.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtConditions.HideSelection = True
		Me.txtConditions.ReadOnly = False
		Me.txtConditions.Maxlength = 0
		Me.txtConditions.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtConditions.MultiLine = False
		Me.txtConditions.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtConditions.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtConditions.TabStop = True
		Me.txtConditions.Visible = True
		Me.txtConditions.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtConditions.Name = "txtConditions"
		grdLC.OcxState = CType(resources.GetObject("grdLC.OcxState"), System.Windows.Forms.AxHost.State)
		Me.grdLC.Size = New System.Drawing.Size(187, 177)
		Me.grdLC.Location = New System.Drawing.Point(16, 64)
		Me.grdLC.TabIndex = 56
		Me.grdLC.Name = "grdLC"
		Me._lblForceUnit_5.Text = "kips"
		Me._lblForceUnit_5.Size = New System.Drawing.Size(25, 17)
		Me._lblForceUnit_5.Location = New System.Drawing.Point(204, 26)
		Me._lblForceUnit_5.TabIndex = 75
		Me._lblForceUnit_5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblForceUnit_5.BackColor = System.Drawing.SystemColors.Control
		Me._lblForceUnit_5.Enabled = True
		Me._lblForceUnit_5.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblForceUnit_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblForceUnit_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblForceUnit_5.UseMnemonic = True
		Me._lblForceUnit_5.Visible = True
		Me._lblForceUnit_5.AutoSize = False
		Me._lblForceUnit_5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblForceUnit_5.Name = "_lblForceUnit_5"
		Me.lblConditions.Text = "Maximum Winch Tension"
		Me.lblConditions.Size = New System.Drawing.Size(129, 17)
		Me.lblConditions.Location = New System.Drawing.Point(16, 24)
		Me.lblConditions.TabIndex = 54
		Me.lblConditions.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblConditions.BackColor = System.Drawing.SystemColors.Control
		Me.lblConditions.Enabled = True
		Me.lblConditions.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblConditions.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblConditions.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblConditions.UseMnemonic = True
		Me.lblConditions.Visible = True
		Me.lblConditions.AutoSize = False
		Me.lblConditions.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblConditions.Name = "lblConditions"
		Me.fraVesselLoc1.Text = "Target Vessel Station :"
		Me.fraVesselLoc1.Size = New System.Drawing.Size(360, 119)
		Me.fraVesselLoc1.Location = New System.Drawing.Point(376, 9)
		Me.fraVesselLoc1.TabIndex = 20
		Me.fraVesselLoc1.BackColor = System.Drawing.SystemColors.Control
		Me.fraVesselLoc1.Enabled = True
		Me.fraVesselLoc1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraVesselLoc1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraVesselLoc1.Visible = True
		Me.fraVesselLoc1.Name = "fraVesselLoc1"
		Me.btnPosition.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnPosition.Text = "<--  &Replace"
		Me.btnPosition.Size = New System.Drawing.Size(77, 27)
		Me.btnPosition.Location = New System.Drawing.Point(261, 82)
		Me.btnPosition.TabIndex = 35
		Me.ToolTip1.SetToolTip(Me.btnPosition, "Replace Current Vessel Station")
		Me.btnPosition.BackColor = System.Drawing.SystemColors.Control
		Me.btnPosition.CausesValidation = True
		Me.btnPosition.Enabled = True
		Me.btnPosition.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnPosition.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnPosition.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnPosition.TabStop = True
		Me.btnPosition.Name = "btnPosition"
		Me._txtVslSt_9.AutoSize = False
		Me._txtVslSt_9.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me._txtVslSt_9.Size = New System.Drawing.Size(49, 19)
		Me._txtVslSt_9.Location = New System.Drawing.Point(287, 53)
		Me._txtVslSt_9.TabIndex = 30
		Me._txtVslSt_9.TabStop = False
		Me._txtVslSt_9.Text = "0"
		Me._txtVslSt_9.AcceptsReturn = True
		Me._txtVslSt_9.BackColor = System.Drawing.SystemColors.Window
		Me._txtVslSt_9.CausesValidation = True
		Me._txtVslSt_9.Enabled = True
		Me._txtVslSt_9.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtVslSt_9.HideSelection = True
		Me._txtVslSt_9.ReadOnly = False
		Me._txtVslSt_9.Maxlength = 0
		Me._txtVslSt_9.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtVslSt_9.MultiLine = False
		Me._txtVslSt_9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtVslSt_9.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtVslSt_9.Visible = True
		Me._txtVslSt_9.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtVslSt_9.Name = "_txtVslSt_9"
		Me._txtVslSt_8.AutoSize = False
		Me._txtVslSt_8.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me._txtVslSt_8.Size = New System.Drawing.Size(73, 19)
		Me._txtVslSt_8.Location = New System.Drawing.Point(135, 55)
		Me._txtVslSt_8.TabIndex = 28
		Me._txtVslSt_8.TabStop = False
		Me._txtVslSt_8.Text = "0"
		Me._txtVslSt_8.AcceptsReturn = True
		Me._txtVslSt_8.BackColor = System.Drawing.SystemColors.Window
		Me._txtVslSt_8.CausesValidation = True
		Me._txtVslSt_8.Enabled = True
		Me._txtVslSt_8.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtVslSt_8.HideSelection = True
		Me._txtVslSt_8.ReadOnly = False
		Me._txtVslSt_8.Maxlength = 0
		Me._txtVslSt_8.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtVslSt_8.MultiLine = False
		Me._txtVslSt_8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtVslSt_8.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtVslSt_8.Visible = True
		Me._txtVslSt_8.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtVslSt_8.Name = "_txtVslSt_8"
		Me._cboWells_1.Size = New System.Drawing.Size(73, 21)
		Me._cboWells_1.Location = New System.Drawing.Point(56, 55)
		Me._cboWells_1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me._cboWells_1.TabIndex = 27
		Me._cboWells_1.BackColor = System.Drawing.SystemColors.Window
		Me._cboWells_1.CausesValidation = True
		Me._cboWells_1.Enabled = True
		Me._cboWells_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._cboWells_1.IntegralHeight = True
		Me._cboWells_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._cboWells_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cboWells_1.Sorted = False
		Me._cboWells_1.TabStop = True
		Me._cboWells_1.Visible = True
		Me._cboWells_1.Name = "_cboWells_1"
		Me._optInputSystem_3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optInputSystem_3.Text = "To"
		Me._optInputSystem_3.Size = New System.Drawing.Size(49, 17)
		Me._optInputSystem_3.Location = New System.Drawing.Point(16, 56)
		Me._optInputSystem_3.TabIndex = 26
		Me._optInputSystem_3.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optInputSystem_3.BackColor = System.Drawing.SystemColors.Control
		Me._optInputSystem_3.CausesValidation = True
		Me._optInputSystem_3.Enabled = True
		Me._optInputSystem_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optInputSystem_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._optInputSystem_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optInputSystem_3.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optInputSystem_3.TabStop = True
		Me._optInputSystem_3.Checked = False
		Me._optInputSystem_3.Visible = True
		Me._optInputSystem_3.Name = "_optInputSystem_3"
		Me._optInputSystem_2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optInputSystem_2.Text = "At Coordinate:"
		Me._optInputSystem_2.Size = New System.Drawing.Size(89, 17)
		Me._optInputSystem_2.Location = New System.Drawing.Point(16, 24)
		Me._optInputSystem_2.TabIndex = 21
		Me._optInputSystem_2.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optInputSystem_2.BackColor = System.Drawing.SystemColors.Control
		Me._optInputSystem_2.CausesValidation = True
		Me._optInputSystem_2.Enabled = True
		Me._optInputSystem_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optInputSystem_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._optInputSystem_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optInputSystem_2.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optInputSystem_2.TabStop = True
		Me._optInputSystem_2.Checked = False
		Me._optInputSystem_2.Visible = True
		Me._optInputSystem_2.Name = "_optInputSystem_2"
		Me._txtVslSt_10.AutoSize = False
		Me._txtVslSt_10.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me._txtVslSt_10.Size = New System.Drawing.Size(49, 19)
		Me._txtVslSt_10.Location = New System.Drawing.Point(159, 86)
		Me._txtVslSt_10.TabIndex = 33
		Me._txtVslSt_10.TabStop = False
		Me._txtVslSt_10.Text = "0"
		Me._txtVslSt_10.AcceptsReturn = True
		Me._txtVslSt_10.BackColor = System.Drawing.SystemColors.Window
		Me._txtVslSt_10.CausesValidation = True
		Me._txtVslSt_10.Enabled = True
		Me._txtVslSt_10.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtVslSt_10.HideSelection = True
		Me._txtVslSt_10.ReadOnly = False
		Me._txtVslSt_10.Maxlength = 0
		Me._txtVslSt_10.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtVslSt_10.MultiLine = False
		Me._txtVslSt_10.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtVslSt_10.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtVslSt_10.Visible = True
		Me._txtVslSt_10.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtVslSt_10.Name = "_txtVslSt_10"
		Me._txtVslSt_7.AutoSize = False
		Me._txtVslSt_7.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me._txtVslSt_7.Size = New System.Drawing.Size(76, 19)
		Me._txtVslSt_7.Location = New System.Drawing.Point(260, 23)
		Me._txtVslSt_7.TabIndex = 25
		Me._txtVslSt_7.TabStop = False
		Me._txtVslSt_7.Text = "0"
		Me._txtVslSt_7.AcceptsReturn = True
		Me._txtVslSt_7.BackColor = System.Drawing.SystemColors.Window
		Me._txtVslSt_7.CausesValidation = True
		Me._txtVslSt_7.Enabled = True
		Me._txtVslSt_7.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtVslSt_7.HideSelection = True
		Me._txtVslSt_7.ReadOnly = False
		Me._txtVslSt_7.Maxlength = 0
		Me._txtVslSt_7.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtVslSt_7.MultiLine = False
		Me._txtVslSt_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtVslSt_7.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtVslSt_7.Visible = True
		Me._txtVslSt_7.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtVslSt_7.Name = "_txtVslSt_7"
		Me._txtVslSt_6.AutoSize = False
		Me._txtVslSt_6.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me._txtVslSt_6.Size = New System.Drawing.Size(73, 19)
		Me._txtVslSt_6.Location = New System.Drawing.Point(135, 23)
		Me._txtVslSt_6.TabIndex = 23
		Me._txtVslSt_6.TabStop = False
		Me._txtVslSt_6.Text = "0"
		Me._txtVslSt_6.AcceptsReturn = True
		Me._txtVslSt_6.BackColor = System.Drawing.SystemColors.Window
		Me._txtVslSt_6.CausesValidation = True
		Me._txtVslSt_6.Enabled = True
		Me._txtVslSt_6.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtVslSt_6.HideSelection = True
		Me._txtVslSt_6.ReadOnly = False
		Me._txtVslSt_6.Maxlength = 0
		Me._txtVslSt_6.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtVslSt_6.MultiLine = False
		Me._txtVslSt_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtVslSt_6.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtVslSt_6.Visible = True
		Me._txtVslSt_6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtVslSt_6.Name = "_txtVslSt_6"
		Me._lblLengthUnit_6.Text = "ft"
		Me._lblLengthUnit_6.Size = New System.Drawing.Size(17, 17)
		Me._lblLengthUnit_6.Location = New System.Drawing.Point(215, 58)
		Me._lblLengthUnit_6.TabIndex = 70
		Me._lblLengthUnit_6.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblLengthUnit_6.BackColor = System.Drawing.SystemColors.Control
		Me._lblLengthUnit_6.Enabled = True
		Me._lblLengthUnit_6.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblLengthUnit_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblLengthUnit_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblLengthUnit_6.UseMnemonic = True
		Me._lblLengthUnit_6.Visible = True
		Me._lblLengthUnit_6.AutoSize = False
		Me._lblLengthUnit_6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblLengthUnit_6.Name = "_lblLengthUnit_6"
		Me._lblLengthUnit_5.Text = "ft"
		Me._lblLengthUnit_5.Size = New System.Drawing.Size(17, 17)
		Me._lblLengthUnit_5.Location = New System.Drawing.Point(336, 25)
		Me._lblLengthUnit_5.TabIndex = 69
		Me._lblLengthUnit_5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblLengthUnit_5.BackColor = System.Drawing.SystemColors.Control
		Me._lblLengthUnit_5.Enabled = True
		Me._lblLengthUnit_5.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblLengthUnit_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblLengthUnit_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblLengthUnit_5.UseMnemonic = True
		Me._lblLengthUnit_5.Visible = True
		Me._lblLengthUnit_5.AutoSize = False
		Me._lblLengthUnit_5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblLengthUnit_5.Name = "_lblLengthUnit_5"
		Me._lblLengthUnit_4.Text = "ft"
		Me._lblLengthUnit_4.Size = New System.Drawing.Size(17, 17)
		Me._lblLengthUnit_4.Location = New System.Drawing.Point(213, 25)
		Me._lblLengthUnit_4.TabIndex = 68
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
		Me._lblVslStUnit_9.Text = "Invalid_string_refer_to_original_code"
		Me._lblVslStUnit_9.Size = New System.Drawing.Size(17, 17)
		Me._lblVslStUnit_9.Location = New System.Drawing.Point(337, 52)
		Me._lblVslStUnit_9.TabIndex = 31
		Me._lblVslStUnit_9.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblVslStUnit_9.BackColor = System.Drawing.SystemColors.Control
		Me._lblVslStUnit_9.Enabled = True
		Me._lblVslStUnit_9.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblVslStUnit_9.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblVslStUnit_9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblVslStUnit_9.UseMnemonic = True
		Me._lblVslStUnit_9.Visible = True
		Me._lblVslStUnit_9.AutoSize = False
		Me._lblVslStUnit_9.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblVslStUnit_9.Name = "_lblVslStUnit_9"
		Me._lblVslSt_7.Text = "Bearing"
		Me._lblVslSt_7.Size = New System.Drawing.Size(41, 17)
		Me._lblVslSt_7.Location = New System.Drawing.Point(244, 55)
		Me._lblVslSt_7.TabIndex = 29
		Me._lblVslSt_7.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblVslSt_7.BackColor = System.Drawing.SystemColors.Control
		Me._lblVslSt_7.Enabled = True
		Me._lblVslSt_7.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblVslSt_7.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblVslSt_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblVslSt_7.UseMnemonic = True
		Me._lblVslSt_7.Visible = True
		Me._lblVslSt_7.AutoSize = False
		Me._lblVslSt_7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblVslSt_7.Name = "_lblVslSt_7"
		Me._lblVslStUnit_10.Text = "Invalid_string_refer_to_original_code"
		Me._lblVslStUnit_10.Size = New System.Drawing.Size(17, 17)
		Me._lblVslStUnit_10.Location = New System.Drawing.Point(212, 88)
		Me._lblVslStUnit_10.TabIndex = 34
		Me._lblVslStUnit_10.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblVslStUnit_10.BackColor = System.Drawing.SystemColors.Control
		Me._lblVslStUnit_10.Enabled = True
		Me._lblVslStUnit_10.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblVslStUnit_10.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblVslStUnit_10.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblVslStUnit_10.UseMnemonic = True
		Me._lblVslStUnit_10.Visible = True
		Me._lblVslStUnit_10.AutoSize = False
		Me._lblVslStUnit_10.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblVslStUnit_10.Name = "_lblVslStUnit_10"
		Me._lblVslSt_8.Text = "With Vessel Heading"
		Me._lblVslSt_8.Size = New System.Drawing.Size(121, 17)
		Me._lblVslSt_8.Location = New System.Drawing.Point(16, 88)
		Me._lblVslSt_8.TabIndex = 32
		Me._lblVslSt_8.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblVslSt_8.BackColor = System.Drawing.SystemColors.Control
		Me._lblVslSt_8.Enabled = True
		Me._lblVslSt_8.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblVslSt_8.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblVslSt_8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblVslSt_8.UseMnemonic = True
		Me._lblVslSt_8.Visible = True
		Me._lblVslSt_8.AutoSize = False
		Me._lblVslSt_8.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblVslSt_8.Name = "_lblVslSt_8"
		Me._lblVslSt_6.Text = "y (N)"
		Me._lblVslSt_6.Size = New System.Drawing.Size(33, 17)
		Me._lblVslSt_6.Location = New System.Drawing.Point(232, 24)
		Me._lblVslSt_6.TabIndex = 24
		Me._lblVslSt_6.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblVslSt_6.BackColor = System.Drawing.SystemColors.Control
		Me._lblVslSt_6.Enabled = True
		Me._lblVslSt_6.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblVslSt_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblVslSt_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblVslSt_6.UseMnemonic = True
		Me._lblVslSt_6.Visible = True
		Me._lblVslSt_6.AutoSize = False
		Me._lblVslSt_6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblVslSt_6.Name = "_lblVslSt_6"
		Me._lblVslSt_5.Text = "x (E)"
		Me._lblVslSt_5.Size = New System.Drawing.Size(26, 17)
		Me._lblVslSt_5.Location = New System.Drawing.Point(109, 25)
		Me._lblVslSt_5.TabIndex = 22
		Me._lblVslSt_5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblVslSt_5.BackColor = System.Drawing.SystemColors.Control
		Me._lblVslSt_5.Enabled = True
		Me._lblVslSt_5.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblVslSt_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblVslSt_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblVslSt_5.UseMnemonic = True
		Me._lblVslSt_5.Visible = True
		Me._lblVslSt_5.AutoSize = False
		Me._lblVslSt_5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblVslSt_5.Name = "_lblVslSt_5"
		Me.fraVesselLoc.Text = "Current Vessel Station :"
		Me.fraVesselLoc.Size = New System.Drawing.Size(360, 121)
		Me.fraVesselLoc.Location = New System.Drawing.Point(8, 8)
		Me.fraVesselLoc.TabIndex = 2
		Me.fraVesselLoc.BackColor = System.Drawing.SystemColors.Control
		Me.fraVesselLoc.Enabled = True
		Me.fraVesselLoc.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraVesselLoc.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraVesselLoc.Visible = True
		Me.fraVesselLoc.Name = "fraVesselLoc"
		Me._optInputSystem_0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optInputSystem_0.Text = "At Coordinate:"
		Me._optInputSystem_0.Size = New System.Drawing.Size(88, 17)
		Me._optInputSystem_0.Location = New System.Drawing.Point(16, 24)
		Me._optInputSystem_0.TabIndex = 3
		Me._optInputSystem_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optInputSystem_0.BackColor = System.Drawing.SystemColors.Control
		Me._optInputSystem_0.CausesValidation = True
		Me._optInputSystem_0.Enabled = True
		Me._optInputSystem_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optInputSystem_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._optInputSystem_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optInputSystem_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optInputSystem_0.TabStop = True
		Me._optInputSystem_0.Checked = False
		Me._optInputSystem_0.Visible = True
		Me._optInputSystem_0.Name = "_optInputSystem_0"
		Me._cboWells_0.Size = New System.Drawing.Size(73, 21)
		Me._cboWells_0.Location = New System.Drawing.Point(56, 53)
		Me._cboWells_0.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me._cboWells_0.TabIndex = 10
		Me._cboWells_0.BackColor = System.Drawing.SystemColors.Window
		Me._cboWells_0.CausesValidation = True
		Me._cboWells_0.Enabled = True
		Me._cboWells_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._cboWells_0.IntegralHeight = True
		Me._cboWells_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._cboWells_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cboWells_0.Sorted = False
		Me._cboWells_0.TabStop = True
		Me._cboWells_0.Visible = True
		Me._cboWells_0.Name = "_cboWells_0"
		Me._txtVslSt_2.AutoSize = False
		Me._txtVslSt_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me._txtVslSt_2.Size = New System.Drawing.Size(73, 19)
		Me._txtVslSt_2.Location = New System.Drawing.Point(133, 54)
		Me._txtVslSt_2.TabIndex = 11
		Me._txtVslSt_2.TabStop = False
		Me._txtVslSt_2.Text = "0"
		Me._txtVslSt_2.AcceptsReturn = True
		Me._txtVslSt_2.BackColor = System.Drawing.SystemColors.Window
		Me._txtVslSt_2.CausesValidation = True
		Me._txtVslSt_2.Enabled = True
		Me._txtVslSt_2.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtVslSt_2.HideSelection = True
		Me._txtVslSt_2.ReadOnly = False
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
		Me._txtVslSt_3.Size = New System.Drawing.Size(58, 19)
		Me._txtVslSt_3.Location = New System.Drawing.Point(278, 54)
		Me._txtVslSt_3.TabIndex = 13
		Me._txtVslSt_3.TabStop = False
		Me._txtVslSt_3.Text = "0"
		Me._txtVslSt_3.AcceptsReturn = True
		Me._txtVslSt_3.BackColor = System.Drawing.SystemColors.Window
		Me._txtVslSt_3.CausesValidation = True
		Me._txtVslSt_3.Enabled = True
		Me._txtVslSt_3.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtVslSt_3.HideSelection = True
		Me._txtVslSt_3.ReadOnly = False
		Me._txtVslSt_3.Maxlength = 0
		Me._txtVslSt_3.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtVslSt_3.MultiLine = False
		Me._txtVslSt_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtVslSt_3.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtVslSt_3.Visible = True
		Me._txtVslSt_3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtVslSt_3.Name = "_txtVslSt_3"
		Me._txtVslSt_0.AutoSize = False
		Me._txtVslSt_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me._txtVslSt_0.Size = New System.Drawing.Size(73, 19)
		Me._txtVslSt_0.Location = New System.Drawing.Point(134, 22)
		Me._txtVslSt_0.TabIndex = 5
		Me._txtVslSt_0.TabStop = False
		Me._txtVslSt_0.Text = "0"
		Me._txtVslSt_0.AcceptsReturn = True
		Me._txtVslSt_0.BackColor = System.Drawing.SystemColors.Window
		Me._txtVslSt_0.CausesValidation = True
		Me._txtVslSt_0.Enabled = True
		Me._txtVslSt_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtVslSt_0.HideSelection = True
		Me._txtVslSt_0.ReadOnly = False
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
		Me._txtVslSt_1.Size = New System.Drawing.Size(80, 19)
		Me._txtVslSt_1.Location = New System.Drawing.Point(257, 23)
		Me._txtVslSt_1.TabIndex = 8
		Me._txtVslSt_1.TabStop = False
		Me._txtVslSt_1.Text = "0"
		Me._txtVslSt_1.AcceptsReturn = True
		Me._txtVslSt_1.BackColor = System.Drawing.SystemColors.Window
		Me._txtVslSt_1.CausesValidation = True
		Me._txtVslSt_1.Enabled = True
		Me._txtVslSt_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtVslSt_1.HideSelection = True
		Me._txtVslSt_1.ReadOnly = False
		Me._txtVslSt_1.Maxlength = 0
		Me._txtVslSt_1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtVslSt_1.MultiLine = False
		Me._txtVslSt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtVslSt_1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtVslSt_1.Visible = True
		Me._txtVslSt_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtVslSt_1.Name = "_txtVslSt_1"
		Me._txtVslSt_4.AutoSize = False
		Me._txtVslSt_4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me._txtVslSt_4.Size = New System.Drawing.Size(49, 19)
		Me._txtVslSt_4.Location = New System.Drawing.Point(157, 86)
		Me._txtVslSt_4.TabIndex = 16
		Me._txtVslSt_4.TabStop = False
		Me._txtVslSt_4.Text = "0"
		Me._txtVslSt_4.AcceptsReturn = True
		Me._txtVslSt_4.BackColor = System.Drawing.SystemColors.Window
		Me._txtVslSt_4.CausesValidation = True
		Me._txtVslSt_4.Enabled = True
		Me._txtVslSt_4.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtVslSt_4.HideSelection = True
		Me._txtVslSt_4.ReadOnly = False
		Me._txtVslSt_4.Maxlength = 0
		Me._txtVslSt_4.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtVslSt_4.MultiLine = False
		Me._txtVslSt_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtVslSt_4.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtVslSt_4.Visible = True
		Me._txtVslSt_4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtVslSt_4.Name = "_txtVslSt_4"
		Me._txtVslSt_5.AutoSize = False
		Me._txtVslSt_5.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me._txtVslSt_5.BackColor = System.Drawing.SystemColors.Control
		Me._txtVslSt_5.Size = New System.Drawing.Size(58, 19)
		Me._txtVslSt_5.Location = New System.Drawing.Point(278, 85)
		Me._txtVslSt_5.ReadOnly = True
		Me._txtVslSt_5.TabIndex = 19
		Me._txtVslSt_5.TabStop = False
		Me._txtVslSt_5.Text = "0"
		Me._txtVslSt_5.AcceptsReturn = True
		Me._txtVslSt_5.CausesValidation = True
		Me._txtVslSt_5.Enabled = True
		Me._txtVslSt_5.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtVslSt_5.HideSelection = True
		Me._txtVslSt_5.Maxlength = 0
		Me._txtVslSt_5.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtVslSt_5.MultiLine = False
		Me._txtVslSt_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtVslSt_5.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtVslSt_5.Visible = True
		Me._txtVslSt_5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtVslSt_5.Name = "_txtVslSt_5"
		Me._optInputSystem_1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optInputSystem_1.Text = "To"
		Me._optInputSystem_1.Size = New System.Drawing.Size(49, 17)
		Me._optInputSystem_1.Location = New System.Drawing.Point(16, 56)
		Me._optInputSystem_1.TabIndex = 9
		Me._optInputSystem_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optInputSystem_1.BackColor = System.Drawing.SystemColors.Control
		Me._optInputSystem_1.CausesValidation = True
		Me._optInputSystem_1.Enabled = True
		Me._optInputSystem_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optInputSystem_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._optInputSystem_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optInputSystem_1.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optInputSystem_1.TabStop = True
		Me._optInputSystem_1.Checked = False
		Me._optInputSystem_1.Visible = True
		Me._optInputSystem_1.Name = "_optInputSystem_1"
		Me._lblLengthUnit_3.Text = "ft"
		Me._lblLengthUnit_3.Size = New System.Drawing.Size(17, 17)
		Me._lblLengthUnit_3.Location = New System.Drawing.Point(340, 87)
		Me._lblLengthUnit_3.TabIndex = 67
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
		Me._lblLengthUnit_2.Text = "ft"
		Me._lblLengthUnit_2.Size = New System.Drawing.Size(13, 17)
		Me._lblLengthUnit_2.Location = New System.Drawing.Point(209, 56)
		Me._lblLengthUnit_2.TabIndex = 66
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
		Me._lblLengthUnit_1.Size = New System.Drawing.Size(12, 17)
		Me._lblLengthUnit_1.Location = New System.Drawing.Point(341, 25)
		Me._lblLengthUnit_1.TabIndex = 65
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
		Me._lblVslSt_2.Text = "Bearing"
		Me._lblVslSt_2.Size = New System.Drawing.Size(41, 17)
		Me._lblVslSt_2.Location = New System.Drawing.Point(236, 56)
		Me._lblVslSt_2.TabIndex = 12
		Me._lblVslSt_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblVslSt_2.BackColor = System.Drawing.SystemColors.Control
		Me._lblVslSt_2.Enabled = True
		Me._lblVslSt_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblVslSt_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblVslSt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblVslSt_2.UseMnemonic = True
		Me._lblVslSt_2.Visible = True
		Me._lblVslSt_2.AutoSize = False
		Me._lblVslSt_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblVslSt_2.Name = "_lblVslSt_2"
		Me._lblVslStUnit_3.Text = "Invalid_string_refer_to_original_code"
		Me._lblVslStUnit_3.Size = New System.Drawing.Size(17, 17)
		Me._lblVslStUnit_3.Location = New System.Drawing.Point(338, 51)
		Me._lblVslStUnit_3.TabIndex = 14
		Me._lblVslStUnit_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblVslStUnit_3.BackColor = System.Drawing.SystemColors.Control
		Me._lblVslStUnit_3.Enabled = True
		Me._lblVslStUnit_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblVslStUnit_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblVslStUnit_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblVslStUnit_3.UseMnemonic = True
		Me._lblVslStUnit_3.Visible = True
		Me._lblVslStUnit_3.AutoSize = False
		Me._lblVslStUnit_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblVslStUnit_3.Name = "_lblVslStUnit_3"
		Me._lblVslSt_4.Text = "Draft"
		Me._lblVslSt_4.Size = New System.Drawing.Size(41, 17)
		Me._lblVslSt_4.Location = New System.Drawing.Point(244, 88)
		Me._lblVslSt_4.TabIndex = 18
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
		Me._lblLengthUnit_0.Location = New System.Drawing.Point(210, 24)
		Me._lblLengthUnit_0.TabIndex = 6
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
		Me._lblVslSt_0.Location = New System.Drawing.Point(110, 25)
		Me._lblVslSt_0.TabIndex = 4
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
		Me._lblVslSt_1.Size = New System.Drawing.Size(31, 17)
		Me._lblVslSt_1.Location = New System.Drawing.Point(232, 24)
		Me._lblVslSt_1.TabIndex = 7
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
		Me._lblVslSt_3.Text = "With Vessel Heading"
		Me._lblVslSt_3.Size = New System.Drawing.Size(105, 17)
		Me._lblVslSt_3.Location = New System.Drawing.Point(16, 88)
		Me._lblVslSt_3.TabIndex = 15
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
		Me._lblVslStUnit_4.Text = "Invalid_string_refer_to_original_code"
		Me._lblVslStUnit_4.Size = New System.Drawing.Size(17, 17)
		Me._lblVslStUnit_4.Location = New System.Drawing.Point(208, 84)
		Me._lblVslStUnit_4.TabIndex = 17
		Me._lblVslStUnit_4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblVslStUnit_4.BackColor = System.Drawing.SystemColors.Control
		Me._lblVslStUnit_4.Enabled = True
		Me._lblVslStUnit_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblVslStUnit_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblVslStUnit_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblVslStUnit_4.UseMnemonic = True
		Me._lblVslStUnit_4.Visible = True
		Me._lblVslStUnit_4.AutoSize = False
		Me._lblVslStUnit_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblVslStUnit_4.Name = "_lblVslStUnit_4"
		Me._lblForceUnit_7.Text = "kips"
		Me._lblForceUnit_7.Size = New System.Drawing.Size(25, 17)
		Me._lblForceUnit_7.Location = New System.Drawing.Point(1, 16)
		Me._lblForceUnit_7.TabIndex = 86
		Me._lblForceUnit_7.Visible = False
		Me._lblForceUnit_7.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblForceUnit_7.BackColor = System.Drawing.SystemColors.Control
		Me._lblForceUnit_7.Enabled = True
		Me._lblForceUnit_7.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblForceUnit_7.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblForceUnit_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblForceUnit_7.UseMnemonic = True
		Me._lblForceUnit_7.AutoSize = False
		Me._lblForceUnit_7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblForceUnit_7.Name = "_lblForceUnit_7"
		Me._lblVelUnit_0.Text = "kn"
		Me._lblVelUnit_0.Size = New System.Drawing.Size(25, 17)
		Me._lblVelUnit_0.Location = New System.Drawing.Point(0, 0)
		Me._lblVelUnit_0.TabIndex = 85
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
		Me._lblGenCmt_3.Text = "2."
		Me._lblGenCmt_3.ForeColor = System.Drawing.Color.FromARGB(0, 0, 128)
		Me._lblGenCmt_3.Size = New System.Drawing.Size(9, 17)
		Me._lblGenCmt_3.Location = New System.Drawing.Point(23, 537)
		Me._lblGenCmt_3.TabIndex = 62
		Me._lblGenCmt_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblGenCmt_3.BackColor = System.Drawing.SystemColors.Control
		Me._lblGenCmt_3.Enabled = True
		Me._lblGenCmt_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblGenCmt_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblGenCmt_3.UseMnemonic = True
		Me._lblGenCmt_3.Visible = True
		Me._lblGenCmt_3.AutoSize = False
		Me._lblGenCmt_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblGenCmt_3.Name = "_lblGenCmt_3"
		Me._lblGenCmt_2.Text = "Mooring System Forces are in global system (E-N system). Yaw moment is in the direction of x (E) to y (N)"
		Me._lblGenCmt_2.ForeColor = System.Drawing.Color.FromARGB(0, 0, 128)
		Me._lblGenCmt_2.Size = New System.Drawing.Size(262, 37)
		Me._lblGenCmt_2.Location = New System.Drawing.Point(40, 538)
		Me._lblGenCmt_2.TabIndex = 63
		Me._lblGenCmt_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblGenCmt_2.BackColor = System.Drawing.SystemColors.Control
		Me._lblGenCmt_2.Enabled = True
		Me._lblGenCmt_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblGenCmt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblGenCmt_2.UseMnemonic = True
		Me._lblGenCmt_2.Visible = True
		Me._lblGenCmt_2.AutoSize = False
		Me._lblGenCmt_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblGenCmt_2.Name = "_lblGenCmt_2"
		Me._lblGenCmt_1.Text = "1."
		Me._lblGenCmt_1.ForeColor = System.Drawing.Color.FromARGB(0, 0, 128)
		Me._lblGenCmt_1.Size = New System.Drawing.Size(9, 17)
		Me._lblGenCmt_1.Location = New System.Drawing.Point(24, 496)
		Me._lblGenCmt_1.TabIndex = 60
		Me._lblGenCmt_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblGenCmt_1.BackColor = System.Drawing.SystemColors.Control
		Me._lblGenCmt_1.Enabled = True
		Me._lblGenCmt_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblGenCmt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblGenCmt_1.UseMnemonic = True
		Me._lblGenCmt_1.Visible = True
		Me._lblGenCmt_1.AutoSize = False
		Me._lblGenCmt_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblGenCmt_1.Name = "_lblGenCmt_1"
		Me._lblGenCmt_0.Text = "Vessel Heading and Bearing Angle measured clockwise from North."
		Me._lblGenCmt_0.ForeColor = System.Drawing.Color.FromARGB(0, 0, 128)
		Me._lblGenCmt_0.Size = New System.Drawing.Size(256, 27)
		Me._lblGenCmt_0.Location = New System.Drawing.Point(40, 496)
		Me._lblGenCmt_0.TabIndex = 61
		Me._lblGenCmt_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblGenCmt_0.BackColor = System.Drawing.SystemColors.Control
		Me._lblGenCmt_0.Enabled = True
		Me._lblGenCmt_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblGenCmt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblGenCmt_0.UseMnemonic = True
		Me._lblGenCmt_0.Visible = True
		Me._lblGenCmt_0.AutoSize = False
		Me._lblGenCmt_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblGenCmt_0.Name = "_lblGenCmt_0"
		Me.Controls.Add(fraReport)
		Me.Controls.Add(_fraExternalLoad_1)
		Me.Controls.Add(btnSave)
		Me.Controls.Add(SysInfo1)
		Me.Controls.Add(btnCancel)
		Me.Controls.Add(_fraExternalLoad_0)
		Me.Controls.Add(fraPayout)
		Me.Controls.Add(fraConditions)
		Me.Controls.Add(fraVesselLoc1)
		Me.Controls.Add(fraVesselLoc)
		Me.Controls.Add(_lblForceUnit_7)
		Me.Controls.Add(_lblVelUnit_0)
		Me.Controls.Add(_lblGenCmt_3)
		Me.Controls.Add(_lblGenCmt_2)
		Me.Controls.Add(_lblGenCmt_1)
		Me.Controls.Add(_lblGenCmt_0)
		Me.fraReport.Controls.Add(btnReport)
		Me.fraReport.Controls.Add(txtClientName)
		Me.fraReport.Controls.Add(txtLocationName)
		Me.fraReport.Controls.Add(lblClientName)
		Me.fraReport.Controls.Add(lblLocation)
		Me._fraExternalLoad_1.Controls.Add(_txtExtLoad_5)
		Me._fraExternalLoad_1.Controls.Add(_txtExtLoad_4)
		Me._fraExternalLoad_1.Controls.Add(_txtExtLoad_3)
		Me._fraExternalLoad_1.Controls.Add(_lblForceUnit_6)
		Me._fraExternalLoad_1.Controls.Add(_lblLengthUnit_8)
		Me._fraExternalLoad_1.Controls.Add(_lblForceUnit_4)
		Me._fraExternalLoad_1.Controls.Add(_lblForceUnit_3)
		Me._fraExternalLoad_1.Controls.Add(_lblExtLoad_5)
		Me._fraExternalLoad_1.Controls.Add(_lblExtLoad_4)
		Me._fraExternalLoad_1.Controls.Add(_lblExtLoad_3)
		Me._fraExternalLoad_1.Controls.Add(_lblExtUnits_3)
		Me._fraExternalLoad_0.Controls.Add(_txtExtLoad_0)
		Me._fraExternalLoad_0.Controls.Add(_txtExtLoad_1)
		Me._fraExternalLoad_0.Controls.Add(_txtExtLoad_2)
		Me._fraExternalLoad_0.Controls.Add(_lblLengthUnit_7)
		Me._fraExternalLoad_0.Controls.Add(_lblForceUnit_2)
		Me._fraExternalLoad_0.Controls.Add(_lblForceUnit_1)
		Me._fraExternalLoad_0.Controls.Add(_lblExtUnits_2)
		Me._fraExternalLoad_0.Controls.Add(_lblForceUnit_0)
		Me._fraExternalLoad_0.Controls.Add(_lblExtLoad_0)
		Me._fraExternalLoad_0.Controls.Add(_lblExtLoad_1)
		Me._fraExternalLoad_0.Controls.Add(_lblExtLoad_2)
		Me.fraPayout.Controls.Add(_btnPayout_0)
		Me.fraPayout.Controls.Add(_btnPayout_1)
		Me.fraPayout.Controls.Add(grdPayout)
		Me.fraConditions.Controls.Add(txtConditions)
		Me.fraConditions.Controls.Add(grdLC)
		Me.fraConditions.Controls.Add(_lblForceUnit_5)
		Me.fraConditions.Controls.Add(lblConditions)
		Me.fraVesselLoc1.Controls.Add(btnPosition)
		Me.fraVesselLoc1.Controls.Add(_txtVslSt_9)
		Me.fraVesselLoc1.Controls.Add(_txtVslSt_8)
		Me.fraVesselLoc1.Controls.Add(_cboWells_1)
		Me.fraVesselLoc1.Controls.Add(_optInputSystem_3)
		Me.fraVesselLoc1.Controls.Add(_optInputSystem_2)
		Me.fraVesselLoc1.Controls.Add(_txtVslSt_10)
		Me.fraVesselLoc1.Controls.Add(_txtVslSt_7)
		Me.fraVesselLoc1.Controls.Add(_txtVslSt_6)
		Me.fraVesselLoc1.Controls.Add(_lblLengthUnit_6)
		Me.fraVesselLoc1.Controls.Add(_lblLengthUnit_5)
		Me.fraVesselLoc1.Controls.Add(_lblLengthUnit_4)
		Me.fraVesselLoc1.Controls.Add(_lblVslStUnit_9)
		Me.fraVesselLoc1.Controls.Add(_lblVslSt_7)
		Me.fraVesselLoc1.Controls.Add(_lblVslStUnit_10)
		Me.fraVesselLoc1.Controls.Add(_lblVslSt_8)
		Me.fraVesselLoc1.Controls.Add(_lblVslSt_6)
		Me.fraVesselLoc1.Controls.Add(_lblVslSt_5)
		Me.fraVesselLoc.Controls.Add(_optInputSystem_0)
		Me.fraVesselLoc.Controls.Add(_cboWells_0)
		Me.fraVesselLoc.Controls.Add(_txtVslSt_2)
		Me.fraVesselLoc.Controls.Add(_txtVslSt_3)
		Me.fraVesselLoc.Controls.Add(_txtVslSt_0)
		Me.fraVesselLoc.Controls.Add(_txtVslSt_1)
		Me.fraVesselLoc.Controls.Add(_txtVslSt_4)
		Me.fraVesselLoc.Controls.Add(_txtVslSt_5)
		Me.fraVesselLoc.Controls.Add(_optInputSystem_1)
		Me.fraVesselLoc.Controls.Add(_lblLengthUnit_3)
		Me.fraVesselLoc.Controls.Add(_lblLengthUnit_2)
		Me.fraVesselLoc.Controls.Add(_lblLengthUnit_1)
		Me.fraVesselLoc.Controls.Add(_lblVslSt_2)
		Me.fraVesselLoc.Controls.Add(_lblVslStUnit_3)
		Me.fraVesselLoc.Controls.Add(_lblVslSt_4)
		Me.fraVesselLoc.Controls.Add(_lblLengthUnit_0)
		Me.fraVesselLoc.Controls.Add(_lblVslSt_0)
		Me.fraVesselLoc.Controls.Add(_lblVslSt_1)
		Me.fraVesselLoc.Controls.Add(_lblVslSt_3)
		Me.fraVesselLoc.Controls.Add(_lblVslStUnit_4)
		Me.btnPayout.SetIndex(_btnPayout_0, CType(0, Short))
		Me.btnPayout.SetIndex(_btnPayout_1, CType(1, Short))
		Me.cboWells.SetIndex(_cboWells_1, CType(1, Short))
		Me.cboWells.SetIndex(_cboWells_0, CType(0, Short))
		Me.fraExternalLoad.SetIndex(_fraExternalLoad_1, CType(1, Short))
		Me.fraExternalLoad.SetIndex(_fraExternalLoad_0, CType(0, Short))
		Me.lblExtLoad.SetIndex(_lblExtLoad_5, CType(5, Short))
		Me.lblExtLoad.SetIndex(_lblExtLoad_4, CType(4, Short))
		Me.lblExtLoad.SetIndex(_lblExtLoad_3, CType(3, Short))
		Me.lblExtLoad.SetIndex(_lblExtLoad_0, CType(0, Short))
		Me.lblExtLoad.SetIndex(_lblExtLoad_1, CType(1, Short))
		Me.lblExtLoad.SetIndex(_lblExtLoad_2, CType(2, Short))
		Me.lblExtUnits.SetIndex(_lblExtUnits_3, CType(3, Short))
		Me.lblExtUnits.SetIndex(_lblExtUnits_2, CType(2, Short))
		Me.lblForceUnit.SetIndex(_lblForceUnit_6, CType(6, Short))
		Me.lblForceUnit.SetIndex(_lblForceUnit_4, CType(4, Short))
		Me.lblForceUnit.SetIndex(_lblForceUnit_3, CType(3, Short))
		Me.lblForceUnit.SetIndex(_lblForceUnit_2, CType(2, Short))
		Me.lblForceUnit.SetIndex(_lblForceUnit_1, CType(1, Short))
		Me.lblForceUnit.SetIndex(_lblForceUnit_0, CType(0, Short))
		Me.lblForceUnit.SetIndex(_lblForceUnit_5, CType(5, Short))
		Me.lblForceUnit.SetIndex(_lblForceUnit_7, CType(7, Short))
		Me.lblGenCmt.SetIndex(_lblGenCmt_3, CType(3, Short))
		Me.lblGenCmt.SetIndex(_lblGenCmt_2, CType(2, Short))
		Me.lblGenCmt.SetIndex(_lblGenCmt_1, CType(1, Short))
		Me.lblGenCmt.SetIndex(_lblGenCmt_0, CType(0, Short))
		Me.lblLengthUnit.SetIndex(_lblLengthUnit_8, CType(8, Short))
		Me.lblLengthUnit.SetIndex(_lblLengthUnit_7, CType(7, Short))
		Me.lblLengthUnit.SetIndex(_lblLengthUnit_6, CType(6, Short))
		Me.lblLengthUnit.SetIndex(_lblLengthUnit_5, CType(5, Short))
		Me.lblLengthUnit.SetIndex(_lblLengthUnit_4, CType(4, Short))
		Me.lblLengthUnit.SetIndex(_lblLengthUnit_3, CType(3, Short))
		Me.lblLengthUnit.SetIndex(_lblLengthUnit_2, CType(2, Short))
		Me.lblLengthUnit.SetIndex(_lblLengthUnit_1, CType(1, Short))
		Me.lblLengthUnit.SetIndex(_lblLengthUnit_0, CType(0, Short))
		Me.lblVelUnit.SetIndex(_lblVelUnit_0, CType(0, Short))
		Me.lblVslSt.SetIndex(_lblVslSt_7, CType(7, Short))
		Me.lblVslSt.SetIndex(_lblVslSt_8, CType(8, Short))
		Me.lblVslSt.SetIndex(_lblVslSt_6, CType(6, Short))
		Me.lblVslSt.SetIndex(_lblVslSt_5, CType(5, Short))
		Me.lblVslSt.SetIndex(_lblVslSt_2, CType(2, Short))
		Me.lblVslSt.SetIndex(_lblVslSt_4, CType(4, Short))
		Me.lblVslSt.SetIndex(_lblVslSt_0, CType(0, Short))
		Me.lblVslSt.SetIndex(_lblVslSt_1, CType(1, Short))
		Me.lblVslSt.SetIndex(_lblVslSt_3, CType(3, Short))
		Me.lblVslStUnit.SetIndex(_lblVslStUnit_9, CType(9, Short))
		Me.lblVslStUnit.SetIndex(_lblVslStUnit_10, CType(10, Short))
		Me.lblVslStUnit.SetIndex(_lblVslStUnit_3, CType(3, Short))
		Me.lblVslStUnit.SetIndex(_lblVslStUnit_4, CType(4, Short))
		Me.optInputSystem.SetIndex(_optInputSystem_3, CType(3, Short))
		Me.optInputSystem.SetIndex(_optInputSystem_2, CType(2, Short))
		Me.optInputSystem.SetIndex(_optInputSystem_0, CType(0, Short))
		Me.optInputSystem.SetIndex(_optInputSystem_1, CType(1, Short))
		Me.txtExtLoad.SetIndex(_txtExtLoad_5, CType(5, Short))
		Me.txtExtLoad.SetIndex(_txtExtLoad_4, CType(4, Short))
		Me.txtExtLoad.SetIndex(_txtExtLoad_3, CType(3, Short))
		Me.txtExtLoad.SetIndex(_txtExtLoad_0, CType(0, Short))
		Me.txtExtLoad.SetIndex(_txtExtLoad_1, CType(1, Short))
		Me.txtExtLoad.SetIndex(_txtExtLoad_2, CType(2, Short))
		Me.txtVslSt.SetIndex(_txtVslSt_9, CType(9, Short))
		Me.txtVslSt.SetIndex(_txtVslSt_8, CType(8, Short))
		Me.txtVslSt.SetIndex(_txtVslSt_10, CType(10, Short))
		Me.txtVslSt.SetIndex(_txtVslSt_7, CType(7, Short))
		Me.txtVslSt.SetIndex(_txtVslSt_6, CType(6, Short))
		Me.txtVslSt.SetIndex(_txtVslSt_2, CType(2, Short))
		Me.txtVslSt.SetIndex(_txtVslSt_3, CType(3, Short))
		Me.txtVslSt.SetIndex(_txtVslSt_0, CType(0, Short))
		Me.txtVslSt.SetIndex(_txtVslSt_1, CType(1, Short))
		Me.txtVslSt.SetIndex(_txtVslSt_4, CType(4, Short))
		Me.txtVslSt.SetIndex(_txtVslSt_5, CType(5, Short))
		CType(Me.txtVslSt, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.txtExtLoad, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.optInputSystem, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lblVslStUnit, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lblVslSt, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lblVelUnit, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lblLengthUnit, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lblGenCmt, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lblForceUnit, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lblExtUnits, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lblExtLoad, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.fraExternalLoad, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cboWells, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.btnPayout, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.grdLC, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.grdPayout, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.SysInfo1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.fraReport.ResumeLayout(False)
		Me._fraExternalLoad_1.ResumeLayout(False)
		Me._fraExternalLoad_0.ResumeLayout(False)
		Me.fraPayout.ResumeLayout(False)
		Me.fraConditions.ResumeLayout(False)
		Me.fraVesselLoc1.ResumeLayout(False)
		Me.fraVesselLoc.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class