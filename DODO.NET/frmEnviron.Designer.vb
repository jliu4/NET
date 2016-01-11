<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmEnviron
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
	Public WithEvents _stbStatus_Panel1 As System.Windows.Forms.ToolStripStatusLabel
	Public WithEvents stbStatus As System.Windows.Forms.StatusStrip
	Public WithEvents btnCancel As System.Windows.Forms.Button
	Public WithEvents btnOK As System.Windows.Forms.Button
	Public WithEvents btnForce As System.Windows.Forms.Button
	Public WithEvents picEnviron As System.Windows.Forms.PictureBox
    Public WithEvents fraForce As System.Windows.Forms.GroupBox
	Public WithEvents _txtVslSt_0 As System.Windows.Forms.TextBox
	Public WithEvents _txtVslSt_1 As System.Windows.Forms.TextBox
	Public WithEvents _lblVslSt_0 As System.Windows.Forms.Label
	Public WithEvents _lblAngleUnit_0 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_1 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_0 As System.Windows.Forms.Label
	Public WithEvents fraVesselLoc As System.Windows.Forms.GroupBox
	Public WithEvents btnProfile As System.Windows.Forms.Button
	Public WithEvents _txtCurr_1 As System.Windows.Forms.TextBox
	Public WithEvents _txtCurr_0 As System.Windows.Forms.TextBox
	Public WithEvents _lblAngleUnit_3 As System.Windows.Forms.Label
	Public WithEvents _lblVelUnit_1 As System.Windows.Forms.Label
	Public WithEvents _lblCurr_1 As System.Windows.Forms.Label
	Public WithEvents _lblCurr_0 As System.Windows.Forms.Label
	Public WithEvents fraCurrent As System.Windows.Forms.GroupBox
	Public WithEvents _txtWave_2 As System.Windows.Forms.TextBox
	Public WithEvents _txtWave_1 As System.Windows.Forms.TextBox
	Public WithEvents _txtWave_0 As System.Windows.Forms.TextBox
	Public WithEvents _lblAngleUnit_2 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_2 As System.Windows.Forms.Label
	Public WithEvents _lblWave_2 As System.Windows.Forms.Label
	Public WithEvents _lblWave_1 As System.Windows.Forms.Label
	Public WithEvents _lblWaveUnit_1 As System.Windows.Forms.Label
	Public WithEvents _lblWave_0 As System.Windows.Forms.Label
	Public WithEvents fraWave As System.Windows.Forms.GroupBox
	Public WithEvents _btrDuration_2 As System.Windows.Forms.RadioButton
	Public WithEvents _btrDuration_1 As System.Windows.Forms.RadioButton
	Public WithEvents _btrDuration_0 As System.Windows.Forms.RadioButton
	Public WithEvents _txtWind_2 As System.Windows.Forms.TextBox
	Public WithEvents _txtWind_1 As System.Windows.Forms.TextBox
	Public WithEvents _txtWind_0 As System.Windows.Forms.TextBox
	Public WithEvents _lblAngleUnit_1 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_1 As System.Windows.Forms.Label
	Public WithEvents _lblWind_2 As System.Windows.Forms.Label
	Public WithEvents _lblWind_1 As System.Windows.Forms.Label
	Public WithEvents _lblWindUnit_1 As System.Windows.Forms.Label
	Public WithEvents _lblWind_0 As System.Windows.Forms.Label
	Public WithEvents _lblVelUnit_0 As System.Windows.Forms.Label
	Public WithEvents fraWind As System.Windows.Forms.GroupBox
	Public WithEvents btnAdd As System.Windows.Forms.Button
	Public WithEvents cboCurEnv As System.Windows.Forms.ComboBox
	Public WithEvents fraCurEnv As System.Windows.Forms.GroupBox
	Public dlgFileOpen As System.Windows.Forms.OpenFileDialog
	Public dlgFileSave As System.Windows.Forms.SaveFileDialog
	Public WithEvents _lblGenCmt_0 As System.Windows.Forms.Label
	Public WithEvents btrDuration As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	Public WithEvents lblAngleUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblCurr As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblGenCmt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblLengthUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblVelUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblVslSt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblWave As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblWaveUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblWind As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblWindUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents txtCurr As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	Public WithEvents txtVslSt As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	Public WithEvents txtWave As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	Public WithEvents txtWind As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	Public WithEvents mnuOpen As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuSave As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuDum As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuClose As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuFile As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEnviron))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.stbStatus = New System.Windows.Forms.StatusStrip()
        Me._stbStatus_Panel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.fraForce = New System.Windows.Forms.GroupBox()
        Me.grdForce = New System.Windows.Forms.DataGridView()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.btnForce = New System.Windows.Forms.Button()
        Me.picEnviron = New System.Windows.Forms.PictureBox()
        Me.fraVesselLoc = New System.Windows.Forms.GroupBox()
        Me._txtVslSt_0 = New System.Windows.Forms.TextBox()
        Me._txtVslSt_1 = New System.Windows.Forms.TextBox()
        Me._lblVslSt_0 = New System.Windows.Forms.Label()
        Me._lblAngleUnit_0 = New System.Windows.Forms.Label()
        Me._lblVslSt_1 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_0 = New System.Windows.Forms.Label()
        Me.fraCurrent = New System.Windows.Forms.GroupBox()
        Me.btnProfile = New System.Windows.Forms.Button()
        Me._txtCurr_1 = New System.Windows.Forms.TextBox()
        Me._txtCurr_0 = New System.Windows.Forms.TextBox()
        Me._lblAngleUnit_3 = New System.Windows.Forms.Label()
        Me._lblVelUnit_1 = New System.Windows.Forms.Label()
        Me._lblCurr_1 = New System.Windows.Forms.Label()
        Me._lblCurr_0 = New System.Windows.Forms.Label()
        Me.fraWave = New System.Windows.Forms.GroupBox()
        Me._txtWave_2 = New System.Windows.Forms.TextBox()
        Me._txtWave_1 = New System.Windows.Forms.TextBox()
        Me._txtWave_0 = New System.Windows.Forms.TextBox()
        Me._lblAngleUnit_2 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_2 = New System.Windows.Forms.Label()
        Me._lblWave_2 = New System.Windows.Forms.Label()
        Me._lblWave_1 = New System.Windows.Forms.Label()
        Me._lblWaveUnit_1 = New System.Windows.Forms.Label()
        Me._lblWave_0 = New System.Windows.Forms.Label()
        Me.fraWind = New System.Windows.Forms.GroupBox()
        Me._btrDuration_2 = New System.Windows.Forms.RadioButton()
        Me._btrDuration_1 = New System.Windows.Forms.RadioButton()
        Me._btrDuration_0 = New System.Windows.Forms.RadioButton()
        Me._txtWind_2 = New System.Windows.Forms.TextBox()
        Me._txtWind_1 = New System.Windows.Forms.TextBox()
        Me._txtWind_0 = New System.Windows.Forms.TextBox()
        Me._lblAngleUnit_1 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_1 = New System.Windows.Forms.Label()
        Me._lblWind_2 = New System.Windows.Forms.Label()
        Me._lblWind_1 = New System.Windows.Forms.Label()
        Me._lblWindUnit_1 = New System.Windows.Forms.Label()
        Me._lblWind_0 = New System.Windows.Forms.Label()
        Me._lblVelUnit_0 = New System.Windows.Forms.Label()
        Me.fraCurEnv = New System.Windows.Forms.GroupBox()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.cboCurEnv = New System.Windows.Forms.ComboBox()
        Me.dlgFileOpen = New System.Windows.Forms.OpenFileDialog()
        Me.dlgFileSave = New System.Windows.Forms.SaveFileDialog()
        Me._lblGenCmt_0 = New System.Windows.Forms.Label()
        Me.btrDuration = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.lblAngleUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblCurr = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblGenCmt = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblLengthUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblVelUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblVslSt = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblWave = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblWaveUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblWind = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblWindUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.txtCurr = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.txtVslSt = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.txtWave = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.txtWind = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip()
        Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuOpen = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuSave = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuDum = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuClose = New System.Windows.Forms.ToolStripMenuItem()
        Me.stbStatus.SuspendLayout()
        Me.fraForce.SuspendLayout()
        CType(Me.grdForce, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picEnviron, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraVesselLoc.SuspendLayout()
        Me.fraCurrent.SuspendLayout()
        Me.fraWave.SuspendLayout()
        Me.fraWind.SuspendLayout()
        Me.fraCurEnv.SuspendLayout()
        CType(Me.btrDuration, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblAngleUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblCurr, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblGenCmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblLengthUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVelUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVslSt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblWave, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblWaveUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblWind, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblWindUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCurr, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtVslSt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtWave, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtWind, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MainMenu1.SuspendLayout()
        Me.SuspendLayout()
        '
        'stbStatus
        '
        Me.stbStatus.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me._stbStatus_Panel1})
        Me.stbStatus.Location = New System.Drawing.Point(0, 503)
        Me.stbStatus.Name = "stbStatus"
        Me.stbStatus.Size = New System.Drawing.Size(595, 25)
        Me.stbStatus.TabIndex = 42
        '
        '_stbStatus_Panel1
        '
        Me._stbStatus_Panel1.AutoSize = False
        Me._stbStatus_Panel1.BorderSides = CType((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me._stbStatus_Panel1.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter
        Me._stbStatus_Panel1.Margin = New System.Windows.Forms.Padding(0)
        Me._stbStatus_Panel1.Name = "_stbStatus_Panel1"
        Me._stbStatus_Panel1.Size = New System.Drawing.Size(96, 25)
        Me._stbStatus_Panel1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.SystemColors.Control
        Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCancel.Location = New System.Drawing.Point(514, 37)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCancel.Size = New System.Drawing.Size(73, 25)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'btnOK
        '
        Me.btnOK.BackColor = System.Drawing.SystemColors.Control
        Me.btnOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnOK.Location = New System.Drawing.Point(514, 7)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnOK.Size = New System.Drawing.Size(73, 25)
        Me.btnOK.TabIndex = 0
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = False
        '
        'fraForce
        '
        Me.fraForce.BackColor = System.Drawing.SystemColors.Control
        Me.fraForce.Controls.Add(Me.grdForce)
        Me.fraForce.Controls.Add(Me.btnForce)
        Me.fraForce.Controls.Add(Me.picEnviron)
        Me.fraForce.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraForce.Location = New System.Drawing.Point(240, 64)
        Me.fraForce.Name = "fraForce"
        Me.fraForce.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraForce.Size = New System.Drawing.Size(265, 369)
        Me.fraForce.TabIndex = 38
        Me.fraForce.TabStop = False
        Me.fraForce.Text = "Environmental Forces :"
        '
        'grdForce
        '
        Me.grdForce.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdForce.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1})
        Me.grdForce.Location = New System.Drawing.Point(16, 198)
        Me.grdForce.Name = "grdForce"
        Me.grdForce.RowHeadersWidth = 75
        Me.grdForce.Size = New System.Drawing.Size(240, 137)
        Me.grdForce.TabIndex = 44
        '
        'Column1
        '
        Me.Column1.HeaderText = "Column1"
        Me.Column1.Name = "Column1"
        '
        'btnForce
        '
        Me.btnForce.BackColor = System.Drawing.SystemColors.Control
        Me.btnForce.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnForce.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnForce.Location = New System.Drawing.Point(184, 341)
        Me.btnForce.Name = "btnForce"
        Me.btnForce.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnForce.Size = New System.Drawing.Size(72, 25)
        Me.btnForce.TabIndex = 43
        Me.btnForce.Text = "Refresh"
        Me.btnForce.UseVisualStyleBackColor = False
        '
        'picEnviron
        '
        Me.picEnviron.BackColor = System.Drawing.SystemColors.Control
        Me.picEnviron.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.picEnviron.Cursor = System.Windows.Forms.Cursors.Default
        Me.picEnviron.ForeColor = System.Drawing.SystemColors.ControlText
        Me.picEnviron.Location = New System.Drawing.Point(12, 20)
        Me.picEnviron.Name = "picEnviron"
        Me.picEnviron.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.picEnviron.Size = New System.Drawing.Size(240, 172)
        Me.picEnviron.TabIndex = 39
        Me.picEnviron.TabStop = False
        '
        'fraVesselLoc
        '
        Me.fraVesselLoc.BackColor = System.Drawing.SystemColors.Control
        Me.fraVesselLoc.Controls.Add(Me._txtVslSt_0)
        Me.fraVesselLoc.Controls.Add(Me._txtVslSt_1)
        Me.fraVesselLoc.Controls.Add(Me._lblVslSt_0)
        Me.fraVesselLoc.Controls.Add(Me._lblAngleUnit_0)
        Me.fraVesselLoc.Controls.Add(Me._lblVslSt_1)
        Me.fraVesselLoc.Controls.Add(Me._lblLengthUnit_0)
        Me.fraVesselLoc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraVesselLoc.Location = New System.Drawing.Point(240, 0)
        Me.fraVesselLoc.Name = "fraVesselLoc"
        Me.fraVesselLoc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraVesselLoc.Size = New System.Drawing.Size(265, 57)
        Me.fraVesselLoc.TabIndex = 5
        Me.fraVesselLoc.TabStop = False
        Me.fraVesselLoc.Text = "Current Vessel Orientation :"
        '
        '_txtVslSt_0
        '
        Me._txtVslSt_0.AcceptsReturn = True
        Me._txtVslSt_0.BackColor = System.Drawing.SystemColors.Control
        Me._txtVslSt_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtVslSt_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtVslSt.SetIndex(Me._txtVslSt_0, CType(0, Short))
        Me._txtVslSt_0.Location = New System.Drawing.Point(72, 24)
        Me._txtVslSt_0.MaxLength = 0
        Me._txtVslSt_0.Name = "_txtVslSt_0"
        Me._txtVslSt_0.ReadOnly = True
        Me._txtVslSt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtVslSt_0.Size = New System.Drawing.Size(49, 20)
        Me._txtVslSt_0.TabIndex = 7
        Me._txtVslSt_0.Text = "0"
        Me._txtVslSt_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtVslSt_1
        '
        Me._txtVslSt_1.AcceptsReturn = True
        Me._txtVslSt_1.BackColor = System.Drawing.SystemColors.Control
        Me._txtVslSt_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtVslSt_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtVslSt.SetIndex(Me._txtVslSt_1, CType(1, Short))
        Me._txtVslSt_1.Location = New System.Drawing.Point(185, 24)
        Me._txtVslSt_1.MaxLength = 0
        Me._txtVslSt_1.Name = "_txtVslSt_1"
        Me._txtVslSt_1.ReadOnly = True
        Me._txtVslSt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtVslSt_1.Size = New System.Drawing.Size(49, 20)
        Me._txtVslSt_1.TabIndex = 10
        Me._txtVslSt_1.Text = "0"
        Me._txtVslSt_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_lblVslSt_0
        '
        Me._lblVslSt_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblVslSt_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVslSt_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVslSt.SetIndex(Me._lblVslSt_0, CType(0, Short))
        Me._lblVslSt_0.Location = New System.Drawing.Point(12, 24)
        Me._lblVslSt_0.Name = "_lblVslSt_0"
        Me._lblVslSt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVslSt_0.Size = New System.Drawing.Size(53, 17)
        Me._lblVslSt_0.TabIndex = 6
        Me._lblVslSt_0.Text = "Heading"
        '
        '_lblAngleUnit_0
        '
        Me._lblAngleUnit_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblAngleUnit_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblAngleUnit_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblAngleUnit.SetIndex(Me._lblAngleUnit_0, CType(0, Short))
        Me._lblAngleUnit_0.Location = New System.Drawing.Point(124, 24)
        Me._lblAngleUnit_0.Name = "_lblAngleUnit_0"
        Me._lblAngleUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblAngleUnit_0.Size = New System.Drawing.Size(17, 17)
        Me._lblAngleUnit_0.TabIndex = 8
        '
        '_lblVslSt_1
        '
        Me._lblVslSt_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblVslSt_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVslSt_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVslSt.SetIndex(Me._lblVslSt_1, CType(1, Short))
        Me._lblVslSt_1.Location = New System.Drawing.Point(157, 24)
        Me._lblVslSt_1.Name = "_lblVslSt_1"
        Me._lblVslSt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVslSt_1.Size = New System.Drawing.Size(33, 17)
        Me._lblVslSt_1.TabIndex = 9
        Me._lblVslSt_1.Text = "Draft"
        '
        '_lblLengthUnit_0
        '
        Me._lblLengthUnit_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_0, CType(0, Short))
        Me._lblLengthUnit_0.Location = New System.Drawing.Point(240, 26)
        Me._lblLengthUnit_0.Name = "_lblLengthUnit_0"
        Me._lblLengthUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_0.Size = New System.Drawing.Size(19, 20)
        Me._lblLengthUnit_0.TabIndex = 11
        Me._lblLengthUnit_0.Text = "ft"
        '
        'fraCurrent
        '
        Me.fraCurrent.BackColor = System.Drawing.SystemColors.Control
        Me.fraCurrent.Controls.Add(Me.btnProfile)
        Me.fraCurrent.Controls.Add(Me._txtCurr_1)
        Me.fraCurrent.Controls.Add(Me._txtCurr_0)
        Me.fraCurrent.Controls.Add(Me._lblAngleUnit_3)
        Me.fraCurrent.Controls.Add(Me._lblVelUnit_1)
        Me.fraCurrent.Controls.Add(Me._lblCurr_1)
        Me.fraCurrent.Controls.Add(Me._lblCurr_0)
        Me.fraCurrent.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraCurrent.Location = New System.Drawing.Point(8, 320)
        Me.fraCurrent.Name = "fraCurrent"
        Me.fraCurrent.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraCurrent.Size = New System.Drawing.Size(225, 113)
        Me.fraCurrent.TabIndex = 32
        Me.fraCurrent.TabStop = False
        Me.fraCurrent.Text = "Current :"
        '
        'btnProfile
        '
        Me.btnProfile.BackColor = System.Drawing.SystemColors.Control
        Me.btnProfile.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnProfile.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnProfile.Location = New System.Drawing.Point(113, 73)
        Me.btnProfile.Name = "btnProfile"
        Me.btnProfile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnProfile.Size = New System.Drawing.Size(63, 25)
        Me.btnProfile.TabIndex = 37
        Me.btnProfile.Text = "&Profile"
        Me.btnProfile.UseVisualStyleBackColor = False
        '
        '_txtCurr_1
        '
        Me._txtCurr_1.AcceptsReturn = True
        Me._txtCurr_1.BackColor = System.Drawing.SystemColors.Window
        Me._txtCurr_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtCurr_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCurr.SetIndex(Me._txtCurr_1, CType(1, Short))
        Me._txtCurr_1.Location = New System.Drawing.Point(114, 48)
        Me._txtCurr_1.MaxLength = 0
        Me._txtCurr_1.Name = "_txtCurr_1"
        Me._txtCurr_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtCurr_1.Size = New System.Drawing.Size(57, 20)
        Me._txtCurr_1.TabIndex = 36
        Me._txtCurr_1.Text = "0"
        Me._txtCurr_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtCurr_0
        '
        Me._txtCurr_0.AcceptsReturn = True
        Me._txtCurr_0.BackColor = System.Drawing.SystemColors.Window
        Me._txtCurr_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtCurr_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCurr.SetIndex(Me._txtCurr_0, CType(0, Short))
        Me._txtCurr_0.Location = New System.Drawing.Point(114, 24)
        Me._txtCurr_0.MaxLength = 0
        Me._txtCurr_0.Name = "_txtCurr_0"
        Me._txtCurr_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtCurr_0.Size = New System.Drawing.Size(57, 20)
        Me._txtCurr_0.TabIndex = 34
        Me._txtCurr_0.Text = "0"
        Me._txtCurr_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_lblAngleUnit_3
        '
        Me._lblAngleUnit_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblAngleUnit_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblAngleUnit_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblAngleUnit.SetIndex(Me._lblAngleUnit_3, CType(3, Short))
        Me._lblAngleUnit_3.Location = New System.Drawing.Point(173, 49)
        Me._lblAngleUnit_3.Name = "_lblAngleUnit_3"
        Me._lblAngleUnit_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblAngleUnit_3.Size = New System.Drawing.Size(17, 17)
        Me._lblAngleUnit_3.TabIndex = 49
        '
        '_lblVelUnit_1
        '
        Me._lblVelUnit_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblVelUnit_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVelUnit_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVelUnit.SetIndex(Me._lblVelUnit_1, CType(1, Short))
        Me._lblVelUnit_1.Location = New System.Drawing.Point(175, 27)
        Me._lblVelUnit_1.Name = "_lblVelUnit_1"
        Me._lblVelUnit_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVelUnit_1.Size = New System.Drawing.Size(34, 17)
        Me._lblVelUnit_1.TabIndex = 46
        Me._lblVelUnit_1.Text = "kn"
        '
        '_lblCurr_1
        '
        Me._lblCurr_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblCurr_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblCurr_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCurr.SetIndex(Me._lblCurr_1, CType(1, Short))
        Me._lblCurr_1.Location = New System.Drawing.Point(16, 48)
        Me._lblCurr_1.Name = "_lblCurr_1"
        Me._lblCurr_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblCurr_1.Size = New System.Drawing.Size(65, 17)
        Me._lblCurr_1.TabIndex = 35
        Me._lblCurr_1.Text = "Heading"
        '
        '_lblCurr_0
        '
        Me._lblCurr_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblCurr_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblCurr_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCurr.SetIndex(Me._lblCurr_0, CType(0, Short))
        Me._lblCurr_0.Location = New System.Drawing.Point(16, 24)
        Me._lblCurr_0.Name = "_lblCurr_0"
        Me._lblCurr_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblCurr_0.Size = New System.Drawing.Size(89, 17)
        Me._lblCurr_0.TabIndex = 33
        Me._lblCurr_0.Text = "Surface Velocity"
        '
        'fraWave
        '
        Me.fraWave.BackColor = System.Drawing.SystemColors.Control
        Me.fraWave.Controls.Add(Me._txtWave_2)
        Me.fraWave.Controls.Add(Me._txtWave_1)
        Me.fraWave.Controls.Add(Me._txtWave_0)
        Me.fraWave.Controls.Add(Me._lblAngleUnit_2)
        Me.fraWave.Controls.Add(Me._lblLengthUnit_2)
        Me.fraWave.Controls.Add(Me._lblWave_2)
        Me.fraWave.Controls.Add(Me._lblWave_1)
        Me.fraWave.Controls.Add(Me._lblWaveUnit_1)
        Me.fraWave.Controls.Add(Me._lblWave_0)
        Me.fraWave.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraWave.Location = New System.Drawing.Point(8, 207)
        Me.fraWave.Name = "fraWave"
        Me.fraWave.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraWave.Size = New System.Drawing.Size(225, 105)
        Me.fraWave.TabIndex = 24
        Me.fraWave.TabStop = False
        Me.fraWave.Text = "Wave :"
        '
        '_txtWave_2
        '
        Me._txtWave_2.AcceptsReturn = True
        Me._txtWave_2.BackColor = System.Drawing.SystemColors.Window
        Me._txtWave_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtWave_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtWave.SetIndex(Me._txtWave_2, CType(2, Short))
        Me._txtWave_2.Location = New System.Drawing.Point(80, 72)
        Me._txtWave_2.MaxLength = 0
        Me._txtWave_2.Name = "_txtWave_2"
        Me._txtWave_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtWave_2.Size = New System.Drawing.Size(57, 20)
        Me._txtWave_2.TabIndex = 31
        Me._txtWave_2.Text = "0"
        Me._txtWave_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtWave_1
        '
        Me._txtWave_1.AcceptsReturn = True
        Me._txtWave_1.BackColor = System.Drawing.SystemColors.Window
        Me._txtWave_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtWave_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtWave.SetIndex(Me._txtWave_1, CType(1, Short))
        Me._txtWave_1.Location = New System.Drawing.Point(80, 48)
        Me._txtWave_1.MaxLength = 0
        Me._txtWave_1.Name = "_txtWave_1"
        Me._txtWave_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtWave_1.Size = New System.Drawing.Size(57, 20)
        Me._txtWave_1.TabIndex = 28
        Me._txtWave_1.Text = "0"
        Me._txtWave_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtWave_0
        '
        Me._txtWave_0.AcceptsReturn = True
        Me._txtWave_0.BackColor = System.Drawing.SystemColors.Window
        Me._txtWave_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtWave_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtWave.SetIndex(Me._txtWave_0, CType(0, Short))
        Me._txtWave_0.Location = New System.Drawing.Point(80, 24)
        Me._txtWave_0.MaxLength = 0
        Me._txtWave_0.Name = "_txtWave_0"
        Me._txtWave_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtWave_0.Size = New System.Drawing.Size(57, 20)
        Me._txtWave_0.TabIndex = 26
        Me._txtWave_0.Text = "0"
        Me._txtWave_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_lblAngleUnit_2
        '
        Me._lblAngleUnit_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblAngleUnit_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblAngleUnit_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblAngleUnit.SetIndex(Me._lblAngleUnit_2, CType(2, Short))
        Me._lblAngleUnit_2.Location = New System.Drawing.Point(140, 72)
        Me._lblAngleUnit_2.Name = "_lblAngleUnit_2"
        Me._lblAngleUnit_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblAngleUnit_2.Size = New System.Drawing.Size(17, 17)
        Me._lblAngleUnit_2.TabIndex = 48
        '
        '_lblLengthUnit_2
        '
        Me._lblLengthUnit_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_2, CType(2, Short))
        Me._lblLengthUnit_2.Location = New System.Drawing.Point(143, 26)
        Me._lblLengthUnit_2.Name = "_lblLengthUnit_2"
        Me._lblLengthUnit_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_2.Size = New System.Drawing.Size(24, 18)
        Me._lblLengthUnit_2.TabIndex = 45
        Me._lblLengthUnit_2.Text = "ft"
        '
        '_lblWave_2
        '
        Me._lblWave_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblWave_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblWave_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblWave.SetIndex(Me._lblWave_2, CType(2, Short))
        Me._lblWave_2.Location = New System.Drawing.Point(16, 72)
        Me._lblWave_2.Name = "_lblWave_2"
        Me._lblWave_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblWave_2.Size = New System.Drawing.Size(65, 17)
        Me._lblWave_2.TabIndex = 30
        Me._lblWave_2.Text = "Heading"
        '
        '_lblWave_1
        '
        Me._lblWave_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblWave_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblWave_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblWave.SetIndex(Me._lblWave_1, CType(1, Short))
        Me._lblWave_1.Location = New System.Drawing.Point(16, 48)
        Me._lblWave_1.Name = "_lblWave_1"
        Me._lblWave_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblWave_1.Size = New System.Drawing.Size(65, 17)
        Me._lblWave_1.TabIndex = 27
        Me._lblWave_1.Text = "Peak Period"
        '
        '_lblWaveUnit_1
        '
        Me._lblWaveUnit_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblWaveUnit_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblWaveUnit_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblWaveUnit.SetIndex(Me._lblWaveUnit_1, CType(1, Short))
        Me._lblWaveUnit_1.Location = New System.Drawing.Point(142, 47)
        Me._lblWaveUnit_1.Name = "_lblWaveUnit_1"
        Me._lblWaveUnit_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblWaveUnit_1.Size = New System.Drawing.Size(25, 17)
        Me._lblWaveUnit_1.TabIndex = 29
        Me._lblWaveUnit_1.Text = "sec"
        '
        '_lblWave_0
        '
        Me._lblWave_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblWave_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblWave_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblWave.SetIndex(Me._lblWave_0, CType(0, Short))
        Me._lblWave_0.Location = New System.Drawing.Point(16, 24)
        Me._lblWave_0.Name = "_lblWave_0"
        Me._lblWave_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblWave_0.Size = New System.Drawing.Size(65, 17)
        Me._lblWave_0.TabIndex = 25
        Me._lblWave_0.Text = "Sig. Height"
        '
        'fraWind
        '
        Me.fraWind.BackColor = System.Drawing.SystemColors.Control
        Me.fraWind.Controls.Add(Me._btrDuration_2)
        Me.fraWind.Controls.Add(Me._btrDuration_1)
        Me.fraWind.Controls.Add(Me._btrDuration_0)
        Me.fraWind.Controls.Add(Me._txtWind_2)
        Me.fraWind.Controls.Add(Me._txtWind_1)
        Me.fraWind.Controls.Add(Me._txtWind_0)
        Me.fraWind.Controls.Add(Me._lblAngleUnit_1)
        Me.fraWind.Controls.Add(Me._lblLengthUnit_1)
        Me.fraWind.Controls.Add(Me._lblWind_2)
        Me.fraWind.Controls.Add(Me._lblWind_1)
        Me.fraWind.Controls.Add(Me._lblWindUnit_1)
        Me.fraWind.Controls.Add(Me._lblWind_0)
        Me.fraWind.Controls.Add(Me._lblVelUnit_0)
        Me.fraWind.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraWind.Location = New System.Drawing.Point(8, 64)
        Me.fraWind.Name = "fraWind"
        Me.fraWind.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraWind.Size = New System.Drawing.Size(225, 137)
        Me.fraWind.TabIndex = 12
        Me.fraWind.TabStop = False
        Me.fraWind.Text = "Wind :"
        '
        '_btrDuration_2
        '
        Me._btrDuration_2.BackColor = System.Drawing.SystemColors.Control
        Me._btrDuration_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._btrDuration_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btrDuration.SetIndex(Me._btrDuration_2, CType(2, Short))
        Me._btrDuration_2.Location = New System.Drawing.Point(136, 72)
        Me._btrDuration_2.Name = "_btrDuration_2"
        Me._btrDuration_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btrDuration_2.Size = New System.Drawing.Size(83, 17)
        Me._btrDuration_2.TabIndex = 21
        Me._btrDuration_2.TabStop = True
        Me._btrDuration_2.Text = "3 seconds"
        Me._btrDuration_2.UseVisualStyleBackColor = False
        '
        '_btrDuration_1
        '
        Me._btrDuration_1.BackColor = System.Drawing.SystemColors.Control
        Me._btrDuration_1.Checked = True
        Me._btrDuration_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._btrDuration_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btrDuration.SetIndex(Me._btrDuration_1, CType(1, Short))
        Me._btrDuration_1.Location = New System.Drawing.Point(72, 72)
        Me._btrDuration_1.Name = "_btrDuration_1"
        Me._btrDuration_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btrDuration_1.Size = New System.Drawing.Size(65, 17)
        Me._btrDuration_1.TabIndex = 20
        Me._btrDuration_1.TabStop = True
        Me._btrDuration_1.Text = "1 minute"
        Me._btrDuration_1.UseVisualStyleBackColor = False
        '
        '_btrDuration_0
        '
        Me._btrDuration_0.BackColor = System.Drawing.SystemColors.Control
        Me._btrDuration_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._btrDuration_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btrDuration.SetIndex(Me._btrDuration_0, CType(0, Short))
        Me._btrDuration_0.Location = New System.Drawing.Point(16, 72)
        Me._btrDuration_0.Name = "_btrDuration_0"
        Me._btrDuration_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btrDuration_0.Size = New System.Drawing.Size(65, 17)
        Me._btrDuration_0.TabIndex = 19
        Me._btrDuration_0.TabStop = True
        Me._btrDuration_0.Text = "1 hour"
        Me._btrDuration_0.UseVisualStyleBackColor = False
        '
        '_txtWind_2
        '
        Me._txtWind_2.AcceptsReturn = True
        Me._txtWind_2.BackColor = System.Drawing.SystemColors.Window
        Me._txtWind_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtWind_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtWind.SetIndex(Me._txtWind_2, CType(2, Short))
        Me._txtWind_2.Location = New System.Drawing.Point(80, 104)
        Me._txtWind_2.MaxLength = 0
        Me._txtWind_2.Name = "_txtWind_2"
        Me._txtWind_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtWind_2.Size = New System.Drawing.Size(57, 20)
        Me._txtWind_2.TabIndex = 23
        Me._txtWind_2.Text = "0"
        Me._txtWind_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtWind_1
        '
        Me._txtWind_1.AcceptsReturn = True
        Me._txtWind_1.BackColor = System.Drawing.SystemColors.Window
        Me._txtWind_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtWind_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtWind.SetIndex(Me._txtWind_1, CType(1, Short))
        Me._txtWind_1.Location = New System.Drawing.Point(80, 48)
        Me._txtWind_1.MaxLength = 0
        Me._txtWind_1.Name = "_txtWind_1"
        Me._txtWind_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtWind_1.Size = New System.Drawing.Size(57, 20)
        Me._txtWind_1.TabIndex = 17
        Me._txtWind_1.Text = "0"
        Me._txtWind_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtWind_0
        '
        Me._txtWind_0.AcceptsReturn = True
        Me._txtWind_0.BackColor = System.Drawing.SystemColors.Window
        Me._txtWind_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtWind_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtWind.SetIndex(Me._txtWind_0, CType(0, Short))
        Me._txtWind_0.Location = New System.Drawing.Point(80, 24)
        Me._txtWind_0.MaxLength = 0
        Me._txtWind_0.Name = "_txtWind_0"
        Me._txtWind_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtWind_0.Size = New System.Drawing.Size(57, 20)
        Me._txtWind_0.TabIndex = 14
        Me._txtWind_0.Text = "0"
        Me._txtWind_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_lblAngleUnit_1
        '
        Me._lblAngleUnit_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblAngleUnit_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblAngleUnit_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblAngleUnit.SetIndex(Me._lblAngleUnit_1, CType(1, Short))
        Me._lblAngleUnit_1.Location = New System.Drawing.Point(140, 104)
        Me._lblAngleUnit_1.Name = "_lblAngleUnit_1"
        Me._lblAngleUnit_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblAngleUnit_1.Size = New System.Drawing.Size(17, 17)
        Me._lblAngleUnit_1.TabIndex = 47
        '
        '_lblLengthUnit_1
        '
        Me._lblLengthUnit_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_1, CType(1, Short))
        Me._lblLengthUnit_1.Location = New System.Drawing.Point(142, 51)
        Me._lblLengthUnit_1.Name = "_lblLengthUnit_1"
        Me._lblLengthUnit_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_1.Size = New System.Drawing.Size(15, 17)
        Me._lblLengthUnit_1.TabIndex = 44
        Me._lblLengthUnit_1.Text = "ft"
        '
        '_lblWind_2
        '
        Me._lblWind_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblWind_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblWind_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblWind.SetIndex(Me._lblWind_2, CType(2, Short))
        Me._lblWind_2.Location = New System.Drawing.Point(16, 104)
        Me._lblWind_2.Name = "_lblWind_2"
        Me._lblWind_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblWind_2.Size = New System.Drawing.Size(65, 17)
        Me._lblWind_2.TabIndex = 22
        Me._lblWind_2.Text = "Heading"
        '
        '_lblWind_1
        '
        Me._lblWind_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblWind_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblWind_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblWind.SetIndex(Me._lblWind_1, CType(1, Short))
        Me._lblWind_1.Location = New System.Drawing.Point(16, 48)
        Me._lblWind_1.Name = "_lblWind_1"
        Me._lblWind_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblWind_1.Size = New System.Drawing.Size(65, 17)
        Me._lblWind_1.TabIndex = 16
        Me._lblWind_1.Text = "measured at"
        '
        '_lblWindUnit_1
        '
        Me._lblWindUnit_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblWindUnit_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblWindUnit_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblWindUnit.SetIndex(Me._lblWindUnit_1, CType(1, Short))
        Me._lblWindUnit_1.Location = New System.Drawing.Point(156, 51)
        Me._lblWindUnit_1.Name = "_lblWindUnit_1"
        Me._lblWindUnit_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblWindUnit_1.Size = New System.Drawing.Size(61, 17)
        Me._lblWindUnit_1.TabIndex = 18
        Me._lblWindUnit_1.Text = "above WL"
        '
        '_lblWind_0
        '
        Me._lblWind_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblWind_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblWind_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblWind.SetIndex(Me._lblWind_0, CType(0, Short))
        Me._lblWind_0.Location = New System.Drawing.Point(16, 24)
        Me._lblWind_0.Name = "_lblWind_0"
        Me._lblWind_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblWind_0.Size = New System.Drawing.Size(65, 17)
        Me._lblWind_0.TabIndex = 13
        Me._lblWind_0.Text = "Velocity"
        '
        '_lblVelUnit_0
        '
        Me._lblVelUnit_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblVelUnit_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVelUnit_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVelUnit.SetIndex(Me._lblVelUnit_0, CType(0, Short))
        Me._lblVelUnit_0.Location = New System.Drawing.Point(144, 24)
        Me._lblVelUnit_0.Name = "_lblVelUnit_0"
        Me._lblVelUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVelUnit_0.Size = New System.Drawing.Size(25, 17)
        Me._lblVelUnit_0.TabIndex = 15
        Me._lblVelUnit_0.Text = "kn"
        '
        'fraCurEnv
        '
        Me.fraCurEnv.BackColor = System.Drawing.SystemColors.Control
        Me.fraCurEnv.Controls.Add(Me.btnAdd)
        Me.fraCurEnv.Controls.Add(Me.cboCurEnv)
        Me.fraCurEnv.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraCurEnv.Location = New System.Drawing.Point(8, 0)
        Me.fraCurEnv.Name = "fraCurEnv"
        Me.fraCurEnv.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraCurEnv.Size = New System.Drawing.Size(225, 57)
        Me.fraCurEnv.TabIndex = 2
        Me.fraCurEnv.TabStop = False
        Me.fraCurEnv.Text = "Current Environment :"
        '
        'btnAdd
        '
        Me.btnAdd.BackColor = System.Drawing.SystemColors.Control
        Me.btnAdd.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnAdd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnAdd.Location = New System.Drawing.Point(144, 20)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnAdd.Size = New System.Drawing.Size(73, 25)
        Me.btnAdd.TabIndex = 4
        Me.btnAdd.Text = "&Add/Update"
        Me.btnAdd.UseVisualStyleBackColor = False
        '
        'cboCurEnv
        '
        Me.cboCurEnv.BackColor = System.Drawing.SystemColors.Window
        Me.cboCurEnv.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboCurEnv.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboCurEnv.Location = New System.Drawing.Point(16, 24)
        Me.cboCurEnv.Name = "cboCurEnv"
        Me.cboCurEnv.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboCurEnv.Size = New System.Drawing.Size(121, 21)
        Me.cboCurEnv.TabIndex = 3
        Me.cboCurEnv.Text = "Current Environment"
        '
        '_lblGenCmt_0
        '
        Me._lblGenCmt_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblGenCmt_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblGenCmt_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblGenCmt.SetIndex(Me._lblGenCmt_0, CType(0, Short))
        Me._lblGenCmt_0.Location = New System.Drawing.Point(24, 440)
        Me._lblGenCmt_0.Name = "_lblGenCmt_0"
        Me._lblGenCmt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblGenCmt_0.Size = New System.Drawing.Size(481, 65)
        Me._lblGenCmt_0.TabIndex = 41
        Me._lblGenCmt_0.Text = resources.GetString("_lblGenCmt_0.Text")
        '
        'btrDuration
        '
        '
        'txtCurr
        '
        '
        'txtWave
        '
        '
        'txtWind
        '
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFile})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(595, 24)
        Me.MainMenu1.TabIndex = 43
        '
        'mnuFile
        '
        Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuOpen, Me.mnuSave, Me.mnuDum, Me.mnuClose})
        Me.mnuFile.Name = "mnuFile"
        Me.mnuFile.Size = New System.Drawing.Size(37, 20)
        Me.mnuFile.Text = "&File"
        '
        'mnuOpen
        '
        Me.mnuOpen.Name = "mnuOpen"
        Me.mnuOpen.Size = New System.Drawing.Size(122, 22)
        Me.mnuOpen.Text = "&Import ..."
        '
        'mnuSave
        '
        Me.mnuSave.Name = "mnuSave"
        Me.mnuSave.Size = New System.Drawing.Size(122, 22)
        Me.mnuSave.Text = "&Export ..."
        '
        'mnuDum
        '
        Me.mnuDum.Name = "mnuDum"
        Me.mnuDum.Size = New System.Drawing.Size(119, 6)
        '
        'mnuClose
        '
        Me.mnuClose.Name = "mnuClose"
        Me.mnuClose.Size = New System.Drawing.Size(122, 22)
        Me.mnuClose.Text = "&Close"
        '
        'frmEnviron
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(595, 528)
        Me.Controls.Add(Me.stbStatus)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.fraForce)
        Me.Controls.Add(Me.fraVesselLoc)
        Me.Controls.Add(Me.fraCurrent)
        Me.Controls.Add(Me.fraWave)
        Me.Controls.Add(Me.fraWind)
        Me.Controls.Add(Me.fraCurEnv)
        Me.Controls.Add(Me._lblGenCmt_0)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(10, 36)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmEnviron"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = " Environment Condition"
        Me.stbStatus.ResumeLayout(False)
        Me.stbStatus.PerformLayout()
        Me.fraForce.ResumeLayout(False)
        CType(Me.grdForce, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picEnviron, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraVesselLoc.ResumeLayout(False)
        Me.fraVesselLoc.PerformLayout()
        Me.fraCurrent.ResumeLayout(False)
        Me.fraCurrent.PerformLayout()
        Me.fraWave.ResumeLayout(False)
        Me.fraWave.PerformLayout()
        Me.fraWind.ResumeLayout(False)
        Me.fraWind.PerformLayout()
        Me.fraCurEnv.ResumeLayout(False)
        CType(Me.btrDuration, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblAngleUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblCurr, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblGenCmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblLengthUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVelUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVslSt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblWave, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblWaveUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblWind, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblWindUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCurr, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtVslSt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtWave, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtWind, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents grdForce As System.Windows.Forms.DataGridView
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
#End Region
End Class