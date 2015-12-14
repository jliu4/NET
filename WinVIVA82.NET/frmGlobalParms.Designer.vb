<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmGlobalParms
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
    Public dlgFilesOpen As System.Windows.Forms.OpenFileDialog
    Public WithEvents txtWaterDensity As System.Windows.Forms.TextBox
    Public WithEvents txtVis As System.Windows.Forms.TextBox
	Public WithEvents lblVisUnits As System.Windows.Forms.Label
	Public WithEvents lblWaterDensity As System.Windows.Forms.Label
    Public WithEvents lblWaterDensityUnits As System.Windows.Forms.Label
	Public WithEvents lblVis As System.Windows.Forms.Label
	Public WithEvents fraParms As System.Windows.Forms.GroupBox
	Public WithEvents btnCancel As System.Windows.Forms.Button
	Public WithEvents btnOK As System.Windows.Forms.Button
	Public WithEvents btnEnglish As System.Windows.Forms.RadioButton
	Public WithEvents btnMetric As System.Windows.Forms.RadioButton
	Public WithEvents fraUnits As System.Windows.Forms.GroupBox
	Public WithEvents cmdBrowse As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
	Public WithEvents optModes As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	Public WithEvents txtModesFile As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmGlobalParms))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.dlgFilesOpen = New System.Windows.Forms.OpenFileDialog
        Me.fraParms = New System.Windows.Forms.GroupBox
        Me.txtWaterDensity = New System.Windows.Forms.TextBox
        Me.txtVis = New System.Windows.Forms.TextBox
        Me.lblVisUnits = New System.Windows.Forms.Label
        Me.lblWaterDensity = New System.Windows.Forms.Label
        Me.lblWaterDensityUnits = New System.Windows.Forms.Label
        Me.lblVis = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.fraUnits = New System.Windows.Forms.GroupBox
        Me.btnEnglish = New System.Windows.Forms.RadioButton
        Me.btnMetric = New System.Windows.Forms.RadioButton
        Me.cmdBrowse = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(Me.components)
        Me.optModes = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.txtModesFile = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.mDamping = New System.Windows.Forms.RadioButton
        Me.cDamping = New System.Windows.Forms.RadioButton
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.btnTwoRisers = New System.Windows.Forms.RadioButton
        Me.btnOneRiser = New System.Windows.Forms.RadioButton
        Me.TwoRisers = New System.Windows.Forms.GroupBox
        Me.chkbtn2 = New System.Windows.Forms.CheckBox
        Me.lblDistance3 = New System.Windows.Forms.Label
        Me.txtRiserRZ = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtRiserFY = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtRiserRY = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtRiserFZ = New System.Windows.Forms.TextBox
        Me.lblDistance4 = New System.Windows.Forms.Label
        Me.lblDistance1 = New System.Windows.Forms.Label
        Me.lblDistance2 = New System.Windows.Forms.Label
        Me.fraParms.SuspendLayout()
        Me.fraUnits.SuspendLayout()
        CType(Me.cmdBrowse, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optModes, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtModesFile, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.TwoRisers.SuspendLayout()
        Me.SuspendLayout()
        '
        'fraParms
        '
        Me.fraParms.BackColor = System.Drawing.SystemColors.Control
        Me.fraParms.Controls.Add(Me.txtWaterDensity)
        Me.fraParms.Controls.Add(Me.txtVis)
        Me.fraParms.Controls.Add(Me.lblVisUnits)
        Me.fraParms.Controls.Add(Me.lblWaterDensity)
        Me.fraParms.Controls.Add(Me.lblWaterDensityUnits)
        Me.fraParms.Controls.Add(Me.lblVis)
        Me.fraParms.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraParms.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraParms.Location = New System.Drawing.Point(12, 89)
        Me.fraParms.Name = "fraParms"
        Me.fraParms.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraParms.Size = New System.Drawing.Size(359, 84)
        Me.fraParms.TabIndex = 9
        Me.fraParms.TabStop = False
        Me.fraParms.Text = "Riser Global Parameters"
        '
        'txtWaterDensity
        '
        Me.txtWaterDensity.AcceptsReturn = True
        Me.txtWaterDensity.BackColor = System.Drawing.SystemColors.Window
        Me.txtWaterDensity.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtWaterDensity.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWaterDensity.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtWaterDensity.Location = New System.Drawing.Point(148, 24)
        Me.txtWaterDensity.MaxLength = 0
        Me.txtWaterDensity.Name = "txtWaterDensity"
        Me.txtWaterDensity.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtWaterDensity.Size = New System.Drawing.Size(149, 20)
        Me.txtWaterDensity.TabIndex = 5
        Me.txtWaterDensity.Text = "Text2"
        Me.txtWaterDensity.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtVis
        '
        Me.txtVis.AcceptsReturn = True
        Me.txtVis.BackColor = System.Drawing.SystemColors.Window
        Me.txtVis.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtVis.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVis.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtVis.Location = New System.Drawing.Point(148, 48)
        Me.txtVis.MaxLength = 0
        Me.txtVis.Name = "txtVis"
        Me.txtVis.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtVis.Size = New System.Drawing.Size(149, 20)
        Me.txtVis.TabIndex = 6
        Me.txtVis.Text = "Text1"
        Me.txtVis.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblVisUnits
        '
        Me.lblVisUnits.BackColor = System.Drawing.SystemColors.Control
        Me.lblVisUnits.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblVisUnits.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVisUnits.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVisUnits.Location = New System.Drawing.Point(304, 48)
        Me.lblVisUnits.Name = "lblVisUnits"
        Me.lblVisUnits.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblVisUnits.Size = New System.Drawing.Size(41, 17)
        Me.lblVisUnits.TabIndex = 13
        Me.lblVisUnits.Text = "m^2/s"
        '
        'lblWaterDensity
        '
        Me.lblWaterDensity.BackColor = System.Drawing.SystemColors.Control
        Me.lblWaterDensity.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblWaterDensity.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWaterDensity.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblWaterDensity.Location = New System.Drawing.Point(24, 24)
        Me.lblWaterDensity.Name = "lblWaterDensity"
        Me.lblWaterDensity.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblWaterDensity.Size = New System.Drawing.Size(129, 17)
        Me.lblWaterDensity.TabIndex = 12
        Me.lblWaterDensity.Text = "Sea Water Density"
        '
        'lblWaterDensityUnits
        '
        Me.lblWaterDensityUnits.BackColor = System.Drawing.SystemColors.Control
        Me.lblWaterDensityUnits.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblWaterDensityUnits.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWaterDensityUnits.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblWaterDensityUnits.Location = New System.Drawing.Point(303, 24)
        Me.lblWaterDensityUnits.Name = "lblWaterDensityUnits"
        Me.lblWaterDensityUnits.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblWaterDensityUnits.Size = New System.Drawing.Size(41, 17)
        Me.lblWaterDensityUnits.TabIndex = 10
        Me.lblWaterDensityUnits.Text = "kg/m^3"
        '
        'lblVis
        '
        Me.lblVis.BackColor = System.Drawing.SystemColors.Control
        Me.lblVis.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblVis.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVis.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVis.Location = New System.Drawing.Point(24, 48)
        Me.lblVis.Name = "lblVis"
        Me.lblVis.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblVis.Size = New System.Drawing.Size(129, 17)
        Me.lblVis.TabIndex = 9
        Me.lblVis.Text = "Sea Water Viscosity"
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.SystemColors.Control
        Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCancel.Location = New System.Drawing.Point(306, 355)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCancel.Size = New System.Drawing.Size(65, 25)
        Me.btnCancel.TabIndex = 14
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'btnOK
        '
        Me.btnOK.BackColor = System.Drawing.SystemColors.Control
        Me.btnOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnOK.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnOK.Location = New System.Drawing.Point(217, 355)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnOK.Size = New System.Drawing.Size(65, 25)
        Me.btnOK.TabIndex = 13
        Me.btnOK.Text = "&Save"
        Me.btnOK.UseVisualStyleBackColor = False
        '
        'fraUnits
        '
        Me.fraUnits.BackColor = System.Drawing.SystemColors.Control
        Me.fraUnits.Controls.Add(Me.btnEnglish)
        Me.fraUnits.Controls.Add(Me.btnMetric)
        Me.fraUnits.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraUnits.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraUnits.Location = New System.Drawing.Point(12, 8)
        Me.fraUnits.Name = "fraUnits"
        Me.fraUnits.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraUnits.Size = New System.Drawing.Size(111, 65)
        Me.fraUnits.TabIndex = 7
        Me.fraUnits.TabStop = False
        Me.fraUnits.Text = "Units"
        '
        'btnEnglish
        '
        Me.btnEnglish.BackColor = System.Drawing.SystemColors.Control
        Me.btnEnglish.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnEnglish.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnEnglish.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnEnglish.Location = New System.Drawing.Point(24, 40)
        Me.btnEnglish.Name = "btnEnglish"
        Me.btnEnglish.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnEnglish.Size = New System.Drawing.Size(73, 17)
        Me.btnEnglish.TabIndex = 3
        Me.btnEnglish.Text = "&English"
        Me.btnEnglish.UseVisualStyleBackColor = False
        '
        'btnMetric
        '
        Me.btnMetric.BackColor = System.Drawing.SystemColors.Control
        Me.btnMetric.Checked = True
        Me.btnMetric.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnMetric.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnMetric.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnMetric.Location = New System.Drawing.Point(24, 20)
        Me.btnMetric.Name = "btnMetric"
        Me.btnMetric.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnMetric.Size = New System.Drawing.Size(65, 17)
        Me.btnMetric.TabIndex = 2
        Me.btnMetric.TabStop = True
        Me.btnMetric.Text = "&Metric"
        Me.btnMetric.UseVisualStyleBackColor = False
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox1.Controls.Add(Me.mDamping)
        Me.GroupBox1.Controls.Add(Me.cDamping)
        Me.GroupBox1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.GroupBox1.Location = New System.Drawing.Point(129, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.GroupBox1.Size = New System.Drawing.Size(242, 65)
        Me.GroupBox1.TabIndex = 8
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Damping Application Method"
        '
        'mDamping
        '
        Me.mDamping.BackColor = System.Drawing.SystemColors.Control
        Me.mDamping.Checked = True
        Me.mDamping.Cursor = System.Windows.Forms.Cursors.Default
        Me.mDamping.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.mDamping.ForeColor = System.Drawing.SystemColors.ControlText
        Me.mDamping.Location = New System.Drawing.Point(24, 20)
        Me.mDamping.Name = "mDamping"
        Me.mDamping.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.mDamping.Size = New System.Drawing.Size(129, 17)
        Me.mDamping.TabIndex = 4
        Me.mDamping.TabStop = True
        Me.mDamping.Text = "&Modal Damping"
        Me.mDamping.UseVisualStyleBackColor = False
        '
        'cDamping
        '
        Me.cDamping.BackColor = System.Drawing.SystemColors.Control
        Me.cDamping.Cursor = System.Windows.Forms.Cursors.Default
        Me.cDamping.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cDamping.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cDamping.Location = New System.Drawing.Point(24, 40)
        Me.cDamping.Name = "cDamping"
        Me.cDamping.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cDamping.Size = New System.Drawing.Size(129, 19)
        Me.cDamping.TabIndex = 25
        Me.cDamping.Text = "&Constant Damping"
        Me.cDamping.UseVisualStyleBackColor = False
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox2.Controls.Add(Me.btnTwoRisers)
        Me.GroupBox2.Controls.Add(Me.btnOneRiser)
        Me.GroupBox2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.GroupBox2.Location = New System.Drawing.Point(12, 193)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.GroupBox2.Size = New System.Drawing.Size(111, 144)
        Me.GroupBox2.TabIndex = 10
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "No. of Risers"
        '
        'btnTwoRisers
        '
        Me.btnTwoRisers.BackColor = System.Drawing.SystemColors.Control
        Me.btnTwoRisers.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnTwoRisers.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTwoRisers.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnTwoRisers.Location = New System.Drawing.Point(24, 40)
        Me.btnTwoRisers.Name = "btnTwoRisers"
        Me.btnTwoRisers.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnTwoRisers.Size = New System.Drawing.Size(65, 17)
        Me.btnTwoRisers.TabIndex = 3
        Me.btnTwoRisers.Text = "&Two"
        Me.btnTwoRisers.UseVisualStyleBackColor = False
        '
        'btnOneRiser
        '
        Me.btnOneRiser.BackColor = System.Drawing.SystemColors.Control
        Me.btnOneRiser.Checked = True
        Me.btnOneRiser.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnOneRiser.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOneRiser.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnOneRiser.Location = New System.Drawing.Point(24, 20)
        Me.btnOneRiser.Name = "btnOneRiser"
        Me.btnOneRiser.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnOneRiser.Size = New System.Drawing.Size(65, 17)
        Me.btnOneRiser.TabIndex = 7
        Me.btnOneRiser.TabStop = True
        Me.btnOneRiser.Text = "&One"
        Me.btnOneRiser.UseVisualStyleBackColor = False
        '
        'TwoRisers
        '
        Me.TwoRisers.BackColor = System.Drawing.SystemColors.Control
        Me.TwoRisers.Controls.Add(Me.lblDistance2)
        Me.TwoRisers.Controls.Add(Me.lblDistance1)
        Me.TwoRisers.Controls.Add(Me.lblDistance4)
        Me.TwoRisers.Controls.Add(Me.chkbtn2)
        Me.TwoRisers.Controls.Add(Me.lblDistance3)
        Me.TwoRisers.Controls.Add(Me.txtRiserRZ)
        Me.TwoRisers.Controls.Add(Me.Label7)
        Me.TwoRisers.Controls.Add(Me.txtRiserFY)
        Me.TwoRisers.Controls.Add(Me.Label6)
        Me.TwoRisers.Controls.Add(Me.txtRiserRY)
        Me.TwoRisers.Controls.Add(Me.Label1)
        Me.TwoRisers.Controls.Add(Me.Label2)
        Me.TwoRisers.Controls.Add(Me.txtRiserFZ)
        Me.TwoRisers.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TwoRisers.ForeColor = System.Drawing.SystemColors.ControlText
        Me.TwoRisers.Location = New System.Drawing.Point(129, 193)
        Me.TwoRisers.Name = "TwoRisers"
        Me.TwoRisers.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TwoRisers.Size = New System.Drawing.Size(242, 144)
        Me.TwoRisers.TabIndex = 11
        Me.TwoRisers.TabStop = False
        Me.TwoRisers.Text = "Risers Location"
        '
        'chkbtn2
        '
        Me.chkbtn2.BackColor = System.Drawing.SystemColors.Control
        Me.chkbtn2.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkbtn2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkbtn2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkbtn2.Location = New System.Drawing.Point(9, 21)
        Me.chkbtn2.Name = "chkbtn2"
        Me.chkbtn2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkbtn2.Size = New System.Drawing.Size(194, 20)
        Me.chkbtn2.TabIndex = 27
        Me.chkbtn2.Text = "&Two Risers are Identical"
        Me.chkbtn2.UseVisualStyleBackColor = False
        '
        'lblDistance3
        '
        Me.lblDistance3.BackColor = System.Drawing.SystemColors.Control
        Me.lblDistance3.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDistance3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDistance3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDistance3.Location = New System.Drawing.Point(209, 81)
        Me.lblDistance3.Name = "lblDistance3"
        Me.lblDistance3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDistance3.Size = New System.Drawing.Size(18, 20)
        Me.lblDistance3.TabIndex = 26
        Me.lblDistance3.Text = "m"
        '
        'txtRiserRZ
        '
        Me.txtRiserRZ.AcceptsReturn = True
        Me.txtRiserRZ.BackColor = System.Drawing.SystemColors.Window
        Me.txtRiserRZ.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRiserRZ.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRiserRZ.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRiserRZ.Location = New System.Drawing.Point(136, 118)
        Me.txtRiserRZ.MaxLength = 0
        Me.txtRiserRZ.Name = "txtRiserRZ"
        Me.txtRiserRZ.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRiserRZ.Size = New System.Drawing.Size(67, 20)
        Me.txtRiserRZ.TabIndex = 12
        Me.txtRiserRZ.Text = "Text4"
        Me.txtRiserRZ.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(122, 57)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(89, 21)
        Me.Label7.TabIndex = 25
        Me.Label7.Text = "Riser Rear"
        '
        'txtRiserFY
        '
        Me.txtRiserFY.AcceptsReturn = True
        Me.txtRiserFY.BackColor = System.Drawing.SystemColors.Window
        Me.txtRiserFY.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRiserFY.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRiserFY.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRiserFY.Location = New System.Drawing.Point(24, 81)
        Me.txtRiserFY.MaxLength = 0
        Me.txtRiserFY.Name = "txtRiserFY"
        Me.txtRiserFY.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRiserFY.Size = New System.Drawing.Size(67, 20)
        Me.txtRiserFY.TabIndex = 9
        Me.txtRiserFY.Text = "Text3"
        Me.txtRiserFY.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(21, 57)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(70, 21)
        Me.Label6.TabIndex = 24
        Me.Label6.Text = "Riser Front"
        '
        'txtRiserRY
        '
        Me.txtRiserRY.AcceptsReturn = True
        Me.txtRiserRY.BackColor = System.Drawing.SystemColors.Window
        Me.txtRiserRY.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRiserRY.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRiserRY.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRiserRY.Location = New System.Drawing.Point(136, 81)
        Me.txtRiserRY.MaxLength = 0
        Me.txtRiserRY.Name = "txtRiserRY"
        Me.txtRiserRY.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRiserRY.Size = New System.Drawing.Size(67, 20)
        Me.txtRiserRY.TabIndex = 11
        Me.txtRiserRY.Text = "Text3"
        Me.txtRiserRY.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(6, 81)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(22, 17)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "Y"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(6, 118)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(19, 20)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "Z"
        '
        'txtRiserFZ
        '
        Me.txtRiserFZ.AcceptsReturn = True
        Me.txtRiserFZ.BackColor = System.Drawing.SystemColors.Window
        Me.txtRiserFZ.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRiserFZ.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRiserFZ.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRiserFZ.Location = New System.Drawing.Point(24, 118)
        Me.txtRiserFZ.MaxLength = 0
        Me.txtRiserFZ.Name = "txtRiserFZ"
        Me.txtRiserFZ.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRiserFZ.Size = New System.Drawing.Size(67, 20)
        Me.txtRiserFZ.TabIndex = 10
        Me.txtRiserFZ.Text = "Text4"
        Me.txtRiserFZ.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblDistance4
        '
        Me.lblDistance4.BackColor = System.Drawing.SystemColors.Control
        Me.lblDistance4.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDistance4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDistance4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDistance4.Location = New System.Drawing.Point(209, 118)
        Me.lblDistance4.Name = "lblDistance4"
        Me.lblDistance4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDistance4.Size = New System.Drawing.Size(18, 20)
        Me.lblDistance4.TabIndex = 28
        Me.lblDistance4.Text = "m"
        '
        'lblDistance1
        '
        Me.lblDistance1.BackColor = System.Drawing.SystemColors.Control
        Me.lblDistance1.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDistance1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDistance1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDistance1.Location = New System.Drawing.Point(101, 81)
        Me.lblDistance1.Name = "lblDistance1"
        Me.lblDistance1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDistance1.Size = New System.Drawing.Size(18, 20)
        Me.lblDistance1.TabIndex = 29
        Me.lblDistance1.Text = "m"
        '
        'lblDistance2
        '
        Me.lblDistance2.BackColor = System.Drawing.SystemColors.Control
        Me.lblDistance2.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDistance2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDistance2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDistance2.Location = New System.Drawing.Point(101, 121)
        Me.lblDistance2.Name = "lblDistance2"
        Me.lblDistance2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDistance2.Size = New System.Drawing.Size(18, 20)
        Me.lblDistance2.TabIndex = 30
        Me.lblDistance2.Text = "m"
        '
        'frmGlobalParms
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(384, 405)
        Me.Controls.Add(Me.TwoRisers)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.fraParms)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.fraUnits)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.Name = "frmGlobalParms"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Global Parameters"
        Me.fraParms.ResumeLayout(False)
        Me.fraParms.PerformLayout()
        Me.fraUnits.ResumeLayout(False)
        CType(Me.cmdBrowse, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optModes, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtModesFile, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.TwoRisers.ResumeLayout(False)
        Me.TwoRisers.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Public WithEvents mDamping As System.Windows.Forms.RadioButton
    Public WithEvents cDamping As System.Windows.Forms.RadioButton
    Public WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Public WithEvents btnTwoRisers As System.Windows.Forms.RadioButton
    Public WithEvents btnOneRiser As System.Windows.Forms.RadioButton
    Public WithEvents TwoRisers As System.Windows.Forms.GroupBox
    Public WithEvents txtRiserFY As System.Windows.Forms.TextBox
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents txtRiserFZ As System.Windows.Forms.TextBox
    Public WithEvents txtRiserRY As System.Windows.Forms.TextBox
    Public WithEvents txtRiserRZ As System.Windows.Forms.TextBox
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents lblDistance3 As System.Windows.Forms.Label
    Public WithEvents chkbtn2 As System.Windows.Forms.CheckBox
    Public WithEvents lblDistance4 As System.Windows.Forms.Label
    Public WithEvents lblDistance2 As System.Windows.Forms.Label
    Public WithEvents lblDistance1 As System.Windows.Forms.Label
#End Region
End Class