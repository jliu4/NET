<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmTransient
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
	Public WithEvents _txtRiser_2 As System.Windows.Forms.TextBox
	Public WithEvents _txtRiser_1 As System.Windows.Forms.TextBox
	Public WithEvents _txtRiser_0 As System.Windows.Forms.TextBox
	Public WithEvents _lblDiaUnit_0 As System.Windows.Forms.Label
	Public WithEvents _lblRiser_2 As System.Windows.Forms.Label
	Public WithEvents _lblMassUnit_0 As System.Windows.Forms.Label
	Public WithEvents _lblRiser_1 As System.Windows.Forms.Label
	Public WithEvents _lblForceUnit_0 As System.Windows.Forms.Label
	Public WithEvents _lblRiser_0 As System.Windows.Forms.Label
	Public WithEvents fraRiser As System.Windows.Forms.GroupBox
	Public WithEvents _txtVslSt_3 As System.Windows.Forms.TextBox
	Public WithEvents _txtVslSt_2 As System.Windows.Forms.TextBox
	Public WithEvents _txtVslSt_0 As System.Windows.Forms.TextBox
	Public WithEvents _txtVslSt_1 As System.Windows.Forms.TextBox
    Public WithEvents txtWell As System.Windows.Forms.TextBox
	Public WithEvents _lblVslSt_3 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_6 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_5 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_2 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_0 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_1 As System.Windows.Forms.Label
	Public WithEvents lblWell As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_4 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_3 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_2 As System.Windows.Forms.Label
	Public WithEvents fraWell As System.Windows.Forms.GroupBox
	Public WithEvents btnDisplayCurves As System.Windows.Forms.Button
	Public WithEvents btnReport As System.Windows.Forms.Button
	Public WithEvents txtEnvironment As System.Windows.Forms.TextBox
	Public WithEvents btnEnvironment As System.Windows.Forms.Button
	Public WithEvents fraEnvironment As System.Windows.Forms.GroupBox
	Public WithEvents txtDuration As System.Windows.Forms.TextBox
	Public WithEvents txtInterval As System.Windows.Forms.TextBox
	Public WithEvents lblDuration As System.Windows.Forms.Label
	Public WithEvents lblInterval As System.Windows.Forms.Label
	Public WithEvents fraTransientTime As System.Windows.Forms.GroupBox
	Public dlgFileOpen As System.Windows.Forms.OpenFileDialog
	Public dlgFileSave As System.Windows.Forms.SaveFileDialog
	Public WithEvents txtMaxOffsetTime As System.Windows.Forms.TextBox
	Public WithEvents txtOffset As System.Windows.Forms.TextBox
	Public WithEvents txtOffsetPWD As System.Windows.Forms.TextBox
	Public WithEvents txtMaxOffset As System.Windows.Forms.TextBox
	Public WithEvents txtMaxOffsetPWD As System.Windows.Forms.TextBox
	Public WithEvents txtOffsetBearing As System.Windows.Forms.TextBox
    Public WithEvents _lblPercent_2 As System.Windows.Forms.Label
	Public WithEvents _lblVslStUnit_4 As System.Windows.Forms.Label
	Public WithEvents _lblPercent_1 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_0 As System.Windows.Forms.Label
	Public WithEvents lblOffset As System.Windows.Forms.Label
	Public WithEvents lblMaxOffset As System.Windows.Forms.Label
	Public WithEvents _lblPercent_0 As System.Windows.Forms.Label
	Public WithEvents lblBearing As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_1 As System.Windows.Forms.Label
	Public WithEvents fraTransientMotion As System.Windows.Forms.GroupBox
	Public WithEvents btnTransient As System.Windows.Forms.Button
    Public WithEvents lblDiaUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblForceUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblLengthUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblMassUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblPercent As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblRiser As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblVslSt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblVslStUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents mnuFilePre As Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents txtRiser As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    Public WithEvents txtVslSt As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    Public WithEvents mnuFileNew As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFileOpen As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuLine1 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents mnuFileSave As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFileSaveAs As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuLine2 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents _mnuFilePre_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuFilePre_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuFilePre_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuFilePre_3 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuLinePreFile As System.Windows.Forms.ToolStripSeparator
    Public WithEvents mnuFileExit As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFile As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuOption As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTransient))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnDisplayCurves = New System.Windows.Forms.Button()
        Me.fraRiser = New System.Windows.Forms.GroupBox()
        Me._txtRiser_2 = New System.Windows.Forms.TextBox()
        Me._txtRiser_1 = New System.Windows.Forms.TextBox()
        Me._txtRiser_0 = New System.Windows.Forms.TextBox()
        Me._lblDiaUnit_0 = New System.Windows.Forms.Label()
        Me._lblRiser_2 = New System.Windows.Forms.Label()
        Me._lblMassUnit_0 = New System.Windows.Forms.Label()
        Me._lblRiser_1 = New System.Windows.Forms.Label()
        Me._lblForceUnit_0 = New System.Windows.Forms.Label()
        Me._lblRiser_0 = New System.Windows.Forms.Label()
        Me.fraWell = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me._txtVslSt_3 = New System.Windows.Forms.TextBox()
        Me._txtVslSt_2 = New System.Windows.Forms.TextBox()
        Me._txtVslSt_0 = New System.Windows.Forms.TextBox()
        Me._txtVslSt_1 = New System.Windows.Forms.TextBox()
        Me.txtWell = New System.Windows.Forms.TextBox()
        Me._lblVslSt_3 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_6 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_5 = New System.Windows.Forms.Label()
        Me._lblVslSt_2 = New System.Windows.Forms.Label()
        Me._lblVslSt_0 = New System.Windows.Forms.Label()
        Me._lblVslSt_1 = New System.Windows.Forms.Label()
        Me.lblWell = New System.Windows.Forms.Label()
        Me._lblLengthUnit_4 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_3 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_2 = New System.Windows.Forms.Label()
        Me.btnReport = New System.Windows.Forms.Button()
        Me.fraEnvironment = New System.Windows.Forms.GroupBox()
        Me.txtEnvironment = New System.Windows.Forms.TextBox()
        Me.btnEnvironment = New System.Windows.Forms.Button()
        Me.fraTransientTime = New System.Windows.Forms.GroupBox()
        Me.txtDuration = New System.Windows.Forms.TextBox()
        Me.txtInterval = New System.Windows.Forms.TextBox()
        Me.lblDuration = New System.Windows.Forms.Label()
        Me.lblInterval = New System.Windows.Forms.Label()
        Me.fraTransientMotion = New System.Windows.Forms.GroupBox()
        Me.grdTM = New System.Windows.Forms.DataGridView()
        Me.txtMaxOffsetTime = New System.Windows.Forms.TextBox()
        Me.txtOffset = New System.Windows.Forms.TextBox()
        Me.txtOffsetPWD = New System.Windows.Forms.TextBox()
        Me.txtMaxOffset = New System.Windows.Forms.TextBox()
        Me.txtMaxOffsetPWD = New System.Windows.Forms.TextBox()
        Me.txtOffsetBearing = New System.Windows.Forms.TextBox()
        Me._lblPercent_2 = New System.Windows.Forms.Label()
        Me._lblVslStUnit_4 = New System.Windows.Forms.Label()
        Me._lblPercent_1 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_0 = New System.Windows.Forms.Label()
        Me.lblOffset = New System.Windows.Forms.Label()
        Me.lblMaxOffset = New System.Windows.Forms.Label()
        Me._lblPercent_0 = New System.Windows.Forms.Label()
        Me.lblBearing = New System.Windows.Forms.Label()
        Me._lblLengthUnit_1 = New System.Windows.Forms.Label()
        Me.dlgFileOpen = New System.Windows.Forms.OpenFileDialog()
        Me.dlgFileSave = New System.Windows.Forms.SaveFileDialog()
        Me.btnTransient = New System.Windows.Forms.Button()
        Me.picPolar = New System.Windows.Forms.PictureBox()
        Me.lblDiaUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblForceUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblLengthUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblMassUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblPercent = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblRiser = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblVslSt = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblVslStUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.mnuFilePre = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me._mnuFilePre_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFilePre_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFilePre_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFilePre_3 = New System.Windows.Forms.ToolStripMenuItem()
        Me.txtRiser = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.txtVslSt = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip()
        Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFileNew = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFileOpen = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLine1 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuFileSave = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFileSaveAs = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLine2 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuLinePreFile = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuFileExit = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuOption = New System.Windows.Forms.ToolStripMenuItem()
        Me.fraRiser.SuspendLayout()
        Me.fraWell.SuspendLayout()
        Me.fraEnvironment.SuspendLayout()
        Me.fraTransientTime.SuspendLayout()
        Me.fraTransientMotion.SuspendLayout()
        CType(Me.grdTM, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picPolar, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblDiaUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblForceUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblLengthUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblMassUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblPercent, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblRiser, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVslSt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVslStUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuFilePre, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtRiser, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtVslSt, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MainMenu1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnDisplayCurves
        '
        Me.btnDisplayCurves.BackColor = System.Drawing.SystemColors.Control
        Me.btnDisplayCurves.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnDisplayCurves.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDisplayCurves.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnDisplayCurves.Location = New System.Drawing.Point(221, 245)
        Me.btnDisplayCurves.Name = "btnDisplayCurves"
        Me.btnDisplayCurves.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnDisplayCurves.Size = New System.Drawing.Size(143, 25)
        Me.btnDisplayCurves.TabIndex = 29
        Me.btnDisplayCurves.Text = "Time History Plots..."
        Me.ToolTip1.SetToolTip(Me.btnDisplayCurves, "View Transient Line Tension Time History")
        Me.btnDisplayCurves.UseVisualStyleBackColor = False
        Me.btnDisplayCurves.Visible = False
        '
        'fraRiser
        '
        Me.fraRiser.BackColor = System.Drawing.SystemColors.Control
        Me.fraRiser.Controls.Add(Me._txtRiser_2)
        Me.fraRiser.Controls.Add(Me._txtRiser_1)
        Me.fraRiser.Controls.Add(Me._txtRiser_0)
        Me.fraRiser.Controls.Add(Me._lblDiaUnit_0)
        Me.fraRiser.Controls.Add(Me._lblRiser_2)
        Me.fraRiser.Controls.Add(Me._lblMassUnit_0)
        Me.fraRiser.Controls.Add(Me._lblRiser_1)
        Me.fraRiser.Controls.Add(Me._lblForceUnit_0)
        Me.fraRiser.Controls.Add(Me._lblRiser_0)
        Me.fraRiser.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraRiser.Location = New System.Drawing.Point(202, 178)
        Me.fraRiser.Name = "fraRiser"
        Me.fraRiser.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraRiser.Size = New System.Drawing.Size(182, 97)
        Me.fraRiser.TabIndex = 46
        Me.fraRiser.TabStop = False
        Me.fraRiser.Text = "Riser"
        '
        '_txtRiser_2
        '
        Me._txtRiser_2.AcceptsReturn = True
        Me._txtRiser_2.BackColor = System.Drawing.SystemColors.Window
        Me._txtRiser_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtRiser_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRiser.SetIndex(Me._txtRiser_2, CType(2, Short))
        Me._txtRiser_2.Location = New System.Drawing.Point(89, 73)
        Me._txtRiser_2.MaxLength = 0
        Me._txtRiser_2.Name = "_txtRiser_2"
        Me._txtRiser_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtRiser_2.Size = New System.Drawing.Size(40, 20)
        Me._txtRiser_2.TabIndex = 53
        Me._txtRiser_2.TabStop = False
        Me._txtRiser_2.Text = "0"
        Me._txtRiser_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtRiser_1
        '
        Me._txtRiser_1.AcceptsReturn = True
        Me._txtRiser_1.BackColor = System.Drawing.SystemColors.Window
        Me._txtRiser_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtRiser_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRiser.SetIndex(Me._txtRiser_1, CType(1, Short))
        Me._txtRiser_1.Location = New System.Drawing.Point(89, 46)
        Me._txtRiser_1.MaxLength = 0
        Me._txtRiser_1.Name = "_txtRiser_1"
        Me._txtRiser_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtRiser_1.Size = New System.Drawing.Size(40, 20)
        Me._txtRiser_1.TabIndex = 50
        Me._txtRiser_1.TabStop = False
        Me._txtRiser_1.Text = "0"
        Me._txtRiser_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtRiser_0
        '
        Me._txtRiser_0.AcceptsReturn = True
        Me._txtRiser_0.BackColor = System.Drawing.SystemColors.Window
        Me._txtRiser_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtRiser_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRiser.SetIndex(Me._txtRiser_0, CType(0, Short))
        Me._txtRiser_0.Location = New System.Drawing.Point(89, 24)
        Me._txtRiser_0.MaxLength = 0
        Me._txtRiser_0.Name = "_txtRiser_0"
        Me._txtRiser_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtRiser_0.Size = New System.Drawing.Size(40, 20)
        Me._txtRiser_0.TabIndex = 47
        Me._txtRiser_0.TabStop = False
        Me._txtRiser_0.Text = "0"
        Me._txtRiser_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_lblDiaUnit_0
        '
        Me._lblDiaUnit_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblDiaUnit_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDiaUnit_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDiaUnit.SetIndex(Me._lblDiaUnit_0, CType(0, Short))
        Me._lblDiaUnit_0.Location = New System.Drawing.Point(144, 72)
        Me._lblDiaUnit_0.Name = "_lblDiaUnit_0"
        Me._lblDiaUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDiaUnit_0.Size = New System.Drawing.Size(28, 17)
        Me._lblDiaUnit_0.TabIndex = 55
        Me._lblDiaUnit_0.Text = "in"
        '
        '_lblRiser_2
        '
        Me._lblRiser_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblRiser_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRiser_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRiser.SetIndex(Me._lblRiser_2, CType(2, Short))
        Me._lblRiser_2.Location = New System.Drawing.Point(8, 73)
        Me._lblRiser_2.Name = "_lblRiser_2"
        Me._lblRiser_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRiser_2.Size = New System.Drawing.Size(75, 16)
        Me._lblRiser_2.TabIndex = 54
        Me._lblRiser_2.Text = "BM Diameter"
        '
        '_lblMassUnit_0
        '
        Me._lblMassUnit_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblMassUnit_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblMassUnit_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblMassUnit.SetIndex(Me._lblMassUnit_0, CType(0, Short))
        Me._lblMassUnit_0.Location = New System.Drawing.Point(144, 46)
        Me._lblMassUnit_0.Name = "_lblMassUnit_0"
        Me._lblMassUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblMassUnit_0.Size = New System.Drawing.Size(28, 17)
        Me._lblMassUnit_0.TabIndex = 52
        Me._lblMassUnit_0.Text = "kips"
        '
        '_lblRiser_1
        '
        Me._lblRiser_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblRiser_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRiser_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRiser.SetIndex(Me._lblRiser_1, CType(1, Short))
        Me._lblRiser_1.Location = New System.Drawing.Point(8, 49)
        Me._lblRiser_1.Name = "_lblRiser_1"
        Me._lblRiser_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRiser_1.Size = New System.Drawing.Size(61, 17)
        Me._lblRiser_1.TabIndex = 51
        Me._lblRiser_1.Text = "Total Mass"
        '
        '_lblForceUnit_0
        '
        Me._lblForceUnit_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblForceUnit_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblForceUnit_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblForceUnit.SetIndex(Me._lblForceUnit_0, CType(0, Short))
        Me._lblForceUnit_0.Location = New System.Drawing.Point(144, 24)
        Me._lblForceUnit_0.Name = "_lblForceUnit_0"
        Me._lblForceUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblForceUnit_0.Size = New System.Drawing.Size(28, 17)
        Me._lblForceUnit_0.TabIndex = 49
        Me._lblForceUnit_0.Text = "kips"
        '
        '_lblRiser_0
        '
        Me._lblRiser_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblRiser_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRiser_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRiser.SetIndex(Me._lblRiser_0, CType(0, Short))
        Me._lblRiser_0.Location = New System.Drawing.Point(8, 25)
        Me._lblRiser_0.Name = "_lblRiser_0"
        Me._lblRiser_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRiser_0.Size = New System.Drawing.Size(71, 19)
        Me._lblRiser_0.TabIndex = 48
        Me._lblRiser_0.Text = "Top Tension"
        '
        'fraWell
        '
        Me.fraWell.BackColor = System.Drawing.SystemColors.Control
        Me.fraWell.Controls.Add(Me.Label1)
        Me.fraWell.Controls.Add(Me._txtVslSt_3)
        Me.fraWell.Controls.Add(Me._txtVslSt_2)
        Me.fraWell.Controls.Add(Me._txtVslSt_0)
        Me.fraWell.Controls.Add(Me._txtVslSt_1)
        Me.fraWell.Controls.Add(Me.txtWell)
        Me.fraWell.Controls.Add(Me._lblVslSt_3)
        Me.fraWell.Controls.Add(Me._lblLengthUnit_6)
        Me.fraWell.Controls.Add(Me._lblLengthUnit_5)
        Me.fraWell.Controls.Add(Me._lblVslSt_2)
        Me.fraWell.Controls.Add(Me._lblVslSt_0)
        Me.fraWell.Controls.Add(Me._lblVslSt_1)
        Me.fraWell.Controls.Add(Me.lblWell)
        Me.fraWell.Controls.Add(Me._lblLengthUnit_4)
        Me.fraWell.Controls.Add(Me._lblLengthUnit_3)
        Me.fraWell.Controls.Add(Me._lblLengthUnit_2)
        Me.fraWell.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraWell.Location = New System.Drawing.Point(7, 93)
        Me.fraWell.Name = "fraWell"
        Me.fraWell.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraWell.Size = New System.Drawing.Size(187, 182)
        Me.fraWell.TabIndex = 30
        Me.fraWell.TabStop = False
        Me.fraWell.Text = "Vessel"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label1.Location = New System.Drawing.Point(11, 140)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(158, 29)
        Me.Label1.TabIndex = 46
        Me.Label1.Text = "* Heading is clockwise from True North"
        '
        '_txtVslSt_3
        '
        Me._txtVslSt_3.AcceptsReturn = True
        Me._txtVslSt_3.BackColor = System.Drawing.SystemColors.Window
        Me._txtVslSt_3.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtVslSt_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtVslSt.SetIndex(Me._txtVslSt_3, CType(3, Short))
        Me._txtVslSt_3.Location = New System.Drawing.Point(68, 88)
        Me._txtVslSt_3.MaxLength = 0
        Me._txtVslSt_3.Name = "_txtVslSt_3"
        Me._txtVslSt_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtVslSt_3.Size = New System.Drawing.Size(70, 20)
        Me._txtVslSt_3.TabIndex = 43
        Me._txtVslSt_3.TabStop = False
        Me._txtVslSt_3.Text = "0"
        Me._txtVslSt_3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtVslSt_2
        '
        Me._txtVslSt_2.AcceptsReturn = True
        Me._txtVslSt_2.BackColor = System.Drawing.SystemColors.Window
        Me._txtVslSt_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtVslSt_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtVslSt.SetIndex(Me._txtVslSt_2, CType(2, Short))
        Me._txtVslSt_2.Location = New System.Drawing.Point(68, 64)
        Me._txtVslSt_2.MaxLength = 0
        Me._txtVslSt_2.Name = "_txtVslSt_2"
        Me._txtVslSt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtVslSt_2.Size = New System.Drawing.Size(71, 20)
        Me._txtVslSt_2.TabIndex = 40
        Me._txtVslSt_2.TabStop = False
        Me._txtVslSt_2.Text = "0"
        Me._txtVslSt_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtVslSt_0
        '
        Me._txtVslSt_0.AcceptsReturn = True
        Me._txtVslSt_0.BackColor = System.Drawing.SystemColors.Window
        Me._txtVslSt_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtVslSt_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtVslSt.SetIndex(Me._txtVslSt_0, CType(0, Short))
        Me._txtVslSt_0.Location = New System.Drawing.Point(39, 15)
        Me._txtVslSt_0.MaxLength = 0
        Me._txtVslSt_0.Name = "_txtVslSt_0"
        Me._txtVslSt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtVslSt_0.Size = New System.Drawing.Size(100, 20)
        Me._txtVslSt_0.TabIndex = 33
        Me._txtVslSt_0.TabStop = False
        Me._txtVslSt_0.Text = "0"
        Me._txtVslSt_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtVslSt_1
        '
        Me._txtVslSt_1.AcceptsReturn = True
        Me._txtVslSt_1.BackColor = System.Drawing.SystemColors.Window
        Me._txtVslSt_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtVslSt_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtVslSt.SetIndex(Me._txtVslSt_1, CType(1, Short))
        Me._txtVslSt_1.Location = New System.Drawing.Point(39, 39)
        Me._txtVslSt_1.MaxLength = 0
        Me._txtVslSt_1.Name = "_txtVslSt_1"
        Me._txtVslSt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtVslSt_1.Size = New System.Drawing.Size(100, 20)
        Me._txtVslSt_1.TabIndex = 32
        Me._txtVslSt_1.TabStop = False
        Me._txtVslSt_1.Text = "0"
        Me._txtVslSt_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtWell
        '
        Me.txtWell.AcceptsReturn = True
        Me.txtWell.BackColor = System.Drawing.SystemColors.Window
        Me.txtWell.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtWell.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtWell.Location = New System.Drawing.Point(76, 111)
        Me.txtWell.MaxLength = 0
        Me.txtWell.Name = "txtWell"
        Me.txtWell.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtWell.Size = New System.Drawing.Size(62, 20)
        Me.txtWell.TabIndex = 31
        Me.txtWell.TabStop = False
        Me.txtWell.Text = "0"
        Me.txtWell.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_lblVslSt_3
        '
        Me._lblVslSt_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblVslSt_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVslSt_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVslSt.SetIndex(Me._lblVslSt_3, CType(3, Short))
        Me._lblVslSt_3.Location = New System.Drawing.Point(8, 89)
        Me._lblVslSt_3.Name = "_lblVslSt_3"
        Me._lblVslSt_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVslSt_3.Size = New System.Drawing.Size(41, 17)
        Me._lblVslSt_3.TabIndex = 45
        Me._lblVslSt_3.Text = "Draft"
        '
        '_lblLengthUnit_6
        '
        Me._lblLengthUnit_6.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_6, CType(6, Short))
        Me._lblLengthUnit_6.Location = New System.Drawing.Point(144, 89)
        Me._lblLengthUnit_6.Name = "_lblLengthUnit_6"
        Me._lblLengthUnit_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_6.Size = New System.Drawing.Size(25, 17)
        Me._lblLengthUnit_6.TabIndex = 44
        Me._lblLengthUnit_6.Text = "ft"
        '
        '_lblLengthUnit_5
        '
        Me._lblLengthUnit_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_5, CType(5, Short))
        Me._lblLengthUnit_5.Location = New System.Drawing.Point(144, 64)
        Me._lblLengthUnit_5.Name = "_lblLengthUnit_5"
        Me._lblLengthUnit_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_5.Size = New System.Drawing.Size(25, 17)
        Me._lblLengthUnit_5.TabIndex = 42
        Me._lblLengthUnit_5.Text = "deg"
        '
        '_lblVslSt_2
        '
        Me._lblVslSt_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblVslSt_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVslSt_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVslSt.SetIndex(Me._lblVslSt_2, CType(2, Short))
        Me._lblVslSt_2.Location = New System.Drawing.Point(8, 68)
        Me._lblVslSt_2.Name = "_lblVslSt_2"
        Me._lblVslSt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVslSt_2.Size = New System.Drawing.Size(65, 17)
        Me._lblVslSt_2.TabIndex = 41
        Me._lblVslSt_2.Text = "Heading"
        '
        '_lblVslSt_0
        '
        Me._lblVslSt_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblVslSt_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVslSt_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVslSt.SetIndex(Me._lblVslSt_0, CType(0, Short))
        Me._lblVslSt_0.Location = New System.Drawing.Point(13, 17)
        Me._lblVslSt_0.Name = "_lblVslSt_0"
        Me._lblVslSt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVslSt_0.Size = New System.Drawing.Size(33, 17)
        Me._lblVslSt_0.TabIndex = 39
        Me._lblVslSt_0.Text = "x (E)"
        '
        '_lblVslSt_1
        '
        Me._lblVslSt_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblVslSt_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVslSt_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVslSt.SetIndex(Me._lblVslSt_1, CType(1, Short))
        Me._lblVslSt_1.Location = New System.Drawing.Point(11, 39)
        Me._lblVslSt_1.Name = "_lblVslSt_1"
        Me._lblVslSt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVslSt_1.Size = New System.Drawing.Size(33, 17)
        Me._lblVslSt_1.TabIndex = 38
        Me._lblVslSt_1.Text = "y (N)"
        '
        'lblWell
        '
        Me.lblWell.BackColor = System.Drawing.SystemColors.Control
        Me.lblWell.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblWell.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblWell.Location = New System.Drawing.Point(8, 112)
        Me.lblWell.Name = "lblWell"
        Me.lblWell.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblWell.Size = New System.Drawing.Size(85, 17)
        Me.lblWell.TabIndex = 37
        Me.lblWell.Text = "Water Depth"
        '
        '_lblLengthUnit_4
        '
        Me._lblLengthUnit_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_4, CType(4, Short))
        Me._lblLengthUnit_4.Location = New System.Drawing.Point(149, 16)
        Me._lblLengthUnit_4.Name = "_lblLengthUnit_4"
        Me._lblLengthUnit_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_4.Size = New System.Drawing.Size(20, 19)
        Me._lblLengthUnit_4.TabIndex = 36
        Me._lblLengthUnit_4.Text = "ft"
        '
        '_lblLengthUnit_3
        '
        Me._lblLengthUnit_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_3, CType(3, Short))
        Me._lblLengthUnit_3.Location = New System.Drawing.Point(149, 39)
        Me._lblLengthUnit_3.Name = "_lblLengthUnit_3"
        Me._lblLengthUnit_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_3.Size = New System.Drawing.Size(20, 16)
        Me._lblLengthUnit_3.TabIndex = 35
        Me._lblLengthUnit_3.Text = "ft"
        '
        '_lblLengthUnit_2
        '
        Me._lblLengthUnit_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_2, CType(2, Short))
        Me._lblLengthUnit_2.Location = New System.Drawing.Point(144, 113)
        Me._lblLengthUnit_2.Name = "_lblLengthUnit_2"
        Me._lblLengthUnit_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_2.Size = New System.Drawing.Size(15, 16)
        Me._lblLengthUnit_2.TabIndex = 34
        Me._lblLengthUnit_2.Text = "ft"
        '
        'btnReport
        '
        Me.btnReport.BackColor = System.Drawing.SystemColors.Control
        Me.btnReport.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnReport.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnReport.Location = New System.Drawing.Point(300, 62)
        Me.btnReport.Name = "btnReport"
        Me.btnReport.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnReport.Size = New System.Drawing.Size(84, 25)
        Me.btnReport.TabIndex = 28
        Me.btnReport.Text = "&Report"
        Me.btnReport.UseVisualStyleBackColor = False
        '
        'fraEnvironment
        '
        Me.fraEnvironment.BackColor = System.Drawing.SystemColors.Control
        Me.fraEnvironment.Controls.Add(Me.txtEnvironment)
        Me.fraEnvironment.Controls.Add(Me.btnEnvironment)
        Me.fraEnvironment.ForeColor = System.Drawing.Color.Black
        Me.fraEnvironment.Location = New System.Drawing.Point(7, 25)
        Me.fraEnvironment.Name = "fraEnvironment"
        Me.fraEnvironment.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraEnvironment.Size = New System.Drawing.Size(278, 62)
        Me.fraEnvironment.TabIndex = 25
        Me.fraEnvironment.TabStop = False
        Me.fraEnvironment.Text = "Environment Condition"
        '
        'txtEnvironment
        '
        Me.txtEnvironment.AcceptsReturn = True
        Me.txtEnvironment.BackColor = System.Drawing.SystemColors.Control
        Me.txtEnvironment.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtEnvironment.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtEnvironment.Location = New System.Drawing.Point(16, 25)
        Me.txtEnvironment.MaxLength = 0
        Me.txtEnvironment.Name = "txtEnvironment"
        Me.txtEnvironment.ReadOnly = True
        Me.txtEnvironment.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtEnvironment.Size = New System.Drawing.Size(194, 20)
        Me.txtEnvironment.TabIndex = 27
        Me.txtEnvironment.Text = "Current Environment"
        '
        'btnEnvironment
        '
        Me.btnEnvironment.BackColor = System.Drawing.SystemColors.Control
        Me.btnEnvironment.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnEnvironment.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnEnvironment.Location = New System.Drawing.Point(216, 22)
        Me.btnEnvironment.Name = "btnEnvironment"
        Me.btnEnvironment.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnEnvironment.Size = New System.Drawing.Size(48, 25)
        Me.btnEnvironment.TabIndex = 26
        Me.btnEnvironment.Text = "Edit..."
        Me.btnEnvironment.UseVisualStyleBackColor = False
        '
        'fraTransientTime
        '
        Me.fraTransientTime.BackColor = System.Drawing.SystemColors.Control
        Me.fraTransientTime.Controls.Add(Me.txtDuration)
        Me.fraTransientTime.Controls.Add(Me.txtInterval)
        Me.fraTransientTime.Controls.Add(Me.lblDuration)
        Me.fraTransientTime.Controls.Add(Me.lblInterval)
        Me.fraTransientTime.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraTransientTime.Location = New System.Drawing.Point(202, 93)
        Me.fraTransientTime.Name = "fraTransientTime"
        Me.fraTransientTime.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraTransientTime.Size = New System.Drawing.Size(182, 79)
        Me.fraTransientTime.TabIndex = 4
        Me.fraTransientTime.TabStop = False
        Me.fraTransientTime.Text = "Time (sec)"
        '
        'txtDuration
        '
        Me.txtDuration.AcceptsReturn = True
        Me.txtDuration.BackColor = System.Drawing.SystemColors.Window
        Me.txtDuration.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDuration.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDuration.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDuration.Location = New System.Drawing.Point(69, 24)
        Me.txtDuration.MaxLength = 0
        Me.txtDuration.Name = "txtDuration"
        Me.txtDuration.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDuration.Size = New System.Drawing.Size(73, 20)
        Me.txtDuration.TabIndex = 6
        Me.txtDuration.Text = "200"
        Me.txtDuration.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtInterval
        '
        Me.txtInterval.AcceptsReturn = True
        Me.txtInterval.BackColor = System.Drawing.SystemColors.Window
        Me.txtInterval.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtInterval.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtInterval.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtInterval.Location = New System.Drawing.Point(69, 47)
        Me.txtInterval.MaxLength = 0
        Me.txtInterval.Name = "txtInterval"
        Me.txtInterval.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtInterval.Size = New System.Drawing.Size(73, 20)
        Me.txtInterval.TabIndex = 5
        Me.txtInterval.Text = "200"
        Me.txtInterval.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblDuration
        '
        Me.lblDuration.BackColor = System.Drawing.SystemColors.Control
        Me.lblDuration.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDuration.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDuration.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDuration.Location = New System.Drawing.Point(23, 25)
        Me.lblDuration.Name = "lblDuration"
        Me.lblDuration.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDuration.Size = New System.Drawing.Size(52, 15)
        Me.lblDuration.TabIndex = 8
        Me.lblDuration.Text = "Duration"
        '
        'lblInterval
        '
        Me.lblInterval.BackColor = System.Drawing.SystemColors.Control
        Me.lblInterval.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblInterval.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInterval.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblInterval.Location = New System.Drawing.Point(23, 48)
        Me.lblInterval.Name = "lblInterval"
        Me.lblInterval.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblInterval.Size = New System.Drawing.Size(50, 15)
        Me.lblInterval.TabIndex = 7
        Me.lblInterval.Text = "Interval"
        '
        'fraTransientMotion
        '
        Me.fraTransientMotion.BackColor = System.Drawing.SystemColors.Control
        Me.fraTransientMotion.Controls.Add(Me.grdTM)
        Me.fraTransientMotion.Controls.Add(Me.txtMaxOffsetTime)
        Me.fraTransientMotion.Controls.Add(Me.btnDisplayCurves)
        Me.fraTransientMotion.Controls.Add(Me.txtOffset)
        Me.fraTransientMotion.Controls.Add(Me.txtOffsetPWD)
        Me.fraTransientMotion.Controls.Add(Me.txtMaxOffset)
        Me.fraTransientMotion.Controls.Add(Me.txtMaxOffsetPWD)
        Me.fraTransientMotion.Controls.Add(Me.txtOffsetBearing)
        Me.fraTransientMotion.Controls.Add(Me._lblPercent_2)
        Me.fraTransientMotion.Controls.Add(Me._lblVslStUnit_4)
        Me.fraTransientMotion.Controls.Add(Me._lblPercent_1)
        Me.fraTransientMotion.Controls.Add(Me._lblLengthUnit_0)
        Me.fraTransientMotion.Controls.Add(Me.lblOffset)
        Me.fraTransientMotion.Controls.Add(Me.lblMaxOffset)
        Me.fraTransientMotion.Controls.Add(Me._lblPercent_0)
        Me.fraTransientMotion.Controls.Add(Me.lblBearing)
        Me.fraTransientMotion.Controls.Add(Me._lblLengthUnit_1)
        Me.fraTransientMotion.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraTransientMotion.Location = New System.Drawing.Point(7, 280)
        Me.fraTransientMotion.Name = "fraTransientMotion"
        Me.fraTransientMotion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraTransientMotion.Size = New System.Drawing.Size(377, 280)
        Me.fraTransientMotion.TabIndex = 2
        Me.fraTransientMotion.TabStop = False
        Me.fraTransientMotion.Text = "Transient Motion"
        '
        'grdTM
        '
        Me.grdTM.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdTM.Location = New System.Drawing.Point(11, 19)
        Me.grdTM.Name = "grdTM"
        Me.grdTM.RowHeadersVisible = False
        Me.grdTM.Size = New System.Drawing.Size(353, 167)
        Me.grdTM.TabIndex = 25
        '
        'txtMaxOffsetTime
        '
        Me.txtMaxOffsetTime.AcceptsReturn = True
        Me.txtMaxOffsetTime.BackColor = System.Drawing.SystemColors.Window
        Me.txtMaxOffsetTime.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMaxOffsetTime.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMaxOffsetTime.Location = New System.Drawing.Point(243, 203)
        Me.txtMaxOffsetTime.MaxLength = 0
        Me.txtMaxOffsetTime.Name = "txtMaxOffsetTime"
        Me.txtMaxOffsetTime.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMaxOffsetTime.Size = New System.Drawing.Size(41, 20)
        Me.txtMaxOffsetTime.TabIndex = 23
        '
        'txtOffset
        '
        Me.txtOffset.AcceptsReturn = True
        Me.txtOffset.BackColor = System.Drawing.SystemColors.Window
        Me.txtOffset.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtOffset.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOffset.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtOffset.Location = New System.Drawing.Point(80, 227)
        Me.txtOffset.MaxLength = 0
        Me.txtOffset.Name = "txtOffset"
        Me.txtOffset.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtOffset.Size = New System.Drawing.Size(45, 20)
        Me.txtOffset.TabIndex = 13
        Me.txtOffset.TabStop = False
        Me.txtOffset.Text = "0"
        Me.txtOffset.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtOffsetPWD
        '
        Me.txtOffsetPWD.AcceptsReturn = True
        Me.txtOffsetPWD.BackColor = System.Drawing.SystemColors.Window
        Me.txtOffsetPWD.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtOffsetPWD.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOffsetPWD.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtOffsetPWD.Location = New System.Drawing.Point(148, 228)
        Me.txtOffsetPWD.MaxLength = 0
        Me.txtOffsetPWD.Name = "txtOffsetPWD"
        Me.txtOffsetPWD.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtOffsetPWD.Size = New System.Drawing.Size(40, 20)
        Me.txtOffsetPWD.TabIndex = 12
        Me.txtOffsetPWD.TabStop = False
        Me.txtOffsetPWD.Text = "0"
        Me.txtOffsetPWD.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtMaxOffset
        '
        Me.txtMaxOffset.AcceptsReturn = True
        Me.txtMaxOffset.BackColor = System.Drawing.SystemColors.Window
        Me.txtMaxOffset.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMaxOffset.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMaxOffset.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMaxOffset.Location = New System.Drawing.Point(80, 202)
        Me.txtMaxOffset.MaxLength = 0
        Me.txtMaxOffset.Name = "txtMaxOffset"
        Me.txtMaxOffset.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMaxOffset.Size = New System.Drawing.Size(45, 20)
        Me.txtMaxOffset.TabIndex = 11
        Me.txtMaxOffset.TabStop = False
        Me.txtMaxOffset.Text = "0"
        Me.txtMaxOffset.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtMaxOffsetPWD
        '
        Me.txtMaxOffsetPWD.AcceptsReturn = True
        Me.txtMaxOffsetPWD.BackColor = System.Drawing.SystemColors.Window
        Me.txtMaxOffsetPWD.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMaxOffsetPWD.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMaxOffsetPWD.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMaxOffsetPWD.Location = New System.Drawing.Point(148, 202)
        Me.txtMaxOffsetPWD.MaxLength = 0
        Me.txtMaxOffsetPWD.Name = "txtMaxOffsetPWD"
        Me.txtMaxOffsetPWD.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMaxOffsetPWD.Size = New System.Drawing.Size(40, 20)
        Me.txtMaxOffsetPWD.TabIndex = 10
        Me.txtMaxOffsetPWD.TabStop = False
        Me.txtMaxOffsetPWD.Text = "0"
        Me.txtMaxOffsetPWD.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtOffsetBearing
        '
        Me.txtOffsetBearing.AcceptsReturn = True
        Me.txtOffsetBearing.BackColor = System.Drawing.SystemColors.Window
        Me.txtOffsetBearing.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtOffsetBearing.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOffsetBearing.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtOffsetBearing.Location = New System.Drawing.Point(80, 252)
        Me.txtOffsetBearing.MaxLength = 0
        Me.txtOffsetBearing.Name = "txtOffsetBearing"
        Me.txtOffsetBearing.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtOffsetBearing.Size = New System.Drawing.Size(45, 20)
        Me.txtOffsetBearing.TabIndex = 9
        Me.txtOffsetBearing.TabStop = False
        Me.txtOffsetBearing.Text = "0"
        Me.txtOffsetBearing.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_lblPercent_2
        '
        Me._lblPercent_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblPercent_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblPercent_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblPercent_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPercent.SetIndex(Me._lblPercent_2, CType(2, Short))
        Me._lblPercent_2.Location = New System.Drawing.Point(287, 207)
        Me._lblPercent_2.Name = "_lblPercent_2"
        Me._lblPercent_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblPercent_2.Size = New System.Drawing.Size(28, 16)
        Me._lblPercent_2.TabIndex = 24
        Me._lblPercent_2.Text = "sec"
        '
        '_lblVslStUnit_4
        '
        Me._lblVslStUnit_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblVslStUnit_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVslStUnit_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVslStUnit.SetIndex(Me._lblVslStUnit_4, CType(4, Short))
        Me._lblVslStUnit_4.Location = New System.Drawing.Point(127, 253)
        Me._lblVslStUnit_4.Name = "_lblVslStUnit_4"
        Me._lblVslStUnit_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVslStUnit_4.Size = New System.Drawing.Size(17, 17)
        Me._lblVslStUnit_4.TabIndex = 21
        '
        '_lblPercent_1
        '
        Me._lblPercent_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblPercent_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblPercent_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblPercent_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPercent.SetIndex(Me._lblPercent_1, CType(1, Short))
        Me._lblPercent_1.Location = New System.Drawing.Point(192, 232)
        Me._lblPercent_1.Name = "_lblPercent_1"
        Me._lblPercent_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblPercent_1.Size = New System.Drawing.Size(36, 13)
        Me._lblPercent_1.TabIndex = 20
        Me._lblPercent_1.Text = "% WD"
        '
        '_lblLengthUnit_0
        '
        Me._lblLengthUnit_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLengthUnit_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_0, CType(0, Short))
        Me._lblLengthUnit_0.Location = New System.Drawing.Point(130, 204)
        Me._lblLengthUnit_0.Name = "_lblLengthUnit_0"
        Me._lblLengthUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_0.Size = New System.Drawing.Size(15, 17)
        Me._lblLengthUnit_0.TabIndex = 19
        Me._lblLengthUnit_0.Text = "ft"
        '
        'lblOffset
        '
        Me.lblOffset.BackColor = System.Drawing.SystemColors.Control
        Me.lblOffset.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblOffset.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOffset.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOffset.Location = New System.Drawing.Point(19, 230)
        Me.lblOffset.Name = "lblOffset"
        Me.lblOffset.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblOffset.Size = New System.Drawing.Size(47, 13)
        Me.lblOffset.TabIndex = 18
        Me.lblOffset.Text = "Offset"
        '
        'lblMaxOffset
        '
        Me.lblMaxOffset.BackColor = System.Drawing.SystemColors.Control
        Me.lblMaxOffset.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMaxOffset.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMaxOffset.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblMaxOffset.Location = New System.Drawing.Point(19, 206)
        Me.lblMaxOffset.Name = "lblMaxOffset"
        Me.lblMaxOffset.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMaxOffset.Size = New System.Drawing.Size(63, 13)
        Me.lblMaxOffset.TabIndex = 17
        Me.lblMaxOffset.Text = "Max Offset"
        '
        '_lblPercent_0
        '
        Me._lblPercent_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblPercent_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblPercent_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblPercent_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPercent.SetIndex(Me._lblPercent_0, CType(0, Short))
        Me._lblPercent_0.Location = New System.Drawing.Point(192, 206)
        Me._lblPercent_0.Name = "_lblPercent_0"
        Me._lblPercent_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblPercent_0.Size = New System.Drawing.Size(52, 13)
        Me._lblPercent_0.TabIndex = 16
        Me._lblPercent_0.Text = "% WD  @"
        '
        'lblBearing
        '
        Me.lblBearing.BackColor = System.Drawing.SystemColors.Control
        Me.lblBearing.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblBearing.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBearing.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblBearing.Location = New System.Drawing.Point(19, 254)
        Me.lblBearing.Name = "lblBearing"
        Me.lblBearing.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblBearing.Size = New System.Drawing.Size(59, 17)
        Me.lblBearing.TabIndex = 15
        Me.lblBearing.Text = "Bearing"
        '
        '_lblLengthUnit_1
        '
        Me._lblLengthUnit_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLengthUnit_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_1, CType(1, Short))
        Me._lblLengthUnit_1.Location = New System.Drawing.Point(129, 230)
        Me._lblLengthUnit_1.Name = "_lblLengthUnit_1"
        Me._lblLengthUnit_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_1.Size = New System.Drawing.Size(15, 17)
        Me._lblLengthUnit_1.TabIndex = 14
        Me._lblLengthUnit_1.Text = "ft"
        '
        'btnTransient
        '
        Me.btnTransient.BackColor = System.Drawing.SystemColors.Control
        Me.btnTransient.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnTransient.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTransient.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnTransient.Location = New System.Drawing.Point(300, 31)
        Me.btnTransient.Name = "btnTransient"
        Me.btnTransient.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnTransient.Size = New System.Drawing.Size(84, 25)
        Me.btnTransient.TabIndex = 1
        Me.btnTransient.Text = "Start"
        Me.btnTransient.UseVisualStyleBackColor = False
        '
        'picPolar
        '
        Me.picPolar.BackColor = System.Drawing.SystemColors.Control
        Me.picPolar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.picPolar.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.picPolar.Cursor = System.Windows.Forms.Cursors.Default
        Me.picPolar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.picPolar.Location = New System.Drawing.Point(390, 31)
        Me.picPolar.Name = "picPolar"
        Me.picPolar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.picPolar.Size = New System.Drawing.Size(475, 525)
        Me.picPolar.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.picPolar.TabIndex = 0
        Me.picPolar.TabStop = False
        '
        'mnuFilePre
        '
        '
        '_mnuFilePre_0
        '
        Me.mnuFilePre.SetIndex(Me._mnuFilePre_0, CType(0, Short))
        Me._mnuFilePre_0.Name = "_mnuFilePre_0"
        Me._mnuFilePre_0.Size = New System.Drawing.Size(114, 22)
        Me._mnuFilePre_0.Visible = False
        '
        '_mnuFilePre_1
        '
        Me.mnuFilePre.SetIndex(Me._mnuFilePre_1, CType(1, Short))
        Me._mnuFilePre_1.Name = "_mnuFilePre_1"
        Me._mnuFilePre_1.Size = New System.Drawing.Size(114, 22)
        Me._mnuFilePre_1.Visible = False
        '
        '_mnuFilePre_2
        '
        Me.mnuFilePre.SetIndex(Me._mnuFilePre_2, CType(2, Short))
        Me._mnuFilePre_2.Name = "_mnuFilePre_2"
        Me._mnuFilePre_2.Size = New System.Drawing.Size(114, 22)
        Me._mnuFilePre_2.Visible = False
        '
        '_mnuFilePre_3
        '
        Me.mnuFilePre.SetIndex(Me._mnuFilePre_3, CType(3, Short))
        Me._mnuFilePre_3.Name = "_mnuFilePre_3"
        Me._mnuFilePre_3.Size = New System.Drawing.Size(114, 22)
        Me._mnuFilePre_3.Visible = False
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFile, Me.mnuOption})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(873, 24)
        Me.MainMenu1.TabIndex = 48
        '
        'mnuFile
        '
        Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFileNew, Me.mnuFileOpen, Me.mnuLine1, Me.mnuFileSave, Me.mnuFileSaveAs, Me.mnuLine2, Me._mnuFilePre_0, Me._mnuFilePre_1, Me._mnuFilePre_2, Me._mnuFilePre_3, Me.mnuLinePreFile, Me.mnuFileExit})
        Me.mnuFile.Name = "mnuFile"
        Me.mnuFile.Size = New System.Drawing.Size(37, 20)
        Me.mnuFile.Text = "&File"
        '
        'mnuFileNew
        '
        Me.mnuFileNew.Name = "mnuFileNew"
        Me.mnuFileNew.Size = New System.Drawing.Size(114, 22)
        Me.mnuFileNew.Text = "&New"
        '
        'mnuFileOpen
        '
        Me.mnuFileOpen.Name = "mnuFileOpen"
        Me.mnuFileOpen.Size = New System.Drawing.Size(114, 22)
        Me.mnuFileOpen.Text = "&Open"
        '
        'mnuLine1
        '
        Me.mnuLine1.Name = "mnuLine1"
        Me.mnuLine1.Size = New System.Drawing.Size(111, 6)
        '
        'mnuFileSave
        '
        Me.mnuFileSave.Name = "mnuFileSave"
        Me.mnuFileSave.Size = New System.Drawing.Size(114, 22)
        Me.mnuFileSave.Text = "&Save"
        '
        'mnuFileSaveAs
        '
        Me.mnuFileSaveAs.Name = "mnuFileSaveAs"
        Me.mnuFileSaveAs.Size = New System.Drawing.Size(114, 22)
        Me.mnuFileSaveAs.Text = "Save &As"
        '
        'mnuLine2
        '
        Me.mnuLine2.Name = "mnuLine2"
        Me.mnuLine2.Size = New System.Drawing.Size(111, 6)
        '
        'mnuLinePreFile
        '
        Me.mnuLinePreFile.Name = "mnuLinePreFile"
        Me.mnuLinePreFile.Size = New System.Drawing.Size(111, 6)
        Me.mnuLinePreFile.Visible = False
        '
        'mnuFileExit
        '
        Me.mnuFileExit.Name = "mnuFileExit"
        Me.mnuFileExit.Size = New System.Drawing.Size(114, 22)
        Me.mnuFileExit.Text = "E&xit"
        '
        'mnuOption
        '
        Me.mnuOption.Name = "mnuOption"
        Me.mnuOption.Size = New System.Drawing.Size(56, 20)
        Me.mnuOption.Text = "&Option"
        '
        'frmTransient
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(873, 568)
        Me.Controls.Add(Me.fraRiser)
        Me.Controls.Add(Me.fraWell)
        Me.Controls.Add(Me.btnReport)
        Me.Controls.Add(Me.fraEnvironment)
        Me.Controls.Add(Me.fraTransientTime)
        Me.Controls.Add(Me.fraTransientMotion)
        Me.Controls.Add(Me.btnTransient)
        Me.Controls.Add(Me.picPolar)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(155, 306)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmTransient"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Transient Analysis"
        Me.fraRiser.ResumeLayout(False)
        Me.fraRiser.PerformLayout()
        Me.fraWell.ResumeLayout(False)
        Me.fraWell.PerformLayout()
        Me.fraEnvironment.ResumeLayout(False)
        Me.fraEnvironment.PerformLayout()
        Me.fraTransientTime.ResumeLayout(False)
        Me.fraTransientTime.PerformLayout()
        Me.fraTransientMotion.ResumeLayout(False)
        Me.fraTransientMotion.PerformLayout()
        CType(Me.grdTM, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picPolar, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblDiaUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblForceUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblLengthUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblMassUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblPercent, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblRiser, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVslSt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVslStUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuFilePre, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtRiser, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtVslSt, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents grdTM As DataGridView
    Friend WithEvents picPolar As PictureBox
    Public WithEvents Label1 As Label
#End Region
End Class