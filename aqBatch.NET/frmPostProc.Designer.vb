<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPostProc
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
    Public WithEvents txtSubTitle As System.Windows.Forms.TextBox
    Public WithEvents _txtInitPos_5 As System.Windows.Forms.TextBox
    Public WithEvents _txtInitPos_4 As System.Windows.Forms.TextBox
    Public WithEvents _txtInitPos_3 As System.Windows.Forms.TextBox
    Public WithEvents _txtInitPos_2 As System.Windows.Forms.TextBox
    Public WithEvents _txtInitPos_1 As System.Windows.Forms.TextBox
    Public WithEvents _txtInitPos_0 As System.Windows.Forms.TextBox
    Public WithEvents _lblAngleUnit_2 As System.Windows.Forms.Label
    Public WithEvents _lblAngleUnit_1 As System.Windows.Forms.Label
    Public WithEvents _lblAngleUnit_0 As System.Windows.Forms.Label
    Public WithEvents _lblLengthUnit_3 As System.Windows.Forms.Label
    Public WithEvents _lblLengthUnit_2 As System.Windows.Forms.Label
    Public WithEvents _lblLengthUnit_1 As System.Windows.Forms.Label
    Public WithEvents _lblInitPos_5 As System.Windows.Forms.Label
    Public WithEvents _lblInitPos_4 As System.Windows.Forms.Label
    Public WithEvents _lblInitPos_3 As System.Windows.Forms.Label
    Public WithEvents _lblInitPos_2 As System.Windows.Forms.Label
    Public WithEvents _lblInitPos_1 As System.Windows.Forms.Label
    Public WithEvents _lblInitPos_0 As System.Windows.Forms.Label
    Public WithEvents fraInitialPosition As System.Windows.Forms.GroupBox
    Public WithEvents txtNumLines As System.Windows.Forms.TextBox
    Public WithEvents txtWaterDepth As System.Windows.Forms.TextBox
    Public WithEvents _lblLengthUnit_0 As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents fraMoor As System.Windows.Forms.GroupBox
    Public WithEvents _optUnit_1 As System.Windows.Forms.RadioButton
    Public WithEvents _optUnit_0 As System.Windows.Forms.RadioButton
    Public WithEvents frmUnit As System.Windows.Forms.GroupBox
    Public WithEvents btnOK As System.Windows.Forms.Button
    Public WithEvents btnCancel As System.Windows.Forms.Button
    Public WithEvents txtAnalysisTitle As System.Windows.Forms.TextBox
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents lblAngleUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblForceUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblInitPos As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblLengthUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents optUnit As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    Public WithEvents txtInitPos As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtSubTitle = New System.Windows.Forms.TextBox()
        Me.fraInitialPosition = New System.Windows.Forms.GroupBox()
        Me._txtInitPos_5 = New System.Windows.Forms.TextBox()
        Me._txtInitPos_4 = New System.Windows.Forms.TextBox()
        Me._txtInitPos_3 = New System.Windows.Forms.TextBox()
        Me._txtInitPos_2 = New System.Windows.Forms.TextBox()
        Me._txtInitPos_1 = New System.Windows.Forms.TextBox()
        Me._txtInitPos_0 = New System.Windows.Forms.TextBox()
        Me._lblAngleUnit_2 = New System.Windows.Forms.Label()
        Me._lblAngleUnit_1 = New System.Windows.Forms.Label()
        Me._lblAngleUnit_0 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_3 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_2 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_1 = New System.Windows.Forms.Label()
        Me._lblInitPos_5 = New System.Windows.Forms.Label()
        Me._lblInitPos_4 = New System.Windows.Forms.Label()
        Me._lblInitPos_3 = New System.Windows.Forms.Label()
        Me._lblInitPos_2 = New System.Windows.Forms.Label()
        Me._lblInitPos_1 = New System.Windows.Forms.Label()
        Me._lblInitPos_0 = New System.Windows.Forms.Label()
        Me.fraMoor = New System.Windows.Forms.GroupBox()
        Me.txtNumLines = New System.Windows.Forms.TextBox()
        Me.txtWaterDepth = New System.Windows.Forms.TextBox()
        Me._lblLengthUnit_0 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.frmUnit = New System.Windows.Forms.GroupBox()
        Me._optUnit_1 = New System.Windows.Forms.RadioButton()
        Me._optUnit_0 = New System.Windows.Forms.RadioButton()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.txtAnalysisTitle = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblAngleUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblForceUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblInitPos = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblLengthUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.optUnit = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.txtInitPos = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.grdLinePreTBS = New System.Windows.Forms.DataGridView()
        Me.fraInitialPosition.SuspendLayout()
        Me.fraMoor.SuspendLayout()
        Me.frmUnit.SuspendLayout()
        CType(Me.lblAngleUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblForceUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblInitPos, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblLengthUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtInitPos, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdLinePreTBS, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtSubTitle
        '
        Me.txtSubTitle.AcceptsReturn = True
        Me.txtSubTitle.BackColor = System.Drawing.SystemColors.Window
        Me.txtSubTitle.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSubTitle.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSubTitle.Location = New System.Drawing.Point(12, 62)
        Me.txtSubTitle.Name = "txtSubTitle"
        Me.txtSubTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSubTitle.Size = New System.Drawing.Size(229, 20)
        Me.txtSubTitle.TabIndex = 1
        '
        'fraInitialPosition
        '
        Me.fraInitialPosition.BackColor = System.Drawing.SystemColors.Control
        Me.fraInitialPosition.Controls.Add(Me._txtInitPos_5)
        Me.fraInitialPosition.Controls.Add(Me._txtInitPos_4)
        Me.fraInitialPosition.Controls.Add(Me._txtInitPos_3)
        Me.fraInitialPosition.Controls.Add(Me._txtInitPos_2)
        Me.fraInitialPosition.Controls.Add(Me._txtInitPos_1)
        Me.fraInitialPosition.Controls.Add(Me._txtInitPos_0)
        Me.fraInitialPosition.Controls.Add(Me._lblAngleUnit_2)
        Me.fraInitialPosition.Controls.Add(Me._lblAngleUnit_1)
        Me.fraInitialPosition.Controls.Add(Me._lblAngleUnit_0)
        Me.fraInitialPosition.Controls.Add(Me._lblLengthUnit_3)
        Me.fraInitialPosition.Controls.Add(Me._lblLengthUnit_2)
        Me.fraInitialPosition.Controls.Add(Me._lblLengthUnit_1)
        Me.fraInitialPosition.Controls.Add(Me._lblInitPos_5)
        Me.fraInitialPosition.Controls.Add(Me._lblInitPos_4)
        Me.fraInitialPosition.Controls.Add(Me._lblInitPos_3)
        Me.fraInitialPosition.Controls.Add(Me._lblInitPos_2)
        Me.fraInitialPosition.Controls.Add(Me._lblInitPos_1)
        Me.fraInitialPosition.Controls.Add(Me._lblInitPos_0)
        Me.fraInitialPosition.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraInitialPosition.Location = New System.Drawing.Point(252, 54)
        Me.fraInitialPosition.Name = "fraInitialPosition"
        Me.fraInitialPosition.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraInitialPosition.Size = New System.Drawing.Size(161, 207)
        Me.fraInitialPosition.TabIndex = 18
        Me.fraInitialPosition.TabStop = False
        Me.fraInitialPosition.Text = "Vessel Initial Position (CG)"
        '
        '_txtInitPos_5
        '
        Me._txtInitPos_5.AcceptsReturn = True
        Me._txtInitPos_5.BackColor = System.Drawing.SystemColors.Window
        Me._txtInitPos_5.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtInitPos_5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtInitPos.SetIndex(Me._txtInitPos_5, CType(5, Short))
        Me._txtInitPos_5.Location = New System.Drawing.Point(47, 149)
        Me._txtInitPos_5.Name = "_txtInitPos_5"
        Me._txtInitPos_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtInitPos_5.Size = New System.Drawing.Size(75, 20)
        Me._txtInitPos_5.TabIndex = 9
        Me._txtInitPos_5.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtInitPos_4
        '
        Me._txtInitPos_4.AcceptsReturn = True
        Me._txtInitPos_4.BackColor = System.Drawing.SystemColors.Window
        Me._txtInitPos_4.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtInitPos_4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtInitPos.SetIndex(Me._txtInitPos_4, CType(4, Short))
        Me._txtInitPos_4.Location = New System.Drawing.Point(47, 125)
        Me._txtInitPos_4.Name = "_txtInitPos_4"
        Me._txtInitPos_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtInitPos_4.Size = New System.Drawing.Size(75, 20)
        Me._txtInitPos_4.TabIndex = 8
        Me._txtInitPos_4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtInitPos_3
        '
        Me._txtInitPos_3.AcceptsReturn = True
        Me._txtInitPos_3.BackColor = System.Drawing.SystemColors.Window
        Me._txtInitPos_3.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtInitPos_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtInitPos.SetIndex(Me._txtInitPos_3, CType(3, Short))
        Me._txtInitPos_3.Location = New System.Drawing.Point(47, 101)
        Me._txtInitPos_3.Name = "_txtInitPos_3"
        Me._txtInitPos_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtInitPos_3.Size = New System.Drawing.Size(75, 20)
        Me._txtInitPos_3.TabIndex = 7
        Me._txtInitPos_3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtInitPos_2
        '
        Me._txtInitPos_2.AcceptsReturn = True
        Me._txtInitPos_2.BackColor = System.Drawing.SystemColors.Window
        Me._txtInitPos_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtInitPos_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtInitPos.SetIndex(Me._txtInitPos_2, CType(2, Short))
        Me._txtInitPos_2.Location = New System.Drawing.Point(47, 77)
        Me._txtInitPos_2.Name = "_txtInitPos_2"
        Me._txtInitPos_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtInitPos_2.Size = New System.Drawing.Size(75, 20)
        Me._txtInitPos_2.TabIndex = 6
        Me._txtInitPos_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtInitPos_1
        '
        Me._txtInitPos_1.AcceptsReturn = True
        Me._txtInitPos_1.BackColor = System.Drawing.SystemColors.Window
        Me._txtInitPos_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtInitPos_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtInitPos.SetIndex(Me._txtInitPos_1, CType(1, Short))
        Me._txtInitPos_1.Location = New System.Drawing.Point(47, 53)
        Me._txtInitPos_1.Name = "_txtInitPos_1"
        Me._txtInitPos_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtInitPos_1.Size = New System.Drawing.Size(75, 20)
        Me._txtInitPos_1.TabIndex = 5
        Me._txtInitPos_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtInitPos_0
        '
        Me._txtInitPos_0.AcceptsReturn = True
        Me._txtInitPos_0.BackColor = System.Drawing.SystemColors.Window
        Me._txtInitPos_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtInitPos_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtInitPos.SetIndex(Me._txtInitPos_0, CType(0, Short))
        Me._txtInitPos_0.Location = New System.Drawing.Point(47, 29)
        Me._txtInitPos_0.Name = "_txtInitPos_0"
        Me._txtInitPos_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtInitPos_0.Size = New System.Drawing.Size(75, 20)
        Me._txtInitPos_0.TabIndex = 4
        Me._txtInitPos_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_lblAngleUnit_2
        '
        Me._lblAngleUnit_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblAngleUnit_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblAngleUnit_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblAngleUnit.SetIndex(Me._lblAngleUnit_2, CType(2, Short))
        Me._lblAngleUnit_2.Location = New System.Drawing.Point(126, 149)
        Me._lblAngleUnit_2.Name = "_lblAngleUnit_2"
        Me._lblAngleUnit_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblAngleUnit_2.Size = New System.Drawing.Size(17, 17)
        Me._lblAngleUnit_2.TabIndex = 33
        '
        '_lblAngleUnit_1
        '
        Me._lblAngleUnit_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblAngleUnit_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblAngleUnit_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblAngleUnit.SetIndex(Me._lblAngleUnit_1, CType(1, Short))
        Me._lblAngleUnit_1.Location = New System.Drawing.Point(126, 125)
        Me._lblAngleUnit_1.Name = "_lblAngleUnit_1"
        Me._lblAngleUnit_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblAngleUnit_1.Size = New System.Drawing.Size(17, 17)
        Me._lblAngleUnit_1.TabIndex = 32
        '
        '_lblAngleUnit_0
        '
        Me._lblAngleUnit_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblAngleUnit_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblAngleUnit_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblAngleUnit.SetIndex(Me._lblAngleUnit_0, CType(0, Short))
        Me._lblAngleUnit_0.Location = New System.Drawing.Point(126, 101)
        Me._lblAngleUnit_0.Name = "_lblAngleUnit_0"
        Me._lblAngleUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblAngleUnit_0.Size = New System.Drawing.Size(17, 17)
        Me._lblAngleUnit_0.TabIndex = 31
        '
        '_lblLengthUnit_3
        '
        Me._lblLengthUnit_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_3, CType(3, Short))
        Me._lblLengthUnit_3.Location = New System.Drawing.Point(128, 83)
        Me._lblLengthUnit_3.Name = "_lblLengthUnit_3"
        Me._lblLengthUnit_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_3.Size = New System.Drawing.Size(26, 20)
        Me._lblLengthUnit_3.TabIndex = 30
        Me._lblLengthUnit_3.Text = "ft"
        '
        '_lblLengthUnit_2
        '
        Me._lblLengthUnit_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_2, CType(2, Short))
        Me._lblLengthUnit_2.Location = New System.Drawing.Point(128, 58)
        Me._lblLengthUnit_2.Name = "_lblLengthUnit_2"
        Me._lblLengthUnit_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_2.Size = New System.Drawing.Size(26, 20)
        Me._lblLengthUnit_2.TabIndex = 29
        Me._lblLengthUnit_2.Text = "ft"
        '
        '_lblLengthUnit_1
        '
        Me._lblLengthUnit_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_1, CType(1, Short))
        Me._lblLengthUnit_1.Location = New System.Drawing.Point(128, 32)
        Me._lblLengthUnit_1.Name = "_lblLengthUnit_1"
        Me._lblLengthUnit_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_1.Size = New System.Drawing.Size(26, 20)
        Me._lblLengthUnit_1.TabIndex = 28
        Me._lblLengthUnit_1.Text = "ft"
        '
        '_lblInitPos_5
        '
        Me._lblInitPos_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblInitPos_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblInitPos_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblInitPos.SetIndex(Me._lblInitPos_5, CType(5, Short))
        Me._lblInitPos_5.Location = New System.Drawing.Point(23, 151)
        Me._lblInitPos_5.Name = "_lblInitPos_5"
        Me._lblInitPos_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblInitPos_5.Size = New System.Drawing.Size(58, 20)
        Me._lblInitPos_5.TabIndex = 24
        Me._lblInitPos_5.Text = "Rz:"
        '
        '_lblInitPos_4
        '
        Me._lblInitPos_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblInitPos_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblInitPos_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblInitPos.SetIndex(Me._lblInitPos_4, CType(4, Short))
        Me._lblInitPos_4.Location = New System.Drawing.Point(23, 127)
        Me._lblInitPos_4.Name = "_lblInitPos_4"
        Me._lblInitPos_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblInitPos_4.Size = New System.Drawing.Size(49, 20)
        Me._lblInitPos_4.TabIndex = 23
        Me._lblInitPos_4.Text = "Ry:"
        '
        '_lblInitPos_3
        '
        Me._lblInitPos_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblInitPos_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblInitPos_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblInitPos.SetIndex(Me._lblInitPos_3, CType(3, Short))
        Me._lblInitPos_3.Location = New System.Drawing.Point(23, 103)
        Me._lblInitPos_3.Name = "_lblInitPos_3"
        Me._lblInitPos_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblInitPos_3.Size = New System.Drawing.Size(54, 20)
        Me._lblInitPos_3.TabIndex = 22
        Me._lblInitPos_3.Text = "Rx:"
        '
        '_lblInitPos_2
        '
        Me._lblInitPos_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblInitPos_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblInitPos_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblInitPos.SetIndex(Me._lblInitPos_2, CType(2, Short))
        Me._lblInitPos_2.Location = New System.Drawing.Point(23, 78)
        Me._lblInitPos_2.Name = "_lblInitPos_2"
        Me._lblInitPos_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblInitPos_2.Size = New System.Drawing.Size(42, 20)
        Me._lblInitPos_2.TabIndex = 21
        Me._lblInitPos_2.Text = "Z:"
        '
        '_lblInitPos_1
        '
        Me._lblInitPos_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblInitPos_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblInitPos_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblInitPos.SetIndex(Me._lblInitPos_1, CType(1, Short))
        Me._lblInitPos_1.Location = New System.Drawing.Point(23, 54)
        Me._lblInitPos_1.Name = "_lblInitPos_1"
        Me._lblInitPos_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblInitPos_1.Size = New System.Drawing.Size(42, 20)
        Me._lblInitPos_1.TabIndex = 20
        Me._lblInitPos_1.Text = "Y:"
        '
        '_lblInitPos_0
        '
        Me._lblInitPos_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblInitPos_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblInitPos_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblInitPos.SetIndex(Me._lblInitPos_0, CType(0, Short))
        Me._lblInitPos_0.Location = New System.Drawing.Point(23, 29)
        Me._lblInitPos_0.Name = "_lblInitPos_0"
        Me._lblInitPos_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblInitPos_0.Size = New System.Drawing.Size(42, 20)
        Me._lblInitPos_0.TabIndex = 19
        Me._lblInitPos_0.Text = "X:"
        '
        'fraMoor
        '
        Me.fraMoor.BackColor = System.Drawing.SystemColors.Control
        Me.fraMoor.Controls.Add(Me.txtNumLines)
        Me.fraMoor.Controls.Add(Me.txtWaterDepth)
        Me.fraMoor.Controls.Add(Me._lblLengthUnit_0)
        Me.fraMoor.Controls.Add(Me.Label4)
        Me.fraMoor.Controls.Add(Me.Label3)
        Me.fraMoor.ForeColor = System.Drawing.SystemColors.WindowText
        Me.fraMoor.Location = New System.Drawing.Point(11, 185)
        Me.fraMoor.Name = "fraMoor"
        Me.fraMoor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraMoor.Size = New System.Drawing.Size(230, 76)
        Me.fraMoor.TabIndex = 16
        Me.fraMoor.TabStop = False
        Me.fraMoor.Text = "Mooring Design Criteria"
        '
        'txtNumLines
        '
        Me.txtNumLines.AcceptsReturn = True
        Me.txtNumLines.BackColor = System.Drawing.SystemColors.Window
        Me.txtNumLines.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNumLines.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNumLines.Location = New System.Drawing.Point(145, 40)
        Me.txtNumLines.Name = "txtNumLines"
        Me.txtNumLines.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNumLines.Size = New System.Drawing.Size(36, 20)
        Me.txtNumLines.TabIndex = 3
        Me.txtNumLines.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtWaterDepth
        '
        Me.txtWaterDepth.AcceptsReturn = True
        Me.txtWaterDepth.BackColor = System.Drawing.SystemColors.Window
        Me.txtWaterDepth.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtWaterDepth.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtWaterDepth.Location = New System.Drawing.Point(108, 18)
        Me.txtWaterDepth.Name = "txtWaterDepth"
        Me.txtWaterDepth.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtWaterDepth.Size = New System.Drawing.Size(73, 20)
        Me.txtWaterDepth.TabIndex = 2
        Me.txtWaterDepth.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_lblLengthUnit_0
        '
        Me._lblLengthUnit_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_0, CType(0, Short))
        Me._lblLengthUnit_0.Location = New System.Drawing.Point(184, 24)
        Me._lblLengthUnit_0.Name = "_lblLengthUnit_0"
        Me._lblLengthUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_0.Size = New System.Drawing.Size(26, 28)
        Me._lblLengthUnit_0.TabIndex = 27
        Me._lblLengthUnit_0.Text = "ft"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(14, 43)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(134, 20)
        Me.Label4.TabIndex = 25
        Me.Label4.Text = "Number of Mooring Lines:"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(16, 21)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(90, 20)
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "Water Depth:"
        '
        'frmUnit
        '
        Me.frmUnit.BackColor = System.Drawing.SystemColors.Control
        Me.frmUnit.Controls.Add(Me._optUnit_1)
        Me.frmUnit.Controls.Add(Me._optUnit_0)
        Me.frmUnit.ForeColor = System.Drawing.SystemColors.WindowText
        Me.frmUnit.Location = New System.Drawing.Point(11, 88)
        Me.frmUnit.Name = "frmUnit"
        Me.frmUnit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.frmUnit.Size = New System.Drawing.Size(231, 89)
        Me.frmUnit.TabIndex = 13
        Me.frmUnit.TabStop = False
        Me.frmUnit.Text = " Result Data Unit"
        '
        '_optUnit_1
        '
        Me._optUnit_1.BackColor = System.Drawing.SystemColors.Control
        Me._optUnit_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optUnit_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optUnit.SetIndex(Me._optUnit_1, CType(1, Short))
        Me._optUnit_1.Location = New System.Drawing.Point(28, 52)
        Me._optUnit_1.Name = "_optUnit_1"
        Me._optUnit_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optUnit_1.Size = New System.Drawing.Size(100, 23)
        Me._optUnit_1.TabIndex = 15
        Me._optUnit_1.TabStop = True
        Me._optUnit_1.Text = "Metric (m, KN)"
        Me._optUnit_1.UseVisualStyleBackColor = False
        '
        '_optUnit_0
        '
        Me._optUnit_0.BackColor = System.Drawing.SystemColors.Control
        Me._optUnit_0.Checked = True
        Me._optUnit_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optUnit_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optUnit.SetIndex(Me._optUnit_0, CType(0, Short))
        Me._optUnit_0.Location = New System.Drawing.Point(28, 25)
        Me._optUnit_0.Name = "_optUnit_0"
        Me._optUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optUnit_0.Size = New System.Drawing.Size(105, 23)
        Me._optUnit_0.TabIndex = 14
        Me._optUnit_0.TabStop = True
        Me._optUnit_0.Text = "English (ft, lbs)"
        Me._optUnit_0.UseVisualStyleBackColor = False
        '
        'btnOK
        '
        Me.btnOK.BackColor = System.Drawing.SystemColors.Control
        Me.btnOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnOK.Location = New System.Drawing.Point(428, 13)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnOK.Size = New System.Drawing.Size(74, 25)
        Me.btnOK.TabIndex = 12
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = False
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.SystemColors.Control
        Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCancel.Location = New System.Drawing.Point(428, 45)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCancel.Size = New System.Drawing.Size(74, 25)
        Me.btnCancel.TabIndex = 11
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'txtAnalysisTitle
        '
        Me.txtAnalysisTitle.AcceptsReturn = True
        Me.txtAnalysisTitle.BackColor = System.Drawing.SystemColors.Window
        Me.txtAnalysisTitle.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAnalysisTitle.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAnalysisTitle.Location = New System.Drawing.Point(11, 17)
        Me.txtAnalysisTitle.Name = "txtAnalysisTitle"
        Me.txtAnalysisTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAnalysisTitle.Size = New System.Drawing.Size(404, 20)
        Me.txtAnalysisTitle.TabIndex = 0
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(11, 46)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(158, 16)
        Me.Label5.TabIndex = 26
        Me.Label5.Text = "Report Environment:"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(11, 4)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(86, 16)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Report Title:"
        '
        'optUnit
        '
        '
        'grdLinePreTBS
        '
        Me.grdLinePreTBS.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdLinePreTBS.Location = New System.Drawing.Point(14, 267)
        Me.grdLinePreTBS.Name = "grdLinePreTBS"
        Me.grdLinePreTBS.Size = New System.Drawing.Size(399, 206)
        Me.grdLinePreTBS.TabIndex = 38
        '
        'frmPostProc
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(513, 530)
        Me.Controls.Add(Me.grdLinePreTBS)
        Me.Controls.Add(Me.txtSubTitle)
        Me.Controls.Add(Me.fraInitialPosition)
        Me.Controls.Add(Me.fraMoor)
        Me.Controls.Add(Me.frmUnit)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.txtAnalysisTitle)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(3, 29)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmPostProc"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Process Mooring Analysis Results"
        Me.fraInitialPosition.ResumeLayout(False)
        Me.fraInitialPosition.PerformLayout()
        Me.fraMoor.ResumeLayout(False)
        Me.fraMoor.PerformLayout()
        Me.frmUnit.ResumeLayout(False)
        CType(Me.lblAngleUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblForceUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblInitPos, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblLengthUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtInitPos, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdLinePreTBS, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents grdLinePreTBS As DataGridView
#End Region
End Class