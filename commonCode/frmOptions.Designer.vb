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
	Public WithEvents btnEnglish As System.Windows.Forms.RadioButton
	Public WithEvents btnMetric As System.Windows.Forms.RadioButton
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
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnOK = New System.Windows.Forms.Button()
        Me.Text2 = New System.Windows.Forms.TextBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.txtNumLineSegments = New System.Windows.Forms.TextBox()
        Me.txtMinSegLength = New System.Windows.Forms.TextBox()
        Me.txtMaxIter = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_3 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.fraIntegration = New System.Windows.Forms.GroupBox()
        Me.txtStartFreq = New System.Windows.Forms.TextBox()
        Me.Text1 = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.fraScopeRef = New System.Windows.Forms.GroupBox()
        Me._optScopeRef_1 = New System.Windows.Forms.RadioButton()
        Me._optScopeRef_0 = New System.Windows.Forms.RadioButton()
        Me.fraTol = New System.Windows.Forms.GroupBox()
        Me.txtVesselHdgTol = New System.Windows.Forms.TextBox()
        Me.txtVesselOffsetTol = New System.Windows.Forms.TextBox()
        Me.txtLineTensionTol = New System.Windows.Forms.TextBox()
        Me.txtAnchorDepthTol = New System.Windows.Forms.TextBox()
        Me.txtSpreadAngleTol = New System.Windows.Forms.TextBox()
        Me.txtScopeTol = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_2 = New System.Windows.Forms.Label()
        Me.lblVesselHdgTol = New System.Windows.Forms.Label()
        Me.lblVesselOffsetTol = New System.Windows.Forms.Label()
        Me.lblForceUnit = New System.Windows.Forms.Label()
        Me._lblLengthUnit_1 = New System.Windows.Forms.Label()
        Me.lblAngleUnit = New System.Windows.Forms.Label()
        Me._lblLengthUnit_0 = New System.Windows.Forms.Label()
        Me.lblForceTol = New System.Windows.Forms.Label()
        Me.lblDepthTol = New System.Windows.Forms.Label()
        Me.lblAngleTol = New System.Windows.Forms.Label()
        Me.lblDistTol = New System.Windows.Forms.Label()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.fraUnit = New System.Windows.Forms.GroupBox()
        Me.btnEnglish = New System.Windows.Forms.RadioButton()
        Me.btnMetric = New System.Windows.Forms.RadioButton()
        Me.Label5 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_4 = New System.Windows.Forms.Label()
        Me.lblLengthUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.optScopeRef = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.optUnit = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.Frame1.SuspendLayout()
        Me.fraIntegration.SuspendLayout()
        Me.fraScopeRef.SuspendLayout()
        Me.fraTol.SuspendLayout()
        Me.fraUnit.SuspendLayout()
        CType(Me.lblLengthUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optScopeRef, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnOK
        '
        Me.btnOK.BackColor = System.Drawing.SystemColors.Control
        Me.btnOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnOK.Location = New System.Drawing.Point(218, 18)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnOK.Size = New System.Drawing.Size(70, 25)
        Me.btnOK.TabIndex = 4
        Me.btnOK.Text = "OK"
        Me.ToolTip1.SetToolTip(Me.btnOK, "Save Final Position")
        Me.btnOK.UseVisualStyleBackColor = False
        '
        'Text2
        '
        Me.Text2.AcceptsReturn = True
        Me.Text2.BackColor = System.Drawing.SystemColors.Window
        Me.Text2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text2.Location = New System.Drawing.Point(204, 455)
        Me.Text2.Name = "Text2"
        Me.Text2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text2.Size = New System.Drawing.Size(45, 20)
        Me.Text2.TabIndex = 36
        Me.Text2.Text = "2"
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.txtNumLineSegments)
        Me.Frame1.Controls.Add(Me.txtMinSegLength)
        Me.Frame1.Controls.Add(Me.txtMaxIter)
        Me.Frame1.Controls.Add(Me.Label4)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.Controls.Add(Me._lblLengthUnit_3)
        Me.Frame1.Controls.Add(Me.Label3)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(8, 295)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(285, 146)
        Me.Frame1.TabIndex = 28
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Capacity Limits"
        '
        'txtNumLineSegments
        '
        Me.txtNumLineSegments.AcceptsReturn = True
        Me.txtNumLineSegments.BackColor = System.Drawing.SystemColors.Window
        Me.txtNumLineSegments.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNumLineSegments.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNumLineSegments.Location = New System.Drawing.Point(210, 71)
        Me.txtNumLineSegments.Name = "txtNumLineSegments"
        Me.txtNumLineSegments.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNumLineSegments.Size = New System.Drawing.Size(45, 20)
        Me.txtNumLineSegments.TabIndex = 34
        Me.txtNumLineSegments.Text = "20"
        '
        'txtMinSegLength
        '
        Me.txtMinSegLength.AcceptsReturn = True
        Me.txtMinSegLength.BackColor = System.Drawing.SystemColors.Window
        Me.txtMinSegLength.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMinSegLength.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMinSegLength.Location = New System.Drawing.Point(211, 21)
        Me.txtMinSegLength.Name = "txtMinSegLength"
        Me.txtMinSegLength.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMinSegLength.Size = New System.Drawing.Size(45, 20)
        Me.txtMinSegLength.TabIndex = 30
        Me.txtMinSegLength.Text = "50"
        '
        'txtMaxIter
        '
        Me.txtMaxIter.AcceptsReturn = True
        Me.txtMaxIter.BackColor = System.Drawing.SystemColors.Window
        Me.txtMaxIter.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMaxIter.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMaxIter.Location = New System.Drawing.Point(211, 44)
        Me.txtMaxIter.Name = "txtMaxIter"
        Me.txtMaxIter.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMaxIter.Size = New System.Drawing.Size(45, 20)
        Me.txtMaxIter.TabIndex = 29
        Me.txtMaxIter.Text = "2000"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(23, 73)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(167, 16)
        Me.Label4.TabIndex = 35
        Me.Label4.Text = "Number of Line Segments:"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(23, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(167, 16)
        Me.Label2.TabIndex = 33
        Me.Label2.Text = "Min. Line Segment Length:"
        '
        '_lblLengthUnit_3
        '
        Me._lblLengthUnit_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_3, CType(3, Short))
        Me._lblLengthUnit_3.Location = New System.Drawing.Point(263, 24)
        Me._lblLengthUnit_3.Name = "_lblLengthUnit_3"
        Me._lblLengthUnit_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_3.Size = New System.Drawing.Size(17, 16)
        Me._lblLengthUnit_3.TabIndex = 32
        Me._lblLengthUnit_3.Text = "ft"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(23, 47)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(100, 16)
        Me.Label3.TabIndex = 31
        Me.Label3.Text = "Max Iteration:"
        '
        'fraIntegration
        '
        Me.fraIntegration.BackColor = System.Drawing.SystemColors.Control
        Me.fraIntegration.Controls.Add(Me.txtStartFreq)
        Me.fraIntegration.Controls.Add(Me.Text1)
        Me.fraIntegration.Controls.Add(Me.Label10)
        Me.fraIntegration.Controls.Add(Me.Label9)
        Me.fraIntegration.Controls.Add(Me.Label8)
        Me.fraIntegration.Controls.Add(Me.Label7)
        Me.fraIntegration.Controls.Add(Me.Label6)
        Me.fraIntegration.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraIntegration.Location = New System.Drawing.Point(302, 99)
        Me.fraIntegration.Name = "fraIntegration"
        Me.fraIntegration.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraIntegration.Size = New System.Drawing.Size(285, 273)
        Me.fraIntegration.TabIndex = 27
        Me.fraIntegration.TabStop = False
        Me.fraIntegration.Text = "Energy Spectrum Integration"
        '
        'txtStartFreq
        '
        Me.txtStartFreq.AcceptsReturn = True
        Me.txtStartFreq.BackColor = System.Drawing.SystemColors.Window
        Me.txtStartFreq.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtStartFreq.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtStartFreq.Location = New System.Drawing.Point(198, 22)
        Me.txtStartFreq.Name = "txtStartFreq"
        Me.txtStartFreq.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtStartFreq.Size = New System.Drawing.Size(45, 20)
        Me.txtStartFreq.TabIndex = 40
        Me.txtStartFreq.Text = "50"
        '
        'Text1
        '
        Me.Text1.AcceptsReturn = True
        Me.Text1.BackColor = System.Drawing.SystemColors.Window
        Me.Text1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text1.Location = New System.Drawing.Point(198, 45)
        Me.Text1.Name = "Text1"
        Me.Text1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text1.Size = New System.Drawing.Size(45, 20)
        Me.Text1.TabIndex = 39
        Me.Text1.Text = "2000"
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Location = New System.Drawing.Point(13, 72)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(167, 16)
        Me.Label10.TabIndex = 45
        Me.Label10.Text = "Wave Frequency Increment:"
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label9.Location = New System.Drawing.Point(16, 104)
        Me.Label9.Name = "Label9"
        Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label9.Size = New System.Drawing.Size(167, 16)
        Me.Label9.TabIndex = 44
        Me.Label9.Text = "Start Low Frequency:"
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(19, 135)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(167, 16)
        Me.Label8.TabIndex = 43
        Me.Label8.Text = "End Low Frequency:"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(9, 26)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(167, 16)
        Me.Label7.TabIndex = 42
        Me.Label7.Text = "Start Wave Frequency:"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(12, 46)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(167, 16)
        Me.Label6.TabIndex = 41
        Me.Label6.Text = "End Wave Frequency:"
        '
        'fraScopeRef
        '
        Me.fraScopeRef.BackColor = System.Drawing.SystemColors.Control
        Me.fraScopeRef.Controls.Add(Me._optScopeRef_1)
        Me.fraScopeRef.Controls.Add(Me._optScopeRef_0)
        Me.fraScopeRef.ForeColor = System.Drawing.Color.Black
        Me.fraScopeRef.Location = New System.Drawing.Point(308, 8)
        Me.fraScopeRef.Name = "fraScopeRef"
        Me.fraScopeRef.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraScopeRef.Size = New System.Drawing.Size(138, 79)
        Me.fraScopeRef.TabIndex = 6
        Me.fraScopeRef.TabStop = False
        Me.fraScopeRef.Text = "Scope Reference"
        Me.fraScopeRef.Visible = False
        '
        '_optScopeRef_1
        '
        Me._optScopeRef_1.BackColor = System.Drawing.SystemColors.Control
        Me._optScopeRef_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optScopeRef_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optScopeRef.SetIndex(Me._optScopeRef_1, CType(1, Short))
        Me._optScopeRef_1.Location = New System.Drawing.Point(20, 44)
        Me._optScopeRef_1.Name = "_optScopeRef_1"
        Me._optScopeRef_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optScopeRef_1.Size = New System.Drawing.Size(98, 22)
        Me._optScopeRef_1.TabIndex = 8
        Me._optScopeRef_1.TabStop = True
        Me._optScopeRef_1.Text = "Vessel Center"
        Me._optScopeRef_1.UseVisualStyleBackColor = False
        '
        '_optScopeRef_0
        '
        Me._optScopeRef_0.BackColor = System.Drawing.SystemColors.Control
        Me._optScopeRef_0.Checked = True
        Me._optScopeRef_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optScopeRef_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optScopeRef.SetIndex(Me._optScopeRef_0, CType(0, Short))
        Me._optScopeRef_0.Location = New System.Drawing.Point(20, 20)
        Me._optScopeRef_0.Name = "_optScopeRef_0"
        Me._optScopeRef_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optScopeRef_0.Size = New System.Drawing.Size(83, 22)
        Me._optScopeRef_0.TabIndex = 7
        Me._optScopeRef_0.TabStop = True
        Me._optScopeRef_0.Text = "Fairlead"
        Me._optScopeRef_0.UseVisualStyleBackColor = False
        '
        'fraTol
        '
        Me.fraTol.BackColor = System.Drawing.SystemColors.Control
        Me.fraTol.Controls.Add(Me.txtVesselHdgTol)
        Me.fraTol.Controls.Add(Me.txtVesselOffsetTol)
        Me.fraTol.Controls.Add(Me.txtLineTensionTol)
        Me.fraTol.Controls.Add(Me.txtAnchorDepthTol)
        Me.fraTol.Controls.Add(Me.txtSpreadAngleTol)
        Me.fraTol.Controls.Add(Me.txtScopeTol)
        Me.fraTol.Controls.Add(Me.Label1)
        Me.fraTol.Controls.Add(Me._lblLengthUnit_2)
        Me.fraTol.Controls.Add(Me.lblVesselHdgTol)
        Me.fraTol.Controls.Add(Me.lblVesselOffsetTol)
        Me.fraTol.Controls.Add(Me.lblForceUnit)
        Me.fraTol.Controls.Add(Me._lblLengthUnit_1)
        Me.fraTol.Controls.Add(Me.lblAngleUnit)
        Me.fraTol.Controls.Add(Me._lblLengthUnit_0)
        Me.fraTol.Controls.Add(Me.lblForceTol)
        Me.fraTol.Controls.Add(Me.lblDepthTol)
        Me.fraTol.Controls.Add(Me.lblAngleTol)
        Me.fraTol.Controls.Add(Me.lblDistTol)
        Me.fraTol.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraTol.Location = New System.Drawing.Point(7, 95)
        Me.fraTol.Name = "fraTol"
        Me.fraTol.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraTol.Size = New System.Drawing.Size(287, 191)
        Me.fraTol.TabIndex = 5
        Me.fraTol.TabStop = False
        Me.fraTol.Text = "Precision/Tolerance"
        '
        'txtVesselHdgTol
        '
        Me.txtVesselHdgTol.AcceptsReturn = True
        Me.txtVesselHdgTol.BackColor = System.Drawing.SystemColors.Window
        Me.txtVesselHdgTol.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtVesselHdgTol.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtVesselHdgTol.Location = New System.Drawing.Point(127, 155)
        Me.txtVesselHdgTol.Name = "txtVesselHdgTol"
        Me.txtVesselHdgTol.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtVesselHdgTol.Size = New System.Drawing.Size(60, 20)
        Me.txtVesselHdgTol.TabIndex = 25
        Me.txtVesselHdgTol.Text = "0.0001"
        '
        'txtVesselOffsetTol
        '
        Me.txtVesselOffsetTol.AcceptsReturn = True
        Me.txtVesselOffsetTol.BackColor = System.Drawing.SystemColors.Window
        Me.txtVesselOffsetTol.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtVesselOffsetTol.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtVesselOffsetTol.Location = New System.Drawing.Point(127, 129)
        Me.txtVesselOffsetTol.Name = "txtVesselOffsetTol"
        Me.txtVesselOffsetTol.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtVesselOffsetTol.Size = New System.Drawing.Size(60, 20)
        Me.txtVesselOffsetTol.TabIndex = 23
        Me.txtVesselOffsetTol.Text = "0.0001"
        '
        'txtLineTensionTol
        '
        Me.txtLineTensionTol.AcceptsReturn = True
        Me.txtLineTensionTol.BackColor = System.Drawing.SystemColors.Window
        Me.txtLineTensionTol.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLineTensionTol.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtLineTensionTol.Location = New System.Drawing.Point(127, 103)
        Me.txtLineTensionTol.Name = "txtLineTensionTol"
        Me.txtLineTensionTol.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtLineTensionTol.Size = New System.Drawing.Size(60, 20)
        Me.txtLineTensionTol.TabIndex = 19
        Me.txtLineTensionTol.Text = "0.5"
        '
        'txtAnchorDepthTol
        '
        Me.txtAnchorDepthTol.AcceptsReturn = True
        Me.txtAnchorDepthTol.BackColor = System.Drawing.SystemColors.Window
        Me.txtAnchorDepthTol.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAnchorDepthTol.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAnchorDepthTol.Location = New System.Drawing.Point(127, 78)
        Me.txtAnchorDepthTol.Name = "txtAnchorDepthTol"
        Me.txtAnchorDepthTol.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAnchorDepthTol.Size = New System.Drawing.Size(60, 20)
        Me.txtAnchorDepthTol.TabIndex = 18
        Me.txtAnchorDepthTol.Text = "0.0005"
        '
        'txtSpreadAngleTol
        '
        Me.txtSpreadAngleTol.AcceptsReturn = True
        Me.txtSpreadAngleTol.BackColor = System.Drawing.SystemColors.Window
        Me.txtSpreadAngleTol.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSpreadAngleTol.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSpreadAngleTol.Location = New System.Drawing.Point(127, 52)
        Me.txtSpreadAngleTol.Name = "txtSpreadAngleTol"
        Me.txtSpreadAngleTol.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSpreadAngleTol.Size = New System.Drawing.Size(60, 20)
        Me.txtSpreadAngleTol.TabIndex = 15
        Me.txtSpreadAngleTol.Text = "0.001"
        '
        'txtScopeTol
        '
        Me.txtScopeTol.AcceptsReturn = True
        Me.txtScopeTol.BackColor = System.Drawing.SystemColors.Window
        Me.txtScopeTol.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtScopeTol.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtScopeTol.Location = New System.Drawing.Point(127, 26)
        Me.txtScopeTol.Name = "txtScopeTol"
        Me.txtScopeTol.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtScopeTol.Size = New System.Drawing.Size(60, 20)
        Me.txtScopeTol.TabIndex = 13
        Me.txtScopeTol.Text = "0.00005"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(196, 157)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(37, 16)
        Me.Label1.TabIndex = 26
        Me.Label1.Text = "deg"
        '
        '_lblLengthUnit_2
        '
        Me._lblLengthUnit_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_2, CType(2, Short))
        Me._lblLengthUnit_2.Location = New System.Drawing.Point(199, 134)
        Me._lblLengthUnit_2.Name = "_lblLengthUnit_2"
        Me._lblLengthUnit_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_2.Size = New System.Drawing.Size(17, 16)
        Me._lblLengthUnit_2.TabIndex = 24
        Me._lblLengthUnit_2.Text = "ft"
        '
        'lblVesselHdgTol
        '
        Me.lblVesselHdgTol.BackColor = System.Drawing.SystemColors.Control
        Me.lblVesselHdgTol.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblVesselHdgTol.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVesselHdgTol.Location = New System.Drawing.Point(34, 160)
        Me.lblVesselHdgTol.Name = "lblVesselHdgTol"
        Me.lblVesselHdgTol.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblVesselHdgTol.Size = New System.Drawing.Size(91, 16)
        Me.lblVesselHdgTol.TabIndex = 22
        Me.lblVesselHdgTol.Text = "Vessel Offset:"
        '
        'lblVesselOffsetTol
        '
        Me.lblVesselOffsetTol.BackColor = System.Drawing.SystemColors.Control
        Me.lblVesselOffsetTol.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblVesselOffsetTol.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVesselOffsetTol.Location = New System.Drawing.Point(34, 134)
        Me.lblVesselOffsetTol.Name = "lblVesselOffsetTol"
        Me.lblVesselOffsetTol.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblVesselOffsetTol.Size = New System.Drawing.Size(91, 16)
        Me.lblVesselOffsetTol.TabIndex = 21
        Me.lblVesselOffsetTol.Text = "Vessel Offset:"
        '
        'lblForceUnit
        '
        Me.lblForceUnit.BackColor = System.Drawing.SystemColors.Control
        Me.lblForceUnit.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblForceUnit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblForceUnit.Location = New System.Drawing.Point(200, 101)
        Me.lblForceUnit.Name = "lblForceUnit"
        Me.lblForceUnit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblForceUnit.Size = New System.Drawing.Size(37, 16)
        Me.lblForceUnit.TabIndex = 20
        Me.lblForceUnit.Text = "kips"
        '
        '_lblLengthUnit_1
        '
        Me._lblLengthUnit_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_1, CType(1, Short))
        Me._lblLengthUnit_1.Location = New System.Drawing.Point(200, 77)
        Me._lblLengthUnit_1.Name = "_lblLengthUnit_1"
        Me._lblLengthUnit_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_1.Size = New System.Drawing.Size(17, 16)
        Me._lblLengthUnit_1.TabIndex = 17
        Me._lblLengthUnit_1.Text = "ft"
        '
        'lblAngleUnit
        '
        Me.lblAngleUnit.BackColor = System.Drawing.SystemColors.Control
        Me.lblAngleUnit.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblAngleUnit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblAngleUnit.Location = New System.Drawing.Point(198, 54)
        Me.lblAngleUnit.Name = "lblAngleUnit"
        Me.lblAngleUnit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblAngleUnit.Size = New System.Drawing.Size(37, 16)
        Me.lblAngleUnit.TabIndex = 16
        Me.lblAngleUnit.Text = "deg"
        '
        '_lblLengthUnit_0
        '
        Me._lblLengthUnit_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_0, CType(0, Short))
        Me._lblLengthUnit_0.Location = New System.Drawing.Point(201, 31)
        Me._lblLengthUnit_0.Name = "_lblLengthUnit_0"
        Me._lblLengthUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_0.Size = New System.Drawing.Size(17, 16)
        Me._lblLengthUnit_0.TabIndex = 14
        Me._lblLengthUnit_0.Text = "ft"
        '
        'lblForceTol
        '
        Me.lblForceTol.BackColor = System.Drawing.SystemColors.Control
        Me.lblForceTol.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblForceTol.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblForceTol.Location = New System.Drawing.Point(34, 108)
        Me.lblForceTol.Name = "lblForceTol"
        Me.lblForceTol.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblForceTol.Size = New System.Drawing.Size(112, 16)
        Me.lblForceTol.TabIndex = 12
        Me.lblForceTol.Text = "Line Tension:"
        '
        'lblDepthTol
        '
        Me.lblDepthTol.BackColor = System.Drawing.SystemColors.Control
        Me.lblDepthTol.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDepthTol.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDepthTol.Location = New System.Drawing.Point(34, 81)
        Me.lblDepthTol.Name = "lblDepthTol"
        Me.lblDepthTol.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDepthTol.Size = New System.Drawing.Size(109, 16)
        Me.lblDepthTol.TabIndex = 11
        Me.lblDepthTol.Text = "Anchor Depth:"
        '
        'lblAngleTol
        '
        Me.lblAngleTol.BackColor = System.Drawing.SystemColors.Control
        Me.lblAngleTol.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblAngleTol.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblAngleTol.Location = New System.Drawing.Point(34, 55)
        Me.lblAngleTol.Name = "lblAngleTol"
        Me.lblAngleTol.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblAngleTol.Size = New System.Drawing.Size(94, 16)
        Me.lblAngleTol.TabIndex = 10
        Me.lblAngleTol.Text = "Spread Angle:"
        '
        'lblDistTol
        '
        Me.lblDistTol.BackColor = System.Drawing.SystemColors.Control
        Me.lblDistTol.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDistTol.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDistTol.Location = New System.Drawing.Point(34, 29)
        Me.lblDistTol.Name = "lblDistTol"
        Me.lblDistTol.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDistTol.Size = New System.Drawing.Size(65, 16)
        Me.lblDistTol.TabIndex = 9
        Me.lblDistTol.Text = "Scope:"
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.SystemColors.Control
        Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCancel.Location = New System.Drawing.Point(219, 53)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCancel.Size = New System.Drawing.Size(70, 25)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'fraUnit
        '
        Me.fraUnit.BackColor = System.Drawing.SystemColors.Control
        Me.fraUnit.Controls.Add(Me.btnEnglish)
        Me.fraUnit.Controls.Add(Me.btnMetric)
        Me.fraUnit.ForeColor = System.Drawing.Color.Black
        Me.fraUnit.Location = New System.Drawing.Point(7, 8)
        Me.fraUnit.Name = "fraUnit"
        Me.fraUnit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraUnit.Size = New System.Drawing.Size(191, 79)
        Me.fraUnit.TabIndex = 0
        Me.fraUnit.TabStop = False
        Me.fraUnit.Text = "Unit System"
        '
        'btnEnglish
        '
        Me.btnEnglish.BackColor = System.Drawing.SystemColors.Control
        Me.btnEnglish.Checked = True
        Me.btnEnglish.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnEnglish.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optUnit.SetIndex(Me.btnEnglish, CType(1, Short))
        Me.btnEnglish.Location = New System.Drawing.Point(6, 19)
        Me.btnEnglish.Name = "btnEnglish"
        Me.btnEnglish.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnEnglish.Size = New System.Drawing.Size(146, 22)
        Me.btnEnglish.TabIndex = 1
        Me.btnEnglish.TabStop = True
        Me.btnEnglish.Text = "&English"
        Me.btnEnglish.UseVisualStyleBackColor = False
        '
        'btnMetric
        '
        Me.btnMetric.BackColor = System.Drawing.SystemColors.Control
        Me.btnMetric.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnMetric.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optUnit.SetIndex(Me.btnMetric, CType(0, Short))
        Me.btnMetric.Location = New System.Drawing.Point(6, 48)
        Me.btnMetric.Name = "btnMetric"
        Me.btnMetric.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnMetric.Size = New System.Drawing.Size(146, 22)
        Me.btnMetric.TabIndex = 2
        Me.btnMetric.TabStop = True
        Me.btnMetric.Text = "&Metric"
        Me.btnMetric.UseVisualStyleBackColor = False
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(16, 457)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(146, 16)
        Me.Label5.TabIndex = 38
        Me.Label5.Text = "Buoy Height:"
        '
        '_lblLengthUnit_4
        '
        Me._lblLengthUnit_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_4, CType(4, Short))
        Me._lblLengthUnit_4.Location = New System.Drawing.Point(256, 458)
        Me._lblLengthUnit_4.Name = "_lblLengthUnit_4"
        Me._lblLengthUnit_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_4.Size = New System.Drawing.Size(17, 16)
        Me._lblLengthUnit_4.TabIndex = 37
        Me._lblLengthUnit_4.Text = "ft"
        '
        'frmOptions
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(458, 90)
        Me.Controls.Add(Me.Text2)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.fraIntegration)
        Me.Controls.Add(Me.fraScopeRef)
        Me.Controls.Add(Me.fraTol)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.fraUnit)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me._lblLengthUnit_4)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Location = New System.Drawing.Point(4, 30)
        Me.Name = "frmOptions"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Options"
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.fraIntegration.ResumeLayout(False)
        Me.fraIntegration.PerformLayout()
        Me.fraScopeRef.ResumeLayout(False)
        Me.fraTol.ResumeLayout(False)
        Me.fraTol.PerformLayout()
        Me.fraUnit.ResumeLayout(False)
        CType(Me.lblLengthUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optScopeRef, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optUnit, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
End Class