<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmRiser
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
    Public WithEvents txtMudHeight As System.Windows.Forms.TextBox
	Public WithEvents txtMudWeight As System.Windows.Forms.TextBox
    Public WithEvents lblMudHeight As System.Windows.Forms.Label
    Public WithEvents _lblMudUnit_0 As System.Windows.Forms.Label
	Public WithEvents lblMudWeight As System.Windows.Forms.Label
	Public WithEvents fraMudData As System.Windows.Forms.GroupBox
	Public WithEvents txtTopLocation As System.Windows.Forms.TextBox
	Public WithEvents txtTopTension As System.Windows.Forms.TextBox
	Public WithEvents _lblAbvWL_1 As System.Windows.Forms.Label
	Public WithEvents lblTopLocation As System.Windows.Forms.Label
	Public WithEvents _lblRiserUnit_1 As System.Windows.Forms.Label
	Public WithEvents lblTopTension As System.Windows.Forms.Label
	Public WithEvents _lblRiserUnit_0 As System.Windows.Forms.Label
	Public WithEvents fraRiserData As System.Windows.Forms.GroupBox
    Public WithEvents btnCancel As System.Windows.Forms.Button
	Public WithEvents btnOK As System.Windows.Forms.Button
	Public WithEvents _btnRiserType_1 As System.Windows.Forms.RadioButton
	Public WithEvents _btnRiserType_0 As System.Windows.Forms.RadioButton
	Public WithEvents fraRiserType As System.Windows.Forms.GroupBox
	Public WithEvents _txtStiffness_1 As System.Windows.Forms.TextBox
	Public WithEvents _btnLowerBC_0 As System.Windows.Forms.RadioButton
	Public WithEvents _btnLowerBC_1 As System.Windows.Forms.RadioButton
	Public WithEvents _btnLowerBC_2 As System.Windows.Forms.RadioButton
	Public WithEvents _btnLowerBC_3 As System.Windows.Forms.RadioButton
	Public WithEvents _lblFJUnit_1 As System.Windows.Forms.Label
	Public WithEvents _lblStiffness_1 As System.Windows.Forms.Label
	Public WithEvents fraLowerBC As System.Windows.Forms.GroupBox
	Public WithEvents _txtStiffness_0 As System.Windows.Forms.TextBox
	Public WithEvents _btnUpperBC_3 As System.Windows.Forms.RadioButton
	Public WithEvents _btnUpperBC_2 As System.Windows.Forms.RadioButton
	Public WithEvents _btnUpperBC_1 As System.Windows.Forms.RadioButton
	Public WithEvents _btnUpperBC_0 As System.Windows.Forms.RadioButton
	Public WithEvents _lblFJUnit_0 As System.Windows.Forms.Label
	Public WithEvents _lblStiffness_0 As System.Windows.Forms.Label
	Public WithEvents fraUpperBC As System.Windows.Forms.GroupBox
	Public WithEvents btnLowerBC As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	Public WithEvents btnRiserType As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	Public WithEvents btnUpperBC As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	Public WithEvents lblAbvWL As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblFJUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblMudUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblRiserUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblStiffness As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents txtStiffness As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmRiser))
        Me.fraMudData = New System.Windows.Forms.GroupBox
        Me.txtMudHeight = New System.Windows.Forms.TextBox
        Me.txtMudWeight = New System.Windows.Forms.TextBox
        Me.lblMudHeight = New System.Windows.Forms.Label
        Me._lblMudUnit_0 = New System.Windows.Forms.Label
        Me.lblMudWeight = New System.Windows.Forms.Label
        Me.fraRiserData = New System.Windows.Forms.GroupBox
        Me.txtTopLocation = New System.Windows.Forms.TextBox
        Me.txtTopTension = New System.Windows.Forms.TextBox
        Me._lblAbvWL_1 = New System.Windows.Forms.Label
        Me.lblTopLocation = New System.Windows.Forms.Label
        Me._lblRiserUnit_1 = New System.Windows.Forms.Label
        Me.lblTopTension = New System.Windows.Forms.Label
        Me._lblRiserUnit_0 = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.fraRiserType = New System.Windows.Forms.GroupBox
        Me._btnRiserType_1 = New System.Windows.Forms.RadioButton
        Me._btnRiserType_0 = New System.Windows.Forms.RadioButton
        Me.fraLowerBC = New System.Windows.Forms.GroupBox
        Me._txtStiffness_1 = New System.Windows.Forms.TextBox
        Me._btnLowerBC_0 = New System.Windows.Forms.RadioButton
        Me._btnLowerBC_1 = New System.Windows.Forms.RadioButton
        Me._btnLowerBC_2 = New System.Windows.Forms.RadioButton
        Me._btnLowerBC_3 = New System.Windows.Forms.RadioButton
        Me._lblFJUnit_1 = New System.Windows.Forms.Label
        Me._lblStiffness_1 = New System.Windows.Forms.Label
        Me.fraUpperBC = New System.Windows.Forms.GroupBox
        Me._txtStiffness_0 = New System.Windows.Forms.TextBox
        Me._btnUpperBC_3 = New System.Windows.Forms.RadioButton
        Me._btnUpperBC_2 = New System.Windows.Forms.RadioButton
        Me._btnUpperBC_1 = New System.Windows.Forms.RadioButton
        Me._btnUpperBC_0 = New System.Windows.Forms.RadioButton
        Me._lblFJUnit_0 = New System.Windows.Forms.Label
        Me._lblStiffness_0 = New System.Windows.Forms.Label
        Me.btnLowerBC = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.btnRiserType = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.btnUpperBC = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.lblAbvWL = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblFJUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblMudUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblRiserUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblStiffness = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.txtStiffness = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.fraMudData.SuspendLayout()
        Me.fraRiserData.SuspendLayout()
        Me.fraRiserType.SuspendLayout()
        Me.fraLowerBC.SuspendLayout()
        Me.fraUpperBC.SuspendLayout()
        CType(Me.btnLowerBC, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnRiserType, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnUpperBC, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblAbvWL, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblFJUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblMudUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblRiserUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblStiffness, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtStiffness, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'fraMudData
        '
        Me.fraMudData.BackColor = System.Drawing.SystemColors.Control
        Me.fraMudData.Controls.Add(Me.txtMudHeight)
        Me.fraMudData.Controls.Add(Me.txtMudWeight)
        Me.fraMudData.Controls.Add(Me.lblMudHeight)
        Me.fraMudData.Controls.Add(Me._lblMudUnit_0)
        Me.fraMudData.Controls.Add(Me.lblMudWeight)
        Me.fraMudData.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraMudData.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraMudData.Location = New System.Drawing.Point(224, 96)
        Me.fraMudData.Name = "fraMudData"
        Me.fraMudData.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraMudData.Size = New System.Drawing.Size(201, 97)
        Me.fraMudData.TabIndex = 56
        Me.fraMudData.TabStop = False
        Me.fraMudData.Text = "Mud Properties"
        '
        'txtMudHeight
        '
        Me.txtMudHeight.AcceptsReturn = True
        Me.txtMudHeight.BackColor = System.Drawing.SystemColors.Window
        Me.txtMudHeight.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMudHeight.Enabled = False
        Me.txtMudHeight.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMudHeight.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMudHeight.Location = New System.Drawing.Point(87, 50)
        Me.txtMudHeight.MaxLength = 0
        Me.txtMudHeight.Name = "txtMudHeight"
        Me.txtMudHeight.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMudHeight.Size = New System.Drawing.Size(66, 20)
        Me.txtMudHeight.TabIndex = 5
        Me.txtMudHeight.Text = "Text1"
        Me.txtMudHeight.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtMudHeight.Visible = False
        '
        'txtMudWeight
        '
        Me.txtMudWeight.AcceptsReturn = True
        Me.txtMudWeight.BackColor = System.Drawing.SystemColors.Window
        Me.txtMudWeight.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMudWeight.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMudWeight.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMudWeight.Location = New System.Drawing.Point(87, 24)
        Me.txtMudWeight.MaxLength = 0
        Me.txtMudWeight.Name = "txtMudWeight"
        Me.txtMudWeight.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMudWeight.Size = New System.Drawing.Size(66, 20)
        Me.txtMudWeight.TabIndex = 4
        Me.txtMudWeight.Text = "Text1"
        Me.txtMudWeight.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblMudHeight
        '
        Me.lblMudHeight.BackColor = System.Drawing.SystemColors.Control
        Me.lblMudHeight.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMudHeight.Enabled = False
        Me.lblMudHeight.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMudHeight.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblMudHeight.Location = New System.Drawing.Point(16, 47)
        Me.lblMudHeight.Name = "lblMudHeight"
        Me.lblMudHeight.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMudHeight.Size = New System.Drawing.Size(65, 38)
        Me.lblMudHeight.TabIndex = 60
        Me.lblMudHeight.Text = "Damping " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Ratio (Zeta)"
        Me.lblMudHeight.Visible = False
        '
        '_lblMudUnit_0
        '
        Me._lblMudUnit_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblMudUnit_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblMudUnit_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblMudUnit_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblMudUnit.SetIndex(Me._lblMudUnit_0, CType(0, Short))
        Me._lblMudUnit_0.Location = New System.Drawing.Point(157, 25)
        Me._lblMudUnit_0.Name = "_lblMudUnit_0"
        Me._lblMudUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblMudUnit_0.Size = New System.Drawing.Size(41, 17)
        Me._lblMudUnit_0.TabIndex = 58
        Me._lblMudUnit_0.Text = "kg/m^3"
        '
        'lblMudWeight
        '
        Me.lblMudWeight.BackColor = System.Drawing.SystemColors.Control
        Me.lblMudWeight.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMudWeight.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMudWeight.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblMudWeight.Location = New System.Drawing.Point(16, 24)
        Me.lblMudWeight.Name = "lblMudWeight"
        Me.lblMudWeight.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMudWeight.Size = New System.Drawing.Size(75, 19)
        Me.lblMudWeight.TabIndex = 57
        Me.lblMudWeight.Text = "Mud Density"
        '
        'fraRiserData
        '
        Me.fraRiserData.BackColor = System.Drawing.SystemColors.Control
        Me.fraRiserData.Controls.Add(Me.txtTopLocation)
        Me.fraRiserData.Controls.Add(Me.txtTopTension)
        Me.fraRiserData.Controls.Add(Me._lblAbvWL_1)
        Me.fraRiserData.Controls.Add(Me.lblTopLocation)
        Me.fraRiserData.Controls.Add(Me._lblRiserUnit_1)
        Me.fraRiserData.Controls.Add(Me.lblTopTension)
        Me.fraRiserData.Controls.Add(Me._lblRiserUnit_0)
        Me.fraRiserData.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraRiserData.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraRiserData.Location = New System.Drawing.Point(8, 96)
        Me.fraRiserData.Name = "fraRiserData"
        Me.fraRiserData.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraRiserData.Size = New System.Drawing.Size(201, 97)
        Me.fraRiserData.TabIndex = 49
        Me.fraRiserData.TabStop = False
        Me.fraRiserData.Text = "Riser Properties"
        '
        'txtTopLocation
        '
        Me.txtTopLocation.AcceptsReturn = True
        Me.txtTopLocation.BackColor = System.Drawing.SystemColors.Window
        Me.txtTopLocation.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTopLocation.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTopLocation.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTopLocation.Location = New System.Drawing.Point(95, 46)
        Me.txtTopLocation.MaxLength = 0
        Me.txtTopLocation.Name = "txtTopLocation"
        Me.txtTopLocation.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTopLocation.Size = New System.Drawing.Size(66, 20)
        Me.txtTopLocation.TabIndex = 3
        Me.txtTopLocation.Text = "Text1"
        Me.txtTopLocation.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtTopTension
        '
        Me.txtTopTension.AcceptsReturn = True
        Me.txtTopTension.BackColor = System.Drawing.SystemColors.Window
        Me.txtTopTension.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTopTension.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTopTension.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTopTension.Location = New System.Drawing.Point(95, 24)
        Me.txtTopTension.MaxLength = 0
        Me.txtTopTension.Name = "txtTopTension"
        Me.txtTopTension.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTopTension.Size = New System.Drawing.Size(66, 20)
        Me.txtTopTension.TabIndex = 2
        Me.txtTopTension.Text = "Text1"
        Me.txtTopTension.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_lblAbvWL_1
        '
        Me._lblAbvWL_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblAbvWL_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblAbvWL_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblAbvWL_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblAbvWL.SetIndex(Me._lblAbvWL_1, CType(1, Short))
        Me._lblAbvWL_1.Location = New System.Drawing.Point(95, 68)
        Me._lblAbvWL_1.Name = "_lblAbvWL_1"
        Me._lblAbvWL_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblAbvWL_1.Size = New System.Drawing.Size(90, 17)
        Me._lblAbvWL_1.TabIndex = 64
        Me._lblAbvWL_1.Text = "above waterline"
        '
        'lblTopLocation
        '
        Me.lblTopLocation.BackColor = System.Drawing.SystemColors.Control
        Me.lblTopLocation.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTopLocation.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTopLocation.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTopLocation.Location = New System.Drawing.Point(16, 47)
        Me.lblTopLocation.Name = "lblTopLocation"
        Me.lblTopLocation.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTopLocation.Size = New System.Drawing.Size(73, 16)
        Me.lblTopLocation.TabIndex = 63
        Me.lblTopLocation.Text = "Top Location"
        '
        '_lblRiserUnit_1
        '
        Me._lblRiserUnit_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblRiserUnit_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRiserUnit_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblRiserUnit_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRiserUnit.SetIndex(Me._lblRiserUnit_1, CType(1, Short))
        Me._lblRiserUnit_1.Location = New System.Drawing.Point(165, 47)
        Me._lblRiserUnit_1.Name = "_lblRiserUnit_1"
        Me._lblRiserUnit_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRiserUnit_1.Size = New System.Drawing.Size(33, 17)
        Me._lblRiserUnit_1.TabIndex = 62
        Me._lblRiserUnit_1.Text = "m"
        '
        'lblTopTension
        '
        Me.lblTopTension.BackColor = System.Drawing.SystemColors.Control
        Me.lblTopTension.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTopTension.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTopTension.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTopTension.Location = New System.Drawing.Point(16, 24)
        Me.lblTopTension.Name = "lblTopTension"
        Me.lblTopTension.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTopTension.Size = New System.Drawing.Size(66, 19)
        Me.lblTopTension.TabIndex = 55
        Me.lblTopTension.Text = "Top Tension"
        '
        '_lblRiserUnit_0
        '
        Me._lblRiserUnit_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblRiserUnit_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRiserUnit_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblRiserUnit_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRiserUnit.SetIndex(Me._lblRiserUnit_0, CType(0, Short))
        Me._lblRiserUnit_0.Location = New System.Drawing.Point(165, 25)
        Me._lblRiserUnit_0.Name = "_lblRiserUnit_0"
        Me._lblRiserUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRiserUnit_0.Size = New System.Drawing.Size(17, 17)
        Me._lblRiserUnit_0.TabIndex = 54
        Me._lblRiserUnit_0.Text = "N"
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.SystemColors.Control
        Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCancel.Location = New System.Drawing.Point(360, 48)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCancel.Size = New System.Drawing.Size(65, 25)
        Me.btnCancel.TabIndex = 25
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'btnOK
        '
        Me.btnOK.BackColor = System.Drawing.SystemColors.Control
        Me.btnOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnOK.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnOK.Location = New System.Drawing.Point(357, 12)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnOK.Size = New System.Drawing.Size(65, 25)
        Me.btnOK.TabIndex = 24
        Me.btnOK.Text = "&Save"
        Me.btnOK.UseVisualStyleBackColor = False
        '
        'fraRiserType
        '
        Me.fraRiserType.BackColor = System.Drawing.SystemColors.Control
        Me.fraRiserType.Controls.Add(Me._btnRiserType_1)
        Me.fraRiserType.Controls.Add(Me._btnRiserType_0)
        Me.fraRiserType.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraRiserType.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraRiserType.Location = New System.Drawing.Point(8, 8)
        Me.fraRiserType.Name = "fraRiserType"
        Me.fraRiserType.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraRiserType.Size = New System.Drawing.Size(201, 81)
        Me.fraRiserType.TabIndex = 28
        Me.fraRiserType.TabStop = False
        Me.fraRiserType.Text = "Riser Analysis Type"
        '
        '_btnRiserType_1
        '
        Me._btnRiserType_1.BackColor = System.Drawing.SystemColors.Control
        Me._btnRiserType_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnRiserType_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._btnRiserType_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnRiserType.SetIndex(Me._btnRiserType_1, CType(1, Short))
        Me._btnRiserType_1.Location = New System.Drawing.Point(16, 48)
        Me._btnRiserType_1.Name = "_btnRiserType_1"
        Me._btnRiserType_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnRiserType_1.Size = New System.Drawing.Size(147, 17)
        Me._btnRiserType_1.TabIndex = 1
        Me._btnRiserType_1.TabStop = True
        Me._btnRiserType_1.Text = "SCR or Lazy-Wave Riser"
        Me._btnRiserType_1.UseVisualStyleBackColor = False
        '
        '_btnRiserType_0
        '
        Me._btnRiserType_0.BackColor = System.Drawing.SystemColors.Control
        Me._btnRiserType_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnRiserType_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._btnRiserType_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnRiserType.SetIndex(Me._btnRiserType_0, CType(0, Short))
        Me._btnRiserType_0.Location = New System.Drawing.Point(16, 24)
        Me._btnRiserType_0.Name = "_btnRiserType_0"
        Me._btnRiserType_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnRiserType_0.Size = New System.Drawing.Size(145, 18)
        Me._btnRiserType_0.TabIndex = 0
        Me._btnRiserType_0.TabStop = True
        Me._btnRiserType_0.Text = "Rigid Riser Analysis"
        Me._btnRiserType_0.UseVisualStyleBackColor = False
        '
        'fraLowerBC
        '
        Me.fraLowerBC.BackColor = System.Drawing.SystemColors.Control
        Me.fraLowerBC.Controls.Add(Me._txtStiffness_1)
        Me.fraLowerBC.Controls.Add(Me._btnLowerBC_0)
        Me.fraLowerBC.Controls.Add(Me._btnLowerBC_1)
        Me.fraLowerBC.Controls.Add(Me._btnLowerBC_2)
        Me.fraLowerBC.Controls.Add(Me._btnLowerBC_3)
        Me.fraLowerBC.Controls.Add(Me._lblFJUnit_1)
        Me.fraLowerBC.Controls.Add(Me._lblStiffness_1)
        Me.fraLowerBC.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraLowerBC.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraLowerBC.Location = New System.Drawing.Point(224, 200)
        Me.fraLowerBC.Name = "fraLowerBC"
        Me.fraLowerBC.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraLowerBC.Size = New System.Drawing.Size(201, 153)
        Me.fraLowerBC.TabIndex = 27
        Me.fraLowerBC.TabStop = False
        Me.fraLowerBC.Text = "Lower Boundary Condition"
        '
        '_txtStiffness_1
        '
        Me._txtStiffness_1.AcceptsReturn = True
        Me._txtStiffness_1.BackColor = System.Drawing.SystemColors.Window
        Me._txtStiffness_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtStiffness_1.Enabled = False
        Me._txtStiffness_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtStiffness_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtStiffness.SetIndex(Me._txtStiffness_1, CType(1, Short))
        Me._txtStiffness_1.Location = New System.Drawing.Point(65, 119)
        Me._txtStiffness_1.MaxLength = 0
        Me._txtStiffness_1.Name = "_txtStiffness_1"
        Me._txtStiffness_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtStiffness_1.Size = New System.Drawing.Size(73, 20)
        Me._txtStiffness_1.TabIndex = 15
        Me._txtStiffness_1.Text = "Text1"
        Me._txtStiffness_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_btnLowerBC_0
        '
        Me._btnLowerBC_0.BackColor = System.Drawing.SystemColors.Control
        Me._btnLowerBC_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnLowerBC_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._btnLowerBC_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnLowerBC.SetIndex(Me._btnLowerBC_0, CType(0, Short))
        Me._btnLowerBC_0.Location = New System.Drawing.Point(16, 24)
        Me._btnLowerBC_0.Name = "_btnLowerBC_0"
        Me._btnLowerBC_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnLowerBC_0.Size = New System.Drawing.Size(145, 16)
        Me._btnLowerBC_0.TabIndex = 11
        Me._btnLowerBC_0.TabStop = True
        Me._btnLowerBC_0.Text = "Pinned Connection"
        Me._btnLowerBC_0.UseVisualStyleBackColor = False
        '
        '_btnLowerBC_1
        '
        Me._btnLowerBC_1.BackColor = System.Drawing.SystemColors.Control
        Me._btnLowerBC_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnLowerBC_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._btnLowerBC_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnLowerBC.SetIndex(Me._btnLowerBC_1, CType(1, Short))
        Me._btnLowerBC_1.Location = New System.Drawing.Point(16, 48)
        Me._btnLowerBC_1.Name = "_btnLowerBC_1"
        Me._btnLowerBC_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnLowerBC_1.Size = New System.Drawing.Size(145, 16)
        Me._btnLowerBC_1.TabIndex = 12
        Me._btnLowerBC_1.TabStop = True
        Me._btnLowerBC_1.Text = "Fixed End"
        Me._btnLowerBC_1.UseVisualStyleBackColor = False
        '
        '_btnLowerBC_2
        '
        Me._btnLowerBC_2.BackColor = System.Drawing.SystemColors.Control
        Me._btnLowerBC_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnLowerBC_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._btnLowerBC_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnLowerBC.SetIndex(Me._btnLowerBC_2, CType(2, Short))
        Me._btnLowerBC_2.Location = New System.Drawing.Point(16, 72)
        Me._btnLowerBC_2.Name = "_btnLowerBC_2"
        Me._btnLowerBC_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnLowerBC_2.Size = New System.Drawing.Size(145, 16)
        Me._btnLowerBC_2.TabIndex = 13
        Me._btnLowerBC_2.TabStop = True
        Me._btnLowerBC_2.Text = "Flex Joint"
        Me._btnLowerBC_2.UseVisualStyleBackColor = False
        '
        '_btnLowerBC_3
        '
        Me._btnLowerBC_3.BackColor = System.Drawing.SystemColors.Control
        Me._btnLowerBC_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnLowerBC_3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._btnLowerBC_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnLowerBC.SetIndex(Me._btnLowerBC_3, CType(3, Short))
        Me._btnLowerBC_3.Location = New System.Drawing.Point(16, 96)
        Me._btnLowerBC_3.Name = "_btnLowerBC_3"
        Me._btnLowerBC_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnLowerBC_3.Size = New System.Drawing.Size(145, 16)
        Me._btnLowerBC_3.TabIndex = 14
        Me._btnLowerBC_3.TabStop = True
        Me._btnLowerBC_3.Text = "Free End (Disconnected)"
        Me._btnLowerBC_3.UseVisualStyleBackColor = False
        '
        '_lblFJUnit_1
        '
        Me._lblFJUnit_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblFJUnit_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblFJUnit_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblFJUnit_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFJUnit.SetIndex(Me._lblFJUnit_1, CType(1, Short))
        Me._lblFJUnit_1.Location = New System.Drawing.Point(144, 120)
        Me._lblFJUnit_1.Name = "_lblFJUnit_1"
        Me._lblFJUnit_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblFJUnit_1.Size = New System.Drawing.Size(49, 17)
        Me._lblFJUnit_1.TabIndex = 53
        Me._lblFJUnit_1.Text = "N-m/deg"
        '
        '_lblStiffness_1
        '
        Me._lblStiffness_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblStiffness_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblStiffness_1.Enabled = False
        Me._lblStiffness_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblStiffness_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblStiffness.SetIndex(Me._lblStiffness_1, CType(1, Short))
        Me._lblStiffness_1.Location = New System.Drawing.Point(16, 120)
        Me._lblStiffness_1.Name = "_lblStiffness_1"
        Me._lblStiffness_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblStiffness_1.Size = New System.Drawing.Size(49, 17)
        Me._lblStiffness_1.TabIndex = 52
        Me._lblStiffness_1.Text = "Stiffness"
        '
        'fraUpperBC
        '
        Me.fraUpperBC.BackColor = System.Drawing.SystemColors.Control
        Me.fraUpperBC.Controls.Add(Me._txtStiffness_0)
        Me.fraUpperBC.Controls.Add(Me._btnUpperBC_3)
        Me.fraUpperBC.Controls.Add(Me._btnUpperBC_2)
        Me.fraUpperBC.Controls.Add(Me._btnUpperBC_1)
        Me.fraUpperBC.Controls.Add(Me._btnUpperBC_0)
        Me.fraUpperBC.Controls.Add(Me._lblFJUnit_0)
        Me.fraUpperBC.Controls.Add(Me._lblStiffness_0)
        Me.fraUpperBC.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraUpperBC.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraUpperBC.Location = New System.Drawing.Point(8, 200)
        Me.fraUpperBC.Name = "fraUpperBC"
        Me.fraUpperBC.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraUpperBC.Size = New System.Drawing.Size(201, 153)
        Me.fraUpperBC.TabIndex = 26
        Me.fraUpperBC.TabStop = False
        Me.fraUpperBC.Text = "Upper Boundary Condition"
        '
        '_txtStiffness_0
        '
        Me._txtStiffness_0.AcceptsReturn = True
        Me._txtStiffness_0.BackColor = System.Drawing.SystemColors.Window
        Me._txtStiffness_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtStiffness_0.Enabled = False
        Me._txtStiffness_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtStiffness_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtStiffness.SetIndex(Me._txtStiffness_0, CType(0, Short))
        Me._txtStiffness_0.Location = New System.Drawing.Point(64, 119)
        Me._txtStiffness_0.MaxLength = 0
        Me._txtStiffness_0.Name = "_txtStiffness_0"
        Me._txtStiffness_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtStiffness_0.Size = New System.Drawing.Size(73, 20)
        Me._txtStiffness_0.TabIndex = 10
        Me._txtStiffness_0.Text = "Text1"
        Me._txtStiffness_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_btnUpperBC_3
        '
        Me._btnUpperBC_3.BackColor = System.Drawing.SystemColors.Control
        Me._btnUpperBC_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnUpperBC_3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._btnUpperBC_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnUpperBC.SetIndex(Me._btnUpperBC_3, CType(3, Short))
        Me._btnUpperBC_3.Location = New System.Drawing.Point(16, 98)
        Me._btnUpperBC_3.Name = "_btnUpperBC_3"
        Me._btnUpperBC_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnUpperBC_3.Size = New System.Drawing.Size(145, 16)
        Me._btnUpperBC_3.TabIndex = 9
        Me._btnUpperBC_3.TabStop = True
        Me._btnUpperBC_3.Text = "Free End (Free Standing)"
        Me._btnUpperBC_3.UseVisualStyleBackColor = False
        '
        '_btnUpperBC_2
        '
        Me._btnUpperBC_2.BackColor = System.Drawing.SystemColors.Control
        Me._btnUpperBC_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnUpperBC_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._btnUpperBC_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnUpperBC.SetIndex(Me._btnUpperBC_2, CType(2, Short))
        Me._btnUpperBC_2.Location = New System.Drawing.Point(16, 68)
        Me._btnUpperBC_2.Name = "_btnUpperBC_2"
        Me._btnUpperBC_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnUpperBC_2.Size = New System.Drawing.Size(145, 24)
        Me._btnUpperBC_2.TabIndex = 8
        Me._btnUpperBC_2.TabStop = True
        Me._btnUpperBC_2.Text = "Flex Joint"
        Me._btnUpperBC_2.UseVisualStyleBackColor = False
        '
        '_btnUpperBC_1
        '
        Me._btnUpperBC_1.BackColor = System.Drawing.SystemColors.Control
        Me._btnUpperBC_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnUpperBC_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._btnUpperBC_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnUpperBC.SetIndex(Me._btnUpperBC_1, CType(1, Short))
        Me._btnUpperBC_1.Location = New System.Drawing.Point(16, 46)
        Me._btnUpperBC_1.Name = "_btnUpperBC_1"
        Me._btnUpperBC_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnUpperBC_1.Size = New System.Drawing.Size(145, 16)
        Me._btnUpperBC_1.TabIndex = 7
        Me._btnUpperBC_1.TabStop = True
        Me._btnUpperBC_1.Text = "Fixed End"
        Me._btnUpperBC_1.UseVisualStyleBackColor = False
        '
        '_btnUpperBC_0
        '
        Me._btnUpperBC_0.BackColor = System.Drawing.SystemColors.Control
        Me._btnUpperBC_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnUpperBC_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._btnUpperBC_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnUpperBC.SetIndex(Me._btnUpperBC_0, CType(0, Short))
        Me._btnUpperBC_0.Location = New System.Drawing.Point(16, 24)
        Me._btnUpperBC_0.Name = "_btnUpperBC_0"
        Me._btnUpperBC_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnUpperBC_0.Size = New System.Drawing.Size(145, 16)
        Me._btnUpperBC_0.TabIndex = 6
        Me._btnUpperBC_0.TabStop = True
        Me._btnUpperBC_0.Text = "Pinned Connection"
        Me._btnUpperBC_0.UseVisualStyleBackColor = False
        '
        '_lblFJUnit_0
        '
        Me._lblFJUnit_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblFJUnit_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblFJUnit_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblFJUnit_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFJUnit.SetIndex(Me._lblFJUnit_0, CType(0, Short))
        Me._lblFJUnit_0.Location = New System.Drawing.Point(144, 120)
        Me._lblFJUnit_0.Name = "_lblFJUnit_0"
        Me._lblFJUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblFJUnit_0.Size = New System.Drawing.Size(49, 17)
        Me._lblFJUnit_0.TabIndex = 51
        Me._lblFJUnit_0.Text = "N-m/deg"
        '
        '_lblStiffness_0
        '
        Me._lblStiffness_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblStiffness_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblStiffness_0.Enabled = False
        Me._lblStiffness_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblStiffness_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblStiffness.SetIndex(Me._lblStiffness_0, CType(0, Short))
        Me._lblStiffness_0.Location = New System.Drawing.Point(16, 120)
        Me._lblStiffness_0.Name = "_lblStiffness_0"
        Me._lblStiffness_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblStiffness_0.Size = New System.Drawing.Size(41, 17)
        Me._lblStiffness_0.TabIndex = 50
        Me._lblStiffness_0.Text = "Stiffness"
        '
        'btnLowerBC
        '
        '
        'btnUpperBC
        '
        '
        'frmRiser
        '
        Me.AcceptButton = Me.btnOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(434, 380)
        Me.Controls.Add(Me.fraMudData)
        Me.Controls.Add(Me.fraRiserData)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.fraRiserType)
        Me.Controls.Add(Me.fraLowerBC)
        Me.Controls.Add(Me.fraUpperBC)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.Name = "frmRiser"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Riser Type and Boundary Conditions"
        Me.fraMudData.ResumeLayout(False)
        Me.fraMudData.PerformLayout()
        Me.fraRiserData.ResumeLayout(False)
        Me.fraRiserData.PerformLayout()
        Me.fraRiserType.ResumeLayout(False)
        Me.fraLowerBC.ResumeLayout(False)
        Me.fraLowerBC.PerformLayout()
        Me.fraUpperBC.ResumeLayout(False)
        Me.fraUpperBC.PerformLayout()
        CType(Me.btnLowerBC, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnRiserType, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnUpperBC, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblAbvWL, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblFJUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblMudUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblRiserUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblStiffness, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtStiffness, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
#End Region
End Class