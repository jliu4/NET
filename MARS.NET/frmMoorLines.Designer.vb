<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmMoorLines
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
	Public WithEvents _txtVslSt_3 As System.Windows.Forms.TextBox
	Public WithEvents cboDraft As System.Windows.Forms.ComboBox
	Public WithEvents _lblLengthUnit_2 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents txtWD As System.Windows.Forms.TextBox
	Public WithEvents _txtVslSt_2 As System.Windows.Forms.TextBox
	Public WithEvents _txtVslSt_1 As System.Windows.Forms.TextBox
	Public WithEvents _txtVslSt_0 As System.Windows.Forms.TextBox
	Public WithEvents _lblVslSt_3 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_11 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_1 As System.Windows.Forms.Label
    Public WithEvents _lblVslSt_2 As System.Windows.Forms.Label
    Public WithEvents _lblVslSt_1 As System.Windows.Forms.Label
    Public WithEvents _lblVslSt_0 As System.Windows.Forms.Label
    Public WithEvents _lblLengthUnit_0 As System.Windows.Forms.Label
    Public WithEvents fraVesselLoc As System.Windows.Forms.GroupBox
    Public WithEvents btnCatenary As System.Windows.Forms.Button
    Public WithEvents _lblSegCmt_0 As System.Windows.Forms.Label
    Public WithEvents _lblSegCmt_1 As System.Windows.Forms.Label
    Public WithEvents fraSegments As System.Windows.Forms.GroupBox
    Public WithEvents btnPayout As System.Windows.Forms.Button
    Public WithEvents btnScope As System.Windows.Forms.Button
    Public WithEvents _txtGeneral_2 As System.Windows.Forms.TextBox
    Public WithEvents _txtGeneral_1 As System.Windows.Forms.TextBox
    Public WithEvents _txtGeneral_0 As System.Windows.Forms.TextBox
    Public WithEvents _lblLengthUnit_4 As System.Windows.Forms.Label
    Public WithEvents _lblLengthUnit_3 As System.Windows.Forms.Label
    Public WithEvents _lblGenCmt_0 As System.Windows.Forms.Label
    Public WithEvents _lblGenCmt_1 As System.Windows.Forms.Label
    Public WithEvents _lblGeneral_2 As System.Windows.Forms.Label
    Public WithEvents _lblGeneral_1 As System.Windows.Forms.Label
    Public WithEvents _lblGeneral_0 As System.Windows.Forms.Label
    Public WithEvents fraGeneral As System.Windows.Forms.GroupBox
    Public WithEvents chkBatch As System.Windows.Forms.CheckBox
    Public WithEvents btnCancel As System.Windows.Forms.Button
    Public WithEvents btnNew As System.Windows.Forms.Button
    Public WithEvents btnRemove As System.Windows.Forms.Button
    Public WithEvents btnSave As System.Windows.Forms.Button
    Public WithEvents btnTopTension As System.Windows.Forms.Button
    Public WithEvents btnPreten As System.Windows.Forms.Button
    Public WithEvents _txtPreTen_1 As System.Windows.Forms.TextBox
    Public WithEvents txtTopTen As System.Windows.Forms.TextBox
    Public WithEvents _txtPreTen_0 As System.Windows.Forms.TextBox
    Public WithEvents _lblForceUnit_2 As System.Windows.Forms.Label
    Public WithEvents _lblForceUnit_1 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents _lblPreTen_1 As System.Windows.Forms.Label
    Public WithEvents lblTopTen As System.Windows.Forms.Label
    Public WithEvents _lblForceUnit_0 As System.Windows.Forms.Label
    Public WithEvents _lblPreTen_0 As System.Windows.Forms.Label
    Public WithEvents fraTopTen As System.Windows.Forms.GroupBox
    Public WithEvents btnAnchor As System.Windows.Forms.Button
    Public WithEvents _txtAnchor_3 As System.Windows.Forms.TextBox
    Public WithEvents _txtAnchor_2 As System.Windows.Forms.TextBox
    Public WithEvents _txtAnchor_1 As System.Windows.Forms.TextBox
    Public WithEvents _txtAnchor_0 As System.Windows.Forms.TextBox
    Public WithEvents _lblAnchorCmt_3 As System.Windows.Forms.Label
    Public WithEvents _lblAnchorCmt_2 As System.Windows.Forms.Label
    Public WithEvents _lblLengthUnit_10 As System.Windows.Forms.Label
    Public WithEvents _lblLengthUnit_9 As System.Windows.Forms.Label
    Public WithEvents _lblLengthUnit_8 As System.Windows.Forms.Label
    Public WithEvents _lblAnchor_3 As System.Windows.Forms.Label
    Public WithEvents _lblAnchorCmt_1 As System.Windows.Forms.Label
    Public WithEvents _lblAnchorCmt_0 As System.Windows.Forms.Label
    Public WithEvents _lblAnchor_2 As System.Windows.Forms.Label
    Public WithEvents _lblAnchor_1 As System.Windows.Forms.Label
    Public WithEvents _lblAnchor_0 As System.Windows.Forms.Label
    Public WithEvents fraAnchor As System.Windows.Forms.GroupBox
    Public WithEvents _txtFL_2 As System.Windows.Forms.TextBox
    Public WithEvents _txtFL_1 As System.Windows.Forms.TextBox
    Public WithEvents _txtFL_0 As System.Windows.Forms.TextBox
    Public WithEvents _lblLengthUnit_7 As System.Windows.Forms.Label
    Public WithEvents _lblLengthUnit_6 As System.Windows.Forms.Label
    Public WithEvents _lblLengthUnit_5 As System.Windows.Forms.Label
    Public WithEvents _lblFL_2 As System.Windows.Forms.Label
    Public WithEvents _lblFLUnit_2 As System.Windows.Forms.Label
    Public WithEvents _lblFLCmt_1 As System.Windows.Forms.Label
    Public WithEvents _lblFLCmt_0 As System.Windows.Forms.Label
    Public WithEvents _lblFLUnit_1 As System.Windows.Forms.Label
    Public WithEvents _lblFL_1 As System.Windows.Forms.Label
    Public WithEvents _lblFLUnit_0 As System.Windows.Forms.Label
    Public WithEvents _lblFL_0 As System.Windows.Forms.Label
    Public WithEvents fraFairLead As System.Windows.Forms.GroupBox
    Public dlgFileOpen As System.Windows.Forms.OpenFileDialog
    Public dlgFileSave As System.Windows.Forms.SaveFileDialog
    Public WithEvents lblAnchor As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblAnchorCmt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblAngleUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblFL As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblFLCmt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblFLUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblForceUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblGenCmt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblGeneral As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblLengthUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblPreTen As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblSegCmt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblVelUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblVslSt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents txtAnchor As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    Public WithEvents txtFL As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    Public WithEvents txtGeneral As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    Public WithEvents txtPreTen As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    Public WithEvents txtVslSt As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    Public WithEvents mnuOpen As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuSave As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuDum As System.Windows.Forms.ToolStripSeparator
    Public WithEvents mnuClose As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFile As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuInsert As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuDelete As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuGridEdit As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMoorLines))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnPayout = New System.Windows.Forms.Button()
        Me.btnScope = New System.Windows.Forms.Button()
        Me._txtGeneral_2 = New System.Windows.Forms.TextBox()
        Me.btnTopTension = New System.Windows.Forms.Button()
        Me.btnPreten = New System.Windows.Forms.Button()
        Me.btnAnchor = New System.Windows.Forms.Button()
        Me._txtAnchor_3 = New System.Windows.Forms.TextBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me._txtVslSt_3 = New System.Windows.Forms.TextBox()
        Me.cboDraft = New System.Windows.Forms.ComboBox()
        Me._lblLengthUnit_2 = New System.Windows.Forms.Label()
        Me.fraVesselLoc = New System.Windows.Forms.GroupBox()
        Me.txtWD = New System.Windows.Forms.TextBox()
        Me._txtVslSt_2 = New System.Windows.Forms.TextBox()
        Me._txtVslSt_1 = New System.Windows.Forms.TextBox()
        Me._txtVslSt_0 = New System.Windows.Forms.TextBox()
        Me._lblVslSt_3 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_11 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_1 = New System.Windows.Forms.Label()
        Me._lblVslSt_2 = New System.Windows.Forms.Label()
        Me._lblVslSt_1 = New System.Windows.Forms.Label()
        Me._lblVslSt_0 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_0 = New System.Windows.Forms.Label()
        Me.btnCatenary = New System.Windows.Forms.Button()
        Me.fraSegments = New System.Windows.Forms.GroupBox()
        Me.grdAnchor = New System.Windows.Forms.DataGridView()
        Me.grdSegments = New System.Windows.Forms.DataGridView()
        Me._lblSegCmt_0 = New System.Windows.Forms.Label()
        Me._lblSegCmt_1 = New System.Windows.Forms.Label()
        Me.fraGeneral = New System.Windows.Forms.GroupBox()
        Me._txtGeneral_1 = New System.Windows.Forms.TextBox()
        Me._txtGeneral_0 = New System.Windows.Forms.TextBox()
        Me._lblLengthUnit_4 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_3 = New System.Windows.Forms.Label()
        Me._lblGenCmt_0 = New System.Windows.Forms.Label()
        Me._lblGenCmt_1 = New System.Windows.Forms.Label()
        Me._lblGeneral_2 = New System.Windows.Forms.Label()
        Me._lblGeneral_1 = New System.Windows.Forms.Label()
        Me._lblGeneral_0 = New System.Windows.Forms.Label()
        Me.chkBatch = New System.Windows.Forms.CheckBox()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnNew = New System.Windows.Forms.Button()
        Me.btnRemove = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.fraTopTen = New System.Windows.Forms.GroupBox()
        Me._txtPreTen_1 = New System.Windows.Forms.TextBox()
        Me.txtTopTen = New System.Windows.Forms.TextBox()
        Me._txtPreTen_0 = New System.Windows.Forms.TextBox()
        Me._lblForceUnit_2 = New System.Windows.Forms.Label()
        Me._lblForceUnit_1 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me._lblPreTen_1 = New System.Windows.Forms.Label()
        Me.lblTopTen = New System.Windows.Forms.Label()
        Me._lblForceUnit_0 = New System.Windows.Forms.Label()
        Me._lblPreTen_0 = New System.Windows.Forms.Label()
        Me.fraAnchor = New System.Windows.Forms.GroupBox()
        Me._txtAnchor_2 = New System.Windows.Forms.TextBox()
        Me._txtAnchor_1 = New System.Windows.Forms.TextBox()
        Me._txtAnchor_0 = New System.Windows.Forms.TextBox()
        Me._lblAnchorCmt_3 = New System.Windows.Forms.Label()
        Me._lblAnchorCmt_2 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_10 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_9 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_8 = New System.Windows.Forms.Label()
        Me._lblAnchor_3 = New System.Windows.Forms.Label()
        Me._lblAnchorCmt_1 = New System.Windows.Forms.Label()
        Me._lblAnchorCmt_0 = New System.Windows.Forms.Label()
        Me._lblAnchor_2 = New System.Windows.Forms.Label()
        Me._lblAnchor_1 = New System.Windows.Forms.Label()
        Me._lblAnchor_0 = New System.Windows.Forms.Label()
        Me.fraFairLead = New System.Windows.Forms.GroupBox()
        Me._txtFL_3 = New System.Windows.Forms.TextBox()
        Me._txtFL_4 = New System.Windows.Forms.TextBox()
        Me._txtFL_2 = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me._txtFL_1 = New System.Windows.Forms.TextBox()
        Me._txtFL_0 = New System.Windows.Forms.TextBox()
        Me._lblLengthUnit_7 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_6 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_5 = New System.Windows.Forms.Label()
        Me._lblFL_2 = New System.Windows.Forms.Label()
        Me._lblFLUnit_2 = New System.Windows.Forms.Label()
        Me._lblFLCmt_1 = New System.Windows.Forms.Label()
        Me._lblFLCmt_0 = New System.Windows.Forms.Label()
        Me._lblFLUnit_1 = New System.Windows.Forms.Label()
        Me._lblFL_1 = New System.Windows.Forms.Label()
        Me._lblFLUnit_0 = New System.Windows.Forms.Label()
        Me._lblFL_0 = New System.Windows.Forms.Label()
        Me.dlgFileOpen = New System.Windows.Forms.OpenFileDialog()
        Me.dlgFileSave = New System.Windows.Forms.SaveFileDialog()
        Me.lblAnchor = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblAnchorCmt = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblAngleUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblFL = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblFLCmt = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblFLUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblForceUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblGenCmt = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblGeneral = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblLengthUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblPreTen = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblSegCmt = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblVelUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblVslSt = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.txtAnchor = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.txtFL = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.txtGeneral = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.txtPreTen = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.txtVslSt = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip()
        Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuOpen = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuSave = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuDum = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuClose = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuGridEdit = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuInsert = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuDelete = New System.Windows.Forms.ToolStripMenuItem()
        Me.tabMoorLines = New System.Windows.Forms.TabControl()
        Me.Tab1 = New System.Windows.Forms.TabPage()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Frame1.SuspendLayout()
        Me.fraVesselLoc.SuspendLayout()
        Me.fraSegments.SuspendLayout()
        CType(Me.grdAnchor, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdSegments, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraGeneral.SuspendLayout()
        Me.fraTopTen.SuspendLayout()
        Me.fraAnchor.SuspendLayout()
        Me.fraFairLead.SuspendLayout()
        CType(Me.lblAnchor, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblAnchorCmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblAngleUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblFL, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblFLCmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblFLUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblForceUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblGenCmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblGeneral, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblLengthUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblPreTen, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblSegCmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVelUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVslSt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtAnchor, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtFL, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtGeneral, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtPreTen, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtVslSt, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MainMenu1.SuspendLayout()
        Me.tabMoorLines.SuspendLayout()
        Me.Tab1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnPayout
        '
        Me.btnPayout.BackColor = System.Drawing.SystemColors.Control
        Me.btnPayout.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnPayout.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnPayout.Location = New System.Drawing.Point(86, 134)
        Me.btnPayout.Name = "btnPayout"
        Me.btnPayout.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnPayout.Size = New System.Drawing.Size(63, 25)
        Me.btnPayout.TabIndex = 23
        Me.btnPayout.Text = "Com&pute "
        Me.ToolTip1.SetToolTip(Me.btnPayout, "Compute Payout & Scope")
        Me.btnPayout.UseVisualStyleBackColor = False
        '
        'btnScope
        '
        Me.btnScope.BackColor = System.Drawing.SystemColors.Control
        Me.btnScope.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnScope.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnScope.Location = New System.Drawing.Point(12, 72)
        Me.btnScope.Name = "btnScope"
        Me.btnScope.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnScope.Size = New System.Drawing.Size(139, 25)
        Me.btnScope.TabIndex = 20
        Me.btnScope.Text = "C&ompute Scope, Spread"
        Me.ToolTip1.SetToolTip(Me.btnScope, "Compute Scope & Spread Angle")
        Me.btnScope.UseVisualStyleBackColor = False
        '
        '_txtGeneral_2
        '
        Me._txtGeneral_2.AcceptsReturn = True
        Me._txtGeneral_2.BackColor = System.Drawing.SystemColors.Window
        Me._txtGeneral_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtGeneral_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtGeneral.SetIndex(Me._txtGeneral_2, CType(2, Short))
        Me._txtGeneral_2.Location = New System.Drawing.Point(88, 112)
        Me._txtGeneral_2.Name = "_txtGeneral_2"
        Me._txtGeneral_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtGeneral_2.Size = New System.Drawing.Size(57, 20)
        Me._txtGeneral_2.TabIndex = 22
        Me._txtGeneral_2.Text = "0"
        Me._txtGeneral_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me._txtGeneral_2, "First Segment Length from Fairlead.")
        '
        'btnTopTension
        '
        Me.btnTopTension.BackColor = System.Drawing.SystemColors.Control
        Me.btnTopTension.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnTopTension.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnTopTension.Location = New System.Drawing.Point(103, 51)
        Me.btnTopTension.Name = "btnTopTension"
        Me.btnTopTension.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnTopTension.Size = New System.Drawing.Size(73, 25)
        Me.btnTopTension.TabIndex = 30
        Me.btnTopTension.Text = "Compu&te"
        Me.ToolTip1.SetToolTip(Me.btnTopTension, "Compute Top Tension")
        Me.btnTopTension.UseVisualStyleBackColor = False
        '
        'btnPreten
        '
        Me.btnPreten.BackColor = System.Drawing.SystemColors.Control
        Me.btnPreten.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnPreten.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnPreten.Location = New System.Drawing.Point(21, 175)
        Me.btnPreten.Name = "btnPreten"
        Me.btnPreten.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnPreten.Size = New System.Drawing.Size(124, 25)
        Me.btnPreten.TabIndex = 36
        Me.btnPreten.Text = "&Set To Top Tension"
        Me.ToolTip1.SetToolTip(Me.btnPreten, "Set Pretension")
        Me.btnPreten.UseVisualStyleBackColor = False
        '
        'btnAnchor
        '
        Me.btnAnchor.BackColor = System.Drawing.SystemColors.Control
        Me.btnAnchor.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnAnchor.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnAnchor.Location = New System.Drawing.Point(74, 118)
        Me.btnAnchor.Name = "btnAnchor"
        Me.btnAnchor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnAnchor.Size = New System.Drawing.Size(75, 28)
        Me.btnAnchor.TabIndex = 58
        Me.btnAnchor.Text = "Comp&ute"
        Me.ToolTip1.SetToolTip(Me.btnAnchor, "Compute Anchor (x,y)")
        Me.btnAnchor.UseVisualStyleBackColor = False
        '
        '_txtAnchor_3
        '
        Me._txtAnchor_3.AcceptsReturn = True
        Me._txtAnchor_3.BackColor = System.Drawing.SystemColors.Window
        Me._txtAnchor_3.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtAnchor_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAnchor.SetIndex(Me._txtAnchor_3, CType(3, Short))
        Me._txtAnchor_3.Location = New System.Drawing.Point(74, 95)
        Me._txtAnchor_3.Name = "_txtAnchor_3"
        Me._txtAnchor_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtAnchor_3.Size = New System.Drawing.Size(57, 20)
        Me._txtAnchor_3.TabIndex = 57
        Me._txtAnchor_3.Text = "0"
        Me._txtAnchor_3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me._txtAnchor_3, "Local slope: default to average.")
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me._txtVslSt_3)
        Me.Frame1.Controls.Add(Me.cboDraft)
        Me.Frame1.Controls.Add(Me._lblLengthUnit_2)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(22, 27)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(190, 100)
        Me.Frame1.TabIndex = 89
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Design Vessel Draft"
        '
        '_txtVslSt_3
        '
        Me._txtVslSt_3.AcceptsReturn = True
        Me._txtVslSt_3.BackColor = System.Drawing.SystemColors.Control
        Me._txtVslSt_3.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtVslSt_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtVslSt.SetIndex(Me._txtVslSt_3, CType(3, Short))
        Me._txtVslSt_3.Location = New System.Drawing.Point(75, 63)
        Me._txtVslSt_3.Name = "_txtVslSt_3"
        Me._txtVslSt_3.ReadOnly = True
        Me._txtVslSt_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtVslSt_3.Size = New System.Drawing.Size(55, 20)
        Me._txtVslSt_3.TabIndex = 91
        Me._txtVslSt_3.TabStop = False
        Me._txtVslSt_3.Text = "0"
        Me._txtVslSt_3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cboDraft
        '
        Me.cboDraft.BackColor = System.Drawing.SystemColors.Window
        Me.cboDraft.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboDraft.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDraft.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboDraft.Items.AddRange(New Object() {"Current Draft", "Survival Draft", "Operating Draft"})
        Me.cboDraft.Location = New System.Drawing.Point(40, 31)
        Me.cboDraft.Name = "cboDraft"
        Me.cboDraft.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboDraft.Size = New System.Drawing.Size(116, 21)
        Me.cboDraft.TabIndex = 90
        '
        '_lblLengthUnit_2
        '
        Me._lblLengthUnit_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_2, CType(2, Short))
        Me._lblLengthUnit_2.Location = New System.Drawing.Point(139, 66)
        Me._lblLengthUnit_2.Name = "_lblLengthUnit_2"
        Me._lblLengthUnit_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_2.Size = New System.Drawing.Size(17, 17)
        Me._lblLengthUnit_2.TabIndex = 92
        Me._lblLengthUnit_2.Text = "ft"
        '
        'fraVesselLoc
        '
        Me.fraVesselLoc.BackColor = System.Drawing.SystemColors.Control
        Me.fraVesselLoc.Controls.Add(Me.txtWD)
        Me.fraVesselLoc.Controls.Add(Me._txtVslSt_2)
        Me.fraVesselLoc.Controls.Add(Me._txtVslSt_1)
        Me.fraVesselLoc.Controls.Add(Me._txtVslSt_0)
        Me.fraVesselLoc.Controls.Add(Me._lblVslSt_3)
        Me.fraVesselLoc.Controls.Add(Me._lblLengthUnit_11)
        Me.fraVesselLoc.Controls.Add(Me._lblLengthUnit_1)
        Me.fraVesselLoc.Controls.Add(Me._lblVslSt_2)
        Me.fraVesselLoc.Controls.Add(Me._lblVslSt_1)
        Me.fraVesselLoc.Controls.Add(Me._lblVslSt_0)
        Me.fraVesselLoc.Controls.Add(Me._lblLengthUnit_0)
        Me.fraVesselLoc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraVesselLoc.Location = New System.Drawing.Point(232, 29)
        Me.fraVesselLoc.Name = "fraVesselLoc"
        Me.fraVesselLoc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraVesselLoc.Size = New System.Drawing.Size(468, 101)
        Me.fraVesselLoc.TabIndex = 3
        Me.fraVesselLoc.TabStop = False
        Me.fraVesselLoc.Text = "Design Vessel Station :"
        '
        'txtWD
        '
        Me.txtWD.AcceptsReturn = True
        Me.txtWD.BackColor = System.Drawing.SystemColors.Control
        Me.txtWD.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtWD.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtWD.Location = New System.Drawing.Point(297, 55)
        Me.txtWD.Name = "txtWD"
        Me.txtWD.ReadOnly = True
        Me.txtWD.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtWD.Size = New System.Drawing.Size(59, 20)
        Me.txtWD.TabIndex = 83
        Me.txtWD.TabStop = False
        Me.txtWD.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtVslSt_2
        '
        Me._txtVslSt_2.AcceptsReturn = True
        Me._txtVslSt_2.BackColor = System.Drawing.SystemColors.Window
        Me._txtVslSt_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtVslSt_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtVslSt.SetIndex(Me._txtVslSt_2, CType(2, Short))
        Me._txtVslSt_2.Location = New System.Drawing.Point(297, 30)
        Me._txtVslSt_2.Name = "_txtVslSt_2"
        Me._txtVslSt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtVslSt_2.Size = New System.Drawing.Size(58, 20)
        Me._txtVslSt_2.TabIndex = 10
        Me._txtVslSt_2.TabStop = False
        Me._txtVslSt_2.Text = "0"
        Me._txtVslSt_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtVslSt_1
        '
        Me._txtVslSt_1.AcceptsReturn = True
        Me._txtVslSt_1.BackColor = System.Drawing.SystemColors.Window
        Me._txtVslSt_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtVslSt_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtVslSt.SetIndex(Me._txtVslSt_1, CType(1, Short))
        Me._txtVslSt_1.Location = New System.Drawing.Point(78, 55)
        Me._txtVslSt_1.Name = "_txtVslSt_1"
        Me._txtVslSt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtVslSt_1.Size = New System.Drawing.Size(121, 20)
        Me._txtVslSt_1.TabIndex = 8
        Me._txtVslSt_1.TabStop = False
        Me._txtVslSt_1.Text = "0"
        Me._txtVslSt_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtVslSt_0
        '
        Me._txtVslSt_0.AcceptsReturn = True
        Me._txtVslSt_0.BackColor = System.Drawing.SystemColors.Window
        Me._txtVslSt_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtVslSt_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtVslSt.SetIndex(Me._txtVslSt_0, CType(0, Short))
        Me._txtVslSt_0.Location = New System.Drawing.Point(77, 30)
        Me._txtVslSt_0.Name = "_txtVslSt_0"
        Me._txtVslSt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtVslSt_0.Size = New System.Drawing.Size(123, 20)
        Me._txtVslSt_0.TabIndex = 5
        Me._txtVslSt_0.TabStop = False
        Me._txtVslSt_0.Text = "0"
        Me._txtVslSt_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_lblVslSt_3
        '
        Me._lblVslSt_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblVslSt_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVslSt_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVslSt.SetIndex(Me._lblVslSt_3, CType(3, Short))
        Me._lblVslSt_3.Location = New System.Drawing.Point(235, 56)
        Me._lblVslSt_3.Name = "_lblVslSt_3"
        Me._lblVslSt_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVslSt_3.Size = New System.Drawing.Size(53, 16)
        Me._lblVslSt_3.TabIndex = 85
        Me._lblVslSt_3.Text = "Depth"
        '
        '_lblLengthUnit_11
        '
        Me._lblLengthUnit_11.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_11.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_11, CType(11, Short))
        Me._lblLengthUnit_11.Location = New System.Drawing.Point(361, 56)
        Me._lblLengthUnit_11.Name = "_lblLengthUnit_11"
        Me._lblLengthUnit_11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_11.Size = New System.Drawing.Size(17, 17)
        Me._lblLengthUnit_11.TabIndex = 84
        Me._lblLengthUnit_11.Text = "ft"
        '
        '_lblLengthUnit_1
        '
        Me._lblLengthUnit_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_1, CType(1, Short))
        Me._lblLengthUnit_1.Location = New System.Drawing.Point(205, 56)
        Me._lblLengthUnit_1.Name = "_lblLengthUnit_1"
        Me._lblLengthUnit_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_1.Size = New System.Drawing.Size(17, 17)
        Me._lblLengthUnit_1.TabIndex = 69
        Me._lblLengthUnit_1.Text = "ft"
        '
        '_lblVslSt_2
        '
        Me._lblVslSt_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblVslSt_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVslSt_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVslSt.SetIndex(Me._lblVslSt_2, CType(2, Short))
        Me._lblVslSt_2.Location = New System.Drawing.Point(234, 32)
        Me._lblVslSt_2.Name = "_lblVslSt_2"
        Me._lblVslSt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVslSt_2.Size = New System.Drawing.Size(58, 17)
        Me._lblVslSt_2.TabIndex = 9
        Me._lblVslSt_2.Text = "Heading"
        '
        '_lblVslSt_1
        '
        Me._lblVslSt_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblVslSt_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVslSt_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVslSt.SetIndex(Me._lblVslSt_1, CType(1, Short))
        Me._lblVslSt_1.Location = New System.Drawing.Point(48, 55)
        Me._lblVslSt_1.Name = "_lblVslSt_1"
        Me._lblVslSt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVslSt_1.Size = New System.Drawing.Size(33, 17)
        Me._lblVslSt_1.TabIndex = 7
        Me._lblVslSt_1.Text = "y (N)"
        '
        '_lblVslSt_0
        '
        Me._lblVslSt_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblVslSt_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVslSt_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVslSt.SetIndex(Me._lblVslSt_0, CType(0, Short))
        Me._lblVslSt_0.Location = New System.Drawing.Point(48, 32)
        Me._lblVslSt_0.Name = "_lblVslSt_0"
        Me._lblVslSt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVslSt_0.Size = New System.Drawing.Size(33, 17)
        Me._lblVslSt_0.TabIndex = 4
        Me._lblVslSt_0.Text = "x (E)"
        '
        '_lblLengthUnit_0
        '
        Me._lblLengthUnit_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_0, CType(0, Short))
        Me._lblLengthUnit_0.Location = New System.Drawing.Point(205, 32)
        Me._lblLengthUnit_0.Name = "_lblLengthUnit_0"
        Me._lblLengthUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_0.Size = New System.Drawing.Size(17, 17)
        Me._lblLengthUnit_0.TabIndex = 6
        Me._lblLengthUnit_0.Text = "ft"
        '
        'btnCatenary
        '
        Me.btnCatenary.BackColor = System.Drawing.SystemColors.Control
        Me.btnCatenary.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCatenary.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCatenary.Location = New System.Drawing.Point(95, 18)
        Me.btnCatenary.Name = "btnCatenary"
        Me.btnCatenary.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCatenary.Size = New System.Drawing.Size(74, 25)
        Me.btnCatenary.TabIndex = 82
        Me.btnCatenary.Text = "&Catenary"
        Me.btnCatenary.UseVisualStyleBackColor = False
        '
        'fraSegments
        '
        Me.fraSegments.BackColor = System.Drawing.SystemColors.Control
        Me.fraSegments.Controls.Add(Me.grdAnchor)
        Me.fraSegments.Controls.Add(Me.grdSegments)
        Me.fraSegments.Controls.Add(Me._lblSegCmt_0)
        Me.fraSegments.Controls.Add(Me._lblSegCmt_1)
        Me.fraSegments.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraSegments.Location = New System.Drawing.Point(12, 315)
        Me.fraSegments.Name = "fraSegments"
        Me.fraSegments.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraSegments.Size = New System.Drawing.Size(786, 206)
        Me.fraSegments.TabIndex = 61
        Me.fraSegments.TabStop = False
        Me.fraSegments.Text = "Properties of Line Segments :"
        '
        'grdAnchor
        '
        Me.grdAnchor.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdAnchor.ColumnHeadersVisible = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grdAnchor.DefaultCellStyle = DataGridViewCellStyle1
        Me.grdAnchor.Location = New System.Drawing.Point(4, 147)
        Me.grdAnchor.Name = "grdAnchor"
        Me.grdAnchor.RowHeadersVisible = False
        Me.grdAnchor.Size = New System.Drawing.Size(759, 34)
        Me.grdAnchor.TabIndex = 69
        '
        'grdSegments
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdSegments.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.grdSegments.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grdSegments.DefaultCellStyle = DataGridViewCellStyle3
        Me.grdSegments.Location = New System.Drawing.Point(3, 16)
        Me.grdSegments.Name = "grdSegments"
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdSegments.RowHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.grdSegments.RowHeadersVisible = False
        Me.grdSegments.Size = New System.Drawing.Size(775, 117)
        Me.grdSegments.TabIndex = 68
        '
        '_lblSegCmt_0
        '
        Me._lblSegCmt_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblSegCmt_0.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSegCmt.SetIndex(Me._lblSegCmt_0, CType(0, Short))
        Me._lblSegCmt_0.Location = New System.Drawing.Point(0, 186)
        Me._lblSegCmt_0.Name = "_lblSegCmt_0"
        Me._lblSegCmt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblSegCmt_0.Size = New System.Drawing.Size(352, 17)
        Me._lblSegCmt_0.TabIndex = 67
        Me._lblSegCmt_0.Text = "Input started from top of the mooring line, i.e. the fair lead"
        '
        '_lblSegCmt_1
        '
        Me._lblSegCmt_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblSegCmt_1.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSegCmt.SetIndex(Me._lblSegCmt_1, CType(1, Short))
        Me._lblSegCmt_1.Location = New System.Drawing.Point(16, 164)
        Me._lblSegCmt_1.Name = "_lblSegCmt_1"
        Me._lblSegCmt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblSegCmt_1.Size = New System.Drawing.Size(9, 17)
        Me._lblSegCmt_1.TabIndex = 66
        Me._lblSegCmt_1.Text = "*"
        '
        'fraGeneral
        '
        Me.fraGeneral.BackColor = System.Drawing.SystemColors.Control
        Me.fraGeneral.Controls.Add(Me.btnPayout)
        Me.fraGeneral.Controls.Add(Me.btnScope)
        Me.fraGeneral.Controls.Add(Me._txtGeneral_2)
        Me.fraGeneral.Controls.Add(Me._txtGeneral_1)
        Me.fraGeneral.Controls.Add(Me._txtGeneral_0)
        Me.fraGeneral.Controls.Add(Me._lblLengthUnit_4)
        Me.fraGeneral.Controls.Add(Me._lblLengthUnit_3)
        Me.fraGeneral.Controls.Add(Me._lblGenCmt_0)
        Me.fraGeneral.Controls.Add(Me._lblGenCmt_1)
        Me.fraGeneral.Controls.Add(Me._lblGeneral_2)
        Me.fraGeneral.Controls.Add(Me._lblGeneral_1)
        Me.fraGeneral.Controls.Add(Me._lblGeneral_0)
        Me.fraGeneral.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraGeneral.Location = New System.Drawing.Point(13, 92)
        Me.fraGeneral.Name = "fraGeneral"
        Me.fraGeneral.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraGeneral.Size = New System.Drawing.Size(185, 217)
        Me.fraGeneral.TabIndex = 15
        Me.fraGeneral.TabStop = False
        Me.fraGeneral.Text = "General Data :"
        '
        '_txtGeneral_1
        '
        Me._txtGeneral_1.AcceptsReturn = True
        Me._txtGeneral_1.BackColor = System.Drawing.SystemColors.Window
        Me._txtGeneral_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtGeneral_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtGeneral.SetIndex(Me._txtGeneral_1, CType(1, Short))
        Me._txtGeneral_1.Location = New System.Drawing.Point(103, 48)
        Me._txtGeneral_1.Name = "_txtGeneral_1"
        Me._txtGeneral_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtGeneral_1.Size = New System.Drawing.Size(57, 20)
        Me._txtGeneral_1.TabIndex = 19
        Me._txtGeneral_1.Text = "0"
        Me._txtGeneral_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtGeneral_0
        '
        Me._txtGeneral_0.AcceptsReturn = True
        Me._txtGeneral_0.BackColor = System.Drawing.SystemColors.Window
        Me._txtGeneral_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtGeneral_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtGeneral.SetIndex(Me._txtGeneral_0, CType(0, Short))
        Me._txtGeneral_0.Location = New System.Drawing.Point(103, 23)
        Me._txtGeneral_0.Name = "_txtGeneral_0"
        Me._txtGeneral_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtGeneral_0.Size = New System.Drawing.Size(57, 20)
        Me._txtGeneral_0.TabIndex = 17
        Me._txtGeneral_0.Text = "0"
        Me._txtGeneral_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_lblLengthUnit_4
        '
        Me._lblLengthUnit_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_4, CType(4, Short))
        Me._lblLengthUnit_4.Location = New System.Drawing.Point(148, 115)
        Me._lblLengthUnit_4.Name = "_lblLengthUnit_4"
        Me._lblLengthUnit_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_4.Size = New System.Drawing.Size(17, 17)
        Me._lblLengthUnit_4.TabIndex = 71
        Me._lblLengthUnit_4.Text = "ft"
        '
        '_lblLengthUnit_3
        '
        Me._lblLengthUnit_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_3, CType(3, Short))
        Me._lblLengthUnit_3.Location = New System.Drawing.Point(162, 51)
        Me._lblLengthUnit_3.Name = "_lblLengthUnit_3"
        Me._lblLengthUnit_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_3.Size = New System.Drawing.Size(17, 17)
        Me._lblLengthUnit_3.TabIndex = 70
        Me._lblLengthUnit_3.Text = "ft"
        '
        '_lblGenCmt_0
        '
        Me._lblGenCmt_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblGenCmt_0.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblGenCmt.SetIndex(Me._lblGenCmt_0, CType(0, Short))
        Me._lblGenCmt_0.Location = New System.Drawing.Point(12, 168)
        Me._lblGenCmt_0.Name = "_lblGenCmt_0"
        Me._lblGenCmt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblGenCmt_0.Size = New System.Drawing.Size(157, 41)
        Me._lblGenCmt_0.TabIndex = 25
        Me._lblGenCmt_0.Text = "Spread Angle measured clockwise from the x-axis of the vessel local system (bow)"
        '
        '_lblGenCmt_1
        '
        Me._lblGenCmt_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblGenCmt_1.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblGenCmt.SetIndex(Me._lblGenCmt_1, CType(1, Short))
        Me._lblGenCmt_1.Location = New System.Drawing.Point(6, 168)
        Me._lblGenCmt_1.Name = "_lblGenCmt_1"
        Me._lblGenCmt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblGenCmt_1.Size = New System.Drawing.Size(9, 17)
        Me._lblGenCmt_1.TabIndex = 24
        Me._lblGenCmt_1.Text = "*"
        '
        '_lblGeneral_2
        '
        Me._lblGeneral_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblGeneral_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblGeneral_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblGeneral.SetIndex(Me._lblGeneral_2, CType(2, Short))
        Me._lblGeneral_2.Location = New System.Drawing.Point(16, 112)
        Me._lblGeneral_2.Name = "_lblGeneral_2"
        Me._lblGeneral_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblGeneral_2.Size = New System.Drawing.Size(65, 17)
        Me._lblGeneral_2.TabIndex = 21
        Me._lblGeneral_2.Text = "Payout"
        '
        '_lblGeneral_1
        '
        Me._lblGeneral_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblGeneral_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblGeneral_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblGeneral.SetIndex(Me._lblGeneral_1, CType(1, Short))
        Me._lblGeneral_1.Location = New System.Drawing.Point(16, 48)
        Me._lblGeneral_1.Name = "_lblGeneral_1"
        Me._lblGeneral_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblGeneral_1.Size = New System.Drawing.Size(65, 17)
        Me._lblGeneral_1.TabIndex = 18
        Me._lblGeneral_1.Text = "Scope"
        '
        '_lblGeneral_0
        '
        Me._lblGeneral_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblGeneral_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblGeneral_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblGeneral.SetIndex(Me._lblGeneral_0, CType(0, Short))
        Me._lblGeneral_0.Location = New System.Drawing.Point(16, 24)
        Me._lblGeneral_0.Name = "_lblGeneral_0"
        Me._lblGeneral_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblGeneral_0.Size = New System.Drawing.Size(81, 20)
        Me._lblGeneral_0.TabIndex = 16
        Me._lblGeneral_0.Text = "Spread Angle"
        '
        'chkBatch
        '
        Me.chkBatch.BackColor = System.Drawing.SystemColors.Control
        Me.chkBatch.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkBatch.Location = New System.Drawing.Point(624, 26)
        Me.chkBatch.Name = "chkBatch"
        Me.chkBatch.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkBatch.Size = New System.Drawing.Size(177, 17)
        Me.chkBatch.TabIndex = 14
        Me.chkBatch.Text = "Edit all mooring lines at same time"
        Me.chkBatch.UseVisualStyleBackColor = False
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.SystemColors.Control
        Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCancel.Location = New System.Drawing.Point(755, 59)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCancel.Size = New System.Drawing.Size(73, 25)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'btnNew
        '
        Me.btnNew.BackColor = System.Drawing.SystemColors.Control
        Me.btnNew.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNew.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNew.Location = New System.Drawing.Point(755, 90)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNew.Size = New System.Drawing.Size(73, 25)
        Me.btnNew.TabIndex = 2
        Me.btnNew.Text = "&New Line"
        Me.btnNew.UseVisualStyleBackColor = False
        '
        'btnRemove
        '
        Me.btnRemove.BackColor = System.Drawing.SystemColors.Control
        Me.btnRemove.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnRemove.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnRemove.Location = New System.Drawing.Point(13, 15)
        Me.btnRemove.Name = "btnRemove"
        Me.btnRemove.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnRemove.Size = New System.Drawing.Size(73, 25)
        Me.btnRemove.TabIndex = 13
        Me.btnRemove.Text = "&Remove"
        Me.btnRemove.UseVisualStyleBackColor = False
        '
        'btnSave
        '
        Me.btnSave.BackColor = System.Drawing.SystemColors.Control
        Me.btnSave.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnSave.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnSave.Location = New System.Drawing.Point(755, 27)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnSave.Size = New System.Drawing.Size(73, 25)
        Me.btnSave.TabIndex = 0
        Me.btnSave.Text = "Sa&ve"
        Me.btnSave.UseVisualStyleBackColor = False
        '
        'fraTopTen
        '
        Me.fraTopTen.BackColor = System.Drawing.SystemColors.Control
        Me.fraTopTen.Controls.Add(Me.btnTopTension)
        Me.fraTopTen.Controls.Add(Me.btnPreten)
        Me.fraTopTen.Controls.Add(Me._txtPreTen_1)
        Me.fraTopTen.Controls.Add(Me.txtTopTen)
        Me.fraTopTen.Controls.Add(Me._txtPreTen_0)
        Me.fraTopTen.Controls.Add(Me._lblForceUnit_2)
        Me.fraTopTen.Controls.Add(Me._lblForceUnit_1)
        Me.fraTopTen.Controls.Add(Me.Label1)
        Me.fraTopTen.Controls.Add(Me._lblPreTen_1)
        Me.fraTopTen.Controls.Add(Me.lblTopTen)
        Me.fraTopTen.Controls.Add(Me._lblForceUnit_0)
        Me.fraTopTen.Controls.Add(Me._lblPreTen_0)
        Me.fraTopTen.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraTopTen.Location = New System.Drawing.Point(204, 92)
        Me.fraTopTen.Name = "fraTopTen"
        Me.fraTopTen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraTopTen.Size = New System.Drawing.Size(205, 217)
        Me.fraTopTen.TabIndex = 26
        Me.fraTopTen.TabStop = False
        Me.fraTopTen.Text = "Top Tensions :"
        '
        '_txtPreTen_1
        '
        Me._txtPreTen_1.AcceptsReturn = True
        Me._txtPreTen_1.BackColor = System.Drawing.SystemColors.Control
        Me._txtPreTen_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtPreTen_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPreTen.SetIndex(Me._txtPreTen_1, CType(1, Short))
        Me._txtPreTen_1.Location = New System.Drawing.Point(88, 152)
        Me._txtPreTen_1.Name = "_txtPreTen_1"
        Me._txtPreTen_1.ReadOnly = True
        Me._txtPreTen_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtPreTen_1.Size = New System.Drawing.Size(57, 20)
        Me._txtPreTen_1.TabIndex = 35
        Me._txtPreTen_1.TabStop = False
        Me._txtPreTen_1.Text = "0"
        Me._txtPreTen_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtTopTen
        '
        Me.txtTopTen.AcceptsReturn = True
        Me.txtTopTen.BackColor = System.Drawing.SystemColors.Window
        Me.txtTopTen.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTopTen.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTopTen.Location = New System.Drawing.Point(103, 26)
        Me.txtTopTen.Name = "txtTopTen"
        Me.txtTopTen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTopTen.Size = New System.Drawing.Size(57, 20)
        Me.txtTopTen.TabIndex = 28
        Me.txtTopTen.Text = "0"
        Me.txtTopTen.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtPreTen_0
        '
        Me._txtPreTen_0.AcceptsReturn = True
        Me._txtPreTen_0.BackColor = System.Drawing.SystemColors.Control
        Me._txtPreTen_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtPreTen_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPreTen.SetIndex(Me._txtPreTen_0, CType(0, Short))
        Me._txtPreTen_0.Location = New System.Drawing.Point(88, 128)
        Me._txtPreTen_0.Name = "_txtPreTen_0"
        Me._txtPreTen_0.ReadOnly = True
        Me._txtPreTen_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtPreTen_0.Size = New System.Drawing.Size(57, 20)
        Me._txtPreTen_0.TabIndex = 33
        Me._txtPreTen_0.TabStop = False
        Me._txtPreTen_0.Text = "0"
        Me._txtPreTen_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_lblForceUnit_2
        '
        Me._lblForceUnit_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblForceUnit_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblForceUnit_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblForceUnit.SetIndex(Me._lblForceUnit_2, CType(2, Short))
        Me._lblForceUnit_2.Location = New System.Drawing.Point(165, 153)
        Me._lblForceUnit_2.Name = "_lblForceUnit_2"
        Me._lblForceUnit_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblForceUnit_2.Size = New System.Drawing.Size(34, 19)
        Me._lblForceUnit_2.TabIndex = 81
        Me._lblForceUnit_2.Text = "kips"
        '
        '_lblForceUnit_1
        '
        Me._lblForceUnit_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblForceUnit_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblForceUnit_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblForceUnit.SetIndex(Me._lblForceUnit_1, CType(1, Short))
        Me._lblForceUnit_1.Location = New System.Drawing.Point(166, 126)
        Me._lblForceUnit_1.Name = "_lblForceUnit_1"
        Me._lblForceUnit_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblForceUnit_1.Size = New System.Drawing.Size(28, 17)
        Me._lblForceUnit_1.TabIndex = 80
        Me._lblForceUnit_1.Text = "kips"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(16, 107)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(65, 17)
        Me.Label1.TabIndex = 31
        Me.Label1.Text = "Pretensions"
        '
        '_lblPreTen_1
        '
        Me._lblPreTen_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblPreTen_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblPreTen_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPreTen.SetIndex(Me._lblPreTen_1, CType(1, Short))
        Me._lblPreTen_1.Location = New System.Drawing.Point(24, 152)
        Me._lblPreTen_1.Name = "_lblPreTen_1"
        Me._lblPreTen_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblPreTen_1.Size = New System.Drawing.Size(57, 17)
        Me._lblPreTen_1.TabIndex = 34
        Me._lblPreTen_1.Text = "Operating"
        '
        'lblTopTen
        '
        Me.lblTopTen.BackColor = System.Drawing.SystemColors.Control
        Me.lblTopTen.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTopTen.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTopTen.Location = New System.Drawing.Point(16, 24)
        Me.lblTopTen.Name = "lblTopTen"
        Me.lblTopTen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTopTen.Size = New System.Drawing.Size(81, 19)
        Me.lblTopTen.TabIndex = 27
        Me.lblTopTen.Text = "Top Tension"
        '
        '_lblForceUnit_0
        '
        Me._lblForceUnit_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblForceUnit_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblForceUnit_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblForceUnit.SetIndex(Me._lblForceUnit_0, CType(0, Short))
        Me._lblForceUnit_0.Location = New System.Drawing.Point(166, 26)
        Me._lblForceUnit_0.Name = "_lblForceUnit_0"
        Me._lblForceUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblForceUnit_0.Size = New System.Drawing.Size(33, 21)
        Me._lblForceUnit_0.TabIndex = 29
        Me._lblForceUnit_0.Text = "kips"
        '
        '_lblPreTen_0
        '
        Me._lblPreTen_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblPreTen_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblPreTen_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPreTen.SetIndex(Me._lblPreTen_0, CType(0, Short))
        Me._lblPreTen_0.Location = New System.Drawing.Point(24, 130)
        Me._lblPreTen_0.Name = "_lblPreTen_0"
        Me._lblPreTen_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblPreTen_0.Size = New System.Drawing.Size(57, 17)
        Me._lblPreTen_0.TabIndex = 32
        Me._lblPreTen_0.Text = "Survival"
        '
        'fraAnchor
        '
        Me.fraAnchor.BackColor = System.Drawing.SystemColors.Control
        Me.fraAnchor.Controls.Add(Me.btnAnchor)
        Me.fraAnchor.Controls.Add(Me._txtAnchor_3)
        Me.fraAnchor.Controls.Add(Me._txtAnchor_2)
        Me.fraAnchor.Controls.Add(Me._txtAnchor_1)
        Me.fraAnchor.Controls.Add(Me._txtAnchor_0)
        Me.fraAnchor.Controls.Add(Me._lblAnchorCmt_3)
        Me.fraAnchor.Controls.Add(Me._lblAnchorCmt_2)
        Me.fraAnchor.Controls.Add(Me._lblLengthUnit_10)
        Me.fraAnchor.Controls.Add(Me._lblLengthUnit_9)
        Me.fraAnchor.Controls.Add(Me._lblLengthUnit_8)
        Me.fraAnchor.Controls.Add(Me._lblAnchor_3)
        Me.fraAnchor.Controls.Add(Me._lblAnchorCmt_1)
        Me.fraAnchor.Controls.Add(Me._lblAnchorCmt_0)
        Me.fraAnchor.Controls.Add(Me._lblAnchor_2)
        Me.fraAnchor.Controls.Add(Me._lblAnchor_1)
        Me.fraAnchor.Controls.Add(Me._lblAnchor_0)
        Me.fraAnchor.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraAnchor.Location = New System.Drawing.Point(609, 92)
        Me.fraAnchor.Name = "fraAnchor"
        Me.fraAnchor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraAnchor.Size = New System.Drawing.Size(192, 225)
        Me.fraAnchor.TabIndex = 49
        Me.fraAnchor.TabStop = False
        Me.fraAnchor.Text = "Anchor Location :"
        '
        '_txtAnchor_2
        '
        Me._txtAnchor_2.AcceptsReturn = True
        Me._txtAnchor_2.BackColor = System.Drawing.SystemColors.Window
        Me._txtAnchor_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtAnchor_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAnchor.SetIndex(Me._txtAnchor_2, CType(2, Short))
        Me._txtAnchor_2.Location = New System.Drawing.Point(73, 66)
        Me._txtAnchor_2.Name = "_txtAnchor_2"
        Me._txtAnchor_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtAnchor_2.Size = New System.Drawing.Size(75, 20)
        Me._txtAnchor_2.TabIndex = 55
        Me._txtAnchor_2.Text = "0"
        Me._txtAnchor_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtAnchor_1
        '
        Me._txtAnchor_1.AcceptsReturn = True
        Me._txtAnchor_1.BackColor = System.Drawing.SystemColors.Window
        Me._txtAnchor_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtAnchor_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAnchor.SetIndex(Me._txtAnchor_1, CType(1, Short))
        Me._txtAnchor_1.Location = New System.Drawing.Point(73, 44)
        Me._txtAnchor_1.Name = "_txtAnchor_1"
        Me._txtAnchor_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtAnchor_1.Size = New System.Drawing.Size(75, 20)
        Me._txtAnchor_1.TabIndex = 53
        Me._txtAnchor_1.Text = "0"
        Me._txtAnchor_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtAnchor_0
        '
        Me._txtAnchor_0.AcceptsReturn = True
        Me._txtAnchor_0.BackColor = System.Drawing.SystemColors.Window
        Me._txtAnchor_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtAnchor_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAnchor.SetIndex(Me._txtAnchor_0, CType(0, Short))
        Me._txtAnchor_0.Location = New System.Drawing.Point(73, 22)
        Me._txtAnchor_0.Name = "_txtAnchor_0"
        Me._txtAnchor_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtAnchor_0.Size = New System.Drawing.Size(76, 20)
        Me._txtAnchor_0.TabIndex = 51
        Me._txtAnchor_0.Text = "0"
        Me._txtAnchor_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_lblAnchorCmt_3
        '
        Me._lblAnchorCmt_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblAnchorCmt_3.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblAnchorCmt.SetIndex(Me._lblAnchorCmt_3, CType(3, Short))
        Me._lblAnchorCmt_3.Location = New System.Drawing.Point(18, 190)
        Me._lblAnchorCmt_3.Name = "_lblAnchorCmt_3"
        Me._lblAnchorCmt_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblAnchorCmt_3.Size = New System.Drawing.Size(9, 17)
        Me._lblAnchorCmt_3.TabIndex = 87
        Me._lblAnchorCmt_3.Text = "*"
        '
        '_lblAnchorCmt_2
        '
        Me._lblAnchorCmt_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblAnchorCmt_2.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblAnchorCmt.SetIndex(Me._lblAnchorCmt_2, CType(2, Short))
        Me._lblAnchorCmt_2.Location = New System.Drawing.Point(32, 189)
        Me._lblAnchorCmt_2.Name = "_lblAnchorCmt_2"
        Me._lblAnchorCmt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblAnchorCmt_2.Size = New System.Drawing.Size(148, 15)
        Me._lblAnchorCmt_2.TabIndex = 86
        Me._lblAnchorCmt_2.Text = "Positive Slope: up towards well"
        '
        '_lblLengthUnit_10
        '
        Me._lblLengthUnit_10.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_10.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_10, CType(10, Short))
        Me._lblLengthUnit_10.Location = New System.Drawing.Point(153, 77)
        Me._lblLengthUnit_10.Name = "_lblLengthUnit_10"
        Me._lblLengthUnit_10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_10.Size = New System.Drawing.Size(17, 17)
        Me._lblLengthUnit_10.TabIndex = 77
        Me._lblLengthUnit_10.Text = "ft"
        '
        '_lblLengthUnit_9
        '
        Me._lblLengthUnit_9.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_9.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_9, CType(9, Short))
        Me._lblLengthUnit_9.Location = New System.Drawing.Point(154, 51)
        Me._lblLengthUnit_9.Name = "_lblLengthUnit_9"
        Me._lblLengthUnit_9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_9.Size = New System.Drawing.Size(17, 17)
        Me._lblLengthUnit_9.TabIndex = 76
        Me._lblLengthUnit_9.Text = "ft"
        '
        '_lblLengthUnit_8
        '
        Me._lblLengthUnit_8.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_8, CType(8, Short))
        Me._lblLengthUnit_8.Location = New System.Drawing.Point(154, 26)
        Me._lblLengthUnit_8.Name = "_lblLengthUnit_8"
        Me._lblLengthUnit_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_8.Size = New System.Drawing.Size(15, 17)
        Me._lblLengthUnit_8.TabIndex = 75
        Me._lblLengthUnit_8.Text = "ft"
        '
        '_lblAnchor_3
        '
        Me._lblAnchor_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblAnchor_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblAnchor_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblAnchor.SetIndex(Me._lblAnchor_3, CType(3, Short))
        Me._lblAnchor_3.Location = New System.Drawing.Point(36, 96)
        Me._lblAnchor_3.Name = "_lblAnchor_3"
        Me._lblAnchor_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblAnchor_3.Size = New System.Drawing.Size(33, 17)
        Me._lblAnchor_3.TabIndex = 56
        Me._lblAnchor_3.Text = "Slope"
        '
        '_lblAnchorCmt_1
        '
        Me._lblAnchorCmt_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblAnchorCmt_1.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblAnchorCmt.SetIndex(Me._lblAnchorCmt_1, CType(1, Short))
        Me._lblAnchorCmt_1.Location = New System.Drawing.Point(17, 159)
        Me._lblAnchorCmt_1.Name = "_lblAnchorCmt_1"
        Me._lblAnchorCmt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblAnchorCmt_1.Size = New System.Drawing.Size(9, 17)
        Me._lblAnchorCmt_1.TabIndex = 59
        Me._lblAnchorCmt_1.Text = "*"
        '
        '_lblAnchorCmt_0
        '
        Me._lblAnchorCmt_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblAnchorCmt_0.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblAnchorCmt.SetIndex(Me._lblAnchorCmt_0, CType(0, Short))
        Me._lblAnchorCmt_0.Location = New System.Drawing.Point(32, 155)
        Me._lblAnchorCmt_0.Name = "_lblAnchorCmt_0"
        Me._lblAnchorCmt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblAnchorCmt_0.Size = New System.Drawing.Size(148, 29)
        Me._lblAnchorCmt_0.TabIndex = 60
        Me._lblAnchorCmt_0.Text = "x,y are in global coordinate system (E-N system)"
        '
        '_lblAnchor_2
        '
        Me._lblAnchor_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblAnchor_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblAnchor_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblAnchor.SetIndex(Me._lblAnchor_2, CType(2, Short))
        Me._lblAnchor_2.Location = New System.Drawing.Point(36, 72)
        Me._lblAnchor_2.Name = "_lblAnchor_2"
        Me._lblAnchor_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblAnchor_2.Size = New System.Drawing.Size(33, 17)
        Me._lblAnchor_2.TabIndex = 54
        Me._lblAnchor_2.Text = "Depth"
        '
        '_lblAnchor_1
        '
        Me._lblAnchor_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblAnchor_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblAnchor_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblAnchor.SetIndex(Me._lblAnchor_1, CType(1, Short))
        Me._lblAnchor_1.Location = New System.Drawing.Point(36, 48)
        Me._lblAnchor_1.Name = "_lblAnchor_1"
        Me._lblAnchor_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblAnchor_1.Size = New System.Drawing.Size(33, 17)
        Me._lblAnchor_1.TabIndex = 52
        Me._lblAnchor_1.Text = "y (N)"
        '
        '_lblAnchor_0
        '
        Me._lblAnchor_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblAnchor_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblAnchor_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblAnchor.SetIndex(Me._lblAnchor_0, CType(0, Short))
        Me._lblAnchor_0.Location = New System.Drawing.Point(36, 24)
        Me._lblAnchor_0.Name = "_lblAnchor_0"
        Me._lblAnchor_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblAnchor_0.Size = New System.Drawing.Size(33, 17)
        Me._lblAnchor_0.TabIndex = 50
        Me._lblAnchor_0.Text = "x (E)"
        '
        'fraFairLead
        '
        Me.fraFairLead.BackColor = System.Drawing.SystemColors.Control
        Me.fraFairLead.Controls.Add(Me._txtFL_3)
        Me.fraFairLead.Controls.Add(Me._txtFL_4)
        Me.fraFairLead.Controls.Add(Me._txtFL_2)
        Me.fraFairLead.Controls.Add(Me.Label10)
        Me.fraFairLead.Controls.Add(Me.Label12)
        Me.fraFairLead.Controls.Add(Me._txtFL_1)
        Me.fraFairLead.Controls.Add(Me._txtFL_0)
        Me.fraFairLead.Controls.Add(Me._lblLengthUnit_7)
        Me.fraFairLead.Controls.Add(Me._lblLengthUnit_6)
        Me.fraFairLead.Controls.Add(Me._lblLengthUnit_5)
        Me.fraFairLead.Controls.Add(Me._lblFL_2)
        Me.fraFairLead.Controls.Add(Me._lblFLUnit_2)
        Me.fraFairLead.Controls.Add(Me._lblFLCmt_1)
        Me.fraFairLead.Controls.Add(Me._lblFLCmt_0)
        Me.fraFairLead.Controls.Add(Me._lblFLUnit_1)
        Me.fraFairLead.Controls.Add(Me._lblFL_1)
        Me.fraFairLead.Controls.Add(Me._lblFLUnit_0)
        Me.fraFairLead.Controls.Add(Me._lblFL_0)
        Me.fraFairLead.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraFairLead.Location = New System.Drawing.Point(415, 92)
        Me.fraFairLead.Name = "fraFairLead"
        Me.fraFairLead.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraFairLead.Size = New System.Drawing.Size(183, 217)
        Me.fraFairLead.TabIndex = 37
        Me.fraFairLead.TabStop = False
        Me.fraFairLead.Text = "Fair Lead Location :"
        '
        '_txtFL_3
        '
        Me._txtFL_3.AcceptsReturn = True
        Me._txtFL_3.BackColor = System.Drawing.SystemColors.Window
        Me._txtFL_3.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtFL_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtFL_3.ImeMode = System.Windows.Forms.ImeMode.Off
        Me._txtFL_3.Location = New System.Drawing.Point(131, 156)
        Me._txtFL_3.Name = "_txtFL_3"
        Me._txtFL_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtFL_3.Size = New System.Drawing.Size(49, 20)
        Me._txtFL_3.TabIndex = 76
        Me._txtFL_3.Text = "0"
        Me._txtFL_3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtFL_4
        '
        Me._txtFL_4.AcceptsReturn = True
        Me._txtFL_4.BackColor = System.Drawing.SystemColors.Window
        Me._txtFL_4.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtFL_4.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtFL_4.ImeMode = System.Windows.Forms.ImeMode.Off
        Me._txtFL_4.Location = New System.Drawing.Point(131, 182)
        Me._txtFL_4.Name = "_txtFL_4"
        Me._txtFL_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtFL_4.Size = New System.Drawing.Size(49, 20)
        Me._txtFL_4.TabIndex = 76
        Me._txtFL_4.Text = "0"
        Me._txtFL_4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtFL_2
        '
        Me._txtFL_2.AcceptsReturn = True
        Me._txtFL_2.BackColor = System.Drawing.SystemColors.Window
        Me._txtFL_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtFL_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFL.SetIndex(Me._txtFL_2, CType(2, Short))
        Me._txtFL_2.Location = New System.Drawing.Point(57, 73)
        Me._txtFL_2.Name = "_txtFL_2"
        Me._txtFL_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtFL_2.Size = New System.Drawing.Size(49, 20)
        Me._txtFL_2.TabIndex = 45
        Me._txtFL_2.Text = "0"
        Me._txtFL_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Location = New System.Drawing.Point(10, 185)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(115, 18)
        Me.Label10.TabIndex = 41
        Me.Label10.Text = "Anchor Aqwa Node"
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.SystemColors.Control
        Me.Label12.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label12.Location = New System.Drawing.Point(10, 159)
        Me.Label12.Name = "Label12"
        Me.Label12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label12.Size = New System.Drawing.Size(117, 20)
        Me.Label12.TabIndex = 38
        Me.Label12.Text = "Fair Lead Aqwa Node"
        '
        '_txtFL_1
        '
        Me._txtFL_1.AcceptsReturn = True
        Me._txtFL_1.BackColor = System.Drawing.SystemColors.Window
        Me._txtFL_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtFL_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFL.SetIndex(Me._txtFL_1, CType(1, Short))
        Me._txtFL_1.Location = New System.Drawing.Point(57, 49)
        Me._txtFL_1.Name = "_txtFL_1"
        Me._txtFL_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtFL_1.Size = New System.Drawing.Size(49, 20)
        Me._txtFL_1.TabIndex = 42
        Me._txtFL_1.Text = "0"
        Me._txtFL_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtFL_0
        '
        Me._txtFL_0.AcceptsReturn = True
        Me._txtFL_0.BackColor = System.Drawing.SystemColors.Window
        Me._txtFL_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtFL_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtFL_0.ImeMode = System.Windows.Forms.ImeMode.Off
        Me.txtFL.SetIndex(Me._txtFL_0, CType(0, Short))
        Me._txtFL_0.Location = New System.Drawing.Point(57, 25)
        Me._txtFL_0.Name = "_txtFL_0"
        Me._txtFL_0.ReadOnly = True
        Me._txtFL_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtFL_0.Size = New System.Drawing.Size(49, 20)
        Me._txtFL_0.TabIndex = 39
        Me._txtFL_0.Text = "0"
        Me._txtFL_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_lblLengthUnit_7
        '
        Me._lblLengthUnit_7.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_7, CType(7, Short))
        Me._lblLengthUnit_7.Location = New System.Drawing.Point(109, 76)
        Me._lblLengthUnit_7.Name = "_lblLengthUnit_7"
        Me._lblLengthUnit_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_7.Size = New System.Drawing.Size(17, 17)
        Me._lblLengthUnit_7.TabIndex = 74
        Me._lblLengthUnit_7.Text = "ft"
        '
        '_lblLengthUnit_6
        '
        Me._lblLengthUnit_6.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_6, CType(6, Short))
        Me._lblLengthUnit_6.Location = New System.Drawing.Point(110, 51)
        Me._lblLengthUnit_6.Name = "_lblLengthUnit_6"
        Me._lblLengthUnit_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_6.Size = New System.Drawing.Size(17, 17)
        Me._lblLengthUnit_6.TabIndex = 73
        Me._lblLengthUnit_6.Text = "ft"
        '
        '_lblLengthUnit_5
        '
        Me._lblLengthUnit_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLengthUnit.SetIndex(Me._lblLengthUnit_5, CType(5, Short))
        Me._lblLengthUnit_5.Location = New System.Drawing.Point(110, 27)
        Me._lblLengthUnit_5.Name = "_lblLengthUnit_5"
        Me._lblLengthUnit_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_5.Size = New System.Drawing.Size(17, 17)
        Me._lblLengthUnit_5.TabIndex = 72
        Me._lblLengthUnit_5.Text = "ft"
        '
        '_lblFL_2
        '
        Me._lblFL_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblFL_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblFL_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFL.SetIndex(Me._lblFL_2, CType(2, Short))
        Me._lblFL_2.Location = New System.Drawing.Point(41, 73)
        Me._lblFL_2.Name = "_lblFL_2"
        Me._lblFL_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblFL_2.Size = New System.Drawing.Size(17, 17)
        Me._lblFL_2.TabIndex = 44
        Me._lblFL_2.Text = "z"
        '
        '_lblFLUnit_2
        '
        Me._lblFLUnit_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblFLUnit_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblFLUnit_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFLUnit.SetIndex(Me._lblFLUnit_2, CType(2, Short))
        Me._lblFLUnit_2.Location = New System.Drawing.Point(128, 75)
        Me._lblFLUnit_2.Name = "_lblFLUnit_2"
        Me._lblFLUnit_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblFLUnit_2.Size = New System.Drawing.Size(67, 18)
        Me._lblFLUnit_2.TabIndex = 46
        Me._lblFLUnit_2.Text = " (abv BL)"
        '
        '_lblFLCmt_1
        '
        Me._lblFLCmt_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblFLCmt_1.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFLCmt.SetIndex(Me._lblFLCmt_1, CType(1, Short))
        Me._lblFLCmt_1.Location = New System.Drawing.Point(6, 102)
        Me._lblFLCmt_1.Name = "_lblFLCmt_1"
        Me._lblFLCmt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblFLCmt_1.Size = New System.Drawing.Size(9, 17)
        Me._lblFLCmt_1.TabIndex = 47
        Me._lblFLCmt_1.Text = "*"
        '
        '_lblFLCmt_0
        '
        Me._lblFLCmt_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblFLCmt_0.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFLCmt.SetIndex(Me._lblFLCmt_0, CType(0, Short))
        Me._lblFLCmt_0.Location = New System.Drawing.Point(26, 102)
        Me._lblFLCmt_0.Name = "_lblFLCmt_0"
        Me._lblFLCmt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblFLCmt_0.Size = New System.Drawing.Size(150, 30)
        Me._lblFLCmt_0.TabIndex = 48
        Me._lblFLCmt_0.Text = "Defined in the vessel local coordinate system"
        '
        '_lblFLUnit_1
        '
        Me._lblFLUnit_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblFLUnit_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblFLUnit_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFLUnit.SetIndex(Me._lblFLUnit_1, CType(1, Short))
        Me._lblFLUnit_1.Location = New System.Drawing.Point(129, 51)
        Me._lblFLUnit_1.Name = "_lblFLUnit_1"
        Me._lblFLUnit_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblFLUnit_1.Size = New System.Drawing.Size(53, 17)
        Me._lblFLUnit_1.TabIndex = 43
        Me._lblFLUnit_1.Text = " (+, port)"
        '
        '_lblFL_1
        '
        Me._lblFL_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblFL_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblFL_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFL.SetIndex(Me._lblFL_1, CType(1, Short))
        Me._lblFL_1.Location = New System.Drawing.Point(41, 49)
        Me._lblFL_1.Name = "_lblFL_1"
        Me._lblFL_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblFL_1.Size = New System.Drawing.Size(17, 17)
        Me._lblFL_1.TabIndex = 41
        Me._lblFL_1.Text = "y"
        '
        '_lblFLUnit_0
        '
        Me._lblFLUnit_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblFLUnit_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblFLUnit_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFLUnit.SetIndex(Me._lblFLUnit_0, CType(0, Short))
        Me._lblFLUnit_0.Location = New System.Drawing.Point(128, 26)
        Me._lblFLUnit_0.Name = "_lblFLUnit_0"
        Me._lblFLUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblFLUnit_0.Size = New System.Drawing.Size(54, 17)
        Me._lblFLUnit_0.TabIndex = 40
        Me._lblFLUnit_0.Text = " (+, fwd)"
        '
        '_lblFL_0
        '
        Me._lblFL_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblFL_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblFL_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFL.SetIndex(Me._lblFL_0, CType(0, Short))
        Me._lblFL_0.Location = New System.Drawing.Point(41, 25)
        Me._lblFL_0.Name = "_lblFL_0"
        Me._lblFL_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblFL_0.Size = New System.Drawing.Size(17, 17)
        Me._lblFL_0.TabIndex = 38
        Me._lblFL_0.Text = "x"
        '
        'txtAnchor
        '
        '
        'txtFL
        '
        '
        'txtGeneral
        '
        '
        'txtVslSt
        '
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFile, Me.mnuGridEdit})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(851, 24)
        Me.MainMenu1.TabIndex = 91
        '
        'mnuFile
        '
        Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuOpen, Me.mnuSave, Me.mnuDum, Me.mnuClose})
        Me.mnuFile.Enabled = False
        Me.mnuFile.Name = "mnuFile"
        Me.mnuFile.Size = New System.Drawing.Size(37, 20)
        Me.mnuFile.Text = "&File"
        Me.mnuFile.Visible = False
        '
        'mnuOpen
        '
        Me.mnuOpen.Name = "mnuOpen"
        Me.mnuOpen.Size = New System.Drawing.Size(119, 22)
        Me.mnuOpen.Text = "&Import..."
        '
        'mnuSave
        '
        Me.mnuSave.Name = "mnuSave"
        Me.mnuSave.Size = New System.Drawing.Size(119, 22)
        Me.mnuSave.Text = "&Export..."
        '
        'mnuDum
        '
        Me.mnuDum.Name = "mnuDum"
        Me.mnuDum.Size = New System.Drawing.Size(116, 6)
        '
        'mnuClose
        '
        Me.mnuClose.Name = "mnuClose"
        Me.mnuClose.Size = New System.Drawing.Size(119, 22)
        Me.mnuClose.Text = "&Close"
        '
        'mnuGridEdit
        '
        Me.mnuGridEdit.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuInsert, Me.mnuDelete})
        Me.mnuGridEdit.Name = "mnuGridEdit"
        Me.mnuGridEdit.Size = New System.Drawing.Size(12, 20)
        Me.mnuGridEdit.Visible = False
        '
        'mnuInsert
        '
        Me.mnuInsert.Name = "mnuInsert"
        Me.mnuInsert.Size = New System.Drawing.Size(157, 22)
        Me.mnuInsert.Text = "&Insert Segment"
        '
        'mnuDelete
        '
        Me.mnuDelete.Name = "mnuDelete"
        Me.mnuDelete.Size = New System.Drawing.Size(157, 22)
        Me.mnuDelete.Text = "&Delete Segment"
        '
        'tabMoorLines
        '
        Me.tabMoorLines.Controls.Add(Me.Tab1)
        Me.tabMoorLines.Location = New System.Drawing.Point(12, 144)
        Me.tabMoorLines.Name = "tabMoorLines"
        Me.tabMoorLines.SelectedIndex = 0
        Me.tabMoorLines.Size = New System.Drawing.Size(812, 550)
        Me.tabMoorLines.TabIndex = 92
        '
        'Tab1
        '
        Me.Tab1.Controls.Add(Me.Panel1)
        Me.Tab1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Tab1.Location = New System.Drawing.Point(4, 22)
        Me.Tab1.Name = "Tab1"
        Me.Tab1.Padding = New System.Windows.Forms.Padding(3)
        Me.Tab1.Size = New System.Drawing.Size(804, 524)
        Me.Tab1.TabIndex = 0
        Me.Tab1.Text = "Line 1"
        Me.Tab1.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.fraGeneral)
        Me.Panel1.Controls.Add(Me.fraSegments)
        Me.Panel1.Controls.Add(Me.chkBatch)
        Me.Panel1.Controls.Add(Me.btnCatenary)
        Me.Panel1.Controls.Add(Me.btnRemove)
        Me.Panel1.Controls.Add(Me.fraAnchor)
        Me.Panel1.Controls.Add(Me.fraFairLead)
        Me.Panel1.Controls.Add(Me.fraTopTen)
        Me.Panel1.Location = New System.Drawing.Point(3, 3)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(795, 521)
        Me.Panel1.TabIndex = 83
        '
        'frmMoorLines
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(851, 706)
        Me.Controls.Add(Me.tabMoorLines)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.fraVesselLoc)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(10, 29)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmMoorLines"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = " Properties of Moorling Lines"
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.fraVesselLoc.ResumeLayout(False)
        Me.fraVesselLoc.PerformLayout()
        Me.fraSegments.ResumeLayout(False)
        CType(Me.grdAnchor, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdSegments, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraGeneral.ResumeLayout(False)
        Me.fraGeneral.PerformLayout()
        Me.fraTopTen.ResumeLayout(False)
        Me.fraTopTen.PerformLayout()
        Me.fraAnchor.ResumeLayout(False)
        Me.fraAnchor.PerformLayout()
        Me.fraFairLead.ResumeLayout(False)
        Me.fraFairLead.PerformLayout()
        CType(Me.lblAnchor, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblAnchorCmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblAngleUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblFL, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblFLCmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblFLUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblForceUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblGenCmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblGeneral, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblLengthUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblPreTen, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblSegCmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVelUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVslSt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtAnchor, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtFL, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtGeneral, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtPreTen, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtVslSt, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        Me.tabMoorLines.ResumeLayout(False)
        Me.Tab1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents grdSegments As DataGridView
    Friend WithEvents grdAnchor As DataGridView
    Friend WithEvents tabMoorLines As TabControl
    Friend WithEvents Tab1 As TabPage
    Public WithEvents Label10 As Label
    Public WithEvents Label12 As Label
    Public WithEvents _txtFL_4 As TextBox
    Public WithEvents _txtFL_3 As TextBox
    Friend WithEvents Panel1 As Panel
#End Region
End Class