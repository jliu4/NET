<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmMain
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
	Public WithEvents txtLocationName As System.Windows.Forms.TextBox
	Public WithEvents txtClientName As System.Windows.Forms.TextBox
	Public WithEvents lblLocation As System.Windows.Forms.Label
	Public WithEvents lblClientName As System.Windows.Forms.Label
	Public WithEvents fraReport As System.Windows.Forms.GroupBox
	Public WithEvents btnMooring As System.Windows.Forms.Button
	Public WithEvents btnEnv As System.Windows.Forms.Button
	Public dlgFileOpen As System.Windows.Forms.OpenFileDialog
	Public dlgFileSave As System.Windows.Forms.SaveFileDialog
	Public dlgFilePrint As System.Windows.Forms.PrintDialog
	Public WithEvents btnUpdate As System.Windows.Forms.Button
    Public WithEvents fraMooring As System.Windows.Forms.GroupBox
    Public WithEvents _btnPlot_1 As System.Windows.Forms.Button
    Public WithEvents _btnPlot_0 As System.Windows.Forms.Button
    Public WithEvents btnWell As System.Windows.Forms.Button
    Public WithEvents _txtWell_2 As System.Windows.Forms.TextBox
    Public WithEvents _txtWell_1 As System.Windows.Forms.TextBox
    Public WithEvents _txtWell_0 As System.Windows.Forms.TextBox
    Public WithEvents cboWells As System.Windows.Forms.ComboBox
    Public WithEvents _lblLengthUnit_2 As System.Windows.Forms.Label
    Public WithEvents _lblLengthUnit_1 As System.Windows.Forms.Label
    Public WithEvents _lblLengthUnit_0 As System.Windows.Forms.Label
    Public WithEvents _lblWell_2 As System.Windows.Forms.Label
    Public WithEvents _lblWell_1 As System.Windows.Forms.Label
    Public WithEvents _lblWell_0 As System.Windows.Forms.Label
    Public WithEvents fraWell As System.Windows.Forms.GroupBox
    Public WithEvents _optInputSystem_1 As System.Windows.Forms.RadioButton
    Public WithEvents _txtVslSt_5 As System.Windows.Forms.TextBox
    Public WithEvents _txtVslSt_4 As System.Windows.Forms.TextBox
    Public WithEvents _txtVslSt_1 As System.Windows.Forms.TextBox
    Public WithEvents _txtVslSt_0 As System.Windows.Forms.TextBox
    Public WithEvents _txtVslSt_3 As System.Windows.Forms.TextBox
    Public WithEvents _txtVslSt_2 As System.Windows.Forms.TextBox
    Public WithEvents _optInputSystem_0 As System.Windows.Forms.RadioButton
    Public WithEvents _lblLengthUnit_6 As System.Windows.Forms.Label
    Public WithEvents _lblLengthUnit_5 As System.Windows.Forms.Label
    Public WithEvents _lblLengthUnit_4 As System.Windows.Forms.Label
    Public WithEvents _lblLengthUnit_3 As System.Windows.Forms.Label
    Public WithEvents _lblVslSt_2 As System.Windows.Forms.Label
    Public WithEvents _lblVslSt_4 As System.Windows.Forms.Label
    Public WithEvents _lblVslSt_1 As System.Windows.Forms.Label
    Public WithEvents _lblVslSt_0 As System.Windows.Forms.Label
    Public WithEvents _lblVslSt_5 As System.Windows.Forms.Label
    Public WithEvents _lblVslSt_3 As System.Windows.Forms.Label
    Public WithEvents fraVesselLoc As System.Windows.Forms.GroupBox
    Public WithEvents lblVessel As System.Windows.Forms.Label
    Public WithEvents mnuFilePre As Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents optInputSystem As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    Public WithEvents mnuFileNew As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFileOpen As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuLine1 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents mnuFileSave As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFileSaveAs As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuLine2 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents mnuFilePrintSetup As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuLine3 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents _mnuFilePre_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuFilePre_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuFilePre_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuFilePre_3 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuLinePreFile As System.Windows.Forms.ToolStripSeparator
    Public WithEvents mnuFileExit As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFile As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuInpProjDes As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuLineInp As System.Windows.Forms.ToolStripSeparator
    Public WithEvents mnuInpEnviron As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuInpMoor As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuInpFC As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuInput As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnu3DPlot As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuRadarPlot As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuMoorLayoutReport As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuPlots As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuMoor As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuTransient As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuMove As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuSep As System.Windows.Forms.ToolStripSeparator
    Public WithEvents mnuAnalysesB As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuLinTools As System.Windows.Forms.ToolStripSeparator
    Public WithEvents mnuTools As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuOptions As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuHelpAbout As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuHelp As System.Windows.Forms.ToolStripMenuItem
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
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.fraReport = New System.Windows.Forms.GroupBox()
        Me.txtLocationName = New System.Windows.Forms.TextBox()
        Me.txtClientName = New System.Windows.Forms.TextBox()
        Me.lblLocation = New System.Windows.Forms.Label()
        Me.lblClientName = New System.Windows.Forms.Label()
        Me.fraMooring = New System.Windows.Forms.GroupBox()
        Me.grdLD = New System.Windows.Forms.DataGridView()
        Me.grdEL = New System.Windows.Forms.DataGridView()
        Me.btnMooring = New System.Windows.Forms.Button()
        Me.btnEnv = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.dlgFileOpen = New System.Windows.Forms.OpenFileDialog()
        Me.dlgFileSave = New System.Windows.Forms.SaveFileDialog()
        Me.dlgFilePrint = New System.Windows.Forms.PrintDialog()
        Me._btnPlot_1 = New System.Windows.Forms.Button()
        Me._btnPlot_0 = New System.Windows.Forms.Button()
        Me.fraWell = New System.Windows.Forms.GroupBox()
        Me.btnWell = New System.Windows.Forms.Button()
        Me._txtWell_2 = New System.Windows.Forms.TextBox()
        Me._txtWell_1 = New System.Windows.Forms.TextBox()
        Me._txtWell_0 = New System.Windows.Forms.TextBox()
        Me.cboWells = New System.Windows.Forms.ComboBox()
        Me._lblLengthUnit_2 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_1 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_0 = New System.Windows.Forms.Label()
        Me._lblWell_2 = New System.Windows.Forms.Label()
        Me._lblWell_1 = New System.Windows.Forms.Label()
        Me._lblWell_0 = New System.Windows.Forms.Label()
        Me.fraVesselLoc = New System.Windows.Forms.GroupBox()
        Me._optInputSystem_1 = New System.Windows.Forms.RadioButton()
        Me._txtVslSt_5 = New System.Windows.Forms.TextBox()
        Me._txtVslSt_4 = New System.Windows.Forms.TextBox()
        Me._txtVslSt_1 = New System.Windows.Forms.TextBox()
        Me._txtVslSt_0 = New System.Windows.Forms.TextBox()
        Me._txtVslSt_3 = New System.Windows.Forms.TextBox()
        Me._txtVslSt_2 = New System.Windows.Forms.TextBox()
        Me._optInputSystem_0 = New System.Windows.Forms.RadioButton()
        Me._lblLengthUnit_6 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_5 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_4 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_3 = New System.Windows.Forms.Label()
        Me._lblVslSt_2 = New System.Windows.Forms.Label()
        Me._lblVslSt_4 = New System.Windows.Forms.Label()
        Me._lblVslSt_1 = New System.Windows.Forms.Label()
        Me._lblVslSt_0 = New System.Windows.Forms.Label()
        Me._lblVslSt_5 = New System.Windows.Forms.Label()
        Me._lblVslSt_3 = New System.Windows.Forms.Label()
        Me.lblVessel = New System.Windows.Forms.Label()
        Me.mnuFilePre = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me._mnuFilePre_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFilePre_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFilePre_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFilePre_3 = New System.Windows.Forms.ToolStripMenuItem()
        Me.optInputSystem = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip()
        Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFileNew = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFileOpen = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLine1 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuFileSave = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFileSaveAs = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLine2 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuFilePrintSetup = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLine3 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuLinePreFile = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuFileExit = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuInput = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuInpProjDes = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLineInp = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuInpEnviron = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuInpMoor = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuInpFC = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuPlots = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnu3DPlot = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuRadarPlot = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuMoorLayoutReport = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuTools = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuMoor = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuTransient = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuMove = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuSep = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuAnalysesB = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLinTools = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuOptions = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHelpAbout = New System.Windows.Forms.ToolStripMenuItem()
        Me.fraReport.SuspendLayout()
        Me.fraMooring.SuspendLayout()
        CType(Me.grdLD, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdEL, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraWell.SuspendLayout()
        Me.fraVesselLoc.SuspendLayout()
        CType(Me.mnuFilePre, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optInputSystem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MainMenu1.SuspendLayout()
        Me.SuspendLayout()
        '
        'fraReport
        '
        Me.fraReport.BackColor = System.Drawing.SystemColors.Control
        Me.fraReport.Controls.Add(Me.txtLocationName)
        Me.fraReport.Controls.Add(Me.txtClientName)
        Me.fraReport.Controls.Add(Me.lblLocation)
        Me.fraReport.Controls.Add(Me.lblClientName)
        Me.fraReport.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraReport.Location = New System.Drawing.Point(7, 509)
        Me.fraReport.Name = "fraReport"
        Me.fraReport.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraReport.Size = New System.Drawing.Size(777, 57)
        Me.fraReport.TabIndex = 40
        Me.fraReport.TabStop = False
        Me.fraReport.Text = "Mooring Layout Report Titles"
        '
        'txtLocationName
        '
        Me.txtLocationName.AcceptsReturn = True
        Me.txtLocationName.BackColor = System.Drawing.SystemColors.Window
        Me.txtLocationName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLocationName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtLocationName.Location = New System.Drawing.Point(467, 23)
        Me.txtLocationName.Name = "txtLocationName"
        Me.txtLocationName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtLocationName.Size = New System.Drawing.Size(242, 20)
        Me.txtLocationName.TabIndex = 42
        Me.txtLocationName.TabStop = False
        Me.txtLocationName.Text = "Location"
        '
        'txtClientName
        '
        Me.txtClientName.AcceptsReturn = True
        Me.txtClientName.BackColor = System.Drawing.SystemColors.Window
        Me.txtClientName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtClientName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtClientName.Location = New System.Drawing.Point(84, 24)
        Me.txtClientName.Name = "txtClientName"
        Me.txtClientName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtClientName.Size = New System.Drawing.Size(270, 20)
        Me.txtClientName.TabIndex = 41
        Me.txtClientName.TabStop = False
        Me.txtClientName.Text = "Noble Drilling Services"
        '
        'lblLocation
        '
        Me.lblLocation.BackColor = System.Drawing.SystemColors.Control
        Me.lblLocation.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblLocation.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLocation.Location = New System.Drawing.Point(386, 26)
        Me.lblLocation.Name = "lblLocation"
        Me.lblLocation.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblLocation.Size = New System.Drawing.Size(87, 17)
        Me.lblLocation.TabIndex = 44
        Me.lblLocation.Text = "Location Name:"
        '
        'lblClientName
        '
        Me.lblClientName.BackColor = System.Drawing.SystemColors.Control
        Me.lblClientName.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblClientName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblClientName.Location = New System.Drawing.Point(5, 27)
        Me.lblClientName.Name = "lblClientName"
        Me.lblClientName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblClientName.Size = New System.Drawing.Size(73, 17)
        Me.lblClientName.TabIndex = 43
        Me.lblClientName.Text = "Prepared for"
        '
        'fraMooring
        '
        Me.fraMooring.BackColor = System.Drawing.SystemColors.Control
        Me.fraMooring.Controls.Add(Me.grdLD)
        Me.fraMooring.Controls.Add(Me.grdEL)
        Me.fraMooring.Controls.Add(Me.btnMooring)
        Me.fraMooring.Controls.Add(Me.btnEnv)
        Me.fraMooring.Controls.Add(Me.btnUpdate)
        Me.fraMooring.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraMooring.Location = New System.Drawing.Point(7, 154)
        Me.fraMooring.Name = "fraMooring"
        Me.fraMooring.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraMooring.Size = New System.Drawing.Size(810, 349)
        Me.fraMooring.TabIndex = 29
        Me.fraMooring.TabStop = False
        Me.fraMooring.Text = "Mooring System Summary"
        '
        'grdLD
        '
        Me.grdLD.AllowUserToAddRows = False
        Me.grdLD.AllowUserToDeleteRows = False
        Me.grdLD.AllowUserToOrderColumns = True
        Me.grdLD.AllowUserToResizeColumns = False
        Me.grdLD.AllowUserToResizeRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.grdLD.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdLD.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.grdLD.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdLD.Cursor = System.Windows.Forms.Cursors.Default
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grdLD.DefaultCellStyle = DataGridViewCellStyle3
        Me.grdLD.Location = New System.Drawing.Point(6, 19)
        Me.grdLD.MultiSelect = False
        Me.grdLD.Name = "grdLD"
        Me.grdLD.ReadOnly = True
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grdLD.RowHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.grdLD.RowHeadersVisible = False
        Me.grdLD.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.grdLD.RowsDefaultCellStyle = DataGridViewCellStyle5
        Me.grdLD.RowTemplate.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.grdLD.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.grdLD.Size = New System.Drawing.Size(798, 255)
        Me.grdLD.TabIndex = 51
        '
        'grdEL
        '
        Me.grdEL.AllowUserToAddRows = False
        Me.grdEL.AllowUserToDeleteRows = False
        Me.grdEL.AllowUserToResizeRows = False
        Me.grdEL.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdEL.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle6
        Me.grdEL.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        DataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grdEL.DefaultCellStyle = DataGridViewCellStyle7
        Me.grdEL.Location = New System.Drawing.Point(6, 289)
        Me.grdEL.Name = "grdEL"
        Me.grdEL.ReadOnly = True
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        DataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdEL.RowHeadersDefaultCellStyle = DataGridViewCellStyle8
        Me.grdEL.RowHeadersVisible = False
        Me.grdEL.RowHeadersWidth = 150
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.grdEL.RowsDefaultCellStyle = DataGridViewCellStyle9
        Me.grdEL.Size = New System.Drawing.Size(381, 36)
        Me.grdEL.TabIndex = 50
        '
        'btnMooring
        '
        Me.btnMooring.BackColor = System.Drawing.SystemColors.Control
        Me.btnMooring.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnMooring.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnMooring.Location = New System.Drawing.Point(631, 300)
        Me.btnMooring.Name = "btnMooring"
        Me.btnMooring.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnMooring.Size = New System.Drawing.Size(79, 25)
        Me.btnMooring.TabIndex = 46
        Me.btnMooring.Text = "Mooring..."
        Me.btnMooring.UseVisualStyleBackColor = False
        '
        'btnEnv
        '
        Me.btnEnv.BackColor = System.Drawing.SystemColors.Control
        Me.btnEnv.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnEnv.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnEnv.Location = New System.Drawing.Point(532, 300)
        Me.btnEnv.Name = "btnEnv"
        Me.btnEnv.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnEnv.Size = New System.Drawing.Size(79, 25)
        Me.btnEnv.TabIndex = 45
        Me.btnEnv.Text = "Enviroment..."
        Me.btnEnv.UseVisualStyleBackColor = False
        '
        'btnUpdate
        '
        Me.btnUpdate.BackColor = System.Drawing.SystemColors.Control
        Me.btnUpdate.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnUpdate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnUpdate.Location = New System.Drawing.Point(725, 300)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnUpdate.Size = New System.Drawing.Size(79, 25)
        Me.btnUpdate.TabIndex = 0
        Me.btnUpdate.Text = "Refresh"
        Me.btnUpdate.UseVisualStyleBackColor = False
        '
        '_btnPlot_1
        '
        Me._btnPlot_1.BackColor = System.Drawing.SystemColors.Control
        Me._btnPlot_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnPlot_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnPlot_1.Location = New System.Drawing.Point(285, 27)
        Me._btnPlot_1.Name = "_btnPlot_1"
        Me._btnPlot_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnPlot_1.Size = New System.Drawing.Size(84, 25)
        Me._btnPlot_1.TabIndex = 3
        Me._btnPlot_1.Text = "Radar Plot"
        Me._btnPlot_1.UseVisualStyleBackColor = False
        '
        '_btnPlot_0
        '
        Me._btnPlot_0.BackColor = System.Drawing.SystemColors.Control
        Me._btnPlot_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnPlot_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnPlot_0.Location = New System.Drawing.Point(218, 28)
        Me._btnPlot_0.Name = "_btnPlot_0"
        Me._btnPlot_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnPlot_0.Size = New System.Drawing.Size(61, 25)
        Me._btnPlot_0.TabIndex = 2
        Me._btnPlot_0.Text = "3D Plot"
        Me._btnPlot_0.UseVisualStyleBackColor = False
        '
        'fraWell
        '
        Me.fraWell.BackColor = System.Drawing.SystemColors.Control
        Me.fraWell.Controls.Add(Me.btnWell)
        Me.fraWell.Controls.Add(Me._txtWell_2)
        Me.fraWell.Controls.Add(Me._txtWell_1)
        Me.fraWell.Controls.Add(Me._txtWell_0)
        Me.fraWell.Controls.Add(Me.cboWells)
        Me.fraWell.Controls.Add(Me._lblLengthUnit_2)
        Me.fraWell.Controls.Add(Me._lblLengthUnit_1)
        Me.fraWell.Controls.Add(Me._lblLengthUnit_0)
        Me.fraWell.Controls.Add(Me._lblWell_2)
        Me.fraWell.Controls.Add(Me._lblWell_1)
        Me.fraWell.Controls.Add(Me._lblWell_0)
        Me.fraWell.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraWell.Location = New System.Drawing.Point(7, 58)
        Me.fraWell.Name = "fraWell"
        Me.fraWell.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraWell.Size = New System.Drawing.Size(362, 89)
        Me.fraWell.TabIndex = 4
        Me.fraWell.TabStop = False
        Me.fraWell.Text = "Current Well Site"
        '
        'btnWell
        '
        Me.btnWell.BackColor = System.Drawing.SystemColors.Control
        Me.btnWell.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnWell.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnWell.Location = New System.Drawing.Point(121, 50)
        Me.btnWell.Name = "btnWell"
        Me.btnWell.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnWell.Size = New System.Drawing.Size(46, 25)
        Me.btnWell.TabIndex = 6
        Me.btnWell.Text = "Edit..."
        Me.btnWell.UseVisualStyleBackColor = False
        '
        '_txtWell_2
        '
        Me._txtWell_2.AcceptsReturn = True
        Me._txtWell_2.BackColor = System.Drawing.SystemColors.Control
        Me._txtWell_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtWell_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtWell_2.Location = New System.Drawing.Point(263, 63)
        Me._txtWell_2.Name = "_txtWell_2"
        Me._txtWell_2.ReadOnly = True
        Me._txtWell_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtWell_2.Size = New System.Drawing.Size(64, 20)
        Me._txtWell_2.TabIndex = 12
        Me._txtWell_2.TabStop = False
        Me._txtWell_2.Text = "0"
        Me._txtWell_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtWell_1
        '
        Me._txtWell_1.AcceptsReturn = True
        Me._txtWell_1.BackColor = System.Drawing.SystemColors.Control
        Me._txtWell_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtWell_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtWell_1.Location = New System.Drawing.Point(239, 39)
        Me._txtWell_1.Name = "_txtWell_1"
        Me._txtWell_1.ReadOnly = True
        Me._txtWell_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtWell_1.Size = New System.Drawing.Size(88, 20)
        Me._txtWell_1.TabIndex = 10
        Me._txtWell_1.TabStop = False
        Me._txtWell_1.Text = "0"
        Me._txtWell_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtWell_0
        '
        Me._txtWell_0.AcceptsReturn = True
        Me._txtWell_0.BackColor = System.Drawing.SystemColors.Control
        Me._txtWell_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtWell_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtWell_0.Location = New System.Drawing.Point(239, 16)
        Me._txtWell_0.Name = "_txtWell_0"
        Me._txtWell_0.ReadOnly = True
        Me._txtWell_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtWell_0.Size = New System.Drawing.Size(88, 20)
        Me._txtWell_0.TabIndex = 8
        Me._txtWell_0.TabStop = False
        Me._txtWell_0.Text = "0"
        Me._txtWell_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cboWells
        '
        Me.cboWells.BackColor = System.Drawing.SystemColors.Window
        Me.cboWells.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboWells.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboWells.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboWells.Location = New System.Drawing.Point(13, 24)
        Me.cboWells.Name = "cboWells"
        Me.cboWells.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboWells.Size = New System.Drawing.Size(154, 21)
        Me.cboWells.TabIndex = 5
        '
        '_lblLengthUnit_2
        '
        Me._lblLengthUnit_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblLengthUnit_2.Location = New System.Drawing.Point(333, 64)
        Me._lblLengthUnit_2.Name = "_lblLengthUnit_2"
        Me._lblLengthUnit_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_2.Size = New System.Drawing.Size(21, 13)
        Me._lblLengthUnit_2.TabIndex = 34
        Me._lblLengthUnit_2.Text = "ft"
        '
        '_lblLengthUnit_1
        '
        Me._lblLengthUnit_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblLengthUnit_1.Location = New System.Drawing.Point(333, 42)
        Me._lblLengthUnit_1.Name = "_lblLengthUnit_1"
        Me._lblLengthUnit_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_1.Size = New System.Drawing.Size(21, 17)
        Me._lblLengthUnit_1.TabIndex = 33
        Me._lblLengthUnit_1.Text = "ft"
        '
        '_lblLengthUnit_0
        '
        Me._lblLengthUnit_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblLengthUnit_0.Location = New System.Drawing.Point(333, 16)
        Me._lblLengthUnit_0.Name = "_lblLengthUnit_0"
        Me._lblLengthUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_0.Size = New System.Drawing.Size(21, 18)
        Me._lblLengthUnit_0.TabIndex = 32
        Me._lblLengthUnit_0.Text = "ft"
        '
        '_lblWell_2
        '
        Me._lblWell_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblWell_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblWell_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblWell_2.Location = New System.Drawing.Point(187, 64)
        Me._lblWell_2.Name = "_lblWell_2"
        Me._lblWell_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblWell_2.Size = New System.Drawing.Size(85, 17)
        Me._lblWell_2.TabIndex = 11
        Me._lblWell_2.Text = "Water Depth"
        '
        '_lblWell_1
        '
        Me._lblWell_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblWell_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblWell_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblWell_1.Location = New System.Drawing.Point(187, 39)
        Me._lblWell_1.Name = "_lblWell_1"
        Me._lblWell_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblWell_1.Size = New System.Drawing.Size(33, 17)
        Me._lblWell_1.TabIndex = 9
        Me._lblWell_1.Text = "y (N)"
        '
        '_lblWell_0
        '
        Me._lblWell_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblWell_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblWell_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblWell_0.Location = New System.Drawing.Point(189, 17)
        Me._lblWell_0.Name = "_lblWell_0"
        Me._lblWell_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblWell_0.Size = New System.Drawing.Size(33, 17)
        Me._lblWell_0.TabIndex = 7
        Me._lblWell_0.Text = "x (E)"
        '
        'fraVesselLoc
        '
        Me.fraVesselLoc.BackColor = System.Drawing.SystemColors.Control
        Me.fraVesselLoc.Controls.Add(Me._optInputSystem_1)
        Me.fraVesselLoc.Controls.Add(Me._txtVslSt_5)
        Me.fraVesselLoc.Controls.Add(Me._txtVslSt_4)
        Me.fraVesselLoc.Controls.Add(Me._txtVslSt_1)
        Me.fraVesselLoc.Controls.Add(Me._txtVslSt_0)
        Me.fraVesselLoc.Controls.Add(Me._txtVslSt_3)
        Me.fraVesselLoc.Controls.Add(Me._txtVslSt_2)
        Me.fraVesselLoc.Controls.Add(Me._optInputSystem_0)
        Me.fraVesselLoc.Controls.Add(Me._lblLengthUnit_6)
        Me.fraVesselLoc.Controls.Add(Me._lblLengthUnit_5)
        Me.fraVesselLoc.Controls.Add(Me._lblLengthUnit_4)
        Me.fraVesselLoc.Controls.Add(Me._lblLengthUnit_3)
        Me.fraVesselLoc.Controls.Add(Me._lblVslSt_2)
        Me.fraVesselLoc.Controls.Add(Me._lblVslSt_4)
        Me.fraVesselLoc.Controls.Add(Me._lblVslSt_1)
        Me.fraVesselLoc.Controls.Add(Me._lblVslSt_0)
        Me.fraVesselLoc.Controls.Add(Me._lblVslSt_5)
        Me.fraVesselLoc.Controls.Add(Me._lblVslSt_3)
        Me.fraVesselLoc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraVesselLoc.Location = New System.Drawing.Point(407, 28)
        Me.fraVesselLoc.Name = "fraVesselLoc"
        Me.fraVesselLoc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraVesselLoc.Size = New System.Drawing.Size(409, 121)
        Me.fraVesselLoc.TabIndex = 13
        Me.fraVesselLoc.TabStop = False
        Me.fraVesselLoc.Text = "Current Vessel Station :"
        '
        '_optInputSystem_1
        '
        Me._optInputSystem_1.BackColor = System.Drawing.SystemColors.Control
        Me._optInputSystem_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optInputSystem_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optInputSystem.SetIndex(Me._optInputSystem_1, CType(1, Short))
        Me._optInputSystem_1.Location = New System.Drawing.Point(16, 56)
        Me._optInputSystem_1.Name = "_optInputSystem_1"
        Me._optInputSystem_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optInputSystem_1.Size = New System.Drawing.Size(65, 17)
        Me._optInputSystem_1.TabIndex = 19
        Me._optInputSystem_1.TabStop = True
        Me._optInputSystem_1.Text = "To Well:"
        Me._optInputSystem_1.UseVisualStyleBackColor = False
        '
        '_txtVslSt_5
        '
        Me._txtVslSt_5.AcceptsReturn = True
        Me._txtVslSt_5.BackColor = System.Drawing.SystemColors.Window
        Me._txtVslSt_5.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtVslSt_5.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtVslSt_5.Location = New System.Drawing.Point(319, 84)
        Me._txtVslSt_5.Name = "_txtVslSt_5"
        Me._txtVslSt_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtVslSt_5.Size = New System.Drawing.Size(49, 20)
        Me._txtVslSt_5.TabIndex = 28
        Me._txtVslSt_5.TabStop = False
        Me._txtVslSt_5.Text = "0"
        Me._txtVslSt_5.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtVslSt_4
        '
        Me._txtVslSt_4.AcceptsReturn = True
        Me._txtVslSt_4.BackColor = System.Drawing.SystemColors.Window
        Me._txtVslSt_4.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtVslSt_4.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtVslSt_4.Location = New System.Drawing.Point(141, 84)
        Me._txtVslSt_4.Name = "_txtVslSt_4"
        Me._txtVslSt_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtVslSt_4.Size = New System.Drawing.Size(71, 20)
        Me._txtVslSt_4.TabIndex = 25
        Me._txtVslSt_4.TabStop = False
        Me._txtVslSt_4.Text = "0"
        Me._txtVslSt_4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtVslSt_1
        '
        Me._txtVslSt_1.AcceptsReturn = True
        Me._txtVslSt_1.BackColor = System.Drawing.SystemColors.Window
        Me._txtVslSt_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtVslSt_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtVslSt_1.Location = New System.Drawing.Point(289, 24)
        Me._txtVslSt_1.Name = "_txtVslSt_1"
        Me._txtVslSt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtVslSt_1.Size = New System.Drawing.Size(79, 20)
        Me._txtVslSt_1.TabIndex = 18
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
        Me._txtVslSt_0.Location = New System.Drawing.Point(141, 24)
        Me._txtVslSt_0.Name = "_txtVslSt_0"
        Me._txtVslSt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtVslSt_0.Size = New System.Drawing.Size(71, 20)
        Me._txtVslSt_0.TabIndex = 16
        Me._txtVslSt_0.TabStop = False
        Me._txtVslSt_0.Text = "0"
        Me._txtVslSt_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtVslSt_3
        '
        Me._txtVslSt_3.AcceptsReturn = True
        Me._txtVslSt_3.BackColor = System.Drawing.SystemColors.Window
        Me._txtVslSt_3.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtVslSt_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtVslSt_3.Location = New System.Drawing.Point(319, 52)
        Me._txtVslSt_3.Name = "_txtVslSt_3"
        Me._txtVslSt_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtVslSt_3.Size = New System.Drawing.Size(49, 20)
        Me._txtVslSt_3.TabIndex = 23
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
        Me._txtVslSt_2.Location = New System.Drawing.Point(141, 54)
        Me._txtVslSt_2.Name = "_txtVslSt_2"
        Me._txtVslSt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtVslSt_2.Size = New System.Drawing.Size(71, 20)
        Me._txtVslSt_2.TabIndex = 21
        Me._txtVslSt_2.TabStop = False
        Me._txtVslSt_2.Text = "0"
        Me._txtVslSt_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_optInputSystem_0
        '
        Me._optInputSystem_0.BackColor = System.Drawing.SystemColors.Control
        Me._optInputSystem_0.Checked = True
        Me._optInputSystem_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optInputSystem_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optInputSystem.SetIndex(Me._optInputSystem_0, CType(0, Short))
        Me._optInputSystem_0.Location = New System.Drawing.Point(16, 24)
        Me._optInputSystem_0.Name = "_optInputSystem_0"
        Me._optInputSystem_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optInputSystem_0.Size = New System.Drawing.Size(92, 17)
        Me._optInputSystem_0.TabIndex = 14
        Me._optInputSystem_0.TabStop = True
        Me._optInputSystem_0.Text = "At Coordinate:"
        Me._optInputSystem_0.UseVisualStyleBackColor = False
        '
        '_lblLengthUnit_6
        '
        Me._lblLengthUnit_6.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblLengthUnit_6.Location = New System.Drawing.Point(377, 89)
        Me._lblLengthUnit_6.Name = "_lblLengthUnit_6"
        Me._lblLengthUnit_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_6.Size = New System.Drawing.Size(22, 17)
        Me._lblLengthUnit_6.TabIndex = 38
        Me._lblLengthUnit_6.Text = "ft"
        '
        '_lblLengthUnit_5
        '
        Me._lblLengthUnit_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblLengthUnit_5.Location = New System.Drawing.Point(216, 55)
        Me._lblLengthUnit_5.Name = "_lblLengthUnit_5"
        Me._lblLengthUnit_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_5.Size = New System.Drawing.Size(14, 16)
        Me._lblLengthUnit_5.TabIndex = 37
        Me._lblLengthUnit_5.Text = "ft"
        '
        '_lblLengthUnit_4
        '
        Me._lblLengthUnit_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblLengthUnit_4.Location = New System.Drawing.Point(377, 28)
        Me._lblLengthUnit_4.Name = "_lblLengthUnit_4"
        Me._lblLengthUnit_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_4.Size = New System.Drawing.Size(16, 15)
        Me._lblLengthUnit_4.TabIndex = 36
        Me._lblLengthUnit_4.Text = "ft"
        '
        '_lblLengthUnit_3
        '
        Me._lblLengthUnit_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblLengthUnit_3.Location = New System.Drawing.Point(217, 26)
        Me._lblLengthUnit_3.Name = "_lblLengthUnit_3"
        Me._lblLengthUnit_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_3.Size = New System.Drawing.Size(13, 17)
        Me._lblLengthUnit_3.TabIndex = 35
        Me._lblLengthUnit_3.Text = "ft"
        '
        '_lblVslSt_2
        '
        Me._lblVslSt_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblVslSt_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVslSt_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVslSt_2.Location = New System.Drawing.Point(93, 56)
        Me._lblVslSt_2.Name = "_lblVslSt_2"
        Me._lblVslSt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVslSt_2.Size = New System.Drawing.Size(49, 17)
        Me._lblVslSt_2.TabIndex = 20
        Me._lblVslSt_2.Text = "Distance "
        '
        '_lblVslSt_4
        '
        Me._lblVslSt_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblVslSt_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVslSt_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVslSt_4.Location = New System.Drawing.Point(16, 88)
        Me._lblVslSt_4.Name = "_lblVslSt_4"
        Me._lblVslSt_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVslSt_4.Size = New System.Drawing.Size(119, 16)
        Me._lblVslSt_4.TabIndex = 24
        Me._lblVslSt_4.Text = "With Vessel Heading"
        '
        '_lblVslSt_1
        '
        Me._lblVslSt_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblVslSt_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVslSt_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVslSt_1.Location = New System.Drawing.Point(250, 24)
        Me._lblVslSt_1.Name = "_lblVslSt_1"
        Me._lblVslSt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVslSt_1.Size = New System.Drawing.Size(33, 17)
        Me._lblVslSt_1.TabIndex = 17
        Me._lblVslSt_1.Text = "y (N)"
        '
        '_lblVslSt_0
        '
        Me._lblVslSt_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblVslSt_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVslSt_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVslSt_0.Location = New System.Drawing.Point(109, 26)
        Me._lblVslSt_0.Name = "_lblVslSt_0"
        Me._lblVslSt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVslSt_0.Size = New System.Drawing.Size(33, 17)
        Me._lblVslSt_0.TabIndex = 15
        Me._lblVslSt_0.Text = "x (E)"
        '
        '_lblVslSt_5
        '
        Me._lblVslSt_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblVslSt_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVslSt_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVslSt_5.Location = New System.Drawing.Point(275, 87)
        Me._lblVslSt_5.Name = "_lblVslSt_5"
        Me._lblVslSt_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVslSt_5.Size = New System.Drawing.Size(41, 17)
        Me._lblVslSt_5.TabIndex = 27
        Me._lblVslSt_5.Text = "Draft"
        '
        '_lblVslSt_3
        '
        Me._lblVslSt_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblVslSt_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVslSt_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVslSt_3.Location = New System.Drawing.Point(263, 54)
        Me._lblVslSt_3.Name = "_lblVslSt_3"
        Me._lblVslSt_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVslSt_3.Size = New System.Drawing.Size(50, 18)
        Me._lblVslSt_3.TabIndex = 22
        Me._lblVslSt_3.Text = "Bearing"
        '
        'lblVessel
        '
        Me.lblVessel.BackColor = System.Drawing.SystemColors.Control
        Me.lblVessel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblVessel.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblVessel.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVessel.Location = New System.Drawing.Point(7, 27)
        Me.lblVessel.Name = "lblVessel"
        Me.lblVessel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblVessel.Size = New System.Drawing.Size(203, 25)
        Me.lblVessel.TabIndex = 1
        Me.lblVessel.Text = "Diamond Baroness"
        Me.lblVessel.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_mnuFilePre_0
        '
        Me._mnuFilePre_0.Enabled = False
        Me.mnuFilePre.SetIndex(Me._mnuFilePre_0, CType(0, Short))
        Me._mnuFilePre_0.Name = "_mnuFilePre_0"
        Me._mnuFilePre_0.Size = New System.Drawing.Size(154, 22)
        Me._mnuFilePre_0.Visible = False
        '
        '_mnuFilePre_1
        '
        Me._mnuFilePre_1.Enabled = False
        Me.mnuFilePre.SetIndex(Me._mnuFilePre_1, CType(1, Short))
        Me._mnuFilePre_1.Name = "_mnuFilePre_1"
        Me._mnuFilePre_1.Size = New System.Drawing.Size(154, 22)
        Me._mnuFilePre_1.Visible = False
        '
        '_mnuFilePre_2
        '
        Me._mnuFilePre_2.Enabled = False
        Me.mnuFilePre.SetIndex(Me._mnuFilePre_2, CType(2, Short))
        Me._mnuFilePre_2.Name = "_mnuFilePre_2"
        Me._mnuFilePre_2.Size = New System.Drawing.Size(154, 22)
        Me._mnuFilePre_2.Visible = False
        '
        '_mnuFilePre_3
        '
        Me._mnuFilePre_3.Enabled = False
        Me.mnuFilePre.SetIndex(Me._mnuFilePre_3, CType(3, Short))
        Me._mnuFilePre_3.Name = "_mnuFilePre_3"
        Me._mnuFilePre_3.Size = New System.Drawing.Size(154, 22)
        Me._mnuFilePre_3.Visible = False
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFile, Me.mnuInput, Me.mnuPlots, Me.mnuTools, Me.mnuOptions, Me.mnuHelp})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(828, 24)
        Me.MainMenu1.TabIndex = 49
        '
        'mnuFile
        '
        Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFileNew, Me.mnuFileOpen, Me.mnuLine1, Me.mnuFileSave, Me.mnuFileSaveAs, Me.mnuLine2, Me.mnuFilePrintSetup, Me.mnuLine3, Me._mnuFilePre_0, Me._mnuFilePre_1, Me._mnuFilePre_2, Me._mnuFilePre_3, Me.mnuLinePreFile, Me.mnuFileExit})
        Me.mnuFile.Name = "mnuFile"
        Me.mnuFile.Size = New System.Drawing.Size(56, 20)
        Me.mnuFile.Text = "&Project"
        '
        'mnuFileNew
        '
        Me.mnuFileNew.Name = "mnuFileNew"
        Me.mnuFileNew.Size = New System.Drawing.Size(154, 22)
        Me.mnuFileNew.Text = "&New"
        '
        'mnuFileOpen
        '
        Me.mnuFileOpen.Name = "mnuFileOpen"
        Me.mnuFileOpen.Size = New System.Drawing.Size(154, 22)
        Me.mnuFileOpen.Text = "&Open ..."
        '
        'mnuLine1
        '
        Me.mnuLine1.Name = "mnuLine1"
        Me.mnuLine1.Size = New System.Drawing.Size(151, 6)
        '
        'mnuFileSave
        '
        Me.mnuFileSave.Name = "mnuFileSave"
        Me.mnuFileSave.Size = New System.Drawing.Size(154, 22)
        Me.mnuFileSave.Text = "&Save"
        '
        'mnuFileSaveAs
        '
        Me.mnuFileSaveAs.Name = "mnuFileSaveAs"
        Me.mnuFileSaveAs.Size = New System.Drawing.Size(154, 22)
        Me.mnuFileSaveAs.Text = "Save &As..."
        '
        'mnuLine2
        '
        Me.mnuLine2.Name = "mnuLine2"
        Me.mnuLine2.Size = New System.Drawing.Size(151, 6)
        '
        'mnuFilePrintSetup
        '
        Me.mnuFilePrintSetup.Name = "mnuFilePrintSetup"
        Me.mnuFilePrintSetup.Size = New System.Drawing.Size(154, 22)
        Me.mnuFilePrintSetup.Text = "&Printer Setup ..."
        '
        'mnuLine3
        '
        Me.mnuLine3.Name = "mnuLine3"
        Me.mnuLine3.Size = New System.Drawing.Size(151, 6)
        '
        'mnuLinePreFile
        '
        Me.mnuLinePreFile.Name = "mnuLinePreFile"
        Me.mnuLinePreFile.Size = New System.Drawing.Size(151, 6)
        Me.mnuLinePreFile.Visible = False
        '
        'mnuFileExit
        '
        Me.mnuFileExit.Name = "mnuFileExit"
        Me.mnuFileExit.Size = New System.Drawing.Size(154, 22)
        Me.mnuFileExit.Text = "E&xit"
        '
        'mnuInput
        '
        Me.mnuInput.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuInpProjDes, Me.mnuLineInp, Me.mnuInpEnviron, Me.mnuInpMoor, Me.mnuInpFC})
        Me.mnuInput.Name = "mnuInput"
        Me.mnuInput.Size = New System.Drawing.Size(39, 20)
        Me.mnuInput.Text = "&Edit"
        '
        'mnuInpProjDes
        '
        Me.mnuInpProjDes.Name = "mnuInpProjDes"
        Me.mnuInpProjDes.Size = New System.Drawing.Size(212, 22)
        Me.mnuInpProjDes.Text = "Vessel..."
        '
        'mnuLineInp
        '
        Me.mnuLineInp.Name = "mnuLineInp"
        Me.mnuLineInp.Size = New System.Drawing.Size(209, 6)
        '
        'mnuInpEnviron
        '
        Me.mnuInpEnviron.Name = "mnuInpEnviron"
        Me.mnuInpEnviron.Size = New System.Drawing.Size(212, 22)
        Me.mnuInpEnviron.Text = "&Environment..."
        '
        'mnuInpMoor
        '
        Me.mnuInpMoor.Name = "mnuInpMoor"
        Me.mnuInpMoor.Size = New System.Drawing.Size(212, 22)
        Me.mnuInpMoor.Text = "&Mooring System..."
        '
        'mnuInpFC
        '
        Me.mnuInpFC.Enabled = False
        Me.mnuInpFC.Name = "mnuInpFC"
        Me.mnuInpFC.Size = New System.Drawing.Size(212, 22)
        Me.mnuInpFC.Text = "&Vessel Force Coefficients..."
        Me.mnuInpFC.Visible = False
        '
        'mnuPlots
        '
        Me.mnuPlots.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnu3DPlot, Me.mnuRadarPlot, Me.mnuMoorLayoutReport})
        Me.mnuPlots.Name = "mnuPlots"
        Me.mnuPlots.Size = New System.Drawing.Size(44, 20)
        Me.mnuPlots.Text = "&View"
        '
        'mnu3DPlot
        '
        Me.mnu3DPlot.Name = "mnu3DPlot"
        Me.mnu3DPlot.Size = New System.Drawing.Size(206, 22)
        Me.mnu3DPlot.Text = "&3D View..."
        '
        'mnuRadarPlot
        '
        Me.mnuRadarPlot.Name = "mnuRadarPlot"
        Me.mnuRadarPlot.Size = New System.Drawing.Size(206, 22)
        Me.mnuRadarPlot.Text = "&Radar View..."
        '
        'mnuMoorLayoutReport
        '
        Me.mnuMoorLayoutReport.Name = "mnuMoorLayoutReport"
        Me.mnuMoorLayoutReport.Size = New System.Drawing.Size(206, 22)
        Me.mnuMoorLayoutReport.Text = "Mooring Layout Report..."
        '
        'mnuTools
        '
        Me.mnuTools.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuMoor, Me.mnuTransient, Me.mnuMove, Me.mnuSep, Me.mnuAnalysesB, Me.mnuLinTools})
        Me.mnuTools.Name = "mnuTools"
        Me.mnuTools.Size = New System.Drawing.Size(65, 20)
        Me.mnuTools.Text = "&Analyses"
        '
        'mnuMoor
        '
        Me.mnuMoor.Name = "mnuMoor"
        Me.mnuMoor.Size = New System.Drawing.Size(258, 22)
        Me.mnuMoor.Text = "Mooring &Analysis..."
        '
        'mnuTransient
        '
        Me.mnuTransient.Name = "mnuTransient"
        Me.mnuTransient.Size = New System.Drawing.Size(258, 22)
        Me.mnuTransient.Text = "&Transient Analysis..."
        '
        'mnuMove
        '
        Me.mnuMove.Name = "mnuMove"
        Me.mnuMove.Size = New System.Drawing.Size(258, 22)
        Me.mnuMove.Text = "Vessel &Moving Around..."
        '
        'mnuSep
        '
        Me.mnuSep.Name = "mnuSep"
        Me.mnuSep.Size = New System.Drawing.Size(255, 6)
        '
        'mnuAnalysesB
        '
        Me.mnuAnalysesB.Name = "mnuAnalysesB"
        Me.mnuAnalysesB.Size = New System.Drawing.Size(258, 22)
        Me.mnuAnalysesB.Text = "&Collinear Multi-heading Analyses..."
        '
        'mnuLinTools
        '
        Me.mnuLinTools.Name = "mnuLinTools"
        Me.mnuLinTools.Size = New System.Drawing.Size(255, 6)
        '
        'mnuOptions
        '
        Me.mnuOptions.Name = "mnuOptions"
        Me.mnuOptions.Size = New System.Drawing.Size(61, 20)
        Me.mnuOptions.Text = "&Options"
        '
        'mnuHelp
        '
        Me.mnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuHelpAbout})
        Me.mnuHelp.Name = "mnuHelp"
        Me.mnuHelp.Size = New System.Drawing.Size(44, 20)
        Me.mnuHelp.Text = "&Help"
        '
        'mnuHelpAbout
        '
        Me.mnuHelpAbout.Name = "mnuHelpAbout"
        Me.mnuHelpAbout.Size = New System.Drawing.Size(151, 22)
        Me.mnuHelpAbout.Text = "&About MARS..."
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(828, 578)
        Me.Controls.Add(Me.fraReport)
        Me.Controls.Add(Me.fraMooring)
        Me.Controls.Add(Me._btnPlot_1)
        Me.Controls.Add(Me._btnPlot_0)
        Me.Controls.Add(Me.fraWell)
        Me.Controls.Add(Me.fraVesselLoc)
        Me.Controls.Add(Me.lblVessel)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(10, 29)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmMain"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = " MARS Console"
        Me.fraReport.ResumeLayout(False)
        Me.fraReport.PerformLayout()
        Me.fraMooring.ResumeLayout(False)
        CType(Me.grdLD, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdEL, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraWell.ResumeLayout(False)
        Me.fraWell.PerformLayout()
        Me.fraVesselLoc.ResumeLayout(False)
        Me.fraVesselLoc.PerformLayout()
        CType(Me.mnuFilePre, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optInputSystem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents grdEL As DataGridView
    Friend WithEvents grdLD As DataGridView
#End Region
End Class