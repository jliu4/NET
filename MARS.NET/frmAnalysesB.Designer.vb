<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmAnalysesB
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
	Public WithEvents txtNumLinesPerGroup As System.Windows.Forms.TextBox
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents fraMoorGroup As System.Windows.Forms.GroupBox
	Public WithEvents txtSubTitle As System.Windows.Forms.TextBox
	Public WithEvents txtReportTitle As System.Windows.Forms.TextBox
	Public WithEvents txtInitYaw As System.Windows.Forms.TextBox
	Public WithEvents txtInitSway As System.Windows.Forms.TextBox
	Public WithEvents txtInitSurge As System.Windows.Forms.TextBox
	Public WithEvents _lblInitPos_2 As System.Windows.Forms.Label
	Public WithEvents _lblInitPos_1 As System.Windows.Forms.Label
	Public WithEvents _lblInitPos_0 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_4 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_3 As System.Windows.Forms.Label
	Public WithEvents _lblAngleUnit_3 As System.Windows.Forms.Label
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents btnReport As System.Windows.Forms.Button
	Public WithEvents lblLocation As System.Windows.Forms.Label
	Public WithEvents lblClientName As System.Windows.Forms.Label
	Public WithEvents fraReport As System.Windows.Forms.GroupBox
	Public dlgFileOpen As System.Windows.Forms.OpenFileDialog
	Public dlgFileSave As System.Windows.Forms.SaveFileDialog
	Public WithEvents btnBrowse As System.Windows.Forms.Button
	Public WithEvents txtFile As System.Windows.Forms.TextBox
	Public WithEvents fraFile As System.Windows.Forms.GroupBox
	Public WithEvents _txtHeadings_2 As System.Windows.Forms.TextBox
	Public WithEvents _txtHeadings_1 As System.Windows.Forms.TextBox
	Public WithEvents _txtHeadings_0 As System.Windows.Forms.TextBox
	Public WithEvents _lblAngleUnit_2 As System.Windows.Forms.Label
	Public WithEvents _lblAngleUnit_1 As System.Windows.Forms.Label
	Public WithEvents _lblHeadings_3 As System.Windows.Forms.Label
	Public WithEvents _lblHeadings_2 As System.Windows.Forms.Label
	Public WithEvents _lblHeadings_1 As System.Windows.Forms.Label
	Public WithEvents _lblHeadings_0 As System.Windows.Forms.Label
	Public WithEvents fraHeadings As System.Windows.Forms.GroupBox
	Public WithEvents btnAnalysis As System.Windows.Forms.Button
	Public WithEvents btnCancel As System.Windows.Forms.Button
	Public WithEvents _txtVslSt_0 As System.Windows.Forms.TextBox
	Public WithEvents _txtVslSt_1 As System.Windows.Forms.TextBox
	Public WithEvents _txtVslSt_2 As System.Windows.Forms.TextBox
	Public WithEvents _txtVslSt_3 As System.Windows.Forms.TextBox
	Public WithEvents _lblLengthUnit_2 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_1 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_4 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_0 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_0 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_1 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_3 As System.Windows.Forms.Label
	Public WithEvents _lblAngleUnit_0 As System.Windows.Forms.Label
	Public WithEvents fraVesselLoc As System.Windows.Forms.GroupBox
	Public WithEvents txtEnvironment As System.Windows.Forms.TextBox
	Public WithEvents btnEnvironment As System.Windows.Forms.Button
	Public WithEvents fraEnvironment As System.Windows.Forms.GroupBox
	Public WithEvents _lblForceUnit_0 As System.Windows.Forms.Label
	Public WithEvents _lblVelUnit_0 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAnalysesB))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.fraMoorGroup = New System.Windows.Forms.GroupBox()
        Me.txtNumLinesPerGroup = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.fraReport = New System.Windows.Forms.GroupBox()
        Me.txtSubTitle = New System.Windows.Forms.TextBox()
        Me.txtReportTitle = New System.Windows.Forms.TextBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.txtInitYaw = New System.Windows.Forms.TextBox()
        Me.txtInitSway = New System.Windows.Forms.TextBox()
        Me.txtInitSurge = New System.Windows.Forms.TextBox()
        Me._lblInitPos_2 = New System.Windows.Forms.Label()
        Me._lblInitPos_1 = New System.Windows.Forms.Label()
        Me._lblInitPos_0 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_4 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_3 = New System.Windows.Forms.Label()
        Me._lblAngleUnit_3 = New System.Windows.Forms.Label()
        Me.btnReport = New System.Windows.Forms.Button()
        Me.lblLocation = New System.Windows.Forms.Label()
        Me.lblClientName = New System.Windows.Forms.Label()
        Me.dlgFileOpen = New System.Windows.Forms.OpenFileDialog()
        Me.dlgFileSave = New System.Windows.Forms.SaveFileDialog()
        Me.fraFile = New System.Windows.Forms.GroupBox()
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.txtFile = New System.Windows.Forms.TextBox()
        Me.fraHeadings = New System.Windows.Forms.GroupBox()
        Me._txtHeadings_2 = New System.Windows.Forms.TextBox()
        Me._txtHeadings_1 = New System.Windows.Forms.TextBox()
        Me._txtHeadings_0 = New System.Windows.Forms.TextBox()
        Me._lblAngleUnit_2 = New System.Windows.Forms.Label()
        Me._lblAngleUnit_1 = New System.Windows.Forms.Label()
        Me._lblHeadings_3 = New System.Windows.Forms.Label()
        Me._lblHeadings_2 = New System.Windows.Forms.Label()
        Me._lblHeadings_1 = New System.Windows.Forms.Label()
        Me._lblHeadings_0 = New System.Windows.Forms.Label()
        Me.btnAnalysis = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.fraVesselLoc = New System.Windows.Forms.GroupBox()
        Me._txtVslSt_0 = New System.Windows.Forms.TextBox()
        Me._txtVslSt_1 = New System.Windows.Forms.TextBox()
        Me._txtVslSt_2 = New System.Windows.Forms.TextBox()
        Me._txtVslSt_3 = New System.Windows.Forms.TextBox()
        Me._lblLengthUnit_2 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_1 = New System.Windows.Forms.Label()
        Me._lblVslSt_4 = New System.Windows.Forms.Label()
        Me._lblLengthUnit_0 = New System.Windows.Forms.Label()
        Me._lblVslSt_0 = New System.Windows.Forms.Label()
        Me._lblVslSt_1 = New System.Windows.Forms.Label()
        Me._lblVslSt_3 = New System.Windows.Forms.Label()
        Me._lblAngleUnit_0 = New System.Windows.Forms.Label()
        Me.fraEnvironment = New System.Windows.Forms.GroupBox()
        Me.txtEnvironment = New System.Windows.Forms.TextBox()
        Me.btnEnvironment = New System.Windows.Forms.Button()
        Me._lblForceUnit_0 = New System.Windows.Forms.Label()
        Me._lblVelUnit_0 = New System.Windows.Forms.Label()
        Me.fraMoorGroup.SuspendLayout()
        Me.fraReport.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.fraFile.SuspendLayout()
        Me.fraHeadings.SuspendLayout()
        Me.fraVesselLoc.SuspendLayout()
        Me.fraEnvironment.SuspendLayout()
        Me.SuspendLayout()
        '
        'fraMoorGroup
        '
        Me.fraMoorGroup.BackColor = System.Drawing.SystemColors.Control
        Me.fraMoorGroup.Controls.Add(Me.txtNumLinesPerGroup)
        Me.fraMoorGroup.Controls.Add(Me.Label1)
        Me.fraMoorGroup.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraMoorGroup.Location = New System.Drawing.Point(376, 299)
        Me.fraMoorGroup.Name = "fraMoorGroup"
        Me.fraMoorGroup.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraMoorGroup.Size = New System.Drawing.Size(149, 59)
        Me.fraMoorGroup.TabIndex = 47
        Me.fraMoorGroup.TabStop = False
        Me.fraMoorGroup.Text = "Grouping Pattern:"
        '
        'txtNumLinesPerGroup
        '
        Me.txtNumLinesPerGroup.AcceptsReturn = True
        Me.txtNumLinesPerGroup.BackColor = System.Drawing.SystemColors.Window
        Me.txtNumLinesPerGroup.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNumLinesPerGroup.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNumLinesPerGroup.Location = New System.Drawing.Point(106, 23)
        Me.txtNumLinesPerGroup.MaxLength = 0
        Me.txtNumLinesPerGroup.Name = "txtNumLinesPerGroup"
        Me.txtNumLinesPerGroup.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNumLinesPerGroup.Size = New System.Drawing.Size(28, 20)
        Me.txtNumLinesPerGroup.TabIndex = 48
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(13, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(78, 33)
        Me.Label1.TabIndex = 49
        Me.Label1.Text = "Number of Lines per Group:"
        '
        'fraReport
        '
        Me.fraReport.BackColor = System.Drawing.SystemColors.Control
        Me.fraReport.Controls.Add(Me.txtSubTitle)
        Me.fraReport.Controls.Add(Me.txtReportTitle)
        Me.fraReport.Controls.Add(Me.Frame2)
        Me.fraReport.Controls.Add(Me.btnReport)
        Me.fraReport.Controls.Add(Me.lblLocation)
        Me.fraReport.Controls.Add(Me.lblClientName)
        Me.fraReport.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraReport.Location = New System.Drawing.Point(235, 78)
        Me.fraReport.Name = "fraReport"
        Me.fraReport.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraReport.Size = New System.Drawing.Size(289, 209)
        Me.fraReport.TabIndex = 31
        Me.fraReport.TabStop = False
        Me.fraReport.Text = "Reporting"
        '
        'txtSubTitle
        '
        Me.txtSubTitle.AcceptsReturn = True
        Me.txtSubTitle.BackColor = System.Drawing.SystemColors.Window
        Me.txtSubTitle.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSubTitle.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSubTitle.Location = New System.Drawing.Point(73, 44)
        Me.txtSubTitle.MaxLength = 0
        Me.txtSubTitle.Name = "txtSubTitle"
        Me.txtSubTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSubTitle.Size = New System.Drawing.Size(199, 19)
        Me.txtSubTitle.TabIndex = 44
        Me.txtSubTitle.TabStop = False
        '
        'txtReportTitle
        '
        Me.txtReportTitle.AcceptsReturn = True
        Me.txtReportTitle.BackColor = System.Drawing.SystemColors.Window
        Me.txtReportTitle.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtReportTitle.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtReportTitle.Location = New System.Drawing.Point(73, 19)
        Me.txtReportTitle.MaxLength = 0
        Me.txtReportTitle.Name = "txtReportTitle"
        Me.txtReportTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtReportTitle.Size = New System.Drawing.Size(199, 19)
        Me.txtReportTitle.TabIndex = 43
        Me.txtReportTitle.TabStop = False
        Me.txtReportTitle.Text = "Mooring Analysis Results"
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.txtInitYaw)
        Me.Frame2.Controls.Add(Me.txtInitSway)
        Me.Frame2.Controls.Add(Me.txtInitSurge)
        Me.Frame2.Controls.Add(Me._lblInitPos_2)
        Me.Frame2.Controls.Add(Me._lblInitPos_1)
        Me.Frame2.Controls.Add(Me._lblInitPos_0)
        Me.Frame2.Controls.Add(Me._lblLengthUnit_4)
        Me.Frame2.Controls.Add(Me._lblLengthUnit_3)
        Me.Frame2.Controls.Add(Me._lblAngleUnit_3)
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(13, 79)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(174, 108)
        Me.Frame2.TabIndex = 33
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Base Vessel Motion:"
        '
        'txtInitYaw
        '
        Me.txtInitYaw.AcceptsReturn = True
        Me.txtInitYaw.BackColor = System.Drawing.SystemColors.Window
        Me.txtInitYaw.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtInitYaw.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtInitYaw.Location = New System.Drawing.Point(77, 74)
        Me.txtInitYaw.MaxLength = 0
        Me.txtInitYaw.Name = "txtInitYaw"
        Me.txtInitYaw.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtInitYaw.Size = New System.Drawing.Size(67, 19)
        Me.txtInitYaw.TabIndex = 36
        Me.txtInitYaw.TabStop = False
        Me.txtInitYaw.Text = "0"
        Me.txtInitYaw.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtInitSway
        '
        Me.txtInitSway.AcceptsReturn = True
        Me.txtInitSway.BackColor = System.Drawing.SystemColors.Window
        Me.txtInitSway.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtInitSway.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtInitSway.Location = New System.Drawing.Point(77, 49)
        Me.txtInitSway.MaxLength = 0
        Me.txtInitSway.Name = "txtInitSway"
        Me.txtInitSway.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtInitSway.Size = New System.Drawing.Size(67, 19)
        Me.txtInitSway.TabIndex = 35
        Me.txtInitSway.TabStop = False
        Me.txtInitSway.Text = "0"
        Me.txtInitSway.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtInitSurge
        '
        Me.txtInitSurge.AcceptsReturn = True
        Me.txtInitSurge.BackColor = System.Drawing.SystemColors.Window
        Me.txtInitSurge.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtInitSurge.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtInitSurge.Location = New System.Drawing.Point(77, 26)
        Me.txtInitSurge.MaxLength = 0
        Me.txtInitSurge.Name = "txtInitSurge"
        Me.txtInitSurge.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtInitSurge.Size = New System.Drawing.Size(67, 19)
        Me.txtInitSurge.TabIndex = 34
        Me.txtInitSurge.TabStop = False
        Me.txtInitSurge.Text = "0"
        Me.txtInitSurge.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_lblInitPos_2
        '
        Me._lblInitPos_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblInitPos_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblInitPos_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblInitPos_2.Location = New System.Drawing.Point(35, 75)
        Me._lblInitPos_2.Name = "_lblInitPos_2"
        Me._lblInitPos_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblInitPos_2.Size = New System.Drawing.Size(37, 17)
        Me._lblInitPos_2.TabIndex = 42
        Me._lblInitPos_2.Text = "Yaw:"
        '
        '_lblInitPos_1
        '
        Me._lblInitPos_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblInitPos_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblInitPos_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblInitPos_1.Location = New System.Drawing.Point(34, 51)
        Me._lblInitPos_1.Name = "_lblInitPos_1"
        Me._lblInitPos_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblInitPos_1.Size = New System.Drawing.Size(37, 17)
        Me._lblInitPos_1.TabIndex = 41
        Me._lblInitPos_1.Text = "Sway:"
        '
        '_lblInitPos_0
        '
        Me._lblInitPos_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblInitPos_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblInitPos_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblInitPos_0.Location = New System.Drawing.Point(35, 27)
        Me._lblInitPos_0.Name = "_lblInitPos_0"
        Me._lblInitPos_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblInitPos_0.Size = New System.Drawing.Size(37, 17)
        Me._lblInitPos_0.TabIndex = 40
        Me._lblInitPos_0.Text = "Surge:"
        '
        '_lblLengthUnit_4
        '
        Me._lblLengthUnit_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblLengthUnit_4.Location = New System.Drawing.Point(149, 48)
        Me._lblLengthUnit_4.Name = "_lblLengthUnit_4"
        Me._lblLengthUnit_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_4.Size = New System.Drawing.Size(17, 17)
        Me._lblLengthUnit_4.TabIndex = 39
        Me._lblLengthUnit_4.Text = "ft"
        '
        '_lblLengthUnit_3
        '
        Me._lblLengthUnit_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblLengthUnit_3.Location = New System.Drawing.Point(149, 28)
        Me._lblLengthUnit_3.Name = "_lblLengthUnit_3"
        Me._lblLengthUnit_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_3.Size = New System.Drawing.Size(17, 17)
        Me._lblLengthUnit_3.TabIndex = 38
        Me._lblLengthUnit_3.Text = "ft"
        '
        '_lblAngleUnit_3
        '
        Me._lblAngleUnit_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblAngleUnit_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblAngleUnit_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblAngleUnit_3.Location = New System.Drawing.Point(148, 72)
        Me._lblAngleUnit_3.Name = "_lblAngleUnit_3"
        Me._lblAngleUnit_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblAngleUnit_3.Size = New System.Drawing.Size(17, 17)
        Me._lblAngleUnit_3.TabIndex = 37
        Me._lblAngleUnit_3.Text = "Invalid_string_refer_to_original_code"
        '
        'btnReport
        '
        Me.btnReport.BackColor = System.Drawing.SystemColors.Control
        Me.btnReport.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnReport.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnReport.Location = New System.Drawing.Point(197, 84)
        Me.btnReport.Name = "btnReport"
        Me.btnReport.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnReport.Size = New System.Drawing.Size(70, 25)
        Me.btnReport.TabIndex = 32
        Me.btnReport.Text = "&Report"
        Me.btnReport.UseVisualStyleBackColor = False
        '
        'lblLocation
        '
        Me.lblLocation.BackColor = System.Drawing.SystemColors.Control
        Me.lblLocation.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblLocation.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLocation.Location = New System.Drawing.Point(13, 46)
        Me.lblLocation.Name = "lblLocation"
        Me.lblLocation.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblLocation.Size = New System.Drawing.Size(62, 17)
        Me.lblLocation.TabIndex = 46
        Me.lblLocation.Text = "Metocean:"
        '
        'lblClientName
        '
        Me.lblClientName.BackColor = System.Drawing.SystemColors.Control
        Me.lblClientName.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblClientName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblClientName.Location = New System.Drawing.Point(13, 22)
        Me.lblClientName.Name = "lblClientName"
        Me.lblClientName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblClientName.Size = New System.Drawing.Size(73, 17)
        Me.lblClientName.TabIndex = 45
        Me.lblClientName.Text = "Main Title:"
        '
        'fraFile
        '
        Me.fraFile.BackColor = System.Drawing.SystemColors.Control
        Me.fraFile.Controls.Add(Me.btnBrowse)
        Me.fraFile.Controls.Add(Me.txtFile)
        Me.fraFile.ForeColor = System.Drawing.Color.Black
        Me.fraFile.Location = New System.Drawing.Point(9, 10)
        Me.fraFile.Name = "fraFile"
        Me.fraFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraFile.Size = New System.Drawing.Size(514, 57)
        Me.fraFile.TabIndex = 24
        Me.fraFile.TabStop = False
        Me.fraFile.Text = "Result File Name"
        '
        'btnBrowse
        '
        Me.btnBrowse.BackColor = System.Drawing.SystemColors.Control
        Me.btnBrowse.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnBrowse.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnBrowse.Location = New System.Drawing.Point(432, 19)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnBrowse.Size = New System.Drawing.Size(65, 25)
        Me.btnBrowse.TabIndex = 26
        Me.btnBrowse.Text = "Browse..."
        Me.btnBrowse.UseVisualStyleBackColor = False
        '
        'txtFile
        '
        Me.txtFile.AcceptsReturn = True
        Me.txtFile.BackColor = System.Drawing.SystemColors.Window
        Me.txtFile.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFile.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFile.Location = New System.Drawing.Point(16, 24)
        Me.txtFile.MaxLength = 0
        Me.txtFile.Name = "txtFile"
        Me.txtFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFile.Size = New System.Drawing.Size(409, 19)
        Me.txtFile.TabIndex = 25
        Me.txtFile.Text = "vessel.sta"
        '
        'fraHeadings
        '
        Me.fraHeadings.BackColor = System.Drawing.SystemColors.Control
        Me.fraHeadings.Controls.Add(Me._txtHeadings_2)
        Me.fraHeadings.Controls.Add(Me._txtHeadings_1)
        Me.fraHeadings.Controls.Add(Me._txtHeadings_0)
        Me.fraHeadings.Controls.Add(Me._lblAngleUnit_2)
        Me.fraHeadings.Controls.Add(Me._lblAngleUnit_1)
        Me.fraHeadings.Controls.Add(Me._lblHeadings_3)
        Me.fraHeadings.Controls.Add(Me._lblHeadings_2)
        Me.fraHeadings.Controls.Add(Me._lblHeadings_1)
        Me.fraHeadings.Controls.Add(Me._lblHeadings_0)
        Me.fraHeadings.ForeColor = System.Drawing.Color.Black
        Me.fraHeadings.Location = New System.Drawing.Point(10, 298)
        Me.fraHeadings.Name = "fraHeadings"
        Me.fraHeadings.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraHeadings.Size = New System.Drawing.Size(355, 60)
        Me.fraHeadings.TabIndex = 16
        Me.fraHeadings.TabStop = False
        Me.fraHeadings.Text = "Colinear Environmental Headings (TN)"
        '
        '_txtHeadings_2
        '
        Me._txtHeadings_2.AcceptsReturn = True
        Me._txtHeadings_2.BackColor = System.Drawing.SystemColors.Window
        Me._txtHeadings_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtHeadings_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtHeadings_2.Location = New System.Drawing.Point(253, 23)
        Me._txtHeadings_2.MaxLength = 0
        Me._txtHeadings_2.Name = "_txtHeadings_2"
        Me._txtHeadings_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtHeadings_2.Size = New System.Drawing.Size(35, 19)
        Me._txtHeadings_2.TabIndex = 21
        Me._txtHeadings_2.TabStop = False
        Me._txtHeadings_2.Text = "9"
        Me._txtHeadings_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtHeadings_1
        '
        Me._txtHeadings_1.AcceptsReturn = True
        Me._txtHeadings_1.BackColor = System.Drawing.SystemColors.Window
        Me._txtHeadings_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtHeadings_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtHeadings_1.Location = New System.Drawing.Point(138, 24)
        Me._txtHeadings_1.MaxLength = 0
        Me._txtHeadings_1.Name = "_txtHeadings_1"
        Me._txtHeadings_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtHeadings_1.Size = New System.Drawing.Size(43, 19)
        Me._txtHeadings_1.TabIndex = 19
        Me._txtHeadings_1.TabStop = False
        Me._txtHeadings_1.Text = "180"
        Me._txtHeadings_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtHeadings_0
        '
        Me._txtHeadings_0.AcceptsReturn = True
        Me._txtHeadings_0.BackColor = System.Drawing.SystemColors.Window
        Me._txtHeadings_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtHeadings_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtHeadings_0.Location = New System.Drawing.Point(49, 24)
        Me._txtHeadings_0.MaxLength = 0
        Me._txtHeadings_0.Name = "_txtHeadings_0"
        Me._txtHeadings_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtHeadings_0.Size = New System.Drawing.Size(45, 19)
        Me._txtHeadings_0.TabIndex = 17
        Me._txtHeadings_0.TabStop = False
        Me._txtHeadings_0.Text = "0"
        Me._txtHeadings_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_lblAngleUnit_2
        '
        Me._lblAngleUnit_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblAngleUnit_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblAngleUnit_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblAngleUnit_2.Location = New System.Drawing.Point(182, 22)
        Me._lblAngleUnit_2.Name = "_lblAngleUnit_2"
        Me._lblAngleUnit_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblAngleUnit_2.Size = New System.Drawing.Size(17, 17)
        Me._lblAngleUnit_2.TabIndex = 30
        Me._lblAngleUnit_2.Text = "Invalid_string_refer_to_original_code"
        '
        '_lblAngleUnit_1
        '
        Me._lblAngleUnit_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblAngleUnit_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblAngleUnit_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblAngleUnit_1.Location = New System.Drawing.Point(95, 23)
        Me._lblAngleUnit_1.Name = "_lblAngleUnit_1"
        Me._lblAngleUnit_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblAngleUnit_1.Size = New System.Drawing.Size(17, 17)
        Me._lblAngleUnit_1.TabIndex = 29
        Me._lblAngleUnit_1.Text = "Invalid_string_refer_to_original_code"
        '
        '_lblHeadings_3
        '
        Me._lblHeadings_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblHeadings_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblHeadings_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblHeadings_3.Location = New System.Drawing.Point(295, 26)
        Me._lblHeadings_3.Name = "_lblHeadings_3"
        Me._lblHeadings_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblHeadings_3.Size = New System.Drawing.Size(49, 17)
        Me._lblHeadings_3.TabIndex = 23
        Me._lblHeadings_3.Text = "headings"
        '
        '_lblHeadings_2
        '
        Me._lblHeadings_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblHeadings_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblHeadings_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblHeadings_2.Location = New System.Drawing.Point(201, 25)
        Me._lblHeadings_2.Name = "_lblHeadings_2"
        Me._lblHeadings_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblHeadings_2.Size = New System.Drawing.Size(49, 17)
        Me._lblHeadings_2.TabIndex = 22
        Me._lblHeadings_2.Text = "with total"
        '
        '_lblHeadings_1
        '
        Me._lblHeadings_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblHeadings_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblHeadings_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblHeadings_1.Location = New System.Drawing.Point(115, 24)
        Me._lblHeadings_1.Name = "_lblHeadings_1"
        Me._lblHeadings_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblHeadings_1.Size = New System.Drawing.Size(21, 17)
        Me._lblHeadings_1.TabIndex = 20
        Me._lblHeadings_1.Text = "to"
        '
        '_lblHeadings_0
        '
        Me._lblHeadings_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblHeadings_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblHeadings_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblHeadings_0.Location = New System.Drawing.Point(17, 25)
        Me._lblHeadings_0.Name = "_lblHeadings_0"
        Me._lblHeadings_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblHeadings_0.Size = New System.Drawing.Size(38, 17)
        Me._lblHeadings_0.TabIndex = 18
        Me._lblHeadings_0.Text = "From"
        '
        'btnAnalysis
        '
        Me.btnAnalysis.BackColor = System.Drawing.SystemColors.Control
        Me.btnAnalysis.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnAnalysis.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnAnalysis.Location = New System.Drawing.Point(532, 45)
        Me.btnAnalysis.Name = "btnAnalysis"
        Me.btnAnalysis.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnAnalysis.Size = New System.Drawing.Size(65, 25)
        Me.btnAnalysis.TabIndex = 15
        Me.btnAnalysis.Text = "St&art"
        Me.btnAnalysis.UseVisualStyleBackColor = False
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.SystemColors.Control
        Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCancel.Location = New System.Drawing.Point(532, 13)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCancel.Size = New System.Drawing.Size(65, 25)
        Me.btnCancel.TabIndex = 14
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'fraVesselLoc
        '
        Me.fraVesselLoc.BackColor = System.Drawing.SystemColors.Control
        Me.fraVesselLoc.Controls.Add(Me._txtVslSt_0)
        Me.fraVesselLoc.Controls.Add(Me._txtVslSt_1)
        Me.fraVesselLoc.Controls.Add(Me._txtVslSt_2)
        Me.fraVesselLoc.Controls.Add(Me._txtVslSt_3)
        Me.fraVesselLoc.Controls.Add(Me._lblLengthUnit_2)
        Me.fraVesselLoc.Controls.Add(Me._lblLengthUnit_1)
        Me.fraVesselLoc.Controls.Add(Me._lblVslSt_4)
        Me.fraVesselLoc.Controls.Add(Me._lblLengthUnit_0)
        Me.fraVesselLoc.Controls.Add(Me._lblVslSt_0)
        Me.fraVesselLoc.Controls.Add(Me._lblVslSt_1)
        Me.fraVesselLoc.Controls.Add(Me._lblVslSt_3)
        Me.fraVesselLoc.Controls.Add(Me._lblAngleUnit_0)
        Me.fraVesselLoc.ForeColor = System.Drawing.Color.Black
        Me.fraVesselLoc.Location = New System.Drawing.Point(10, 146)
        Me.fraVesselLoc.Name = "fraVesselLoc"
        Me.fraVesselLoc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraVesselLoc.Size = New System.Drawing.Size(214, 141)
        Me.fraVesselLoc.TabIndex = 3
        Me.fraVesselLoc.TabStop = False
        Me.fraVesselLoc.Text = "Current Vessel Station :"
        '
        '_txtVslSt_0
        '
        Me._txtVslSt_0.AcceptsReturn = True
        Me._txtVslSt_0.BackColor = System.Drawing.SystemColors.Control
        Me._txtVslSt_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtVslSt_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtVslSt_0.Location = New System.Drawing.Point(71, 26)
        Me._txtVslSt_0.MaxLength = 0
        Me._txtVslSt_0.Name = "_txtVslSt_0"
        Me._txtVslSt_0.ReadOnly = True
        Me._txtVslSt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtVslSt_0.Size = New System.Drawing.Size(104, 19)
        Me._txtVslSt_0.TabIndex = 7
        Me._txtVslSt_0.TabStop = False
        Me._txtVslSt_0.Text = "0"
        Me._txtVslSt_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtVslSt_1
        '
        Me._txtVslSt_1.AcceptsReturn = True
        Me._txtVslSt_1.BackColor = System.Drawing.SystemColors.Control
        Me._txtVslSt_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtVslSt_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtVslSt_1.Location = New System.Drawing.Point(71, 53)
        Me._txtVslSt_1.MaxLength = 0
        Me._txtVslSt_1.Name = "_txtVslSt_1"
        Me._txtVslSt_1.ReadOnly = True
        Me._txtVslSt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtVslSt_1.Size = New System.Drawing.Size(104, 19)
        Me._txtVslSt_1.TabIndex = 6
        Me._txtVslSt_1.TabStop = False
        Me._txtVslSt_1.Text = "0"
        Me._txtVslSt_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtVslSt_2
        '
        Me._txtVslSt_2.AcceptsReturn = True
        Me._txtVslSt_2.BackColor = System.Drawing.SystemColors.Control
        Me._txtVslSt_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtVslSt_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtVslSt_2.Location = New System.Drawing.Point(118, 79)
        Me._txtVslSt_2.MaxLength = 0
        Me._txtVslSt_2.Name = "_txtVslSt_2"
        Me._txtVslSt_2.ReadOnly = True
        Me._txtVslSt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtVslSt_2.Size = New System.Drawing.Size(57, 19)
        Me._txtVslSt_2.TabIndex = 5
        Me._txtVslSt_2.TabStop = False
        Me._txtVslSt_2.Text = "0"
        Me._txtVslSt_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtVslSt_3
        '
        Me._txtVslSt_3.AcceptsReturn = True
        Me._txtVslSt_3.BackColor = System.Drawing.SystemColors.Control
        Me._txtVslSt_3.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtVslSt_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtVslSt_3.Location = New System.Drawing.Point(118, 106)
        Me._txtVslSt_3.MaxLength = 0
        Me._txtVslSt_3.Name = "_txtVslSt_3"
        Me._txtVslSt_3.ReadOnly = True
        Me._txtVslSt_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtVslSt_3.Size = New System.Drawing.Size(57, 19)
        Me._txtVslSt_3.TabIndex = 4
        Me._txtVslSt_3.TabStop = False
        Me._txtVslSt_3.Text = "0"
        Me._txtVslSt_3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_lblLengthUnit_2
        '
        Me._lblLengthUnit_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblLengthUnit_2.Location = New System.Drawing.Point(183, 108)
        Me._lblLengthUnit_2.Name = "_lblLengthUnit_2"
        Me._lblLengthUnit_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_2.Size = New System.Drawing.Size(17, 17)
        Me._lblLengthUnit_2.TabIndex = 28
        Me._lblLengthUnit_2.Text = "ft"
        '
        '_lblLengthUnit_1
        '
        Me._lblLengthUnit_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblLengthUnit_1.Location = New System.Drawing.Point(183, 53)
        Me._lblLengthUnit_1.Name = "_lblLengthUnit_1"
        Me._lblLengthUnit_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_1.Size = New System.Drawing.Size(17, 17)
        Me._lblLengthUnit_1.TabIndex = 27
        Me._lblLengthUnit_1.Text = "ft"
        '
        '_lblVslSt_4
        '
        Me._lblVslSt_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblVslSt_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVslSt_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVslSt_4.Location = New System.Drawing.Point(29, 107)
        Me._lblVslSt_4.Name = "_lblVslSt_4"
        Me._lblVslSt_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVslSt_4.Size = New System.Drawing.Size(25, 17)
        Me._lblVslSt_4.TabIndex = 13
        Me._lblVslSt_4.Text = "Draft"
        '
        '_lblLengthUnit_0
        '
        Me._lblLengthUnit_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblLengthUnit_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLengthUnit_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblLengthUnit_0.Location = New System.Drawing.Point(183, 26)
        Me._lblLengthUnit_0.Name = "_lblLengthUnit_0"
        Me._lblLengthUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLengthUnit_0.Size = New System.Drawing.Size(17, 17)
        Me._lblLengthUnit_0.TabIndex = 12
        Me._lblLengthUnit_0.Text = "ft"
        '
        '_lblVslSt_0
        '
        Me._lblVslSt_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblVslSt_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVslSt_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVslSt_0.Location = New System.Drawing.Point(29, 26)
        Me._lblVslSt_0.Name = "_lblVslSt_0"
        Me._lblVslSt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVslSt_0.Size = New System.Drawing.Size(33, 17)
        Me._lblVslSt_0.TabIndex = 11
        Me._lblVslSt_0.Text = "x (E)"
        '
        '_lblVslSt_1
        '
        Me._lblVslSt_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblVslSt_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVslSt_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVslSt_1.Location = New System.Drawing.Point(29, 53)
        Me._lblVslSt_1.Name = "_lblVslSt_1"
        Me._lblVslSt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVslSt_1.Size = New System.Drawing.Size(33, 17)
        Me._lblVslSt_1.TabIndex = 10
        Me._lblVslSt_1.Text = "y (N)"
        '
        '_lblVslSt_3
        '
        Me._lblVslSt_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblVslSt_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVslSt_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVslSt_3.Location = New System.Drawing.Point(29, 80)
        Me._lblVslSt_3.Name = "_lblVslSt_3"
        Me._lblVslSt_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVslSt_3.Size = New System.Drawing.Size(49, 17)
        Me._lblVslSt_3.TabIndex = 9
        Me._lblVslSt_3.Text = "Heading"
        '
        '_lblAngleUnit_0
        '
        Me._lblAngleUnit_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblAngleUnit_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblAngleUnit_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblAngleUnit_0.Location = New System.Drawing.Point(183, 81)
        Me._lblAngleUnit_0.Name = "_lblAngleUnit_0"
        Me._lblAngleUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblAngleUnit_0.Size = New System.Drawing.Size(17, 17)
        Me._lblAngleUnit_0.TabIndex = 8
        Me._lblAngleUnit_0.Text = "Invalid_string_refer_to_original_code"
        '
        'fraEnvironment
        '
        Me.fraEnvironment.BackColor = System.Drawing.SystemColors.Control
        Me.fraEnvironment.Controls.Add(Me.txtEnvironment)
        Me.fraEnvironment.Controls.Add(Me.btnEnvironment)
        Me.fraEnvironment.ForeColor = System.Drawing.Color.Black
        Me.fraEnvironment.Location = New System.Drawing.Point(9, 79)
        Me.fraEnvironment.Name = "fraEnvironment"
        Me.fraEnvironment.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraEnvironment.Size = New System.Drawing.Size(216, 58)
        Me.fraEnvironment.TabIndex = 0
        Me.fraEnvironment.TabStop = False
        Me.fraEnvironment.Text = "Environment Condition"
        '
        'txtEnvironment
        '
        Me.txtEnvironment.AcceptsReturn = True
        Me.txtEnvironment.BackColor = System.Drawing.SystemColors.Control
        Me.txtEnvironment.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtEnvironment.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtEnvironment.Location = New System.Drawing.Point(16, 24)
        Me.txtEnvironment.MaxLength = 0
        Me.txtEnvironment.Name = "txtEnvironment"
        Me.txtEnvironment.ReadOnly = True
        Me.txtEnvironment.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtEnvironment.Size = New System.Drawing.Size(131, 19)
        Me.txtEnvironment.TabIndex = 2
        Me.txtEnvironment.Text = "Current Environment"
        '
        'btnEnvironment
        '
        Me.btnEnvironment.BackColor = System.Drawing.SystemColors.Control
        Me.btnEnvironment.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnEnvironment.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnEnvironment.Location = New System.Drawing.Point(154, 22)
        Me.btnEnvironment.Name = "btnEnvironment"
        Me.btnEnvironment.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnEnvironment.Size = New System.Drawing.Size(50, 25)
        Me.btnEnvironment.TabIndex = 1
        Me.btnEnvironment.Text = "Edit..."
        Me.btnEnvironment.UseVisualStyleBackColor = False
        '
        '_lblForceUnit_0
        '
        Me._lblForceUnit_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblForceUnit_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblForceUnit_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblForceUnit_0.Location = New System.Drawing.Point(1, 16)
        Me._lblForceUnit_0.Name = "_lblForceUnit_0"
        Me._lblForceUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblForceUnit_0.Size = New System.Drawing.Size(25, 17)
        Me._lblForceUnit_0.TabIndex = 51
        Me._lblForceUnit_0.Text = "kips"
        Me._lblForceUnit_0.Visible = False
        '
        '_lblVelUnit_0
        '
        Me._lblVelUnit_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblVelUnit_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVelUnit_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVelUnit_0.Location = New System.Drawing.Point(0, 0)
        Me._lblVelUnit_0.Name = "_lblVelUnit_0"
        Me._lblVelUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVelUnit_0.Size = New System.Drawing.Size(25, 17)
        Me._lblVelUnit_0.TabIndex = 50
        Me._lblVelUnit_0.Text = "kn"
        Me._lblVelUnit_0.Visible = False
        '
        'frmAnalysesB
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(609, 371)
        Me.Controls.Add(Me.fraMoorGroup)
        Me.Controls.Add(Me.fraReport)
        Me.Controls.Add(Me.fraFile)
        Me.Controls.Add(Me.fraHeadings)
        Me.Controls.Add(Me.btnAnalysis)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.fraVesselLoc)
        Me.Controls.Add(Me.fraEnvironment)
        Me.Controls.Add(Me._lblForceUnit_0)
        Me.Controls.Add(Me._lblVelUnit_0)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(4, 23)
        Me.Name = "frmAnalysesB"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Mooring Analyses with Multiple Environmental Headings"
        Me.fraMoorGroup.ResumeLayout(False)
        Me.fraReport.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        Me.fraFile.ResumeLayout(False)
        Me.fraHeadings.ResumeLayout(False)
        Me.fraVesselLoc.ResumeLayout(False)
        Me.fraEnvironment.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
#End Region
End Class