<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmOtherRiserProperties
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
    Public WithEvents _cmdBrowse_3 As System.Windows.Forms.Button
    Public WithEvents _txtModesFile_3 As System.Windows.Forms.TextBox
    Public WithEvents _cmdBrowse_2 As System.Windows.Forms.Button
    Public WithEvents _txtModesFile_2 As System.Windows.Forms.TextBox
    Public WithEvents _cmdBrowse_1 As System.Windows.Forms.Button
    Public WithEvents _cmdBrowse_0 As System.Windows.Forms.Button
    Public WithEvents _txtModesFile_1 As System.Windows.Forms.TextBox
    Public WithEvents _txtModesFile_0 As System.Windows.Forms.TextBox
    Public WithEvents _optModes_2 As System.Windows.Forms.RadioButton
    Public WithEvents _optModes_1 As System.Windows.Forms.RadioButton
    Public WithEvents _optModes_0 As System.Windows.Forms.RadioButton
    Public WithEvents frmModes As System.Windows.Forms.GroupBox
    Public WithEvents btnCancel As System.Windows.Forms.Button
    Public WithEvents btnOK As System.Windows.Forms.Button
    Public WithEvents cmdBrowse As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
    Public WithEvents optModes As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    Public WithEvents txtModesFile As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmOtherRiserProperties))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.dlgFilesOpen = New System.Windows.Forms.OpenFileDialog
        Me.frmModes = New System.Windows.Forms.GroupBox
        Me._cmdBrowse_3 = New System.Windows.Forms.Button
        Me._txtModesFile_3 = New System.Windows.Forms.TextBox
        Me._cmdBrowse_2 = New System.Windows.Forms.Button
        Me._txtModesFile_2 = New System.Windows.Forms.TextBox
        Me._cmdBrowse_1 = New System.Windows.Forms.Button
        Me._cmdBrowse_0 = New System.Windows.Forms.Button
        Me._txtModesFile_1 = New System.Windows.Forms.TextBox
        Me._txtModesFile_0 = New System.Windows.Forms.TextBox
        Me._optModes_2 = New System.Windows.Forms.RadioButton
        Me._optModes_1 = New System.Windows.Forms.RadioButton
        Me._optModes_0 = New System.Windows.Forms.RadioButton
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.cmdBrowse = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(Me.components)
        Me.optModes = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.txtModesFile = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.Frame1 = New System.Windows.Forms.GroupBox
        Me._lblEquation_1 = New System.Windows.Forms.Label
        Me._lblEquation_0 = New System.Windows.Forms.Label
        Me.cboFatigueCurves = New System.Windows.Forms.ComboBox
        Me.txtFatConstA = New System.Windows.Forms.TextBox
        Me.txtFatConstB = New System.Windows.Forms.TextBox
        Me.txtSCF = New System.Windows.Forms.TextBox
        Me._lblEquation_2 = New System.Windows.Forms.Label
        Me.lblUnit = New System.Windows.Forms.Label
        Me.lblFatConstA = New System.Windows.Forms.Label
        Me.lblFatConstB = New System.Windows.Forms.Label
        Me.lblSCF = New System.Windows.Forms.Label
        Me.fraParms = New System.Windows.Forms.GroupBox
        Me.txtDampingCoef = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.btnNoHiRe = New System.Windows.Forms.RadioButton
        Me.btnHiRe = New System.Windows.Forms.RadioButton
        Me.frmModes.SuspendLayout()
        CType(Me.cmdBrowse, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optModes, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtModesFile, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame1.SuspendLayout()
        Me.fraParms.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'frmModes
        '
        Me.frmModes.BackColor = System.Drawing.SystemColors.Control
        Me.frmModes.Controls.Add(Me._cmdBrowse_3)
        Me.frmModes.Controls.Add(Me._txtModesFile_3)
        Me.frmModes.Controls.Add(Me._cmdBrowse_2)
        Me.frmModes.Controls.Add(Me._txtModesFile_2)
        Me.frmModes.Controls.Add(Me._cmdBrowse_1)
        Me.frmModes.Controls.Add(Me._cmdBrowse_0)
        Me.frmModes.Controls.Add(Me._txtModesFile_1)
        Me.frmModes.Controls.Add(Me._txtModesFile_0)
        Me.frmModes.Controls.Add(Me._optModes_2)
        Me.frmModes.Controls.Add(Me._optModes_1)
        Me.frmModes.Controls.Add(Me._optModes_0)
        Me.frmModes.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.frmModes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.frmModes.Location = New System.Drawing.Point(12, 249)
        Me.frmModes.Name = "frmModes"
        Me.frmModes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.frmModes.Size = New System.Drawing.Size(314, 262)
        Me.frmModes.TabIndex = 21
        Me.frmModes.TabStop = False
        Me.frmModes.Text = "Natural Frequencies and Modes"
        '
        '_cmdBrowse_3
        '
        Me._cmdBrowse_3.BackColor = System.Drawing.SystemColors.Control
        Me._cmdBrowse_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._cmdBrowse_3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._cmdBrowse_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBrowse.SetIndex(Me._cmdBrowse_3, CType(3, Short))
        Me._cmdBrowse_3.Location = New System.Drawing.Point(232, 222)
        Me._cmdBrowse_3.Name = "_cmdBrowse_3"
        Me._cmdBrowse_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._cmdBrowse_3.Size = New System.Drawing.Size(65, 25)
        Me._cmdBrowse_3.TabIndex = 25
        Me._cmdBrowse_3.Text = "Browse ..."
        Me._cmdBrowse_3.UseVisualStyleBackColor = False
        '
        '_txtModesFile_3
        '
        Me._txtModesFile_3.AcceptsReturn = True
        Me._txtModesFile_3.BackColor = System.Drawing.SystemColors.Window
        Me._txtModesFile_3.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtModesFile_3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtModesFile_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtModesFile.SetIndex(Me._txtModesFile_3, CType(3, Short))
        Me._txtModesFile_3.Location = New System.Drawing.Point(137, 227)
        Me._txtModesFile_3.MaxLength = 0
        Me._txtModesFile_3.Name = "_txtModesFile_3"
        Me._txtModesFile_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtModesFile_3.Size = New System.Drawing.Size(89, 20)
        Me._txtModesFile_3.TabIndex = 13
        Me._txtModesFile_3.Text = "Curvature File Name"
        '
        '_cmdBrowse_2
        '
        Me._cmdBrowse_2.BackColor = System.Drawing.SystemColors.Control
        Me._cmdBrowse_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._cmdBrowse_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._cmdBrowse_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBrowse.SetIndex(Me._cmdBrowse_2, CType(2, Short))
        Me._cmdBrowse_2.Location = New System.Drawing.Point(232, 191)
        Me._cmdBrowse_2.Name = "_cmdBrowse_2"
        Me._cmdBrowse_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._cmdBrowse_2.Size = New System.Drawing.Size(65, 25)
        Me._cmdBrowse_2.TabIndex = 23
        Me._cmdBrowse_2.Text = "Browse ..."
        Me._cmdBrowse_2.UseVisualStyleBackColor = False
        '
        '_txtModesFile_2
        '
        Me._txtModesFile_2.AcceptsReturn = True
        Me._txtModesFile_2.BackColor = System.Drawing.SystemColors.Window
        Me._txtModesFile_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtModesFile_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtModesFile_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtModesFile.SetIndex(Me._txtModesFile_2, CType(2, Short))
        Me._txtModesFile_2.Location = New System.Drawing.Point(137, 191)
        Me._txtModesFile_2.MaxLength = 0
        Me._txtModesFile_2.Name = "_txtModesFile_2"
        Me._txtModesFile_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtModesFile_2.Size = New System.Drawing.Size(89, 20)
        Me._txtModesFile_2.TabIndex = 12
        Me._txtModesFile_2.Text = "Modes File Name"
        '
        '_cmdBrowse_1
        '
        Me._cmdBrowse_1.BackColor = System.Drawing.SystemColors.Control
        Me._cmdBrowse_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._cmdBrowse_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._cmdBrowse_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBrowse.SetIndex(Me._cmdBrowse_1, CType(1, Short))
        Me._cmdBrowse_1.Location = New System.Drawing.Point(232, 154)
        Me._cmdBrowse_1.Name = "_cmdBrowse_1"
        Me._cmdBrowse_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._cmdBrowse_1.Size = New System.Drawing.Size(65, 25)
        Me._cmdBrowse_1.TabIndex = 21
        Me._cmdBrowse_1.Text = "Browse ..."
        Me._cmdBrowse_1.UseVisualStyleBackColor = False
        '
        '_cmdBrowse_0
        '
        Me._cmdBrowse_0.BackColor = System.Drawing.SystemColors.Control
        Me._cmdBrowse_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._cmdBrowse_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._cmdBrowse_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBrowse.SetIndex(Me._cmdBrowse_0, CType(0, Short))
        Me._cmdBrowse_0.Location = New System.Drawing.Point(232, 87)
        Me._cmdBrowse_0.Name = "_cmdBrowse_0"
        Me._cmdBrowse_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._cmdBrowse_0.Size = New System.Drawing.Size(65, 25)
        Me._cmdBrowse_0.TabIndex = 20
        Me._cmdBrowse_0.Text = "Browse ..."
        Me._cmdBrowse_0.UseVisualStyleBackColor = False
        '
        '_txtModesFile_1
        '
        Me._txtModesFile_1.AcceptsReturn = True
        Me._txtModesFile_1.BackColor = System.Drawing.SystemColors.Window
        Me._txtModesFile_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtModesFile_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtModesFile_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtModesFile.SetIndex(Me._txtModesFile_1, CType(1, Short))
        Me._txtModesFile_1.Location = New System.Drawing.Point(137, 156)
        Me._txtModesFile_1.MaxLength = 0
        Me._txtModesFile_1.Name = "_txtModesFile_1"
        Me._txtModesFile_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtModesFile_1.Size = New System.Drawing.Size(89, 20)
        Me._txtModesFile_1.TabIndex = 11
        Me._txtModesFile_1.Text = "Freq File Name"
        '
        '_txtModesFile_0
        '
        Me._txtModesFile_0.AcceptsReturn = True
        Me._txtModesFile_0.BackColor = System.Drawing.SystemColors.Window
        Me._txtModesFile_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtModesFile_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtModesFile_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtModesFile.SetIndex(Me._txtModesFile_0, CType(0, Short))
        Me._txtModesFile_0.Location = New System.Drawing.Point(137, 87)
        Me._txtModesFile_0.MaxLength = 0
        Me._txtModesFile_0.Name = "_txtModesFile_0"
        Me._txtModesFile_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtModesFile_0.Size = New System.Drawing.Size(89, 20)
        Me._txtModesFile_0.TabIndex = 10
        Me._txtModesFile_0.Text = "Freq File Name"
        '
        '_optModes_2
        '
        Me._optModes_2.BackColor = System.Drawing.SystemColors.Control
        Me._optModes_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._optModes_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._optModes_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optModes.SetIndex(Me._optModes_2, CType(2, Short))
        Me._optModes_2.Location = New System.Drawing.Point(16, 128)
        Me._optModes_2.Name = "_optModes_2"
        Me._optModes_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optModes_2.Size = New System.Drawing.Size(282, 20)
        Me._optModes_2.TabIndex = 17
        Me._optModes_2.TabStop = True
        Me._optModes_2.Text = "Import natural frequencies and modes from files"
        Me._optModes_2.UseVisualStyleBackColor = False
        '
        '_optModes_1
        '
        Me._optModes_1.BackColor = System.Drawing.SystemColors.Control
        Me._optModes_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optModes_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._optModes_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optModes.SetIndex(Me._optModes_1, CType(1, Short))
        Me._optModes_1.Location = New System.Drawing.Point(16, 61)
        Me._optModes_1.Name = "_optModes_1"
        Me._optModes_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optModes_1.Size = New System.Drawing.Size(241, 20)
        Me._optModes_1.TabIndex = 16
        Me._optModes_1.TabStop = True
        Me._optModes_1.Text = "Import natural frequencies from file"
        Me._optModes_1.UseVisualStyleBackColor = False
        '
        '_optModes_0
        '
        Me._optModes_0.BackColor = System.Drawing.SystemColors.Control
        Me._optModes_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optModes_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._optModes_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optModes.SetIndex(Me._optModes_0, CType(0, Short))
        Me._optModes_0.Location = New System.Drawing.Point(16, 35)
        Me._optModes_0.Name = "_optModes_0"
        Me._optModes_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optModes_0.Size = New System.Drawing.Size(241, 20)
        Me._optModes_0.TabIndex = 15
        Me._optModes_0.TabStop = True
        Me._optModes_0.Text = "Let WinViva calculate natural frequencies"
        Me._optModes_0.UseVisualStyleBackColor = False
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.SystemColors.Control
        Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCancel.Location = New System.Drawing.Point(261, 526)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCancel.Size = New System.Drawing.Size(65, 25)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'btnOK
        '
        Me.btnOK.BackColor = System.Drawing.SystemColors.Control
        Me.btnOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnOK.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnOK.Location = New System.Drawing.Point(188, 526)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnOK.Size = New System.Drawing.Size(65, 25)
        Me.btnOK.TabIndex = 0
        Me.btnOK.Text = "&Save"
        Me.btnOK.UseVisualStyleBackColor = False
        '
        'cmdBrowse
        '
        '
        'optModes
        '
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me._lblEquation_1)
        Me.Frame1.Controls.Add(Me._lblEquation_0)
        Me.Frame1.Controls.Add(Me.cboFatigueCurves)
        Me.Frame1.Controls.Add(Me.txtFatConstA)
        Me.Frame1.Controls.Add(Me.txtFatConstB)
        Me.Frame1.Controls.Add(Me.txtSCF)
        Me.Frame1.Controls.Add(Me._lblEquation_2)
        Me.Frame1.Controls.Add(Me.lblUnit)
        Me.Frame1.Controls.Add(Me.lblFatConstA)
        Me.Frame1.Controls.Add(Me.lblFatConstB)
        Me.Frame1.Controls.Add(Me.lblSCF)
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(12, 74)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(314, 169)
        Me.Frame1.TabIndex = 20
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Fatigue Parameters"
        '
        '_lblEquation_1
        '
        Me._lblEquation_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblEquation_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEquation_1.Font = New System.Drawing.Font("Arial", 6.0!)
        Me._lblEquation_1.ForeColor = System.Drawing.SystemColors.Highlight
        Me._lblEquation_1.Location = New System.Drawing.Point(267, 103)
        Me._lblEquation_1.Name = "_lblEquation_1"
        Me._lblEquation_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEquation_1.Size = New System.Drawing.Size(14, 16)
        Me._lblEquation_1.TabIndex = 10
        Me._lblEquation_1.Text = "B"
        '
        '_lblEquation_0
        '
        Me._lblEquation_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblEquation_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEquation_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblEquation_0.ForeColor = System.Drawing.SystemColors.Highlight
        Me._lblEquation_0.Location = New System.Drawing.Point(18, 103)
        Me._lblEquation_0.Name = "_lblEquation_0"
        Me._lblEquation_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEquation_0.Size = New System.Drawing.Size(276, 30)
        Me._lblEquation_0.TabIndex = 9
        Me._lblEquation_0.Text = "* Cycles to Fatigue Failure = ( A / Maximum Stress )"
        '
        'cboFatigueCurves
        '
        Me.cboFatigueCurves.BackColor = System.Drawing.SystemColors.Window
        Me.cboFatigueCurves.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboFatigueCurves.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboFatigueCurves.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboFatigueCurves.Location = New System.Drawing.Point(176, 21)
        Me.cboFatigueCurves.Name = "cboFatigueCurves"
        Me.cboFatigueCurves.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboFatigueCurves.Size = New System.Drawing.Size(105, 22)
        Me.cboFatigueCurves.TabIndex = 6
        Me.cboFatigueCurves.Text = "Fatigue Curve"
        '
        'txtFatConstA
        '
        Me.txtFatConstA.AcceptsReturn = True
        Me.txtFatConstA.BackColor = System.Drawing.SystemColors.Window
        Me.txtFatConstA.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFatConstA.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFatConstA.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFatConstA.Location = New System.Drawing.Point(200, 47)
        Me.txtFatConstA.MaxLength = 0
        Me.txtFatConstA.Name = "txtFatConstA"
        Me.txtFatConstA.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFatConstA.Size = New System.Drawing.Size(81, 20)
        Me.txtFatConstA.TabIndex = 7
        Me.txtFatConstA.Text = "Text1"
        Me.txtFatConstA.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtFatConstB
        '
        Me.txtFatConstB.AcceptsReturn = True
        Me.txtFatConstB.BackColor = System.Drawing.SystemColors.Window
        Me.txtFatConstB.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFatConstB.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFatConstB.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFatConstB.Location = New System.Drawing.Point(200, 70)
        Me.txtFatConstB.MaxLength = 0
        Me.txtFatConstB.Name = "txtFatConstB"
        Me.txtFatConstB.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFatConstB.Size = New System.Drawing.Size(81, 20)
        Me.txtFatConstB.TabIndex = 8
        Me.txtFatConstB.Text = "Text1"
        Me.txtFatConstB.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtSCF
        '
        Me.txtSCF.AcceptsReturn = True
        Me.txtSCF.BackColor = System.Drawing.SystemColors.Window
        Me.txtSCF.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSCF.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSCF.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSCF.Location = New System.Drawing.Point(200, 136)
        Me.txtSCF.MaxLength = 0
        Me.txtSCF.Name = "txtSCF"
        Me.txtSCF.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSCF.Size = New System.Drawing.Size(81, 20)
        Me.txtSCF.TabIndex = 9
        Me.txtSCF.Text = "Text1"
        Me.txtSCF.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_lblEquation_2
        '
        Me._lblEquation_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblEquation_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEquation_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblEquation_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblEquation_2.Location = New System.Drawing.Point(18, 24)
        Me._lblEquation_2.Name = "_lblEquation_2"
        Me._lblEquation_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEquation_2.Size = New System.Drawing.Size(95, 17)
        Me._lblEquation_2.TabIndex = 13
        Me._lblEquation_2.Text = "Fatigue Curve:"
        '
        'lblUnit
        '
        Me.lblUnit.BackColor = System.Drawing.SystemColors.Control
        Me.lblUnit.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblUnit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUnit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblUnit.Location = New System.Drawing.Point(283, 50)
        Me.lblUnit.Name = "lblUnit"
        Me.lblUnit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblUnit.Size = New System.Drawing.Size(22, 17)
        Me.lblUnit.TabIndex = 11
        Me.lblUnit.Text = "MPa"
        '
        'lblFatConstA
        '
        Me.lblFatConstA.BackColor = System.Drawing.SystemColors.Control
        Me.lblFatConstA.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFatConstA.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFatConstA.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFatConstA.Location = New System.Drawing.Point(39, 48)
        Me.lblFatConstA.Name = "lblFatConstA"
        Me.lblFatConstA.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFatConstA.Size = New System.Drawing.Size(120, 18)
        Me.lblFatConstA.TabIndex = 8
        Me.lblFatConstA.Text = "Fatigue Constant A"
        '
        'lblFatConstB
        '
        Me.lblFatConstB.BackColor = System.Drawing.SystemColors.Control
        Me.lblFatConstB.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFatConstB.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFatConstB.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFatConstB.Location = New System.Drawing.Point(39, 72)
        Me.lblFatConstB.Name = "lblFatConstB"
        Me.lblFatConstB.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFatConstB.Size = New System.Drawing.Size(120, 18)
        Me.lblFatConstB.TabIndex = 7
        Me.lblFatConstB.Text = "Fatigue Constant B"
        '
        'lblSCF
        '
        Me.lblSCF.BackColor = System.Drawing.SystemColors.Control
        Me.lblSCF.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSCF.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSCF.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblSCF.Location = New System.Drawing.Point(15, 136)
        Me.lblSCF.Name = "lblSCF"
        Me.lblSCF.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSCF.Size = New System.Drawing.Size(178, 19)
        Me.lblSCF.TabIndex = 6
        Me.lblSCF.Text = "Stress Concentration Factor (SCF)"
        '
        'fraParms
        '
        Me.fraParms.BackColor = System.Drawing.SystemColors.Control
        Me.fraParms.Controls.Add(Me.txtDampingCoef)
        Me.fraParms.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraParms.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraParms.Location = New System.Drawing.Point(12, 12)
        Me.fraParms.Name = "fraParms"
        Me.fraParms.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraParms.Size = New System.Drawing.Size(159, 56)
        Me.fraParms.TabIndex = 18
        Me.fraParms.TabStop = False
        Me.fraParms.Text = "Structural Damping Ratio"
        '
        'txtDampingCoef
        '
        Me.txtDampingCoef.AcceptsReturn = True
        Me.txtDampingCoef.BackColor = System.Drawing.SystemColors.Window
        Me.txtDampingCoef.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDampingCoef.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDampingCoef.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDampingCoef.Location = New System.Drawing.Point(6, 29)
        Me.txtDampingCoef.MaxLength = 0
        Me.txtDampingCoef.Name = "txtDampingCoef"
        Me.txtDampingCoef.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDampingCoef.Size = New System.Drawing.Size(113, 20)
        Me.txtDampingCoef.TabIndex = 4
        Me.txtDampingCoef.Text = "Text2"
        Me.txtDampingCoef.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox1.Controls.Add(Me.btnNoHiRe)
        Me.GroupBox1.Controls.Add(Me.btnHiRe)
        Me.GroupBox1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.GroupBox1.Location = New System.Drawing.Point(177, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.GroupBox1.Size = New System.Drawing.Size(149, 56)
        Me.GroupBox1.TabIndex = 19
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Use High Reynolds Number Database"
        '
        'btnNoHiRe
        '
        Me.btnNoHiRe.BackColor = System.Drawing.SystemColors.Control
        Me.btnNoHiRe.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNoHiRe.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNoHiRe.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNoHiRe.Location = New System.Drawing.Point(83, 29)
        Me.btnNoHiRe.Name = "btnNoHiRe"
        Me.btnNoHiRe.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNoHiRe.Size = New System.Drawing.Size(46, 21)
        Me.btnNoHiRe.TabIndex = 1
        Me.btnNoHiRe.TabStop = True
        Me.btnNoHiRe.Text = "No"
        Me.btnNoHiRe.UseVisualStyleBackColor = False
        '
        'btnHiRe
        '
        Me.btnHiRe.BackColor = System.Drawing.SystemColors.Control
        Me.btnHiRe.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnHiRe.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnHiRe.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnHiRe.Location = New System.Drawing.Point(6, 29)
        Me.btnHiRe.Name = "btnHiRe"
        Me.btnHiRe.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnHiRe.Size = New System.Drawing.Size(56, 21)
        Me.btnHiRe.TabIndex = 5
        Me.btnHiRe.TabStop = True
        Me.btnHiRe.Text = "Yes"
        Me.btnHiRe.UseVisualStyleBackColor = False
        '
        'frmOtherRiserProperties
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(339, 563)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.fraParms)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.frmModes)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.Name = "frmOtherRiserProperties"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Other Riser Properties"
        Me.frmModes.ResumeLayout(False)
        Me.frmModes.PerformLayout()
        CType(Me.cmdBrowse, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optModes, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtModesFile, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.fraParms.ResumeLayout(False)
        Me.fraParms.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents _lblEquation_1 As System.Windows.Forms.Label
    Public WithEvents _lblEquation_0 As System.Windows.Forms.Label
    Public WithEvents cboFatigueCurves As System.Windows.Forms.ComboBox
    Public WithEvents txtFatConstA As System.Windows.Forms.TextBox
    Public WithEvents txtFatConstB As System.Windows.Forms.TextBox
    Public WithEvents txtSCF As System.Windows.Forms.TextBox
    Public WithEvents _lblEquation_2 As System.Windows.Forms.Label
    Public WithEvents lblUnit As System.Windows.Forms.Label
    Public WithEvents lblFatConstA As System.Windows.Forms.Label
    Public WithEvents lblFatConstB As System.Windows.Forms.Label
    Public WithEvents lblSCF As System.Windows.Forms.Label
    Public WithEvents fraParms As System.Windows.Forms.GroupBox
    Public WithEvents txtDampingCoef As System.Windows.Forms.TextBox
    Public WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Public WithEvents btnNoHiRe As System.Windows.Forms.RadioButton
    Public WithEvents btnHiRe As System.Windows.Forms.RadioButton
#End Region
End Class