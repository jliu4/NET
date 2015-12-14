<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmAbout
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
	Public WithEvents picIcon As System.Windows.Forms.PictureBox
	Public WithEvents cmdOK As System.Windows.Forms.Button
    Public WithEvents Line1 As System.Windows.Forms.Label
	Public WithEvents _lblDisclaimer_2 As System.Windows.Forms.Label
	Public WithEvents _lblDisclaimer_1 As System.Windows.Forms.Label
	Public WithEvents lblTitle As System.Windows.Forms.Label
	Public WithEvents lblVersion As System.Windows.Forms.Label
	Public WithEvents _lblDisclaimer_0 As System.Windows.Forms.Label
	Public WithEvents lblDisclaimer As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAbout))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.picIcon = New System.Windows.Forms.PictureBox
        Me.cmdOK = New System.Windows.Forms.Button
        Me.Line1 = New System.Windows.Forms.Label
        Me._lblDisclaimer_2 = New System.Windows.Forms.Label
        Me._lblDisclaimer_1 = New System.Windows.Forms.Label
        Me.lblTitle = New System.Windows.Forms.Label
        Me.lblVersion = New System.Windows.Forms.Label
        Me._lblDisclaimer_0 = New System.Windows.Forms.Label
        Me.lblDisclaimer = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.cmdSysInfo = New System.Windows.Forms.Button
        CType(Me.picIcon, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblDisclaimer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'picIcon
        '
        Me.picIcon.BackColor = System.Drawing.SystemColors.Control
        Me.picIcon.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.picIcon.Cursor = System.Windows.Forms.Cursors.Default
        Me.picIcon.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.picIcon.ForeColor = System.Drawing.SystemColors.ControlText
        Me.picIcon.Image = CType(resources.GetObject("picIcon.Image"), System.Drawing.Image)
        Me.picIcon.Location = New System.Drawing.Point(16, 16)
        Me.picIcon.Name = "picIcon"
        Me.picIcon.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.picIcon.Size = New System.Drawing.Size(36, 36)
        Me.picIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.picIcon.TabIndex = 1
        Me.picIcon.TabStop = False
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOK.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdOK.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOK.Location = New System.Drawing.Point(274, 112)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOK.Size = New System.Drawing.Size(94, 25)
        Me.cmdOK.TabIndex = 0
        Me.cmdOK.Text = "OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'Line1
        '
        Me.Line1.BackColor = System.Drawing.SystemColors.WindowText
        Me.Line1.Location = New System.Drawing.Point(10, 96)
        Me.Line1.Name = "Line1"
        Me.Line1.Size = New System.Drawing.Size(360, 1)
        Me.Line1.TabIndex = 4
        '
        '_lblDisclaimer_2
        '
        Me._lblDisclaimer_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblDisclaimer_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDisclaimer_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDisclaimer_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDisclaimer.SetIndex(Me._lblDisclaimer_2, CType(2, Short))
        Me._lblDisclaimer_2.Location = New System.Drawing.Point(16, 136)
        Me._lblDisclaimer_2.Name = "_lblDisclaimer_2"
        Me._lblDisclaimer_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDisclaimer_2.Size = New System.Drawing.Size(239, 39)
        Me._lblDisclaimer_2.TabIndex = 7
        Me._lblDisclaimer_2.Text = "Copyright of VIVA: Drs. Michael Triantafyllow and George Triantafyllou"
        '
        '_lblDisclaimer_1
        '
        Me._lblDisclaimer_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblDisclaimer_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDisclaimer_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDisclaimer_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDisclaimer.SetIndex(Me._lblDisclaimer_1, CType(1, Short))
        Me._lblDisclaimer_1.Location = New System.Drawing.Point(74, 112)
        Me._lblDisclaimer_1.Name = "_lblDisclaimer_1"
        Me._lblDisclaimer_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDisclaimer_1.Size = New System.Drawing.Size(194, 12)
        Me._lblDisclaimer_1.TabIndex = 6
        Me._lblDisclaimer_1.Text = "Genesis Engineering, 1998 - 2015"
        '
        'lblTitle
        '
        Me.lblTitle.BackColor = System.Drawing.SystemColors.Control
        Me.lblTitle.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTitle.Font = New System.Drawing.Font("Arial", 13.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblTitle.Location = New System.Drawing.Point(70, 16)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTitle.Size = New System.Drawing.Size(259, 32)
        Me.lblTitle.TabIndex = 4
        Me.lblTitle.Text = "WinVIVA"
        '
        'lblVersion
        '
        Me.lblVersion.BackColor = System.Drawing.SystemColors.Control
        Me.lblVersion.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblVersion.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVersion.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVersion.Location = New System.Drawing.Point(70, 52)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblVersion.Size = New System.Drawing.Size(259, 23)
        Me.lblVersion.TabIndex = 5
        Me.lblVersion.Text = "Version: 8.2.0.0"
        '
        '_lblDisclaimer_0
        '
        Me._lblDisclaimer_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblDisclaimer_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDisclaimer_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDisclaimer_0.ForeColor = System.Drawing.Color.Black
        Me.lblDisclaimer.SetIndex(Me._lblDisclaimer_0, CType(0, Short))
        Me._lblDisclaimer_0.Location = New System.Drawing.Point(16, 112)
        Me._lblDisclaimer_0.Name = "_lblDisclaimer_0"
        Me._lblDisclaimer_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDisclaimer_0.Size = New System.Drawing.Size(60, 39)
        Me._lblDisclaimer_0.TabIndex = 3
        Me._lblDisclaimer_0.Text = "Copyright:"
        '
        'cmdSysInfo
        '
        Me.cmdSysInfo.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSysInfo.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSysInfo.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSysInfo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSysInfo.Location = New System.Drawing.Point(274, 144)
        Me.cmdSysInfo.Name = "cmdSysInfo"
        Me.cmdSysInfo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSysInfo.Size = New System.Drawing.Size(94, 25)
        Me.cmdSysInfo.TabIndex = 2
        Me.cmdSysInfo.Text = "&System Info..."
        Me.cmdSysInfo.UseVisualStyleBackColor = False
        '
        'frmAbout
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.cmdOK
        Me.ClientSize = New System.Drawing.Size(380, 184)
        Me.Controls.Add(Me.lblTitle)
        Me.Controls.Add(Me.picIcon)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.cmdSysInfo)
        Me.Controls.Add(Me.Line1)
        Me.Controls.Add(Me._lblDisclaimer_2)
        Me.Controls.Add(Me._lblDisclaimer_1)
        Me.Controls.Add(Me.lblVersion)
        Me.Controls.Add(Me._lblDisclaimer_0)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(156, 129)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAbout"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = " About WinVIVA"
        CType(Me.picIcon, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblDisclaimer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents cmdSysInfo As System.Windows.Forms.Button
#End Region 
End Class