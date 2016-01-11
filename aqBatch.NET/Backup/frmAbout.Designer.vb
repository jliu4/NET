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
	Public WithEvents cmdSysInfo As System.Windows.Forms.Button
	Public WithEvents _lblDescription_1 As System.Windows.Forms.Label
	Public WithEvents lblVersion As System.Windows.Forms.Label
	Public WithEvents lblDisclaimer As System.Windows.Forms.Label
	Public WithEvents _Line1_1 As System.Windows.Forms.Label
	Public WithEvents _Line1_0 As System.Windows.Forms.Label
	Public WithEvents Line1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblDescription As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAbout))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.picIcon = New System.Windows.Forms.PictureBox
		Me.cmdOK = New System.Windows.Forms.Button
		Me.cmdSysInfo = New System.Windows.Forms.Button
		Me._lblDescription_1 = New System.Windows.Forms.Label
		Me.lblVersion = New System.Windows.Forms.Label
		Me.lblDisclaimer = New System.Windows.Forms.Label
		Me._Line1_1 = New System.Windows.Forms.Label
		Me._Line1_0 = New System.Windows.Forms.Label
		Me.Line1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.lblDescription = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.Line1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lblDescription, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = " About AQWA Batch Runner"
		Me.ClientSize = New System.Drawing.Size(378, 171)
		Me.Location = New System.Drawing.Point(156, 129)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmAbout"
		Me.picIcon.Size = New System.Drawing.Size(36, 36)
		Me.picIcon.Location = New System.Drawing.Point(14, 37)
		Me.picIcon.Image = CType(resources.GetObject("picIcon.Image"), System.Drawing.Image)
		Me.picIcon.TabIndex = 1
		Me.picIcon.Dock = System.Windows.Forms.DockStyle.None
		Me.picIcon.BackColor = System.Drawing.SystemColors.Control
		Me.picIcon.CausesValidation = True
		Me.picIcon.Enabled = True
		Me.picIcon.ForeColor = System.Drawing.SystemColors.ControlText
		Me.picIcon.Cursor = System.Windows.Forms.Cursors.Default
		Me.picIcon.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picIcon.TabStop = True
		Me.picIcon.Visible = True
		Me.picIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.picIcon.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.picIcon.Name = "picIcon"
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdOK
		Me.cmdOK.Text = "OK"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Size = New System.Drawing.Size(84, 23)
		Me.cmdOK.Location = New System.Drawing.Point(284, 112)
		Me.cmdOK.TabIndex = 0
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.cmdSysInfo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdSysInfo.Text = "&System Info..."
		Me.cmdSysInfo.Size = New System.Drawing.Size(83, 23)
		Me.cmdSysInfo.Location = New System.Drawing.Point(284, 139)
		Me.cmdSysInfo.TabIndex = 2
		Me.cmdSysInfo.Visible = False
		Me.cmdSysInfo.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSysInfo.CausesValidation = True
		Me.cmdSysInfo.Enabled = True
		Me.cmdSysInfo.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSysInfo.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSysInfo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSysInfo.TabStop = True
		Me.cmdSysInfo.Name = "cmdSysInfo"
		Me._lblDescription_1.Text = "Invalid_string_refer_to_original_code"
		Me._lblDescription_1.ForeColor = System.Drawing.Color.Black
		Me._lblDescription_1.Size = New System.Drawing.Size(259, 30)
		Me._lblDescription_1.Location = New System.Drawing.Point(60, 43)
		Me._lblDescription_1.TabIndex = 5
		Me._lblDescription_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblDescription_1.BackColor = System.Drawing.SystemColors.Control
		Me._lblDescription_1.Enabled = True
		Me._lblDescription_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblDescription_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblDescription_1.UseMnemonic = True
		Me._lblDescription_1.Visible = True
		Me._lblDescription_1.AutoSize = False
		Me._lblDescription_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblDescription_1.Name = "_lblDescription_1"
		Me.lblVersion.Text = "Version 1.00"
		Me.lblVersion.Size = New System.Drawing.Size(259, 15)
		Me.lblVersion.Location = New System.Drawing.Point(12, 13)
		Me.lblVersion.TabIndex = 4
		Me.lblVersion.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblVersion.BackColor = System.Drawing.SystemColors.Control
		Me.lblVersion.Enabled = True
		Me.lblVersion.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblVersion.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblVersion.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblVersion.UseMnemonic = True
		Me.lblVersion.Visible = True
		Me.lblVersion.AutoSize = False
		Me.lblVersion.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblVersion.Name = "lblVersion"
		Me.lblDisclaimer.Text = "Warning: This computer program is protected by copyright law."
		Me.lblDisclaimer.ForeColor = System.Drawing.Color.Black
		Me.lblDisclaimer.Size = New System.Drawing.Size(258, 36)
		Me.lblDisclaimer.Location = New System.Drawing.Point(13, 122)
		Me.lblDisclaimer.TabIndex = 3
		Me.lblDisclaimer.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblDisclaimer.BackColor = System.Drawing.SystemColors.Control
		Me.lblDisclaimer.Enabled = True
		Me.lblDisclaimer.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblDisclaimer.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblDisclaimer.UseMnemonic = True
		Me.lblDisclaimer.Visible = True
		Me.lblDisclaimer.AutoSize = False
		Me.lblDisclaimer.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblDisclaimer.Name = "lblDisclaimer"
		Me._Line1_1.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me._Line1_1.Visible = True
		Me._Line1_1.Location = New System.Drawing.Point(2, 131)
		Me._Line1_1.Size = New System.Drawing.Size(307, 1)
		Me._Line1_1.Name = "_Line1_1"
		Me._Line1_0.BackColor = System.Drawing.Color.White
		Me._Line1_0.Visible = True
		Me._Line1_0.Location = New System.Drawing.Point(3, 132)
		Me._Line1_0.Size = New System.Drawing.Size(306, 1)
		Me._Line1_0.Name = "_Line1_0"
		Me.Controls.Add(picIcon)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(cmdSysInfo)
		Me.Controls.Add(_lblDescription_1)
		Me.Controls.Add(lblVersion)
		Me.Controls.Add(lblDisclaimer)
		Me.Controls.Add(_Line1_1)
		Me.Controls.Add(_Line1_0)
		Me.Line1.SetIndex(_Line1_1, CType(1, Short))
		Me.Line1.SetIndex(_Line1_0, CType(0, Short))
		Me.lblDescription.SetIndex(_lblDescription_1, CType(1, Short))
		CType(Me.lblDescription, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Line1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class