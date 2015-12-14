<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmProjDesc
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
	Public WithEvents btnSetMarsDir As System.Windows.Forms.Button
	Public WithEvents Drive1 As Microsoft.VisualBasic.Compatibility.VB6.DriveListBox
	Public WithEvents Dir1 As Microsoft.VisualBasic.Compatibility.VB6.DirListBox
	Public WithEvents txtVesselName As System.Windows.Forms.TextBox
	Public WithEvents File1 As Microsoft.VisualBasic.Compatibility.VB6.FileListBox
	Public WithEvents btnCancel As System.Windows.Forms.Button
	Public WithEvents btnOK As System.Windows.Forms.Button
	Public WithEvents txtProjDesc As System.Windows.Forms.TextBox
	Public WithEvents txtProjName As System.Windows.Forms.TextBox
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents lblMarsDir As System.Windows.Forms.Label
	Public WithEvents _lblGenCmt_0 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents lblProjDesc As System.Windows.Forms.Label
	Public WithEvents lblProjName As System.Windows.Forms.Label
	Public WithEvents lblGenCmt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmProjDesc))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.btnSetMarsDir = New System.Windows.Forms.Button
		Me.Drive1 = New Microsoft.VisualBasic.Compatibility.VB6.DriveListBox
		Me.Dir1 = New Microsoft.VisualBasic.Compatibility.VB6.DirListBox
		Me.txtVesselName = New System.Windows.Forms.TextBox
		Me.File1 = New Microsoft.VisualBasic.Compatibility.VB6.FileListBox
		Me.btnCancel = New System.Windows.Forms.Button
		Me.btnOK = New System.Windows.Forms.Button
		Me.txtProjDesc = New System.Windows.Forms.TextBox
		Me.txtProjName = New System.Windows.Forms.TextBox
		Me.Label3 = New System.Windows.Forms.Label
		Me.lblMarsDir = New System.Windows.Forms.Label
		Me._lblGenCmt_0 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.lblProjDesc = New System.Windows.Forms.Label
		Me.lblProjName = New System.Windows.Forms.Label
		Me.lblGenCmt = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.lblGenCmt, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = " Project"
		Me.ClientSize = New System.Drawing.Size(508, 360)
		Me.Location = New System.Drawing.Point(3, 22)
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
		Me.Name = "frmProjDesc"
		Me.btnSetMarsDir.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnSetMarsDir.Text = "Set Working Directory"
		Me.btnSetMarsDir.Size = New System.Drawing.Size(127, 27)
		Me.btnSetMarsDir.Location = New System.Drawing.Point(129, 279)
		Me.btnSetMarsDir.TabIndex = 13
		Me.btnSetMarsDir.BackColor = System.Drawing.SystemColors.Control
		Me.btnSetMarsDir.CausesValidation = True
		Me.btnSetMarsDir.Enabled = True
		Me.btnSetMarsDir.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnSetMarsDir.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnSetMarsDir.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnSetMarsDir.TabStop = True
		Me.btnSetMarsDir.Name = "btnSetMarsDir"
		Me.Drive1.Size = New System.Drawing.Size(249, 21)
		Me.Drive1.Location = New System.Drawing.Point(8, 152)
		Me.Drive1.TabIndex = 11
		Me.Drive1.BackColor = System.Drawing.SystemColors.Window
		Me.Drive1.CausesValidation = True
		Me.Drive1.Enabled = True
		Me.Drive1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Drive1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Drive1.TabStop = True
		Me.Drive1.Visible = True
		Me.Drive1.Name = "Drive1"
		Me.Dir1.Size = New System.Drawing.Size(249, 96)
		Me.Dir1.Location = New System.Drawing.Point(8, 176)
		Me.Dir1.TabIndex = 10
		Me.Dir1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Dir1.BackColor = System.Drawing.SystemColors.Window
		Me.Dir1.CausesValidation = True
		Me.Dir1.Enabled = True
		Me.Dir1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Dir1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Dir1.TabStop = True
		Me.Dir1.Visible = True
		Me.Dir1.Name = "Dir1"
		Me.txtVesselName.AutoSize = False
		Me.txtVesselName.BackColor = System.Drawing.SystemColors.Control
		Me.txtVesselName.Size = New System.Drawing.Size(153, 19)
		Me.txtVesselName.Location = New System.Drawing.Point(272, 24)
		Me.txtVesselName.ReadOnly = True
		Me.txtVesselName.TabIndex = 9
		Me.txtVesselName.AcceptsReturn = True
		Me.txtVesselName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtVesselName.CausesValidation = True
		Me.txtVesselName.Enabled = True
		Me.txtVesselName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtVesselName.HideSelection = True
		Me.txtVesselName.Maxlength = 0
		Me.txtVesselName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtVesselName.MultiLine = False
		Me.txtVesselName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtVesselName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtVesselName.TabStop = True
		Me.txtVesselName.Visible = True
		Me.txtVesselName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtVesselName.Name = "txtVesselName"
		Me.File1.Size = New System.Drawing.Size(153, 188)
		Me.File1.Location = New System.Drawing.Point(272, 63)
		Me.File1.Normal = False
		Me.File1.Pattern = "*.rao"
		Me.File1.TabIndex = 8
		Me.File1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.File1.Archive = True
		Me.File1.BackColor = System.Drawing.SystemColors.Window
		Me.File1.CausesValidation = True
		Me.File1.Enabled = True
		Me.File1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.File1.Hidden = False
		Me.File1.Cursor = System.Windows.Forms.Cursors.Default
		Me.File1.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.File1.ReadOnly = True
		Me.File1.System = False
		Me.File1.TabStop = True
		Me.File1.TopIndex = 0
		Me.File1.Visible = True
		Me.File1.Name = "File1"
		Me.btnCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.btnCancel
		Me.btnCancel.Text = "Cancel"
		Me.btnCancel.Size = New System.Drawing.Size(65, 25)
		Me.btnCancel.Location = New System.Drawing.Point(436, 54)
		Me.btnCancel.TabIndex = 1
		Me.btnCancel.BackColor = System.Drawing.SystemColors.Control
		Me.btnCancel.CausesValidation = True
		Me.btnCancel.Enabled = True
		Me.btnCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnCancel.TabStop = True
		Me.btnCancel.Name = "btnCancel"
		Me.btnOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnOK.Text = "OK"
		Me.btnOK.Size = New System.Drawing.Size(65, 25)
		Me.btnOK.Location = New System.Drawing.Point(436, 24)
		Me.btnOK.TabIndex = 0
		Me.btnOK.BackColor = System.Drawing.SystemColors.Control
		Me.btnOK.CausesValidation = True
		Me.btnOK.Enabled = True
		Me.btnOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnOK.TabStop = True
		Me.btnOK.Name = "btnOK"
		Me.txtProjDesc.AutoSize = False
		Me.txtProjDesc.Size = New System.Drawing.Size(249, 81)
		Me.txtProjDesc.Location = New System.Drawing.Point(8, 64)
		Me.txtProjDesc.Maxlength = 4096
		Me.txtProjDesc.MultiLine = True
		Me.txtProjDesc.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
		Me.txtProjDesc.TabIndex = 5
		Me.txtProjDesc.AcceptsReturn = True
		Me.txtProjDesc.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtProjDesc.BackColor = System.Drawing.SystemColors.Window
		Me.txtProjDesc.CausesValidation = True
		Me.txtProjDesc.Enabled = True
		Me.txtProjDesc.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtProjDesc.HideSelection = True
		Me.txtProjDesc.ReadOnly = False
		Me.txtProjDesc.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtProjDesc.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtProjDesc.TabStop = True
		Me.txtProjDesc.Visible = True
		Me.txtProjDesc.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtProjDesc.Name = "txtProjDesc"
		Me.txtProjName.AutoSize = False
		Me.txtProjName.Size = New System.Drawing.Size(249, 19)
		Me.txtProjName.Location = New System.Drawing.Point(8, 24)
		Me.txtProjName.TabIndex = 3
		Me.txtProjName.AcceptsReturn = True
		Me.txtProjName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtProjName.BackColor = System.Drawing.SystemColors.Window
		Me.txtProjName.CausesValidation = True
		Me.txtProjName.Enabled = True
		Me.txtProjName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtProjName.HideSelection = True
		Me.txtProjName.ReadOnly = False
		Me.txtProjName.Maxlength = 0
		Me.txtProjName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtProjName.MultiLine = False
		Me.txtProjName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtProjName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtProjName.TabStop = True
		Me.txtProjName.Visible = True
		Me.txtProjName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtProjName.Name = "txtProjName"
		Me.Label3.Text = "Working Directory:"
		Me.Label3.Size = New System.Drawing.Size(98, 17)
		Me.Label3.Location = New System.Drawing.Point(7, 297)
		Me.Label3.TabIndex = 15
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label3.BackColor = System.Drawing.SystemColors.Control
		Me.Label3.Enabled = True
		Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label3.UseMnemonic = True
		Me.Label3.Visible = True
		Me.Label3.AutoSize = False
		Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label3.Name = "Label3"
		Me.lblMarsDir.Size = New System.Drawing.Size(422, 24)
		Me.lblMarsDir.Location = New System.Drawing.Point(7, 313)
		Me.lblMarsDir.TabIndex = 14
		Me.lblMarsDir.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblMarsDir.BackColor = System.Drawing.SystemColors.Control
		Me.lblMarsDir.Enabled = True
		Me.lblMarsDir.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblMarsDir.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblMarsDir.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblMarsDir.UseMnemonic = True
		Me.lblMarsDir.Visible = True
		Me.lblMarsDir.AutoSize = False
		Me.lblMarsDir.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lblMarsDir.Name = "lblMarsDir"
		Me._lblGenCmt_0.Text = "The selected vessel data will be copied to the MARS working directory."
		Me._lblGenCmt_0.ForeColor = System.Drawing.Color.FromARGB(0, 0, 128)
		Me._lblGenCmt_0.Size = New System.Drawing.Size(157, 40)
		Me._lblGenCmt_0.Location = New System.Drawing.Point(271, 260)
		Me._lblGenCmt_0.TabIndex = 12
		Me._lblGenCmt_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblGenCmt_0.BackColor = System.Drawing.SystemColors.Control
		Me._lblGenCmt_0.Enabled = True
		Me._lblGenCmt_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblGenCmt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblGenCmt_0.UseMnemonic = True
		Me._lblGenCmt_0.Visible = True
		Me._lblGenCmt_0.AutoSize = False
		Me._lblGenCmt_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblGenCmt_0.Name = "_lblGenCmt_0"
		Me.Label2.Text = "Select Vessel:"
		Me.Label2.Size = New System.Drawing.Size(81, 17)
		Me.Label2.Location = New System.Drawing.Point(272, 48)
		Me.Label2.TabIndex = 7
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label2.BackColor = System.Drawing.SystemColors.Control
		Me.Label2.Enabled = True
		Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me.Label1.Text = "Current Vessel Name"
		Me.Label1.Size = New System.Drawing.Size(121, 17)
		Me.Label1.Location = New System.Drawing.Point(272, 8)
		Me.Label1.TabIndex = 6
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.BackColor = System.Drawing.SystemColors.Control
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.lblProjDesc.Text = "Project Description"
		Me.lblProjDesc.Size = New System.Drawing.Size(97, 17)
		Me.lblProjDesc.Location = New System.Drawing.Point(7, 48)
		Me.lblProjDesc.TabIndex = 4
		Me.lblProjDesc.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblProjDesc.BackColor = System.Drawing.SystemColors.Control
		Me.lblProjDesc.Enabled = True
		Me.lblProjDesc.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblProjDesc.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblProjDesc.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblProjDesc.UseMnemonic = True
		Me.lblProjDesc.Visible = True
		Me.lblProjDesc.AutoSize = False
		Me.lblProjDesc.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblProjDesc.Name = "lblProjDesc"
		Me.lblProjName.Text = "Project Name"
		Me.lblProjName.Size = New System.Drawing.Size(73, 17)
		Me.lblProjName.Location = New System.Drawing.Point(8, 8)
		Me.lblProjName.TabIndex = 2
		Me.lblProjName.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblProjName.BackColor = System.Drawing.SystemColors.Control
		Me.lblProjName.Enabled = True
		Me.lblProjName.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblProjName.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblProjName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblProjName.UseMnemonic = True
		Me.lblProjName.Visible = True
		Me.lblProjName.AutoSize = False
		Me.lblProjName.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblProjName.Name = "lblProjName"
		Me.Controls.Add(btnSetMarsDir)
		Me.Controls.Add(Drive1)
		Me.Controls.Add(Dir1)
		Me.Controls.Add(txtVesselName)
		Me.Controls.Add(File1)
		Me.Controls.Add(btnCancel)
		Me.Controls.Add(btnOK)
		Me.Controls.Add(txtProjDesc)
		Me.Controls.Add(txtProjName)
		Me.Controls.Add(Label3)
		Me.Controls.Add(lblMarsDir)
		Me.Controls.Add(_lblGenCmt_0)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		Me.Controls.Add(lblProjDesc)
		Me.Controls.Add(lblProjName)
		Me.lblGenCmt.SetIndex(_lblGenCmt_0, CType(0, Short))
		CType(Me.lblGenCmt, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class