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
	Public WithEvents btnCancel As System.Windows.Forms.Button
	Public WithEvents btnOK As System.Windows.Forms.Button
	Public WithEvents txtProjDesc As System.Windows.Forms.TextBox
	Public WithEvents txtProjName As System.Windows.Forms.TextBox
	Public WithEvents lblProjDesc As System.Windows.Forms.Label
	Public WithEvents lblProjName As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmProjDesc))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.btnCancel = New System.Windows.Forms.Button
		Me.btnOK = New System.Windows.Forms.Button
		Me.txtProjDesc = New System.Windows.Forms.TextBox
		Me.txtProjName = New System.Windows.Forms.TextBox
		Me.lblProjDesc = New System.Windows.Forms.Label
		Me.lblProjName = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Project Description"
		Me.ClientSize = New System.Drawing.Size(265, 243)
		Me.Location = New System.Drawing.Point(3, 22)
		Me.Icon = CType(resources.GetObject("frmProjDesc.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.btnCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnCancel.Text = "Cancel"
		Me.btnCancel.Size = New System.Drawing.Size(65, 25)
		Me.btnCancel.Location = New System.Drawing.Point(192, 208)
		Me.btnCancel.TabIndex = 3
		Me.btnCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.btnOK.Location = New System.Drawing.Point(120, 208)
		Me.btnOK.TabIndex = 2
		Me.btnOK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.btnOK.BackColor = System.Drawing.SystemColors.Control
		Me.btnOK.CausesValidation = True
		Me.btnOK.Enabled = True
		Me.btnOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnOK.TabStop = True
		Me.btnOK.Name = "btnOK"
		Me.txtProjDesc.AutoSize = False
		Me.txtProjDesc.Size = New System.Drawing.Size(249, 129)
		Me.txtProjDesc.Location = New System.Drawing.Point(8, 72)
		Me.txtProjDesc.Maxlength = 4096
		Me.txtProjDesc.MultiLine = True
		Me.txtProjDesc.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
		Me.txtProjDesc.TabIndex = 1
		Me.txtProjDesc.Text = "Text1"
		Me.txtProjDesc.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.txtProjName.TabIndex = 0
		Me.txtProjName.Text = "Text1"
		Me.txtProjName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.lblProjDesc.Text = "Project Description"
		Me.lblProjDesc.Size = New System.Drawing.Size(97, 17)
		Me.lblProjDesc.Location = New System.Drawing.Point(8, 56)
		Me.lblProjDesc.TabIndex = 5
		Me.lblProjDesc.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.lblProjName.TabIndex = 4
		Me.lblProjName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.Controls.Add(btnCancel)
		Me.Controls.Add(btnOK)
		Me.Controls.Add(txtProjDesc)
		Me.Controls.Add(txtProjName)
		Me.Controls.Add(lblProjDesc)
		Me.Controls.Add(lblProjName)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class