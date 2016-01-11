<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmCurrProfile
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
	Public WithEvents btnCopyAll As System.Windows.Forms.Button
	Public WithEvents btnCopy As System.Windows.Forms.Button
	Public WithEvents txtSameAsCase As System.Windows.Forms.TextBox
	Public WithEvents txtEdit As System.Windows.Forms.TextBox
	Public WithEvents btnCancel As System.Windows.Forms.Button
	Public WithEvents bntOK As System.Windows.Forms.Button
	Public WithEvents txtNumProfilePts As System.Windows.Forms.TextBox
	Public WithEvents grdProfile As AxMSFlexGridLib.AxMSFlexGrid
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents lblCaseNo As System.Windows.Forms.Label
	Public WithEvents _lblCurr_2 As System.Windows.Forms.Label
	Public WithEvents lblCurr As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmCurrProfile))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.btnCopyAll = New System.Windows.Forms.Button
		Me.btnCopy = New System.Windows.Forms.Button
		Me.txtSameAsCase = New System.Windows.Forms.TextBox
		Me.txtEdit = New System.Windows.Forms.TextBox
		Me.btnCancel = New System.Windows.Forms.Button
		Me.bntOK = New System.Windows.Forms.Button
		Me.txtNumProfilePts = New System.Windows.Forms.TextBox
		Me.grdProfile = New AxMSFlexGridLib.AxMSFlexGrid
		Me.Label1 = New System.Windows.Forms.Label
		Me.lblCaseNo = New System.Windows.Forms.Label
		Me._lblCurr_2 = New System.Windows.Forms.Label
		Me.lblCurr = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.grdProfile, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lblCurr, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "Current Profile"
		Me.ClientSize = New System.Drawing.Size(284, 375)
		Me.Location = New System.Drawing.Point(3, 29)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmCurrProfile"
		Me.btnCopyAll.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnCopyAll.Text = "Copy All"
		Me.btnCopyAll.Size = New System.Drawing.Size(65, 25)
		Me.btnCopyAll.Location = New System.Drawing.Point(136, 32)
		Me.btnCopyAll.TabIndex = 10
		Me.btnCopyAll.BackColor = System.Drawing.SystemColors.Control
		Me.btnCopyAll.CausesValidation = True
		Me.btnCopyAll.Enabled = True
		Me.btnCopyAll.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnCopyAll.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnCopyAll.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnCopyAll.TabStop = True
		Me.btnCopyAll.Name = "btnCopyAll"
		Me.btnCopy.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnCopy.Text = "Copy Case"
		Me.btnCopy.Size = New System.Drawing.Size(78, 25)
		Me.btnCopy.Location = New System.Drawing.Point(18, 31)
		Me.btnCopy.TabIndex = 8
		Me.btnCopy.BackColor = System.Drawing.SystemColors.Control
		Me.btnCopy.CausesValidation = True
		Me.btnCopy.Enabled = True
		Me.btnCopy.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnCopy.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnCopy.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnCopy.TabStop = True
		Me.btnCopy.Name = "btnCopy"
		Me.txtSameAsCase.AutoSize = False
		Me.txtSameAsCase.Size = New System.Drawing.Size(24, 19)
		Me.txtSameAsCase.Location = New System.Drawing.Point(102, 34)
		Me.txtSameAsCase.TabIndex = 7
		Me.txtSameAsCase.Text = "1"
		Me.txtSameAsCase.AcceptsReturn = True
		Me.txtSameAsCase.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtSameAsCase.BackColor = System.Drawing.SystemColors.Window
		Me.txtSameAsCase.CausesValidation = True
		Me.txtSameAsCase.Enabled = True
		Me.txtSameAsCase.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtSameAsCase.HideSelection = True
		Me.txtSameAsCase.ReadOnly = False
		Me.txtSameAsCase.Maxlength = 0
		Me.txtSameAsCase.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtSameAsCase.MultiLine = False
		Me.txtSameAsCase.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtSameAsCase.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtSameAsCase.TabStop = True
		Me.txtSameAsCase.Visible = True
		Me.txtSameAsCase.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtSameAsCase.Name = "txtSameAsCase"
		Me.txtEdit.AutoSize = False
		Me.txtEdit.Size = New System.Drawing.Size(59, 18)
		Me.txtEdit.Location = New System.Drawing.Point(190, 195)
		Me.txtEdit.TabIndex = 6
		Me.txtEdit.Visible = False
		Me.txtEdit.AcceptsReturn = True
		Me.txtEdit.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtEdit.BackColor = System.Drawing.SystemColors.Window
		Me.txtEdit.CausesValidation = True
		Me.txtEdit.Enabled = True
		Me.txtEdit.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtEdit.HideSelection = True
		Me.txtEdit.ReadOnly = False
		Me.txtEdit.Maxlength = 0
		Me.txtEdit.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtEdit.MultiLine = False
		Me.txtEdit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtEdit.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtEdit.TabStop = True
		Me.txtEdit.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtEdit.Name = "txtEdit"
		Me.btnCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.btnCancel
		Me.btnCancel.Text = "Cancel"
		Me.btnCancel.Size = New System.Drawing.Size(58, 25)
		Me.btnCancel.Location = New System.Drawing.Point(218, 33)
		Me.btnCancel.TabIndex = 3
		Me.btnCancel.BackColor = System.Drawing.SystemColors.Control
		Me.btnCancel.CausesValidation = True
		Me.btnCancel.Enabled = True
		Me.btnCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnCancel.TabStop = True
		Me.btnCancel.Name = "btnCancel"
		Me.bntOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.bntOK.Text = "OK"
		Me.bntOK.Size = New System.Drawing.Size(58, 25)
		Me.bntOK.Location = New System.Drawing.Point(218, 6)
		Me.bntOK.TabIndex = 2
		Me.bntOK.BackColor = System.Drawing.SystemColors.Control
		Me.bntOK.CausesValidation = True
		Me.bntOK.Enabled = True
		Me.bntOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.bntOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.bntOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.bntOK.TabStop = True
		Me.bntOK.Name = "bntOK"
		Me.txtNumProfilePts.AutoSize = False
		Me.txtNumProfilePts.Size = New System.Drawing.Size(43, 20)
		Me.txtNumProfilePts.Location = New System.Drawing.Point(158, 69)
		Me.txtNumProfilePts.TabIndex = 0
		Me.txtNumProfilePts.Text = "1"
		Me.txtNumProfilePts.AcceptsReturn = True
		Me.txtNumProfilePts.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtNumProfilePts.BackColor = System.Drawing.SystemColors.Window
		Me.txtNumProfilePts.CausesValidation = True
		Me.txtNumProfilePts.Enabled = True
		Me.txtNumProfilePts.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtNumProfilePts.HideSelection = True
		Me.txtNumProfilePts.ReadOnly = False
		Me.txtNumProfilePts.Maxlength = 0
		Me.txtNumProfilePts.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtNumProfilePts.MultiLine = False
		Me.txtNumProfilePts.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtNumProfilePts.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtNumProfilePts.TabStop = True
		Me.txtNumProfilePts.Visible = True
		Me.txtNumProfilePts.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtNumProfilePts.Name = "txtNumProfilePts"
		grdProfile.OcxState = CType(resources.GetObject("grdProfile.OcxState"), System.Windows.Forms.AxHost.State)
		Me.grdProfile.Size = New System.Drawing.Size(237, 208)
		Me.grdProfile.Location = New System.Drawing.Point(20, 100)
		Me.grdProfile.TabIndex = 5
		Me.grdProfile.Name = "grdProfile"
		Me.Label1.Text = "Must include seabed depth as the  last row."
		Me.Label1.ForeColor = System.Drawing.Color.Red
		Me.Label1.Size = New System.Drawing.Size(230, 17)
		Me.Label1.Location = New System.Drawing.Point(14, 341)
		Me.Label1.TabIndex = 9
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.BackColor = System.Drawing.SystemColors.Control
		Me.Label1.Enabled = True
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.lblCaseNo.Size = New System.Drawing.Size(160, 18)
		Me.lblCaseNo.Location = New System.Drawing.Point(19, 7)
		Me.lblCaseNo.TabIndex = 4
		Me.lblCaseNo.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblCaseNo.BackColor = System.Drawing.SystemColors.Control
		Me.lblCaseNo.Enabled = True
		Me.lblCaseNo.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblCaseNo.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblCaseNo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblCaseNo.UseMnemonic = True
		Me.lblCaseNo.Visible = True
		Me.lblCaseNo.AutoSize = False
		Me.lblCaseNo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lblCaseNo.Name = "lblCaseNo"
		Me._lblCurr_2.Text = "Sub-surface Profile Points:"
		Me._lblCurr_2.Size = New System.Drawing.Size(133, 17)
		Me._lblCurr_2.Location = New System.Drawing.Point(24, 72)
		Me._lblCurr_2.TabIndex = 1
		Me._lblCurr_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblCurr_2.BackColor = System.Drawing.SystemColors.Control
		Me._lblCurr_2.Enabled = True
		Me._lblCurr_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblCurr_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblCurr_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblCurr_2.UseMnemonic = True
		Me._lblCurr_2.Visible = True
		Me._lblCurr_2.AutoSize = False
		Me._lblCurr_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblCurr_2.Name = "_lblCurr_2"
		Me.Controls.Add(btnCopyAll)
		Me.Controls.Add(btnCopy)
		Me.Controls.Add(txtSameAsCase)
		Me.Controls.Add(txtEdit)
		Me.Controls.Add(btnCancel)
		Me.Controls.Add(bntOK)
		Me.Controls.Add(txtNumProfilePts)
		Me.Controls.Add(grdProfile)
		Me.Controls.Add(Label1)
		Me.Controls.Add(lblCaseNo)
		Me.Controls.Add(_lblCurr_2)
		Me.lblCurr.SetIndex(_lblCurr_2, CType(2, Short))
		CType(Me.lblCurr, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.grdProfile, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class