<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmProgress
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
	Public WithEvents prgProgress As System.Windows.Forms.ProgressBar
	Public WithEvents prgProgressBatch As System.Windows.Forms.ProgressBar
	Public WithEvents lblBatchPercent As System.Windows.Forms.Label
	Public WithEvents lblBatchStage As System.Windows.Forms.Label
	Public WithEvents CurrentStage As System.Windows.Forms.Label
	Public WithEvents lblProgress As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmProgress))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.btnCancel = New System.Windows.Forms.Button
		Me.prgProgress = New System.Windows.Forms.ProgressBar
		Me.prgProgressBatch = New System.Windows.Forms.ProgressBar
		Me.lblBatchPercent = New System.Windows.Forms.Label
		Me.lblBatchStage = New System.Windows.Forms.Label
		Me.CurrentStage = New System.Windows.Forms.Label
		Me.lblProgress = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = " Quasi-Static Analysis"
		Me.ClientSize = New System.Drawing.Size(304, 123)
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
		Me.Name = "frmProgress"
		Me.btnCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnCancel.Text = "Cancel"
		Me.btnCancel.Size = New System.Drawing.Size(81, 25)
		Me.btnCancel.Location = New System.Drawing.Point(112, 92)
		Me.btnCancel.TabIndex = 0
		Me.btnCancel.BackColor = System.Drawing.SystemColors.Control
		Me.btnCancel.CausesValidation = True
		Me.btnCancel.Enabled = True
		Me.btnCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnCancel.TabStop = True
		Me.btnCancel.Name = "btnCancel"
		Me.prgProgress.Size = New System.Drawing.Size(273, 17)
		Me.prgProgress.Location = New System.Drawing.Point(14, 34)
		Me.prgProgress.TabIndex = 3
		Me.prgProgress.Name = "prgProgress"
		Me.prgProgressBatch.Size = New System.Drawing.Size(273, 17)
		Me.prgProgressBatch.Location = New System.Drawing.Point(14, 73)
		Me.prgProgressBatch.TabIndex = 4
		Me.prgProgressBatch.Visible = False
		Me.prgProgressBatch.Name = "prgProgressBatch"
		Me.lblBatchPercent.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblBatchPercent.Text = "10% Completed"
		Me.lblBatchPercent.Size = New System.Drawing.Size(89, 17)
		Me.lblBatchPercent.Location = New System.Drawing.Point(199, 56)
		Me.lblBatchPercent.TabIndex = 6
		Me.lblBatchPercent.Visible = False
		Me.lblBatchPercent.BackColor = System.Drawing.SystemColors.Control
		Me.lblBatchPercent.Enabled = True
		Me.lblBatchPercent.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblBatchPercent.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblBatchPercent.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblBatchPercent.UseMnemonic = True
		Me.lblBatchPercent.AutoSize = False
		Me.lblBatchPercent.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblBatchPercent.Name = "lblBatchPercent"
		Me.lblBatchStage.Text = "Equilibrium location ..."
		Me.lblBatchStage.Size = New System.Drawing.Size(177, 17)
		Me.lblBatchStage.Location = New System.Drawing.Point(14, 56)
		Me.lblBatchStage.TabIndex = 5
		Me.lblBatchStage.Visible = False
		Me.lblBatchStage.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblBatchStage.BackColor = System.Drawing.SystemColors.Control
		Me.lblBatchStage.Enabled = True
		Me.lblBatchStage.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblBatchStage.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblBatchStage.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblBatchStage.UseMnemonic = True
		Me.lblBatchStage.AutoSize = False
		Me.lblBatchStage.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblBatchStage.Name = "lblBatchStage"
		Me.CurrentStage.Text = "Equilibrium location ..."
		Me.CurrentStage.Size = New System.Drawing.Size(177, 17)
		Me.CurrentStage.Location = New System.Drawing.Point(15, 16)
		Me.CurrentStage.TabIndex = 2
		Me.CurrentStage.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.CurrentStage.BackColor = System.Drawing.SystemColors.Control
		Me.CurrentStage.Enabled = True
		Me.CurrentStage.ForeColor = System.Drawing.SystemColors.ControlText
		Me.CurrentStage.Cursor = System.Windows.Forms.Cursors.Default
		Me.CurrentStage.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.CurrentStage.UseMnemonic = True
		Me.CurrentStage.Visible = True
		Me.CurrentStage.AutoSize = False
		Me.CurrentStage.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.CurrentStage.Name = "CurrentStage"
		Me.lblProgress.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblProgress.Text = "10% Completed"
		Me.lblProgress.Size = New System.Drawing.Size(89, 17)
		Me.lblProgress.Location = New System.Drawing.Point(200, 16)
		Me.lblProgress.TabIndex = 1
		Me.lblProgress.BackColor = System.Drawing.SystemColors.Control
		Me.lblProgress.Enabled = True
		Me.lblProgress.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblProgress.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblProgress.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblProgress.UseMnemonic = True
		Me.lblProgress.Visible = True
		Me.lblProgress.AutoSize = False
		Me.lblProgress.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblProgress.Name = "lblProgress"
		Me.Controls.Add(btnCancel)
		Me.Controls.Add(prgProgress)
		Me.Controls.Add(prgProgressBatch)
		Me.Controls.Add(lblBatchPercent)
		Me.Controls.Add(lblBatchStage)
		Me.Controls.Add(CurrentStage)
		Me.Controls.Add(lblProgress)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class