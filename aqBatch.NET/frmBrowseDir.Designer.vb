<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmBrowseDir
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
	Public WithEvents File1 As Microsoft.VisualBasic.Compatibility.VB6.FileListBox
	Public WithEvents btnCancel As System.Windows.Forms.Button
	Public WithEvents btnOK As System.Windows.Forms.Button
	Public WithEvents Dir1 As Microsoft.VisualBasic.Compatibility.VB6.DirListBox
	Public WithEvents Drive1 As Microsoft.VisualBasic.Compatibility.VB6.DriveListBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmBrowseDir))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.File1 = New Microsoft.VisualBasic.Compatibility.VB6.FileListBox
		Me.btnCancel = New System.Windows.Forms.Button
		Me.btnOK = New System.Windows.Forms.Button
		Me.Dir1 = New Microsoft.VisualBasic.Compatibility.VB6.DirListBox
		Me.Drive1 = New Microsoft.VisualBasic.Compatibility.VB6.DriveListBox
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Browse Working Directory"
		Me.ClientSize = New System.Drawing.Size(526, 253)
		Me.Location = New System.Drawing.Point(3, 29)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmBrowseDir"
		Me.File1.Archive = False
		Me.File1.BackColor = System.Drawing.SystemColors.Control
		Me.File1.CausesValidation = False
		Me.File1.Enabled = False
		Me.File1.Size = New System.Drawing.Size(118, 188)
		Me.File1.Location = New System.Drawing.Point(317, 41)
		Me.File1.Pattern = "a*run.*"
		Me.File1.TabIndex = 4
		Me.File1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.File1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.File1.Hidden = False
		Me.File1.Cursor = System.Windows.Forms.Cursors.Default
		Me.File1.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.File1.Normal = True
		Me.File1.ReadOnly = True
		Me.File1.System = False
		Me.File1.TabStop = True
		Me.File1.TopIndex = 0
		Me.File1.Visible = True
		Me.File1.Name = "File1"
		Me.btnCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.btnCancel
		Me.btnCancel.Text = "Cancel"
		Me.btnCancel.Size = New System.Drawing.Size(71, 26)
		Me.btnCancel.Location = New System.Drawing.Point(445, 48)
		Me.btnCancel.TabIndex = 3
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
		Me.btnOK.Size = New System.Drawing.Size(71, 26)
		Me.btnOK.Location = New System.Drawing.Point(445, 18)
		Me.btnOK.TabIndex = 2
		Me.btnOK.BackColor = System.Drawing.SystemColors.Control
		Me.btnOK.CausesValidation = True
		Me.btnOK.Enabled = True
		Me.btnOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnOK.TabStop = True
		Me.btnOK.Name = "btnOK"
		Me.Dir1.Size = New System.Drawing.Size(303, 216)
		Me.Dir1.Location = New System.Drawing.Point(8, 14)
		Me.Dir1.TabIndex = 1
		Me.Dir1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Dir1.BackColor = System.Drawing.SystemColors.Window
		Me.Dir1.CausesValidation = True
		Me.Dir1.Enabled = True
		Me.Dir1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Dir1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Dir1.TabStop = True
		Me.Dir1.Visible = True
		Me.Dir1.Name = "Dir1"
		Me.Drive1.Size = New System.Drawing.Size(118, 21)
		Me.Drive1.Location = New System.Drawing.Point(317, 15)
		Me.Drive1.TabIndex = 0
		Me.Drive1.BackColor = System.Drawing.SystemColors.Window
		Me.Drive1.CausesValidation = True
		Me.Drive1.Enabled = True
		Me.Drive1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Drive1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Drive1.TabStop = True
		Me.Drive1.Visible = True
		Me.Drive1.Name = "Drive1"
		Me.Controls.Add(File1)
		Me.Controls.Add(btnCancel)
		Me.Controls.Add(btnOK)
		Me.Controls.Add(Dir1)
		Me.Controls.Add(Drive1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class