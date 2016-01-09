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
    Public WithEvents btnCancel As System.Windows.Forms.Button
    Public WithEvents bntOK As System.Windows.Forms.Button
    Public WithEvents txtNumProfilePts As System.Windows.Forms.TextBox
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents lblCaseNo As System.Windows.Forms.Label
    Public WithEvents _lblCurr_2 As System.Windows.Forms.Label
    Public WithEvents lblCurr As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnCopyAll = New System.Windows.Forms.Button()
        Me.btnCopy = New System.Windows.Forms.Button()
        Me.txtSameAsCase = New System.Windows.Forms.TextBox()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.bntOK = New System.Windows.Forms.Button()
        Me.txtNumProfilePts = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblCaseNo = New System.Windows.Forms.Label()
        Me._lblCurr_2 = New System.Windows.Forms.Label()
        Me.lblCurr = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.grdProfile = New System.Windows.Forms.DataGridView()
        CType(Me.lblCurr, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdProfile, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnCopyAll
        '
        Me.btnCopyAll.BackColor = System.Drawing.SystemColors.Control
        Me.btnCopyAll.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCopyAll.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCopyAll.Location = New System.Drawing.Point(136, 32)
        Me.btnCopyAll.Name = "btnCopyAll"
        Me.btnCopyAll.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCopyAll.Size = New System.Drawing.Size(65, 25)
        Me.btnCopyAll.TabIndex = 10
        Me.btnCopyAll.Text = "Copy All"
        Me.btnCopyAll.UseVisualStyleBackColor = False
        '
        'btnCopy
        '
        Me.btnCopy.BackColor = System.Drawing.SystemColors.Control
        Me.btnCopy.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCopy.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCopy.Location = New System.Drawing.Point(18, 31)
        Me.btnCopy.Name = "btnCopy"
        Me.btnCopy.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCopy.Size = New System.Drawing.Size(78, 25)
        Me.btnCopy.TabIndex = 8
        Me.btnCopy.Text = "Copy Case"
        Me.btnCopy.UseVisualStyleBackColor = False
        '
        'txtSameAsCase
        '
        Me.txtSameAsCase.AcceptsReturn = True
        Me.txtSameAsCase.BackColor = System.Drawing.SystemColors.Window
        Me.txtSameAsCase.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSameAsCase.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSameAsCase.Location = New System.Drawing.Point(102, 34)
        Me.txtSameAsCase.Name = "txtSameAsCase"
        Me.txtSameAsCase.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSameAsCase.Size = New System.Drawing.Size(24, 20)
        Me.txtSameAsCase.TabIndex = 7
        Me.txtSameAsCase.Text = "1"
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.SystemColors.Control
        Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCancel.Location = New System.Drawing.Point(218, 33)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCancel.Size = New System.Drawing.Size(58, 25)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'bntOK
        '
        Me.bntOK.BackColor = System.Drawing.SystemColors.Control
        Me.bntOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.bntOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.bntOK.Location = New System.Drawing.Point(218, 6)
        Me.bntOK.Name = "bntOK"
        Me.bntOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.bntOK.Size = New System.Drawing.Size(58, 25)
        Me.bntOK.TabIndex = 2
        Me.bntOK.Text = "OK"
        Me.bntOK.UseVisualStyleBackColor = False
        '
        'txtNumProfilePts
        '
        Me.txtNumProfilePts.AcceptsReturn = True
        Me.txtNumProfilePts.BackColor = System.Drawing.SystemColors.Window
        Me.txtNumProfilePts.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNumProfilePts.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNumProfilePts.Location = New System.Drawing.Point(158, 69)
        Me.txtNumProfilePts.Name = "txtNumProfilePts"
        Me.txtNumProfilePts.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNumProfilePts.Size = New System.Drawing.Size(43, 20)
        Me.txtNumProfilePts.TabIndex = 0
        Me.txtNumProfilePts.Text = "1"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.Color.Red
        Me.Label1.Location = New System.Drawing.Point(14, 341)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(230, 17)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Must include seabed depth as the  last row."
        '
        'lblCaseNo
        '
        Me.lblCaseNo.BackColor = System.Drawing.SystemColors.Control
        Me.lblCaseNo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCaseNo.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCaseNo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCaseNo.Location = New System.Drawing.Point(19, 7)
        Me.lblCaseNo.Name = "lblCaseNo"
        Me.lblCaseNo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCaseNo.Size = New System.Drawing.Size(160, 18)
        Me.lblCaseNo.TabIndex = 4
        '
        '_lblCurr_2
        '
        Me._lblCurr_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblCurr_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblCurr_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCurr.SetIndex(Me._lblCurr_2, CType(2, Short))
        Me._lblCurr_2.Location = New System.Drawing.Point(24, 72)
        Me._lblCurr_2.Name = "_lblCurr_2"
        Me._lblCurr_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblCurr_2.Size = New System.Drawing.Size(133, 17)
        Me._lblCurr_2.TabIndex = 1
        Me._lblCurr_2.Text = "Sub-surface Profile Points:"
        '
        'grdProfile
        '
        Me.grdProfile.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.grdProfile.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdProfile.Location = New System.Drawing.Point(17, 95)
        Me.grdProfile.Name = "grdProfile"
        Me.grdProfile.RowHeadersVisible = False
        Me.grdProfile.Size = New System.Drawing.Size(240, 223)
        Me.grdProfile.TabIndex = 11
        '
        'frmCurrProfile
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(284, 375)
        Me.Controls.Add(Me.grdProfile)
        Me.Controls.Add(Me.btnCopyAll)
        Me.Controls.Add(Me.btnCopy)
        Me.Controls.Add(Me.txtSameAsCase)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.bntOK)
        Me.Controls.Add(Me.txtNumProfilePts)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblCaseNo)
        Me.Controls.Add(Me._lblCurr_2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Location = New System.Drawing.Point(3, 29)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmCurrProfile"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Current Profile"
        CType(Me.lblCurr, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdProfile, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents grdProfile As DataGridView
#End Region
End Class