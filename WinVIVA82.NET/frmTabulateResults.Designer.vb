<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmTabulateResults
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
    Public WithEvents btnCancel As System.Windows.Forms.Button
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTabulateResults))
        Me.btnCancel = New System.Windows.Forms.Button
        Me.tabResults = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.grdResults = New System.Windows.Forms.DataGridView
        Me.FileDialogOpen = New System.Windows.Forms.OpenFileDialog
        Me.FileDialogSave = New System.Windows.Forms.SaveFileDialog
        Me.DetailResults = New System.Windows.Forms.GroupBox
        Me.Save = New System.Windows.Forms.Button
        Me.AC = New System.Windows.Forms.RadioButton
        Me.AM = New System.Windows.Forms.RadioButton
        Me.MM = New System.Windows.Forms.RadioButton
        Me.tabResults.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.grdResults, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.DetailResults.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.SystemColors.Control
        Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCancel.Location = New System.Drawing.Point(641, 464)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCancel.Size = New System.Drawing.Size(65, 25)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Return"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'tabResults
        '
        Me.tabResults.Controls.Add(Me.TabPage1)
        Me.tabResults.Location = New System.Drawing.Point(27, 12)
        Me.tabResults.Name = "tabResults"
        Me.tabResults.SelectedIndex = 0
        Me.tabResults.Size = New System.Drawing.Size(676, 443)
        Me.tabResults.TabIndex = 7
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.grdResults)
        Me.TabPage1.Location = New System.Drawing.Point(4, 23)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(668, 416)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Fatigue Summary"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'grdResults
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdResults.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.grdResults.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdResults.Location = New System.Drawing.Point(6, 6)
        Me.grdResults.Name = "grdResults"
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdResults.RowHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.grdResults.Size = New System.Drawing.Size(656, 407)
        Me.grdResults.TabIndex = 0
        '
        'FileDialogOpen
        '
        Me.FileDialogOpen.FileName = "FileDialogOpen"
        '
        'FileDialogSave
        '
        Me.FileDialogSave.FileName = "FileDialogSave"
        '
        'DetailResults
        '
        Me.DetailResults.Controls.Add(Me.Save)
        Me.DetailResults.Controls.Add(Me.AC)
        Me.DetailResults.Controls.Add(Me.AM)
        Me.DetailResults.Controls.Add(Me.MM)
        Me.DetailResults.Location = New System.Drawing.Point(31, 457)
        Me.DetailResults.Name = "DetailResults"
        Me.DetailResults.Size = New System.Drawing.Size(337, 42)
        Me.DetailResults.TabIndex = 8
        Me.DetailResults.TabStop = False
        Me.DetailResults.Text = "Detail Results"
        '
        'Save
        '
        Me.Save.Location = New System.Drawing.Point(245, 14)
        Me.Save.Name = "Save"
        Me.Save.Size = New System.Drawing.Size(75, 23)
        Me.Save.TabIndex = 9
        Me.Save.Text = "Save"
        Me.Save.UseVisualStyleBackColor = True
        '
        'AC
        '
        Me.AC.AutoSize = True
        Me.AC.Location = New System.Drawing.Point(140, 16)
        Me.AC.Name = "AC"
        Me.AC.Size = New System.Drawing.Size(99, 18)
        Me.AC.TabIndex = 3
        Me.AC.Text = "All by Category"
        Me.AC.UseVisualStyleBackColor = True
        '
        'AM
        '
        Me.AM.AutoSize = True
        Me.AM.Location = New System.Drawing.Point(53, 16)
        Me.AM.Name = "AM"
        Me.AM.Size = New System.Drawing.Size(81, 18)
        Me.AM.TabIndex = 2
        Me.AM.Text = "All by Mode"
        Me.AM.UseVisualStyleBackColor = True
        '
        'MM
        '
        Me.MM.AutoSize = True
        Me.MM.Checked = True
        Me.MM.Location = New System.Drawing.Point(6, 16)
        Me.MM.Name = "MM"
        Me.MM.Size = New System.Drawing.Size(41, 18)
        Me.MM.TabIndex = 1
        Me.MM.TabStop = True
        Me.MM.Text = "MM"
        Me.MM.UseVisualStyleBackColor = True
        '
        'frmTabulateResults
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(715, 503)
        Me.Controls.Add(Me.DetailResults)
        Me.Controls.Add(Me.tabResults)
        Me.Controls.Add(Me.btnCancel)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.ImeMode = System.Windows.Forms.ImeMode.[On]
        Me.Location = New System.Drawing.Point(10, 48)
        Me.MaximizeBox = False
        Me.Name = "frmTabulateResults"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Tabular Results"
        Me.tabResults.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.grdResults, System.ComponentModel.ISupportInitialize).EndInit()
        Me.DetailResults.ResumeLayout(False)
        Me.DetailResults.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Public WithEvents grdResults As System.Windows.Forms.DataGridView
    Public WithEvents tabResults As System.Windows.Forms.TabControl
    Friend WithEvents FileDialogOpen As System.Windows.Forms.OpenFileDialog
    Friend WithEvents FileDialogSave As System.Windows.Forms.SaveFileDialog
    Friend WithEvents DetailResults As System.Windows.Forms.GroupBox
    Friend WithEvents MM As System.Windows.Forms.RadioButton
    Friend WithEvents AM As System.Windows.Forms.RadioButton
    Friend WithEvents AC As System.Windows.Forms.RadioButton
    Friend WithEvents Save As System.Windows.Forms.Button
#End Region
End Class