<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmAuxLines
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
	Public WithEvents chkUseInEICalc As System.Windows.Forms.CheckBox
    Public WithEvents btnCancel As System.Windows.Forms.Button
	Public WithEvents btnOK As System.Windows.Forms.Button
    Public WithEvents mnuAddRow As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuDeleteRow As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuGridEdit As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAuxLines))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.chkUseInEICalc = New System.Windows.Forms.CheckBox
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip
        Me.mnuGridEdit = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuAddRow = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuDeleteRow = New System.Windows.Forms.ToolStripMenuItem
        Me.ContextMenuStripGridEdit = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.pasteFromClipboard = New System.Windows.Forms.ToolStripMenuItem
        Me.copyToExcel = New System.Windows.Forms.ToolStripMenuItem
        Me.grdAuxLines = New System.Windows.Forms.DataGridView
        Me.MainMenu1.SuspendLayout()
        Me.ContextMenuStripGridEdit.SuspendLayout()
        CType(Me.grdAuxLines, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'chkUseInEICalc
        '
        Me.chkUseInEICalc.BackColor = System.Drawing.SystemColors.Control
        Me.chkUseInEICalc.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkUseInEICalc.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkUseInEICalc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkUseInEICalc.Location = New System.Drawing.Point(12, 184)
        Me.chkUseInEICalc.Name = "chkUseInEICalc"
        Me.chkUseInEICalc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkUseInEICalc.Size = New System.Drawing.Size(233, 17)
        Me.chkUseInEICalc.TabIndex = 1
        Me.chkUseInEICalc.Text = "&Use auxiliary lines in EI calculations"
        Me.chkUseInEICalc.UseVisualStyleBackColor = False
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.SystemColors.Control
        Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCancel.Location = New System.Drawing.Point(245, 207)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCancel.Size = New System.Drawing.Size(65, 25)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'btnOK
        '
        Me.btnOK.BackColor = System.Drawing.SystemColors.Control
        Me.btnOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnOK.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnOK.Location = New System.Drawing.Point(169, 207)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnOK.Size = New System.Drawing.Size(65, 25)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "&Save"
        Me.btnOK.UseVisualStyleBackColor = False
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuGridEdit})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(323, 24)
        Me.MainMenu1.TabIndex = 5
        '
        'mnuGridEdit
        '
        Me.mnuGridEdit.Name = "mnuGridEdit"
        Me.mnuGridEdit.Size = New System.Drawing.Size(12, 20)
        Me.mnuGridEdit.Visible = False
        '
        'mnuAddRow
        '
        Me.mnuAddRow.Name = "mnuAddRow"
        Me.mnuAddRow.Size = New System.Drawing.Size(133, 22)
        Me.mnuAddRow.Text = "Insert Row"
        '
        'mnuDeleteRow
        '
        Me.mnuDeleteRow.Name = "mnuDeleteRow"
        Me.mnuDeleteRow.Size = New System.Drawing.Size(133, 22)
        Me.mnuDeleteRow.Text = "Delete Row"
        '
        'ContextMenuStripGridEdit
        '
        Me.ContextMenuStripGridEdit.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuAddRow, Me.mnuDeleteRow, Me.pasteFromClipboard, Me.copyToExcel})
        Me.ContextMenuStripGridEdit.Name = "ContextMenuStrip1"
        Me.ContextMenuStripGridEdit.Size = New System.Drawing.Size(134, 92)
        '
        'pasteFromClipboard
        '
        Me.pasteFromClipboard.Name = "pasteFromClipboard"
        Me.pasteFromClipboard.Size = New System.Drawing.Size(133, 22)
        Me.pasteFromClipboard.Text = "Paste"
        '
        'copyToExcel
        '
        Me.copyToExcel.Name = "copyToExcel"
        Me.copyToExcel.Size = New System.Drawing.Size(133, 22)
        Me.copyToExcel.Text = "Copy"
        '
        'grdAuxLines
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdAuxLines.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.grdAuxLines.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdAuxLines.ContextMenuStrip = Me.ContextMenuStripGridEdit
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grdAuxLines.DefaultCellStyle = DataGridViewCellStyle2
        Me.grdAuxLines.Location = New System.Drawing.Point(12, 27)
        Me.grdAuxLines.Name = "grdAuxLines"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdAuxLines.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.grdAuxLines.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.grdAuxLines.Size = New System.Drawing.Size(298, 147)
        Me.grdAuxLines.TabIndex = 6
        '
        'frmAuxLines
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(323, 245)
        Me.Controls.Add(Me.chkUseInEICalc)
        Me.Controls.Add(Me.MainMenu1)
        Me.Controls.Add(Me.grdAuxLines)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(10, 29)
        Me.MaximizeBox = False
        Me.Name = "frmAuxLines"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Auxiliary Line Data"
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        Me.ContextMenuStripGridEdit.ResumeLayout(False)
        CType(Me.grdAuxLines, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ContextMenuStripGridEdit As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents grdAuxLines As System.Windows.Forms.DataGridView
    Friend WithEvents pasteFromClipboard As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents copyToExcel As System.Windows.Forms.ToolStripMenuItem
#End Region
End Class