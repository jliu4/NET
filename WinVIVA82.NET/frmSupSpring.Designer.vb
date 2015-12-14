<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmSupSpring
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
    Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSupSpring))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.grdSupSpring = New System.Windows.Forms.DataGridView
        Me.ContextMenuStripGridEdit = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.mnuAddRow = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuDeleteRow = New System.Windows.Forms.ToolStripMenuItem
        Me.pasteFromClipBoard = New System.Windows.Forms.ToolStripMenuItem
        Me.CopyToExcel = New System.Windows.Forms.ToolStripMenuItem
        CType(Me.grdSupSpring, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuStripGridEdit.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.SystemColors.Control
        Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCancel.Location = New System.Drawing.Point(249, 205)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCancel.Size = New System.Drawing.Size(65, 25)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'btnOK
        '
        Me.btnOK.BackColor = System.Drawing.SystemColors.Control
        Me.btnOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnOK.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnOK.Location = New System.Drawing.Point(178, 205)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnOK.Size = New System.Drawing.Size(65, 25)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "&Save"
        Me.btnOK.UseVisualStyleBackColor = False
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(41, 172)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(233, 33)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Start from the uppermost support and proceed towards the bottom."
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(7, 172)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(51, 33)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "*Note:"
        '
        'grdSupSpring
        '
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdSupSpring.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.grdSupSpring.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdSupSpring.ContextMenuStrip = Me.ContextMenuStripGridEdit
        Me.grdSupSpring.Location = New System.Drawing.Point(7, 26)
        Me.grdSupSpring.Name = "grdSupSpring"
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdSupSpring.RowHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.grdSupSpring.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.grdSupSpring.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdSupSpring.Size = New System.Drawing.Size(308, 130)
        Me.grdSupSpring.TabIndex = 6
        '
        'ContextMenuStripGridEdit
        '
        Me.ContextMenuStripGridEdit.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuAddRow, Me.mnuDeleteRow, Me.pasteFromClipBoard, Me.CopyToExcel})
        Me.ContextMenuStripGridEdit.Name = "ContextMenuStripGridEdit"
        Me.ContextMenuStripGridEdit.Size = New System.Drawing.Size(134, 92)
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
        'pasteFromClipBoard
        '
        Me.pasteFromClipBoard.Name = "pasteFromClipBoard"
        Me.pasteFromClipBoard.Size = New System.Drawing.Size(133, 22)
        Me.pasteFromClipBoard.Text = "Paste"
        '
        'CopyToExcel
        '
        Me.CopyToExcel.Name = "CopyToExcel"
        Me.CopyToExcel.Size = New System.Drawing.Size(133, 22)
        Me.CopyToExcel.Text = "Copy"
        '
        'frmSupSpring
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(326, 242)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.grdSupSpring)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(10, 29)
        Me.MaximizeBox = False
        Me.Name = "frmSupSpring"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Lateral Supports"
        CType(Me.grdSupSpring, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuStripGridEdit.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents grdSupSpring As System.Windows.Forms.DataGridView
    Friend WithEvents ContextMenuStripGridEdit As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents mnuAddRow As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuDeleteRow As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents pasteFromClipBoard As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CopyToExcel As System.Windows.Forms.ToolStripMenuItem
#End Region 
End Class