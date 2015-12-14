<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmSegments
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
    Public WithEvents _lblNote_0 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents _lblNote_1 As System.Windows.Forms.Label
	Public WithEvents lblNote As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSegments))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me._lblNote_0 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me._lblNote_1 = New System.Windows.Forms.Label
        Me.lblNote = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.mnuAddRow = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuDeleteRow = New System.Windows.Forms.ToolStripMenuItem
        Me.ContextMenuStripGridEdit = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.pasteFromClipboard = New System.Windows.Forms.ToolStripMenuItem
        Me.copyToExcel = New System.Windows.Forms.ToolStripMenuItem
        Me.AddToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.RdBttnStiffness = New System.Windows.Forms.RadioButton
        Me.RdBttnYoungsModules = New System.Windows.Forms.RadioButton
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.txtVectoAngle = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.grdSegments = New System.Windows.Forms.DataGridView
        CType(Me.lblNote, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuStripGridEdit.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.grdSegments, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.SystemColors.Control
        Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCancel.Location = New System.Drawing.Point(771, 419)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCancel.Size = New System.Drawing.Size(65, 24)
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
        Me.btnOK.Location = New System.Drawing.Point(675, 419)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnOK.Size = New System.Drawing.Size(65, 24)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "&Save"
        Me.btnOK.UseVisualStyleBackColor = False
        '
        '_lblNote_0
        '
        Me._lblNote_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblNote_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblNote_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblNote_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblNote.SetIndex(Me._lblNote_0, CType(0, Short))
        Me._lblNote_0.Location = New System.Drawing.Point(11, 410)
        Me._lblNote_0.Name = "_lblNote_0"
        Me._lblNote_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblNote_0.Size = New System.Drawing.Size(41, 16)
        Me._lblNote_0.TabIndex = 6
        Me._lblNote_0.Text = "* Note : "
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(48, 427)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(621, 19)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "2. Joint weights include weight of riser, auxiliary lines and buoyancy modules."
        '
        '_lblNote_1
        '
        Me._lblNote_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblNote_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblNote_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblNote_1.ForeColor = System.Drawing.Color.Red
        Me.lblNote.SetIndex(Me._lblNote_1, CType(1, Short))
        Me._lblNote_1.Location = New System.Drawing.Point(48, 410)
        Me._lblNote_1.Name = "_lblNote_1"
        Me._lblNote_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblNote_1.Size = New System.Drawing.Size(601, 16)
        Me._lblNote_1.TabIndex = 3
        Me._lblNote_1.Text = "1. Starting from the bottom of the riser, then moving toward the water surface."
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
        Me.ContextMenuStripGridEdit.Name = "ContextMenuStripGridEdit"
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
        'AddToolStripMenuItem
        '
        Me.AddToolStripMenuItem.Name = "AddToolStripMenuItem"
        Me.AddToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.AddToolStripMenuItem.Text = "Add"
        '
        'RdBttnStiffness
        '
        Me.RdBttnStiffness.AutoSize = True
        Me.RdBttnStiffness.Checked = True
        Me.RdBttnStiffness.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RdBttnStiffness.Location = New System.Drawing.Point(15, 19)
        Me.RdBttnStiffness.Name = "RdBttnStiffness"
        Me.RdBttnStiffness.Size = New System.Drawing.Size(176, 18)
        Me.RdBttnStiffness.TabIndex = 10
        Me.RdBttnStiffness.TabStop = True
        Me.RdBttnStiffness.Text = "Input segment stiffness directly"
        Me.RdBttnStiffness.UseVisualStyleBackColor = True
        '
        'RdBttnYoungsModules
        '
        Me.RdBttnYoungsModules.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RdBttnYoungsModules.Location = New System.Drawing.Point(15, 41)
        Me.RdBttnYoungsModules.Name = "RdBttnYoungsModules"
        Me.RdBttnYoungsModules.Size = New System.Drawing.Size(255, 34)
        Me.RdBttnYoungsModules.TabIndex = 11
        Me.RdBttnYoungsModules.Text = "Input segment Modulus of Elasticity, calculate stiffness from segment cross secti" & _
            "on"
        Me.RdBttnYoungsModules.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        Me.RdBttnYoungsModules.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.RdBttnStiffness)
        Me.GroupBox1.Controls.Add(Me.RdBttnYoungsModules)
        Me.GroupBox1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(8, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(276, 81)
        Me.GroupBox1.TabIndex = 12
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Segment Stiffness"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtVectoAngle)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(290, 12)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(239, 81)
        Me.GroupBox2.TabIndex = 13
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Vetco Section"
        '
        'txtVectoAngle
        '
        Me.txtVectoAngle.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVectoAngle.Location = New System.Drawing.Point(127, 28)
        Me.txtVectoAngle.Name = "txtVectoAngle"
        Me.txtVectoAngle.Size = New System.Drawing.Size(54, 20)
        Me.txtVectoAngle.TabIndex = 1
        Me.txtVectoAngle.Text = "0"
        Me.txtVectoAngle.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(187, 28)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(41, 14)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "degree"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(17, 34)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(104, 14)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Vecto Section Angle"
        '
        'grdSegments
        '
        Me.grdSegments.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdSegments.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.grdSegments.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdSegments.ContextMenuStrip = Me.ContextMenuStripGridEdit
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grdSegments.DefaultCellStyle = DataGridViewCellStyle2
        Me.grdSegments.Location = New System.Drawing.Point(8, 111)
        Me.grdSegments.Name = "grdSegments"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdSegments.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.grdSegments.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.grdSegments.Size = New System.Drawing.Size(828, 293)
        Me.grdSegments.TabIndex = 14
        '
        'frmSegments
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(848, 455)
        Me.ContextMenuStrip = Me.ContextMenuStripGridEdit
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me._lblNote_0)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me._lblNote_1)
        Me.Controls.Add(Me.grdSegments)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(-9, 158)
        Me.MaximizeBox = False
        Me.Name = "frmSegments"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Riser Segment Data"
        CType(Me.lblNote, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuStripGridEdit.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.grdSegments, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents mnuAddRow As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuDeleteRow As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ContextMenuStripGridEdit As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents AddToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RdBttnStiffness As System.Windows.Forms.RadioButton
    Friend WithEvents RdBttnYoungsModules As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtVectoAngle As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents grdSegments As System.Windows.Forms.DataGridView
    Friend WithEvents pasteFromClipboard As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents copyToExcel As System.Windows.Forms.ToolStripMenuItem
#End Region
End Class