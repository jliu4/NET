<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmCurrent
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
    Public WithEvents fraProfile As System.Windows.Forms.GroupBox
    Public WithEvents btnCancel As System.Windows.Forms.Button
	Public WithEvents btnOK As System.Windows.Forms.Button
	Public dlgCurrentProfileOpen As System.Windows.Forms.OpenFileDialog
	Public dlgCurrentProfileSave As System.Windows.Forms.SaveFileDialog
	Public WithEvents mnuFileOpen As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuFileSaveAs As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuLine As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuExit As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuFile As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuAddRow As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuDeleteRow As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.fraProfile = New System.Windows.Forms.GroupBox()
        Me.grdCurrent = New System.Windows.Forms.DataGridView()
        Me.ContextMenuStripGridEdit = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.mnuAddRow = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuDeleteRow = New System.Windows.Forms.ToolStripMenuItem()
        Me.pasteFromClipBoard = New System.Windows.Forms.ToolStripMenuItem()
        Me.copyToExcel = New System.Windows.Forms.ToolStripMenuItem()
        Me.picCurrent = New System.Windows.Forms.PictureBox()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.dlgCurrentProfileOpen = New System.Windows.Forms.OpenFileDialog()
        Me.dlgCurrentProfileSave = New System.Windows.Forms.SaveFileDialog()
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip()
        Me.mnuGridEdit = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFileOpen = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFileSaveAs = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLine = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuExit = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.PrintDialogCurrent = New System.Windows.Forms.PrintDialog()
        Me.PrintDocumentCurrent = New System.Drawing.Printing.PrintDocument()
        Me.fraProfile.SuspendLayout()
        CType(Me.grdCurrent, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuStripGridEdit.SuspendLayout()
        CType(Me.picCurrent, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MainMenu1.SuspendLayout()
        Me.SuspendLayout()
        '
        'fraProfile
        '
        Me.fraProfile.BackColor = System.Drawing.SystemColors.Control
        Me.fraProfile.Controls.Add(Me.grdCurrent)
        Me.fraProfile.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraProfile.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraProfile.Location = New System.Drawing.Point(7, 27)
        Me.fraProfile.Name = "fraProfile"
        Me.fraProfile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraProfile.Size = New System.Drawing.Size(265, 385)
        Me.fraProfile.TabIndex = 5
        Me.fraProfile.TabStop = False
        Me.fraProfile.Text = "Current Profile"
        '
        'grdCurrent
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdCurrent.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.grdCurrent.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdCurrent.ContextMenuStrip = Me.ContextMenuStripGridEdit
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.Format = "N2"
        DataGridViewCellStyle2.NullValue = Nothing
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grdCurrent.DefaultCellStyle = DataGridViewCellStyle2
        Me.grdCurrent.Location = New System.Drawing.Point(6, 32)
        Me.grdCurrent.Name = "grdCurrent"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdCurrent.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.grdCurrent.RowHeadersVisible = False
        Me.grdCurrent.Size = New System.Drawing.Size(239, 347)
        Me.grdCurrent.TabIndex = 7
        '
        'ContextMenuStripGridEdit
        '
        Me.ContextMenuStripGridEdit.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuAddRow, Me.mnuDeleteRow, Me.pasteFromClipBoard, Me.copyToExcel})
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
        'copyToExcel
        '
        Me.copyToExcel.Name = "copyToExcel"
        Me.copyToExcel.Size = New System.Drawing.Size(133, 22)
        Me.copyToExcel.Text = "Copy"
        '
        'picCurrent
        '
        Me.picCurrent.BackColor = System.Drawing.Color.White
        Me.picCurrent.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.picCurrent.Cursor = System.Windows.Forms.Cursors.Default
        Me.picCurrent.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.picCurrent.ForeColor = System.Drawing.SystemColors.ControlText
        Me.picCurrent.Location = New System.Drawing.Point(278, 27)
        Me.picCurrent.Name = "picCurrent"
        Me.picCurrent.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.picCurrent.Size = New System.Drawing.Size(425, 417)
        Me.picCurrent.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.picCurrent.TabIndex = 4
        Me.picCurrent.TabStop = False
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.SystemColors.Control
        Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCancel.Location = New System.Drawing.Point(93, 430)
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
        Me.btnOK.Location = New System.Drawing.Point(13, 429)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnOK.Size = New System.Drawing.Size(65, 25)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "&Save"
        Me.btnOK.UseVisualStyleBackColor = False
        '
        'MainMenu1
        '
        Me.MainMenu1.BackColor = System.Drawing.SystemColors.MenuBar
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuGridEdit, Me.mnuFile})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(715, 24)
        Me.MainMenu1.TabIndex = 6
        '
        'mnuGridEdit
        '
        Me.mnuGridEdit.Name = "mnuGridEdit"
        Me.mnuGridEdit.Size = New System.Drawing.Size(12, 20)
        Me.mnuGridEdit.Visible = False
        '
        'mnuFile
        '
        Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFileOpen, Me.mnuFileSaveAs, Me.mnuLine, Me.mnuExit})
        Me.mnuFile.Name = "mnuFile"
        Me.mnuFile.Size = New System.Drawing.Size(37, 20)
        Me.mnuFile.Text = "&File"
        '
        'mnuFileOpen
        '
        Me.mnuFileOpen.Name = "mnuFileOpen"
        Me.mnuFileOpen.Size = New System.Drawing.Size(119, 22)
        Me.mnuFileOpen.Text = "&Import..."
        '
        'mnuFileSaveAs
        '
        Me.mnuFileSaveAs.Name = "mnuFileSaveAs"
        Me.mnuFileSaveAs.Size = New System.Drawing.Size(119, 22)
        Me.mnuFileSaveAs.Text = "&Export..."
        '
        'mnuLine
        '
        Me.mnuLine.Name = "mnuLine"
        Me.mnuLine.Size = New System.Drawing.Size(116, 6)
        '
        'mnuExit
        '
        Me.mnuExit.Name = "mnuExit"
        Me.mnuExit.Size = New System.Drawing.Size(119, 22)
        Me.mnuExit.Text = "&Close"
        '
        'ToolStripMenuItem2
        '
        Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
        Me.ToolStripMenuItem2.Size = New System.Drawing.Size(80, 22)
        Me.ToolStripMenuItem2.Text = "1"
        '
        'PrintDialogCurrent
        '
        Me.PrintDialogCurrent.UseEXDialog = True
        '
        'frmCurrent
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CausesValidation = False
        Me.ClientSize = New System.Drawing.Size(715, 467)
        Me.ContextMenuStrip = Me.ContextMenuStripGridEdit
        Me.Controls.Add(Me.fraProfile)
        Me.Controls.Add(Me.picCurrent)
        Me.Controls.Add(Me.MainMenu1)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(10, 48)
        Me.MaximizeBox = False
        Me.Name = "frmCurrent"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Current Profile"
        Me.fraProfile.ResumeLayout(False)
        CType(Me.grdCurrent, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuStripGridEdit.ResumeLayout(False)
        CType(Me.picCurrent, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ContextMenuStripGridEdit As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ToolStripMenuItem2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuGridEdit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PrintDialogCurrent As System.Windows.Forms.PrintDialog
    Friend WithEvents PrintDocumentCurrent As System.Drawing.Printing.PrintDocument
    Friend WithEvents picCurrent As System.Windows.Forms.PictureBox
    Friend WithEvents pasteFromClipBoard As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents copyToExcel As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents grdCurrent As System.Windows.Forms.DataGridView
#End Region
End Class