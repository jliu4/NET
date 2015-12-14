<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmStaticSolution
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
	Public WithEvents btnEstimateTension As System.Windows.Forms.Button
	Public WithEvents btnCurrentVel As System.Windows.Forms.Button
	Public WithEvents btnCancel As System.Windows.Forms.Button
	Public WithEvents btnOK As System.Windows.Forms.Button
    Public dlgImportVIVOOpen As System.Windows.Forms.OpenFileDialog
	Public WithEvents _lblNote_3 As System.Windows.Forms.Label
	Public WithEvents _lblNote_2 As System.Windows.Forms.Label
	Public WithEvents _lblNote_1 As System.Windows.Forms.Label
	Public WithEvents _lblNote_0 As System.Windows.Forms.Label
	Public WithEvents lblNote As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents mnuOpen As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuFileClose As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuFile As System.Windows.Forms.ToolStripMenuItem
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmStaticSolution))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnEstimateTension = New System.Windows.Forms.Button
        Me.btnCurrentVel = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.dlgImportVIVOOpen = New System.Windows.Forms.OpenFileDialog
        Me._lblNote_3 = New System.Windows.Forms.Label
        Me._lblNote_2 = New System.Windows.Forms.Label
        Me._lblNote_1 = New System.Windows.Forms.Label
        Me._lblNote_0 = New System.Windows.Forms.Label
        Me.lblNote = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip
        Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuOpen = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuFileClose = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuGridEdit = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuAddRow = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuDeleteRow = New System.Windows.Forms.ToolStripMenuItem
        Me.grdStaticSolution = New System.Windows.Forms.DataGridView
        Me.ContextMenuStripGridEdit = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.pasteFromClipboard = New System.Windows.Forms.ToolStripMenuItem
        Me.copyToExcel = New System.Windows.Forms.ToolStripMenuItem
        CType(Me.lblNote, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MainMenu1.SuspendLayout()
        CType(Me.grdStaticSolution, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuStripGridEdit.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnEstimateTension
        '
        Me.btnEstimateTension.BackColor = System.Drawing.SystemColors.Control
        Me.btnEstimateTension.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnEstimateTension.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnEstimateTension.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnEstimateTension.Location = New System.Drawing.Point(135, 365)
        Me.btnEstimateTension.Name = "btnEstimateTension"
        Me.btnEstimateTension.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnEstimateTension.Size = New System.Drawing.Size(145, 25)
        Me.btnEstimateTension.TabIndex = 2
        Me.btnEstimateTension.Text = "Change in Static &Tension"
        Me.btnEstimateTension.UseVisualStyleBackColor = False
        '
        'btnCurrentVel
        '
        Me.btnCurrentVel.BackColor = System.Drawing.SystemColors.Control
        Me.btnCurrentVel.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCurrentVel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCurrentVel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCurrentVel.Location = New System.Drawing.Point(16, 365)
        Me.btnCurrentVel.Name = "btnCurrentVel"
        Me.btnCurrentVel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCurrentVel.Size = New System.Drawing.Size(113, 25)
        Me.btnCurrentVel.TabIndex = 1
        Me.btnCurrentVel.Text = "Current &Velocity"
        Me.btnCurrentVel.UseVisualStyleBackColor = False
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.SystemColors.Control
        Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCancel.Location = New System.Drawing.Point(674, 365)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCancel.Size = New System.Drawing.Size(65, 25)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'btnOK
        '
        Me.btnOK.BackColor = System.Drawing.SystemColors.Control
        Me.btnOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnOK.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnOK.Location = New System.Drawing.Point(674, 329)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnOK.Size = New System.Drawing.Size(65, 25)
        Me.btnOK.TabIndex = 3
        Me.btnOK.Text = "&Save"
        Me.btnOK.UseVisualStyleBackColor = False
        '
        '_lblNote_3
        '
        Me._lblNote_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblNote_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblNote_3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblNote_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblNote.SetIndex(Me._lblNote_3, CType(3, Short))
        Me._lblNote_3.Location = New System.Drawing.Point(60, 337)
        Me._lblNote_3.Name = "_lblNote_3"
        Me._lblNote_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblNote_3.Size = New System.Drawing.Size(369, 17)
        Me._lblNote_3.TabIndex = 9
        Me._lblNote_3.Text = "3. The riser in-plane angle is measured from horizontal."
        '
        '_lblNote_2
        '
        Me._lblNote_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblNote_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblNote_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblNote_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblNote.SetIndex(Me._lblNote_2, CType(2, Short))
        Me._lblNote_2.Location = New System.Drawing.Point(60, 320)
        Me._lblNote_2.Name = "_lblNote_2"
        Me._lblNote_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblNote_2.Size = New System.Drawing.Size(317, 17)
        Me._lblNote_2.TabIndex = 8
        Me._lblNote_2.Text = "2. The distance along the riser is measured from the bottom end."
        '
        '_lblNote_1
        '
        Me._lblNote_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblNote_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblNote_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblNote_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblNote.SetIndex(Me._lblNote_1, CType(1, Short))
        Me._lblNote_1.Location = New System.Drawing.Point(60, 303)
        Me._lblNote_1.Name = "_lblNote_1"
        Me._lblNote_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblNote_1.Size = New System.Drawing.Size(369, 17)
        Me._lblNote_1.TabIndex = 7
        Me._lblNote_1.Text = "1. Points start from the top of the riser."
        '
        '_lblNote_0
        '
        Me._lblNote_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblNote_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblNote_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblNote_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblNote.SetIndex(Me._lblNote_0, CType(0, Short))
        Me._lblNote_0.Location = New System.Drawing.Point(13, 298)
        Me._lblNote_0.Name = "_lblNote_0"
        Me._lblNote_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblNote_0.Size = New System.Drawing.Size(41, 17)
        Me._lblNote_0.TabIndex = 6
        Me._lblNote_0.Text = "* Note :"
        '
        'MainMenu1
        '
        Me.MainMenu1.BackColor = System.Drawing.SystemColors.MenuBar
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFile, Me.mnuGridEdit})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(754, 24)
        Me.MainMenu1.TabIndex = 10
        '
        'mnuFile
        '
        Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuOpen, Me.mnuFileClose})
        Me.mnuFile.Name = "mnuFile"
        Me.mnuFile.Size = New System.Drawing.Size(37, 20)
        Me.mnuFile.Text = "&File"
        '
        'mnuOpen
        '
        Me.mnuOpen.Name = "mnuOpen"
        Me.mnuOpen.Size = New System.Drawing.Size(189, 22)
        Me.mnuOpen.Text = "&Import Static Solution"
        '
        'mnuFileClose
        '
        Me.mnuFileClose.Name = "mnuFileClose"
        Me.mnuFileClose.Size = New System.Drawing.Size(189, 22)
        Me.mnuFileClose.Text = "&Close"
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
        'grdStaticSolution
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdStaticSolution.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.grdStaticSolution.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdStaticSolution.ContextMenuStrip = Me.ContextMenuStripGridEdit
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grdStaticSolution.DefaultCellStyle = DataGridViewCellStyle2
        Me.grdStaticSolution.Location = New System.Drawing.Point(12, 27)
        Me.grdStaticSolution.Name = "grdStaticSolution"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdStaticSolution.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.grdStaticSolution.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.grdStaticSolution.RowTemplate.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.grdStaticSolution.RowTemplate.DefaultCellStyle.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdStaticSolution.Size = New System.Drawing.Size(727, 261)
        Me.grdStaticSolution.TabIndex = 11
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
        'frmStaticSolution
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(754, 402)
        Me.Controls.Add(Me._lblNote_3)
        Me.Controls.Add(Me._lblNote_2)
        Me.Controls.Add(Me.btnEstimateTension)
        Me.Controls.Add(Me._lblNote_1)
        Me.Controls.Add(Me.btnCurrentVel)
        Me.Controls.Add(Me._lblNote_0)
        Me.Controls.Add(Me.MainMenu1)
        Me.Controls.Add(Me.grdStaticSolution)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(10, 48)
        Me.MaximizeBox = False
        Me.Name = "frmStaticSolution"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Data for SCR/Lazy-Wave Riser"
        CType(Me.lblNote, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        CType(Me.grdStaticSolution, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuStripGridEdit.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents grdStaticSolution As System.Windows.Forms.DataGridView
    Friend WithEvents ContextMenuStripGridEdit As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents pasteFromClipboard As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents copyToExcel As System.Windows.Forms.ToolStripMenuItem
   
#End Region 
End Class