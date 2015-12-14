<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmComputation
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
	Public WithEvents _sbrStatus_Panel1 As System.Windows.Forms.ToolStripStatusLabel
	Public WithEvents sbrStatus As System.Windows.Forms.StatusStrip
	Public WithEvents txtIterations As System.Windows.Forms.TextBox
	Public WithEvents txtNumPoints As System.Windows.Forms.TextBox
	Public WithEvents lblIterations As System.Windows.Forms.Label
	Public WithEvents lblNumPoints As System.Windows.Forms.Label
	Public WithEvents fraParameters As System.Windows.Forms.GroupBox
	Public WithEvents optDOSBoxVisible As System.Windows.Forms.RadioButton
	Public WithEvents optDOSBoxHidden As System.Windows.Forms.RadioButton
	Public WithEvents fraShowDOSBox As System.Windows.Forms.GroupBox
	Public WithEvents btnCancel As System.Windows.Forms.Button
	Public WithEvents btnWriteFiles As System.Windows.Forms.Button
	Public WithEvents btnExecute As System.Windows.Forms.Button
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmComputation))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.sbrStatus = New System.Windows.Forms.StatusStrip
        Me._sbrStatus_Panel1 = New System.Windows.Forms.ToolStripStatusLabel
        Me.fraParameters = New System.Windows.Forms.GroupBox
        Me.txtIterations = New System.Windows.Forms.TextBox
        Me.txtNumPoints = New System.Windows.Forms.TextBox
        Me.lblIterations = New System.Windows.Forms.Label
        Me.lblNumPoints = New System.Windows.Forms.Label
        Me.fraShowDOSBox = New System.Windows.Forms.GroupBox
        Me.optDOSBoxVisible = New System.Windows.Forms.RadioButton
        Me.optDOSBoxHidden = New System.Windows.Forms.RadioButton
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnWriteFiles = New System.Windows.Forms.Button
        Me.btnExecute = New System.Windows.Forms.Button
        Me.sbrStatus.SuspendLayout()
        Me.fraParameters.SuspendLayout()
        Me.fraShowDOSBox.SuspendLayout()
        Me.SuspendLayout()
        '
        'sbrStatus
        '
        Me.sbrStatus.Enabled = False
        Me.sbrStatus.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.sbrStatus.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me._sbrStatus_Panel1})
        Me.sbrStatus.Location = New System.Drawing.Point(0, 139)
        Me.sbrStatus.Name = "sbrStatus"
        Me.sbrStatus.Size = New System.Drawing.Size(433, 22)
        Me.sbrStatus.TabIndex = 11
        Me.sbrStatus.Text = "Standby..."
        '
        '_sbrStatus_Panel1
        '
        Me._sbrStatus_Panel1.AutoSize = False
        Me._sbrStatus_Panel1.BorderSides = CType((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) _
                    Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) _
                    Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me._sbrStatus_Panel1.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter
        Me._sbrStatus_Panel1.Margin = New System.Windows.Forms.Padding(0)
        Me._sbrStatus_Panel1.Name = "_sbrStatus_Panel1"
        Me._sbrStatus_Panel1.Size = New System.Drawing.Size(96, 22)
        Me._sbrStatus_Panel1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'fraParameters
        '
        Me.fraParameters.BackColor = System.Drawing.SystemColors.Control
        Me.fraParameters.Controls.Add(Me.txtIterations)
        Me.fraParameters.Controls.Add(Me.txtNumPoints)
        Me.fraParameters.Controls.Add(Me.lblIterations)
        Me.fraParameters.Controls.Add(Me.lblNumPoints)
        Me.fraParameters.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraParameters.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraParameters.Location = New System.Drawing.Point(8, 8)
        Me.fraParameters.Name = "fraParameters"
        Me.fraParameters.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraParameters.Size = New System.Drawing.Size(300, 81)
        Me.fraParameters.TabIndex = 8
        Me.fraParameters.TabStop = False
        Me.fraParameters.Text = "Computation Parameters"
        '
        'txtIterations
        '
        Me.txtIterations.AcceptsReturn = True
        Me.txtIterations.BackColor = System.Drawing.SystemColors.Window
        Me.txtIterations.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtIterations.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtIterations.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtIterations.Location = New System.Drawing.Point(236, 23)
        Me.txtIterations.MaxLength = 0
        Me.txtIterations.Name = "txtIterations"
        Me.txtIterations.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtIterations.Size = New System.Drawing.Size(49, 20)
        Me.txtIterations.TabIndex = 3
        Me.txtIterations.Text = "Text1"
        Me.txtIterations.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtNumPoints
        '
        Me.txtNumPoints.AcceptsReturn = True
        Me.txtNumPoints.BackColor = System.Drawing.SystemColors.Window
        Me.txtNumPoints.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNumPoints.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNumPoints.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNumPoints.Location = New System.Drawing.Point(236, 47)
        Me.txtNumPoints.MaxLength = 0
        Me.txtNumPoints.Name = "txtNumPoints"
        Me.txtNumPoints.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNumPoints.Size = New System.Drawing.Size(49, 20)
        Me.txtNumPoints.TabIndex = 4
        Me.txtNumPoints.Text = "Text1"
        Me.txtNumPoints.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblIterations
        '
        Me.lblIterations.BackColor = System.Drawing.SystemColors.Control
        Me.lblIterations.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblIterations.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblIterations.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblIterations.Location = New System.Drawing.Point(16, 24)
        Me.lblIterations.Name = "lblIterations"
        Me.lblIterations.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblIterations.Size = New System.Drawing.Size(214, 17)
        Me.lblIterations.TabIndex = 10
        Me.lblIterations.Text = "Number of Iterations (>= 30) : "
        '
        'lblNumPoints
        '
        Me.lblNumPoints.BackColor = System.Drawing.SystemColors.Control
        Me.lblNumPoints.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblNumPoints.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNumPoints.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblNumPoints.Location = New System.Drawing.Point(16, 48)
        Me.lblNumPoints.Name = "lblNumPoints"
        Me.lblNumPoints.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblNumPoints.Size = New System.Drawing.Size(214, 16)
        Me.lblNumPoints.TabIndex = 9
        Me.lblNumPoints.Text = "Number of Points on Riser (200--6001) : "
        '
        'fraShowDOSBox
        '
        Me.fraShowDOSBox.BackColor = System.Drawing.SystemColors.Control
        Me.fraShowDOSBox.Controls.Add(Me.optDOSBoxVisible)
        Me.fraShowDOSBox.Controls.Add(Me.optDOSBoxHidden)
        Me.fraShowDOSBox.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraShowDOSBox.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraShowDOSBox.Location = New System.Drawing.Point(314, 8)
        Me.fraShowDOSBox.Name = "fraShowDOSBox"
        Me.fraShowDOSBox.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraShowDOSBox.Size = New System.Drawing.Size(105, 81)
        Me.fraShowDOSBox.TabIndex = 7
        Me.fraShowDOSBox.TabStop = False
        Me.fraShowDOSBox.Text = "DOS Windows"
        '
        'optDOSBoxVisible
        '
        Me.optDOSBoxVisible.BackColor = System.Drawing.SystemColors.Control
        Me.optDOSBoxVisible.Cursor = System.Windows.Forms.Cursors.Default
        Me.optDOSBoxVisible.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optDOSBoxVisible.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optDOSBoxVisible.Location = New System.Drawing.Point(16, 44)
        Me.optDOSBoxVisible.Name = "optDOSBoxVisible"
        Me.optDOSBoxVisible.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optDOSBoxVisible.Size = New System.Drawing.Size(60, 25)
        Me.optDOSBoxVisible.TabIndex = 6
        Me.optDOSBoxVisible.TabStop = True
        Me.optDOSBoxVisible.Text = "&Visible"
        Me.optDOSBoxVisible.UseVisualStyleBackColor = False
        '
        'optDOSBoxHidden
        '
        Me.optDOSBoxHidden.BackColor = System.Drawing.SystemColors.Control
        Me.optDOSBoxHidden.Checked = True
        Me.optDOSBoxHidden.Cursor = System.Windows.Forms.Cursors.Default
        Me.optDOSBoxHidden.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optDOSBoxHidden.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optDOSBoxHidden.Location = New System.Drawing.Point(16, 24)
        Me.optDOSBoxHidden.Name = "optDOSBoxHidden"
        Me.optDOSBoxHidden.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optDOSBoxHidden.Size = New System.Drawing.Size(60, 17)
        Me.optDOSBoxHidden.TabIndex = 5
        Me.optDOSBoxHidden.TabStop = True
        Me.optDOSBoxHidden.Text = "&Hidden"
        Me.optDOSBoxHidden.UseVisualStyleBackColor = False
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.SystemColors.Control
        Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCancel.Location = New System.Drawing.Point(354, 104)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCancel.Size = New System.Drawing.Size(65, 25)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "Return"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'btnWriteFiles
        '
        Me.btnWriteFiles.BackColor = System.Drawing.SystemColors.Control
        Me.btnWriteFiles.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnWriteFiles.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnWriteFiles.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnWriteFiles.Location = New System.Drawing.Point(104, 104)
        Me.btnWriteFiles.Name = "btnWriteFiles"
        Me.btnWriteFiles.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnWriteFiles.Size = New System.Drawing.Size(91, 25)
        Me.btnWriteFiles.TabIndex = 1
        Me.btnWriteFiles.Text = "Generate &Input"
        Me.btnWriteFiles.UseVisualStyleBackColor = False
        '
        'btnExecute
        '
        Me.btnExecute.BackColor = System.Drawing.SystemColors.Control
        Me.btnExecute.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnExecute.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExecute.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnExecute.Location = New System.Drawing.Point(8, 104)
        Me.btnExecute.Name = "btnExecute"
        Me.btnExecute.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnExecute.Size = New System.Drawing.Size(91, 25)
        Me.btnExecute.TabIndex = 0
        Me.btnExecute.Text = "&Run"
        Me.btnExecute.UseVisualStyleBackColor = False
        '
        'frmComputation
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(433, 161)
        Me.Controls.Add(Me.sbrStatus)
        Me.Controls.Add(Me.fraParameters)
        Me.Controls.Add(Me.fraShowDOSBox)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnWriteFiles)
        Me.Controls.Add(Me.btnExecute)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.Name = "frmComputation"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Computation Parameters and Execution"
        Me.sbrStatus.ResumeLayout(False)
        Me.sbrStatus.PerformLayout()
        Me.fraParameters.ResumeLayout(False)
        Me.fraParameters.PerformLayout()
        Me.fraShowDOSBox.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region 
End Class