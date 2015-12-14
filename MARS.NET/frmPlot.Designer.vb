<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPlot
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
	Public WithEvents _txtPicControls_1 As System.Windows.Forms.TextBox
	Public WithEvents btnClose As System.Windows.Forms.Button
	Public WithEvents _btnPlot_0 As System.Windows.Forms.Button
	Public WithEvents _btnPlot_1 As System.Windows.Forms.Button
	Public WithEvents _txtPicControls_0 As System.Windows.Forms.TextBox
	Public WithEvents _btnZoom_1 As System.Windows.Forms.Button
	Public WithEvents _btnZoom_0 As System.Windows.Forms.Button
	Public WithEvents hPan As System.Windows.Forms.HScrollBar
	Public WithEvents hScroll_Renamed As System.Windows.Forms.HScrollBar
	Public WithEvents vPan As System.Windows.Forms.VScrollBar
	Public WithEvents vScroll_Renamed As System.Windows.Forms.VScrollBar
	Public WithEvents ThreeD As System.Windows.Forms.PictureBox
    Public WithEvents lblVessel As System.Windows.Forms.Label
    Public WithEvents btnPlot As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
	Public WithEvents btnZoom As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
	Public WithEvents lblForceUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblVelUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents txtPicControls As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPlot))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me._btnPlot_0 = New System.Windows.Forms.Button()
        Me._btnPlot_1 = New System.Windows.Forms.Button()
        Me._txtPicControls_1 = New System.Windows.Forms.TextBox()
        Me.btnClose = New System.Windows.Forms.Button()
        Me._txtPicControls_0 = New System.Windows.Forms.TextBox()
        Me._btnZoom_1 = New System.Windows.Forms.Button()
        Me._btnZoom_0 = New System.Windows.Forms.Button()
        Me.hPan = New System.Windows.Forms.HScrollBar()
        Me.hScroll_Renamed = New System.Windows.Forms.HScrollBar()
        Me.vPan = New System.Windows.Forms.VScrollBar()
        Me.vScroll_Renamed = New System.Windows.Forms.VScrollBar()
        Me.ThreeD = New System.Windows.Forms.PictureBox()
        Me.lblVessel = New System.Windows.Forms.Label()
        Me.btnPlot = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(Me.components)
        Me.btnZoom = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(Me.components)
        Me.lblForceUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblVelUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.txtPicControls = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        CType(Me.ThreeD, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnPlot, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnZoom, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblForceUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVelUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtPicControls, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        '_btnPlot_0
        '
        Me._btnPlot_0.BackColor = System.Drawing.SystemColors.Control
        Me._btnPlot_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnPlot_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnPlot.SetIndex(Me._btnPlot_0, CType(0, Short))
        Me._btnPlot_0.Location = New System.Drawing.Point(376, 536)
        Me._btnPlot_0.Name = "_btnPlot_0"
        Me._btnPlot_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnPlot_0.Size = New System.Drawing.Size(57, 25)
        Me._btnPlot_0.TabIndex = 5
        Me._btnPlot_0.Text = "3D Plot"
        Me.ToolTip1.SetToolTip(Me._btnPlot_0, "Show 3D plot")
        Me._btnPlot_0.UseVisualStyleBackColor = False
        '
        '_btnPlot_1
        '
        Me._btnPlot_1.BackColor = System.Drawing.SystemColors.Control
        Me._btnPlot_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnPlot_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnPlot.SetIndex(Me._btnPlot_1, CType(1, Short))
        Me._btnPlot_1.Location = New System.Drawing.Point(432, 536)
        Me._btnPlot_1.Name = "_btnPlot_1"
        Me._btnPlot_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnPlot_1.Size = New System.Drawing.Size(65, 25)
        Me._btnPlot_1.TabIndex = 6
        Me._btnPlot_1.Text = "Radar Plot"
        Me.ToolTip1.SetToolTip(Me._btnPlot_1, "Show radar plot")
        Me._btnPlot_1.UseVisualStyleBackColor = False
        '
        '_txtPicControls_1
        '
        Me._txtPicControls_1.AcceptsReturn = True
        Me._txtPicControls_1.BackColor = System.Drawing.SystemColors.Window
        Me._txtPicControls_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtPicControls_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPicControls.SetIndex(Me._txtPicControls_1, CType(1, Short))
        Me._txtPicControls_1.Location = New System.Drawing.Point(272, 536)
        Me._txtPicControls_1.Name = "_txtPicControls_1"
        Me._txtPicControls_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtPicControls_1.Size = New System.Drawing.Size(33, 21)
        Me._txtPicControls_1.TabIndex = 12
        Me._txtPicControls_1.TabStop = False
        Me._txtPicControls_1.Text = "0"
        Me._txtPicControls_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'btnClose
        '
        Me.btnClose.BackColor = System.Drawing.SystemColors.Control
        Me.btnClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnClose.Location = New System.Drawing.Point(432, 8)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnClose.Size = New System.Drawing.Size(65, 25)
        Me.btnClose.TabIndex = 0
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = False
        '
        '_txtPicControls_0
        '
        Me._txtPicControls_0.AcceptsReturn = True
        Me._txtPicControls_0.BackColor = System.Drawing.SystemColors.Window
        Me._txtPicControls_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtPicControls_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPicControls.SetIndex(Me._txtPicControls_0, CType(0, Short))
        Me._txtPicControls_0.Location = New System.Drawing.Point(192, 536)
        Me._txtPicControls_0.Name = "_txtPicControls_0"
        Me._txtPicControls_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtPicControls_0.Size = New System.Drawing.Size(33, 21)
        Me._txtPicControls_0.TabIndex = 10
        Me._txtPicControls_0.TabStop = False
        Me._txtPicControls_0.Text = "0"
        Me._txtPicControls_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_btnZoom_1
        '
        Me._btnZoom_1.BackColor = System.Drawing.SystemColors.Control
        Me._btnZoom_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnZoom_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnZoom.SetIndex(Me._btnZoom_1, CType(1, Short))
        Me._btnZoom_1.Location = New System.Drawing.Point(64, 536)
        Me._btnZoom_1.Name = "_btnZoom_1"
        Me._btnZoom_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnZoom_1.Size = New System.Drawing.Size(73, 25)
        Me._btnZoom_1.TabIndex = 8
        Me._btnZoom_1.Text = "Zoom Out"
        Me._btnZoom_1.UseVisualStyleBackColor = False
        '
        '_btnZoom_0
        '
        Me._btnZoom_0.BackColor = System.Drawing.SystemColors.Control
        Me._btnZoom_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnZoom_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnZoom.SetIndex(Me._btnZoom_0, CType(0, Short))
        Me._btnZoom_0.Location = New System.Drawing.Point(8, 536)
        Me._btnZoom_0.Name = "_btnZoom_0"
        Me._btnZoom_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnZoom_0.Size = New System.Drawing.Size(57, 25)
        Me._btnZoom_0.TabIndex = 7
        Me._btnZoom_0.Text = "Zoom In"
        Me._btnZoom_0.UseVisualStyleBackColor = False
        '
        'hPan
        '
        Me.hPan.Cursor = System.Windows.Forms.Cursors.Default
        Me.hPan.Location = New System.Drawing.Point(8, 512)
        Me.hPan.Maximum = 109
        Me.hPan.Name = "hPan"
        Me.hPan.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.hPan.Size = New System.Drawing.Size(473, 17)
        Me.hPan.TabIndex = 4
        '
        'hScroll_Renamed
        '
        Me.hScroll_Renamed.Cursor = System.Windows.Forms.Cursors.Default
        Me.hScroll_Renamed.Location = New System.Drawing.Point(168, 534)
        Me.hScroll_Renamed.Maximum = 189
        Me.hScroll_Renamed.Minimum = -180
        Me.hScroll_Renamed.Name = "hScroll_Renamed"
        Me.hScroll_Renamed.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.hScroll_Renamed.Size = New System.Drawing.Size(25, 25)
        Me.hScroll_Renamed.TabIndex = 9
        '
        'vPan
        '
        Me.vPan.Cursor = System.Windows.Forms.Cursors.Default
        Me.vPan.Location = New System.Drawing.Point(480, 40)
        Me.vPan.Maximum = 109
        Me.vPan.Name = "vPan"
        Me.vPan.Size = New System.Drawing.Size(17, 472)
        Me.vPan.TabIndex = 3
        '
        'vScroll_Renamed
        '
        Me.vScroll_Renamed.Cursor = System.Windows.Forms.Cursors.Default
        Me.vScroll_Renamed.Location = New System.Drawing.Point(248, 534)
        Me.vScroll_Renamed.Maximum = 99
        Me.vScroll_Renamed.Name = "vScroll_Renamed"
        Me.vScroll_Renamed.Size = New System.Drawing.Size(25, 25)
        Me.vScroll_Renamed.TabIndex = 11
        Me.vScroll_Renamed.Value = 90
        '
        'ThreeD
        '
        Me.ThreeD.BackColor = System.Drawing.SystemColors.Control
        Me.ThreeD.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.ThreeD.Cursor = System.Windows.Forms.Cursors.Default
        Me.ThreeD.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ThreeD.Location = New System.Drawing.Point(8, 40)
        Me.ThreeD.Name = "ThreeD"
        Me.ThreeD.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ThreeD.Size = New System.Drawing.Size(473, 473)
        Me.ThreeD.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.ThreeD.TabIndex = 2
        Me.ThreeD.TabStop = False
        '
        'lblVessel
        '
        Me.lblVessel.BackColor = System.Drawing.SystemColors.Control
        Me.lblVessel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblVessel.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblVessel.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVessel.Location = New System.Drawing.Point(8, 8)
        Me.lblVessel.Name = "lblVessel"
        Me.lblVessel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblVessel.Size = New System.Drawing.Size(209, 25)
        Me.lblVessel.TabIndex = 1
        Me.lblVessel.Text = "GlobalSantaFe"
        Me.lblVessel.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btnPlot
        '
        '
        'btnZoom
        '
        '
        'txtPicControls
        '
        '
        'frmPlot
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(509, 571)
        Me.Controls.Add(Me._txtPicControls_1)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me._btnPlot_0)
        Me.Controls.Add(Me._btnPlot_1)
        Me.Controls.Add(Me._txtPicControls_0)
        Me.Controls.Add(Me._btnZoom_1)
        Me.Controls.Add(Me._btnZoom_0)
        Me.Controls.Add(Me.hPan)
        Me.Controls.Add(Me.hScroll_Renamed)
        Me.Controls.Add(Me.vPan)
        Me.Controls.Add(Me.vScroll_Renamed)
        Me.Controls.Add(Me.ThreeD)
        Me.Controls.Add(Me.lblVessel)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(128, 48)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmPlot"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = " MARS - Plot"
        CType(Me.ThreeD, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnPlot, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnZoom, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblForceUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVelUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtPicControls, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
End Class