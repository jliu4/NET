<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPlotResults
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
	Public WithEvents btnPrint As System.Windows.Forms.Button
	Public WithEvents btnCancel As System.Windows.Forms.Button
	Public WithEvents cmbMode As System.Windows.Forms.ComboBox
    'Public WithEvents tabPlot As AxMSComctlLib.AxTabStrip
    'Public WithEvents tabPlot As AxMSComctlLib.AxTabStrip
    Public WithEvents lblMode As System.Windows.Forms.Label
    Public WithEvents mnuPlotPrint As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuExit As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuPlot As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPlotResults))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnPrint = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.cmbMode = New System.Windows.Forms.ComboBox
        Me.tabPlot = New AxMSComctlLib.AxTabStrip
        Me.lblMode = New System.Windows.Forms.Label
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip
        Me.mnuPlot = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuPlotCopy = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuPlotPrint = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuExit = New System.Windows.Forms.ToolStripMenuItem
        Me.picResults = New System.Windows.Forms.PictureBox
        Me.dlgPrint = New System.Windows.Forms.PrintDialog
        Me.PrintDocResults = New System.Drawing.Printing.PrintDocument
        CType(Me.tabPlot, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MainMenu1.SuspendLayout()
        CType(Me.picResults, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnPrint
        '
        Me.btnPrint.BackColor = System.Drawing.SystemColors.Control
        Me.btnPrint.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnPrint.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrint.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnPrint.Location = New System.Drawing.Point(569, 502)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnPrint.Size = New System.Drawing.Size(65, 25)
        Me.btnPrint.TabIndex = 2
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.UseVisualStyleBackColor = False
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.SystemColors.Control
        Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCancel.Location = New System.Drawing.Point(641, 502)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCancel.Size = New System.Drawing.Size(65, 25)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Return"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'cmbMode
        '
        Me.cmbMode.BackColor = System.Drawing.SystemColors.Window
        Me.cmbMode.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbMode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbMode.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbMode.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbMode.Location = New System.Drawing.Point(86, 503)
        Me.cmbMode.Name = "cmbMode"
        Me.cmbMode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbMode.Size = New System.Drawing.Size(65, 22)
        Me.cmbMode.TabIndex = 1
        '
        'tabPlot
        '
        Me.tabPlot.Location = New System.Drawing.Point(9, 27)
        Me.tabPlot.Name = "tabPlot"
        Me.tabPlot.OcxState = CType(resources.GetObject("tabPlot.OcxState"), System.Windows.Forms.AxHost.State)
        Me.tabPlot.Size = New System.Drawing.Size(697, 469)
        Me.tabPlot.TabIndex = 0
        '
        'lblMode
        '
        Me.lblMode.BackColor = System.Drawing.SystemColors.Control
        Me.lblMode.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMode.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMode.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblMode.Location = New System.Drawing.Point(6, 506)
        Me.lblMode.Name = "lblMode"
        Me.lblMode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMode.Size = New System.Drawing.Size(74, 17)
        Me.lblMode.TabIndex = 4
        Me.lblMode.Text = "Mode Number"
        '
        'MainMenu1
        '
        Me.MainMenu1.BackColor = System.Drawing.SystemColors.MenuBar
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuPlot})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(715, 24)
        Me.MainMenu1.TabIndex = 6
        '
        'mnuPlot
        '
        Me.mnuPlot.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuPlotCopy, Me.mnuPlotPrint, Me.mnuExit})
        Me.mnuPlot.Name = "mnuPlot"
        Me.mnuPlot.Size = New System.Drawing.Size(37, 20)
        Me.mnuPlot.Text = "P&lot"
        '
        'mnuPlotCopy
        '
        Me.mnuPlotCopy.Name = "mnuPlotCopy"
        Me.mnuPlotCopy.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.C), System.Windows.Forms.Keys)
        Me.mnuPlotCopy.Size = New System.Drawing.Size(160, 22)
        Me.mnuPlotCopy.Text = "&Copy"
        '
        'mnuPlotPrint
        '
        Me.mnuPlotPrint.Name = "mnuPlotPrint"
        Me.mnuPlotPrint.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.P), System.Windows.Forms.Keys)
        Me.mnuPlotPrint.Size = New System.Drawing.Size(160, 22)
        Me.mnuPlotPrint.Text = "&Print Plots"
        '
        'mnuExit
        '
        Me.mnuExit.Name = "mnuExit"
        Me.mnuExit.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.X), System.Windows.Forms.Keys)
        Me.mnuExit.Size = New System.Drawing.Size(160, 22)
        Me.mnuExit.Text = "Clos&e"
        '
        'picResults
        '
        Me.picResults.BackColor = System.Drawing.SystemColors.Window
        Me.picResults.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.picResults.Location = New System.Drawing.Point(31, 63)
        Me.picResults.Name = "picResults"
        Me.picResults.Size = New System.Drawing.Size(650, 403)
        Me.picResults.TabIndex = 7
        Me.picResults.TabStop = False
        '
        'dlgPrint
        '
        Me.dlgPrint.UseEXDialog = True
        '
        'PrintDocResults
        '
        '
        'frmPlotResults
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(715, 540)
        Me.Controls.Add(Me.picResults)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.cmbMode)
        Me.Controls.Add(Me.tabPlot)
        Me.Controls.Add(Me.lblMode)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(10, 48)
        Me.MaximizeBox = False
        Me.Name = "frmPlotResults"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Plot Results"
        CType(Me.tabPlot, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        CType(Me.picResults, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents picResults As System.Windows.Forms.PictureBox
    Friend WithEvents mnuPlotCopy As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents dlgPrint As System.Windows.Forms.PrintDialog
    Friend WithEvents PrintDocResults As System.Drawing.Printing.PrintDocument
#End Region
End Class