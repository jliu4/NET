<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.btnFileProject = New System.Windows.Forms.Button
        Me.FileDialogOpen = New System.Windows.Forms.OpenFileDialog
        Me.tbcMain = New System.Windows.Forms.TabControl
        Me.tbpgProject = New System.Windows.Forms.TabPage
        Me.LblProject = New System.Windows.Forms.Label
        Me.tbpgCurrents = New System.Windows.Forms.TabPage
        Me.dgvCurrents = New System.Windows.Forms.DataGridView
        Me.btnCurrentFile = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnBatch = New System.Windows.Forms.Button
        Me.btnResults = New System.Windows.Forms.Button
        Me.tbcMain.SuspendLayout()
        Me.tbpgProject.SuspendLayout()
        Me.tbpgCurrents.SuspendLayout()
        CType(Me.dgvCurrents, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnFileProject
        '
        Me.btnFileProject.Location = New System.Drawing.Point(18, 525)
        Me.btnFileProject.Name = "btnFileProject"
        Me.btnFileProject.Size = New System.Drawing.Size(131, 29)
        Me.btnFileProject.TabIndex = 2
        Me.btnFileProject.Text = "Load WinVIVA &Project"
        Me.btnFileProject.UseVisualStyleBackColor = True
        '
        'tbcMain
        '
        Me.tbcMain.Controls.Add(Me.tbpgProject)
        Me.tbcMain.Controls.Add(Me.tbpgCurrents)
        Me.tbcMain.Location = New System.Drawing.Point(18, 12)
        Me.tbcMain.Name = "tbcMain"
        Me.tbcMain.SelectedIndex = 0
        Me.tbcMain.Size = New System.Drawing.Size(757, 507)
        Me.tbcMain.TabIndex = 3
        '
        'tbpgProject
        '
        Me.tbpgProject.Controls.Add(Me.LblProject)
        Me.tbpgProject.Location = New System.Drawing.Point(4, 22)
        Me.tbpgProject.Name = "tbpgProject"
        Me.tbpgProject.Padding = New System.Windows.Forms.Padding(3)
        Me.tbpgProject.Size = New System.Drawing.Size(749, 481)
        Me.tbpgProject.TabIndex = 0
        Me.tbpgProject.Text = "Project Description"
        Me.tbpgProject.UseVisualStyleBackColor = True
        '
        'LblProject
        '
        Me.LblProject.BackColor = System.Drawing.Color.White
        Me.LblProject.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.LblProject.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LblProject.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.LblProject.Location = New System.Drawing.Point(3, 3)
        Me.LblProject.Name = "LblProject"
        Me.LblProject.Size = New System.Drawing.Size(743, 475)
        Me.LblProject.TabIndex = 4
        '
        'tbpgCurrents
        '
        Me.tbpgCurrents.Controls.Add(Me.dgvCurrents)
        Me.tbpgCurrents.Location = New System.Drawing.Point(4, 22)
        Me.tbpgCurrents.Name = "tbpgCurrents"
        Me.tbpgCurrents.Padding = New System.Windows.Forms.Padding(3)
        Me.tbpgCurrents.Size = New System.Drawing.Size(749, 481)
        Me.tbpgCurrents.TabIndex = 1
        Me.tbpgCurrents.Text = "Batch Currents (knots)"
        Me.tbpgCurrents.UseVisualStyleBackColor = True
        '
        'dgvCurrents
        '
        Me.dgvCurrents.AllowUserToAddRows = False
        Me.dgvCurrents.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvCurrents.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.dgvCurrents.Location = New System.Drawing.Point(3, 3)
        Me.dgvCurrents.Name = "dgvCurrents"
        Me.dgvCurrents.Size = New System.Drawing.Size(740, 475)
        Me.dgvCurrents.TabIndex = 0
        '
        'btnCurrentFile
        '
        Me.btnCurrentFile.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCurrentFile.Location = New System.Drawing.Point(155, 525)
        Me.btnCurrentFile.Name = "btnCurrentFile"
        Me.btnCurrentFile.Size = New System.Drawing.Size(131, 29)
        Me.btnCurrentFile.TabIndex = 2
        Me.btnCurrentFile.Text = "Load C&urrent Profiles"
        Me.btnCurrentFile.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(637, 525)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(131, 29)
        Me.btnClose.TabIndex = 2
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnBatch
        '
        Me.btnBatch.Location = New System.Drawing.Point(292, 525)
        Me.btnBatch.Name = "btnBatch"
        Me.btnBatch.Size = New System.Drawing.Size(131, 29)
        Me.btnBatch.TabIndex = 2
        Me.btnBatch.Text = "Batch &Run"
        Me.btnBatch.UseVisualStyleBackColor = True
        '
        'btnResults
        '
        Me.btnResults.Location = New System.Drawing.Point(429, 525)
        Me.btnResults.Name = "btnResults"
        Me.btnResults.Size = New System.Drawing.Size(131, 29)
        Me.btnResults.TabIndex = 2
        Me.btnResults.Text = "R&esults"
        Me.btnResults.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(792, 566)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnResults)
        Me.Controls.Add(Me.btnBatch)
        Me.Controls.Add(Me.btnCurrentFile)
        Me.Controls.Add(Me.btnFileProject)
        Me.Controls.Add(Me.tbcMain)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "WinVIVA Batch"
        Me.tbcMain.ResumeLayout(False)
        Me.tbpgProject.ResumeLayout(False)
        Me.tbpgCurrents.ResumeLayout(False)
        CType(Me.dgvCurrents, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btnFileProject As System.Windows.Forms.Button
    Friend WithEvents FileDialogOpen As System.Windows.Forms.OpenFileDialog
    Friend WithEvents tbcMain As System.Windows.Forms.TabControl
    Friend WithEvents tbpgProject As System.Windows.Forms.TabPage
    Friend WithEvents tbpgCurrents As System.Windows.Forms.TabPage
    Friend WithEvents LblProject As System.Windows.Forms.Label
    Friend WithEvents dgvCurrents As System.Windows.Forms.DataGridView
    Friend WithEvents btnCurrentFile As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnBatch As System.Windows.Forms.Button
    Friend WithEvents btnResults As System.Windows.Forms.Button

End Class
