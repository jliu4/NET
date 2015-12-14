<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmMainMenu
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
    Public WithEvents refreshMainMenu As System.Windows.Forms.PictureBox
    Public FileDialogOpen As System.Windows.Forms.OpenFileDialog
    Public FileDialogSave As System.Windows.Forms.SaveFileDialog
    Public FileDialogPrint As System.Windows.Forms.PrintDialog
    Public WithEvents _lblCopyright_0 As System.Windows.Forms.Label
    Public WithEvents _lblCopyright_1 As System.Windows.Forms.Label
    Public WithEvents lblCopyright As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents mnuFilePre As Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuManual As Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuFileNew As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFileOpen As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFileSave As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFileSaveAs As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFilePD As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFilePrintSetup As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuLineFilePre As System.Windows.Forms.ToolStripSeparator
    Public WithEvents _mnuFilePre_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuFilePre_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuFilePre_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuFilePre_3 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuExitLine As System.Windows.Forms.ToolStripSeparator
    Public WithEvents mnuFileExit As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFile As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuAPGlobal As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuAPCurrentProfiles As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuAP As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuRiserFRiserTypeandBC As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuRiserFAuxline As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuRiserFSupSpring As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuRiserFStaticSolution As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuRiserF As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuManual_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuManual_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuManualLine As System.Windows.Forms.ToolStripSeparator
    Public WithEvents mnuHelpAbout As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuHelp As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMainMenu))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.refreshMainMenu = New System.Windows.Forms.PictureBox
        Me.FileDialogOpen = New System.Windows.Forms.OpenFileDialog
        Me.FileDialogSave = New System.Windows.Forms.SaveFileDialog
        Me.FileDialogPrint = New System.Windows.Forms.PrintDialog
        Me._lblCopyright_0 = New System.Windows.Forms.Label
        Me._lblCopyright_1 = New System.Windows.Forms.Label
        Me.lblCopyright = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.mnuFilePre = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me._mnuFilePre_0 = New System.Windows.Forms.ToolStripMenuItem
        Me._mnuFilePre_1 = New System.Windows.Forms.ToolStripMenuItem
        Me._mnuFilePre_2 = New System.Windows.Forms.ToolStripMenuItem
        Me._mnuFilePre_3 = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuManual = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me._mnuManual_0 = New System.Windows.Forms.ToolStripMenuItem
        Me._mnuManual_1 = New System.Windows.Forms.ToolStripMenuItem
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip
        Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuFileNew = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuFileOpen = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuFileSave = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuFileSaveAs = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuFilePD = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuFilePrintSetup = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuLineFilePre = New System.Windows.Forms.ToolStripSeparator
        Me.mnuExitLine = New System.Windows.Forms.ToolStripSeparator
        Me.mnuFileExit = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuAP = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuAPGlobal = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuAPCurrentProfiles = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuRefreshMainMenu = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuRiserF = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuRiserFRiserTypeandBC = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuRiserFSegment = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuRiserFAuxline = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuRiserFSupSpring = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuRiserFStaticSolution = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuRiserFOtherProperties = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuRiserR = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuRiserRRiserTypeandBC = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuRiserRSegment = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuRiserRAuxLine = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuRiserRSupSpring = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuRiserRStaticSolution = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuRiserROtherProperties = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuComputation = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuAPComputation = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuResultsF = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuResultsR = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuBatch = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuTools = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuDatabase = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuManualLine = New System.Windows.Forms.ToolStripSeparator
        Me.mnuHelpAbout = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuReference = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuVIVOfFlexibleStructures = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuFatigueCalculation = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuModelingStrakes = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuOMAE1 = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuOMAE2 = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuQuestionAnswers = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuVortexInducedVibrationsOfLongCyclindricalStructures = New System.Windows.Forms.ToolStripMenuItem
        CType(Me.refreshMainMenu, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblCopyright, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuFilePre, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuManual, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MainMenu1.SuspendLayout()
        Me.SuspendLayout()
        '
        'refreshMainMenu
        '
        Me.refreshMainMenu.BackColor = System.Drawing.SystemColors.Control
        Me.refreshMainMenu.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.refreshMainMenu.Cursor = System.Windows.Forms.Cursors.Default
        Me.refreshMainMenu.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.refreshMainMenu.ForeColor = System.Drawing.SystemColors.ControlText
        Me.refreshMainMenu.Image = CType(resources.GetObject("refreshMainMenu.Image"), System.Drawing.Image)
        Me.refreshMainMenu.Location = New System.Drawing.Point(14, 27)
        Me.refreshMainMenu.Name = "refreshMainMenu"
        Me.refreshMainMenu.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.refreshMainMenu.Size = New System.Drawing.Size(743, 397)
        Me.refreshMainMenu.TabIndex = 1
        Me.refreshMainMenu.TabStop = False
        '
        '_lblCopyright_0
        '
        Me._lblCopyright_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblCopyright_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblCopyright_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblCopyright_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCopyright.SetIndex(Me._lblCopyright_0, CType(0, Short))
        Me._lblCopyright_0.Location = New System.Drawing.Point(11, 437)
        Me._lblCopyright_0.Name = "_lblCopyright_0"
        Me._lblCopyright_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblCopyright_0.Size = New System.Drawing.Size(385, 17)
        Me._lblCopyright_0.TabIndex = 2
        Me._lblCopyright_0.Text = "WinVIVA: (c) Genesis Engineering LLC, 1998 - 2015."
        '
        '_lblCopyright_1
        '
        Me._lblCopyright_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblCopyright_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblCopyright_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblCopyright_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCopyright.SetIndex(Me._lblCopyright_1, CType(1, Short))
        Me._lblCopyright_1.Location = New System.Drawing.Point(11, 454)
        Me._lblCopyright_1.Name = "_lblCopyright_1"
        Me._lblCopyright_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblCopyright_1.Size = New System.Drawing.Size(740, 50)
        Me._lblCopyright_1.TabIndex = 0
        Me._lblCopyright_1.Text = resources.GetString("_lblCopyright_1.Text")
        '
        'mnuFilePre
        '
        '
        '_mnuFilePre_0
        '
        Me._mnuFilePre_0.Enabled = False
        Me.mnuFilePre.SetIndex(Me._mnuFilePre_0, CType(0, Short))
        Me._mnuFilePre_0.Name = "_mnuFilePre_0"
        Me._mnuFilePre_0.Size = New System.Drawing.Size(174, 22)
        Me._mnuFilePre_0.Visible = False
        '
        '_mnuFilePre_1
        '
        Me._mnuFilePre_1.Enabled = False
        Me.mnuFilePre.SetIndex(Me._mnuFilePre_1, CType(1, Short))
        Me._mnuFilePre_1.Name = "_mnuFilePre_1"
        Me._mnuFilePre_1.Size = New System.Drawing.Size(174, 22)
        Me._mnuFilePre_1.Visible = False
        '
        '_mnuFilePre_2
        '
        Me._mnuFilePre_2.Enabled = False
        Me.mnuFilePre.SetIndex(Me._mnuFilePre_2, CType(2, Short))
        Me._mnuFilePre_2.Name = "_mnuFilePre_2"
        Me._mnuFilePre_2.Size = New System.Drawing.Size(174, 22)
        Me._mnuFilePre_2.Visible = False
        '
        '_mnuFilePre_3
        '
        Me._mnuFilePre_3.Enabled = False
        Me.mnuFilePre.SetIndex(Me._mnuFilePre_3, CType(3, Short))
        Me._mnuFilePre_3.Name = "_mnuFilePre_3"
        Me._mnuFilePre_3.Size = New System.Drawing.Size(174, 22)
        Me._mnuFilePre_3.Visible = False
        '
        'mnuManual
        '
        '
        '_mnuManual_0
        '
        Me.mnuManual.SetIndex(Me._mnuManual_0, CType(0, Short))
        Me._mnuManual_0.Name = "_mnuManual_0"
        Me._mnuManual_0.Size = New System.Drawing.Size(168, 22)
        Me._mnuManual_0.Text = "&WinVIVA Manual"
        '
        '_mnuManual_1
        '
        Me.mnuManual.SetIndex(Me._mnuManual_1, CType(1, Short))
        Me._mnuManual_1.Name = "_mnuManual_1"
        Me._mnuManual_1.Size = New System.Drawing.Size(168, 22)
        Me._mnuManual_1.Text = "&DOS VIVA Manual"
        '
        'MainMenu1
        '
        Me.MainMenu1.BackColor = System.Drawing.SystemColors.MenuBar
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFile, Me.mnuAP, Me.mnuRiserF, Me.mnuRiserR, Me.mnuComputation, Me.mnuTools, Me.mnuHelp})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(771, 24)
        Me.MainMenu1.TabIndex = 3
        '
        'mnuFile
        '
        Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFileNew, Me.mnuFileOpen, Me.mnuFileSave, Me.mnuFileSaveAs, Me.mnuFilePD, Me.mnuFilePrintSetup, Me.mnuLineFilePre, Me._mnuFilePre_0, Me._mnuFilePre_1, Me._mnuFilePre_2, Me._mnuFilePre_3, Me.mnuExitLine, Me.mnuFileExit})
        Me.mnuFile.Name = "mnuFile"
        Me.mnuFile.Size = New System.Drawing.Size(37, 20)
        Me.mnuFile.Text = "&File"
        '
        'mnuFileNew
        '
        Me.mnuFileNew.Name = "mnuFileNew"
        Me.mnuFileNew.Size = New System.Drawing.Size(174, 22)
        Me.mnuFileNew.Text = "&New Project"
        '
        'mnuFileOpen
        '
        Me.mnuFileOpen.Name = "mnuFileOpen"
        Me.mnuFileOpen.Size = New System.Drawing.Size(174, 22)
        Me.mnuFileOpen.Text = "&Open Project"
        '
        'mnuFileSave
        '
        Me.mnuFileSave.Name = "mnuFileSave"
        Me.mnuFileSave.Size = New System.Drawing.Size(174, 22)
        Me.mnuFileSave.Text = "&Save Project"
        '
        'mnuFileSaveAs
        '
        Me.mnuFileSaveAs.Name = "mnuFileSaveAs"
        Me.mnuFileSaveAs.Size = New System.Drawing.Size(174, 22)
        Me.mnuFileSaveAs.Text = "Save Project &As..."
        '
        'mnuFilePD
        '
        Me.mnuFilePD.Name = "mnuFilePD"
        Me.mnuFilePD.Size = New System.Drawing.Size(174, 22)
        Me.mnuFilePD.Text = "Project &Description"
        '
        'mnuFilePrintSetup
        '
        Me.mnuFilePrintSetup.Name = "mnuFilePrintSetup"
        Me.mnuFilePrintSetup.Size = New System.Drawing.Size(174, 22)
        Me.mnuFilePrintSetup.Text = "&Print Setup"
        '
        'mnuLineFilePre
        '
        Me.mnuLineFilePre.Name = "mnuLineFilePre"
        Me.mnuLineFilePre.Size = New System.Drawing.Size(171, 6)
        Me.mnuLineFilePre.Visible = False
        '
        'mnuExitLine
        '
        Me.mnuExitLine.Name = "mnuExitLine"
        Me.mnuExitLine.Size = New System.Drawing.Size(171, 6)
        '
        'mnuFileExit
        '
        Me.mnuFileExit.Name = "mnuFileExit"
        Me.mnuFileExit.Size = New System.Drawing.Size(174, 22)
        Me.mnuFileExit.Text = "&Exit"
        '
        'mnuAP
        '
        Me.mnuAP.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuAPGlobal, Me.mnuAPCurrentProfiles, Me.mnuRefreshMainMenu})
        Me.mnuAP.Name = "mnuAP"
        Me.mnuAP.Size = New System.Drawing.Size(124, 20)
        Me.mnuAP.Text = "&Analysis Parameters"
        '
        'mnuAPGlobal
        '
        Me.mnuAPGlobal.Name = "mnuAPGlobal"
        Me.mnuAPGlobal.Size = New System.Drawing.Size(177, 22)
        Me.mnuAPGlobal.Text = "&Global Parameters"
        '
        'mnuAPCurrentProfiles
        '
        Me.mnuAPCurrentProfiles.Name = "mnuAPCurrentProfiles"
        Me.mnuAPCurrentProfiles.Size = New System.Drawing.Size(177, 22)
        Me.mnuAPCurrentProfiles.Text = "Current &Profile"
        '
        'mnuRefreshMainMenu
        '
        Me.mnuRefreshMainMenu.Name = "mnuRefreshMainMenu"
        Me.mnuRefreshMainMenu.Size = New System.Drawing.Size(177, 22)
        Me.mnuRefreshMainMenu.Text = "Refresh Main Menu"
        '
        'mnuRiserF
        '
        Me.mnuRiserF.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuRiserFRiserTypeandBC, Me.mnuRiserFSegment, Me.mnuRiserFAuxline, Me.mnuRiserFSupSpring, Me.mnuRiserFStaticSolution, Me.mnuRiserFOtherProperties})
        Me.mnuRiserF.Name = "mnuRiserF"
        Me.mnuRiserF.Size = New System.Drawing.Size(75, 20)
        Me.mnuRiserF.Text = "Riser &Front"
        '
        'mnuRiserFRiserTypeandBC
        '
        Me.mnuRiserFRiserTypeandBC.Name = "mnuRiserFRiserTypeandBC"
        Me.mnuRiserFRiserTypeandBC.Size = New System.Drawing.Size(266, 22)
        Me.mnuRiserFRiserTypeandBC.Text = "&Riser Type and Boundary Conditions"
        '
        'mnuRiserFSegment
        '
        Me.mnuRiserFSegment.Name = "mnuRiserFSegment"
        Me.mnuRiserFSegment.Size = New System.Drawing.Size(266, 22)
        Me.mnuRiserFSegment.Text = "&Segment Data"
        '
        'mnuRiserFAuxline
        '
        Me.mnuRiserFAuxline.Name = "mnuRiserFAuxline"
        Me.mnuRiserFAuxline.Size = New System.Drawing.Size(266, 22)
        Me.mnuRiserFAuxline.Text = "&Auxilliary Line Data"
        '
        'mnuRiserFSupSpring
        '
        Me.mnuRiserFSupSpring.Name = "mnuRiserFSupSpring"
        Me.mnuRiserFSupSpring.Size = New System.Drawing.Size(266, 22)
        Me.mnuRiserFSupSpring.Text = "&Intermediate Lateral Supports"
        '
        'mnuRiserFStaticSolution
        '
        Me.mnuRiserFStaticSolution.Enabled = False
        Me.mnuRiserFStaticSolution.Name = "mnuRiserFStaticSolution"
        Me.mnuRiserFStaticSolution.Size = New System.Drawing.Size(266, 22)
        Me.mnuRiserFStaticSolution.Text = "SCR/&Lazy-Wave Static Solution"
        '
        'mnuRiserFOtherProperties
        '
        Me.mnuRiserFOtherProperties.Name = "mnuRiserFOtherProperties"
        Me.mnuRiserFOtherProperties.Size = New System.Drawing.Size(266, 22)
        Me.mnuRiserFOtherProperties.Text = "Other Riser Properties"
        '
        'mnuRiserR
        '
        Me.mnuRiserR.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuRiserRRiserTypeandBC, Me.mnuRiserRSegment, Me.mnuRiserRAuxLine, Me.mnuRiserRSupSpring, Me.mnuRiserRStaticSolution, Me.mnuRiserROtherProperties})
        Me.mnuRiserR.Name = "mnuRiserR"
        Me.mnuRiserR.Size = New System.Drawing.Size(70, 20)
        Me.mnuRiserR.Text = "Riser &Rear"
        '
        'mnuRiserRRiserTypeandBC
        '
        Me.mnuRiserRRiserTypeandBC.Name = "mnuRiserRRiserTypeandBC"
        Me.mnuRiserRRiserTypeandBC.Size = New System.Drawing.Size(266, 22)
        Me.mnuRiserRRiserTypeandBC.Text = "&Riser Type and Boundary Conditions"
        '
        'mnuRiserRSegment
        '
        Me.mnuRiserRSegment.Name = "mnuRiserRSegment"
        Me.mnuRiserRSegment.Size = New System.Drawing.Size(266, 22)
        Me.mnuRiserRSegment.Text = "&Segment Data"
        '
        'mnuRiserRAuxLine
        '
        Me.mnuRiserRAuxLine.Name = "mnuRiserRAuxLine"
        Me.mnuRiserRAuxLine.Size = New System.Drawing.Size(266, 22)
        Me.mnuRiserRAuxLine.Text = "&Auxilliary Line Data"
        '
        'mnuRiserRSupSpring
        '
        Me.mnuRiserRSupSpring.Name = "mnuRiserRSupSpring"
        Me.mnuRiserRSupSpring.Size = New System.Drawing.Size(266, 22)
        Me.mnuRiserRSupSpring.Text = "&Intermediate Lateral Supports"
        '
        'mnuRiserRStaticSolution
        '
        Me.mnuRiserRStaticSolution.Enabled = False
        Me.mnuRiserRStaticSolution.Name = "mnuRiserRStaticSolution"
        Me.mnuRiserRStaticSolution.Size = New System.Drawing.Size(266, 22)
        Me.mnuRiserRStaticSolution.Text = "SCR/&Lazy-Wave Static Solution"
        '
        'mnuRiserROtherProperties
        '
        Me.mnuRiserROtherProperties.Name = "mnuRiserROtherProperties"
        Me.mnuRiserROtherProperties.Size = New System.Drawing.Size(266, 22)
        Me.mnuRiserROtherProperties.Text = "Other Riser Properties"
        '
        'mnuComputation
        '
        Me.mnuComputation.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuAPComputation, Me.mnuResultsF, Me.mnuResultsR, Me.mnuBatch})
        Me.mnuComputation.Name = "mnuComputation"
        Me.mnuComputation.Size = New System.Drawing.Size(90, 20)
        Me.mnuComputation.Text = "&Computation"
        '
        'mnuAPComputation
        '
        Me.mnuAPComputation.Name = "mnuAPComputation"
        Me.mnuAPComputation.Size = New System.Drawing.Size(170, 22)
        Me.mnuAPComputation.Text = "&Calculating Now"
        '
        'mnuResultsF
        '
        Me.mnuResultsF.Name = "mnuResultsF"
        Me.mnuResultsF.Size = New System.Drawing.Size(170, 22)
        Me.mnuResultsF.Text = "Results Riser &Front"
        '
        'mnuResultsR
        '
        Me.mnuResultsR.Name = "mnuResultsR"
        Me.mnuResultsR.Size = New System.Drawing.Size(170, 22)
        Me.mnuResultsR.Text = "Results Riser &Rear"
        '
        'mnuBatch
        '
        Me.mnuBatch.Name = "mnuBatch"
        Me.mnuBatch.Size = New System.Drawing.Size(170, 22)
        Me.mnuBatch.Text = "&Batch Processing"
        '
        'mnuTools
        '
        Me.mnuTools.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuDatabase})
        Me.mnuTools.Name = "mnuTools"
        Me.mnuTools.Size = New System.Drawing.Size(48, 20)
        Me.mnuTools.Text = "Tools"
        '
        'mnuDatabase
        '
        Me.mnuDatabase.Name = "mnuDatabase"
        Me.mnuDatabase.Size = New System.Drawing.Size(199, 22)
        Me.mnuDatabase.Text = "&Database Configuration"
        '
        'mnuHelp
        '
        Me.mnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuManual_0, Me._mnuManual_1, Me.mnuManualLine, Me.mnuHelpAbout, Me.mnuReference})
        Me.mnuHelp.Name = "mnuHelp"
        Me.mnuHelp.Size = New System.Drawing.Size(44, 20)
        Me.mnuHelp.Text = "&Help"
        '
        'mnuManualLine
        '
        Me.mnuManualLine.Name = "mnuManualLine"
        Me.mnuManualLine.Size = New System.Drawing.Size(165, 6)
        '
        'mnuHelpAbout
        '
        Me.mnuHelpAbout.Name = "mnuHelpAbout"
        Me.mnuHelpAbout.Size = New System.Drawing.Size(168, 22)
        Me.mnuHelpAbout.Text = "&About WinVIVA"
        '
        'mnuReference
        '
        Me.mnuReference.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuVIVOfFlexibleStructures, Me.mnuFatigueCalculation, Me.mnuModelingStrakes, Me.mnuOMAE1, Me.mnuOMAE2, Me.mnuQuestionAnswers, Me.mnuVortexInducedVibrationsOfLongCyclindricalStructures})
        Me.mnuReference.Name = "mnuReference"
        Me.mnuReference.Size = New System.Drawing.Size(168, 22)
        Me.mnuReference.Text = "Reference"
        '
        'mnuVIVOfFlexibleStructures
        '
        Me.mnuVIVOfFlexibleStructures.Name = "mnuVIVOfFlexibleStructures"
        Me.mnuVIVOfFlexibleStructures.Size = New System.Drawing.Size(409, 22)
        Me.mnuVIVOfFlexibleStructures.Text = "3D VIV of Flexible Structures (2001)"
        '
        'mnuFatigueCalculation
        '
        Me.mnuFatigueCalculation.Name = "mnuFatigueCalculation"
        Me.mnuFatigueCalculation.Size = New System.Drawing.Size(409, 22)
        Me.mnuFatigueCalculation.Text = "Fatigue Calculation in VIVA (2010)"
        '
        'mnuModelingStrakes
        '
        Me.mnuModelingStrakes.Name = "mnuModelingStrakes"
        Me.mnuModelingStrakes.Size = New System.Drawing.Size(409, 22)
        Me.mnuModelingStrakes.Text = "Modeling Strakes and Foil Sections in VIVA (2002)"
        '
        'mnuOMAE1
        '
        Me.mnuOMAE1.Name = "mnuOMAE1"
        Me.mnuOMAE1.Size = New System.Drawing.Size(409, 22)
        Me.mnuOMAE1.Text = "OMAE2011-49820 (2011)"
        '
        'mnuOMAE2
        '
        Me.mnuOMAE2.Name = "mnuOMAE2"
        Me.mnuOMAE2.Size = New System.Drawing.Size(409, 22)
        Me.mnuOMAE2.Text = "OMAE2011-50192 (2011)"
        '
        'mnuQuestionAnswers
        '
        Me.mnuQuestionAnswers.Name = "mnuQuestionAnswers"
        Me.mnuQuestionAnswers.Size = New System.Drawing.Size(409, 22)
        Me.mnuQuestionAnswers.Text = "Questions & Answers about VIVA (2003)"
        '
        'mnuVortexInducedVibrationsOfLongCyclindricalStructures
        '
        Me.mnuVortexInducedVibrationsOfLongCyclindricalStructures.Name = "mnuVortexInducedVibrationsOfLongCyclindricalStructures"
        Me.mnuVortexInducedVibrationsOfLongCyclindricalStructures.Size = New System.Drawing.Size(409, 22)
        Me.mnuVortexInducedVibrationsOfLongCyclindricalStructures.Text = "Vortex Induced Vibrations of Long Cyclindrical Structures (2002)"
        '
        'frmMainMenu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(771, 508)
        Me.Controls.Add(Me.refreshMainMenu)
        Me.Controls.Add(Me._lblCopyright_0)
        Me.Controls.Add(Me._lblCopyright_1)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(11, 30)
        Me.MaximizeBox = False
        Me.Name = "frmMainMenu"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "WinVIVA - Main Menu"
        CType(Me.refreshMainMenu, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblCopyright, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuFilePre, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuManual, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents mnuComputation As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuAPComputation As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuBatch As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuResultsF As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuTools As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuDatabase As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuRiserFOtherProperties As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuRiserR As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuRiserRRiserTypeandBC As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuRiserRSegment As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuRiserRAuxLine As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuRiserRSupSpring As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuRiserRStaticSolution As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuRiserFSegment As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuRiserROtherProperties As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuResultsR As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuRefreshMainMenu As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuReference As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuVIVOfFlexibleStructures As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuFatigueCalculation As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuModelingStrakes As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuOMAE1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuOMAE2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuQuestionAnswers As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuVortexInducedVibrationsOfLongCyclindricalStructures As System.Windows.Forms.ToolStripMenuItem
#End Region
End Class