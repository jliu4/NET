Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmMain

    Inherits System.Windows.Forms.Form
    ' frmMain           form of main window
    ' Version 1.0
    ' 2015, Copyright Genesis, Jin Liu

    ' for messages
    Private msgNotSavedWarning As String
    Private Const msgInputWarning As String = "The input file has not been read properly. " & " It may be corrupted. Please check data closely."
    Private Const msgNoFileWarning As String = "Cannot find the specifiied input file."
    Private msgOutputWarning As String

    ' for form operations
    Private NumProj As Short

    Private CheckingGrid As Boolean
    Private FirstLoad As Boolean
    Public PlotShow As Boolean

    ' summary
    ' for grid
    Private LDLabels(11) As String
    Private ELLabels(4) As String

    Private NumLines As Short
    Private FMGlob As Force

    Private Sub btnEnv_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnEnv.Click
        mnuInpEnviron_Click(mnuInpEnviron, New System.EventArgs())
    End Sub

    ' form load and unload

    Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        IsMetricUnit = False

        Dim i As Short
        Dim FileOpen_Renamed As Boolean

        On Error GoTo ErrorHandler

        InitProject()
        FileOpen_Renamed = False

        NumProj = 1
        With CurProj
            CurVessel = .Vessel
            CurField = .WellSites
            CurVessel.Name = .Vessel.Name
            lblVessel.Text = UCase(CurVessel.Name)

            .Title = "Project" & NumProj
            Text = " MARS Console - " & .Title
            .Saved = True

        End With

        If Len(Dir(MarsDir & IniFile)) > 0 Then
            FileOpen(10, MarsDir & IniFile, OpenMode.Input, OpenAccess.Read)
            FileOpen_Renamed = True
            If Defaults.InputData(10) Then
                UpdateFileMenu()
                UpdateMenu()
                CurProj.Directory = Defaults.WorkDir
            End If
            FileClose(10)
            FileOpen_Renamed = False
        End If

        FMGlob = New Force
        NumLines = CurVessel.MoorSystem.MoorLineCount

        PlotShow = False
        InitSummary()
        UpdateFileMenu()
        UpdateMenu()
        LoadData(True)
        RefreshUnitLabels(Me)
        Exit Sub

ErrorHandler:
        If FileOpen_Renamed Then FileClose(10)
        Cursor = System.Windows.Forms.Cursors.Default
        Me.Close()

    End Sub

    Private Sub frmMain_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason

        Dim i As Short

        If CurProj.Saved = False Then
            If CurProj.Title = "" Then
                msgNotSavedWarning = "Do you want to save changes made to " & " the current project ?"
            Else
                msgNotSavedWarning = "Do you want to save changes made to " & CurProj.Title & " ?"
            End If

            Select Case MsgBox(msgNotSavedWarning, MsgBoxStyle.YesNoCancel, msgTitle)
                Case MsgBoxResult.Yes
                    '           save
                    mnuFileSave_Click(mnuFileSave, New System.EventArgs())

                    Defaults.AddPreFile(CurProj.Directory & CurProj.FileName)
                Case MsgBoxResult.No
                    If CurProj.FileName <> "" Then Defaults.AddPreFile(CurProj.Directory & CurProj.FileName)
                Case MsgBoxResult.Cancel
                    '           do nothing
                    Cancel = True
            End Select
        Else
            With CurProj
                If .FileName <> "" Then Defaults.AddPreFile(.Directory & .FileName)
            End With
        End If

        ' Loop through the forms collection and unload
        ' each form.
        'TODO JLIU
        If Not Cancel Then
            On Error Resume Next
            Defaults.WorkDir = CurProj.Directory
            FileOpen(20, MarsDir & IniFile, OpenMode.Output)

            FileClose(20)

            'Destroy existing case and leave
            CurProj = Nothing
            FMGlob = Nothing
        End If

        eventArgs.Cancel = Cancel
    End Sub

    ' buttons

    Private Sub btnPlot_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnPlot.Click
        Dim Index As Short = btnPlot.GetIndex(eventSender)

        With frmPlot
            Select Case Index
                Case 0
                    If Not PlotShow Then .Show()
                    .PlotIn3D = True
                Case 1
                    If Not PlotShow Then .Show()
                    .PlotIn3D = False
            End Select
            .UpdatePicture()
        End With

    End Sub

    Private Sub btnMooring_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnMooring.Click

        frmMoorLines.Show()
        NumLines = CurVessel.MoorSystem.MoorLineCount
        UpdateMenu()
        RefreshData()

    End Sub

    Private Sub btnWell_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnWell.Click

        Dim NumWells As Short
        Dim i As Short
        frmWells.Show()

        NumWells = CurField.Count
        With CurField
            cboWells.Items.Clear()
            For i = 1 To NumWells
                cboWells.Items.Add(.Item(i).NameID)
            Next i
            cboWells.SelectedIndex = .CurWellNo - 1
            With .Item(.CurWellNo)
                txtWell(0).Text = Format(.Xg * LFactor, "0.0")
                txtWell(1).Text = Format(.Yg * LFactor, "0.0")
                If .Depth > 0# Then CurVessel.WaterDepth = .Depth
                txtWell(2).Text = Format(CurVessel.WaterDepth * LFactor, "0.0")
            End With
        End With

    End Sub

    Private Sub btnUpdate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnUpdate.Click
        SaveLC()
        InitSummary()
        LoadData(False)
    End Sub

    Private Sub frmMain_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ' make sure all forms are unloaded first
        ' any referenced form would remain in memory even if invisible
        Dim frm As System.Windows.Forms.Form
        For Each frm In My.Application.OpenForms
            If Not frm Is Me Then frm.Close()
        Next frm
        End
    End Sub

    Public Sub mnu3DPlot_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnu3DPlot.Click
        Call btnPlot_Click(btnPlot.Item(0), New System.EventArgs())
    End Sub

    ' menus
    Public Sub mnuFileNew_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileNew.Click

        If CurProj.Saved = False Then
            If CurProj.Title = "" Then
                msgNotSavedWarning = "Do you want to save changes made to " & " the current project ?"
            Else
                msgNotSavedWarning = "Do you want to save changes made to " & CurProj.Title & " ?"
            End If

            Select Case MsgBox(msgNotSavedWarning, MsgBoxStyle.YesNoCancel, msgTitle)
                Case MsgBoxResult.Yes
                    '           save
                    mnuFileSave_Click(mnuFileSave, New System.EventArgs())
                    If CurProj.FileName <> "" Then Defaults.AddPreFile(CurProj.Directory & CurProj.FileName)
                Case MsgBoxResult.No
                    If CurProj.FileName <> "" Then Defaults.AddPreFile(CurProj.Directory & CurProj.FileName)
                Case MsgBoxResult.Cancel
                    '           do nothing
                    Exit Sub
            End Select
        Else
            With CurProj
                If .FileName <> "" Then Defaults.AddPreFile(.Directory & .FileName)
            End With
        End If
        UpdateFileMenu()

        CurProj = Nothing
        ClearScreenData()
        frmMain_Load(Me, New System.EventArgs())

        VB6.ShowForm(frmProjDesc, 1, Me)

    End Sub

    Public Sub mnuFileOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileOpen.Click

        Dim FileOpen_Renamed As Boolean
        '   should the user cancel the dialog box, exit
        On Error GoTo ErrHandler

        FileOpen_Renamed = False

        '   save the current project
        If CurProj.Saved = False Then
            If CurProj.Title = "" Then
                msgNotSavedWarning = "Do you want to save changes made to " & " the current project ?"
            Else
                msgNotSavedWarning = "Do you want to save changes made to " & CurProj.Title & " ?"
            End If

            Select Case MsgBox(msgNotSavedWarning, MsgBoxStyle.YesNoCancel, msgTitle)
                Case MsgBoxResult.Yes
                    '           save
                    mnuFileSave_Click(mnuFileSave, New System.EventArgs())
                    If CurProj.FileName <> "" Then Defaults.AddPreFile(CurProj.Directory & CurProj.FileName)
                Case MsgBoxResult.No
                    If CurProj.FileName <> "" Then Defaults.AddPreFile(CurProj.Directory & CurProj.FileName)
                Case MsgBoxResult.Cancel
                    '           do nothing
                    Exit Sub
            End Select
        Else
            With CurProj
                If .FileName <> "" Then Defaults.AddPreFile(.Directory & .FileName)
            End With
        End If
        UpdateFileMenu()

        '   get input file name
        Dim CatFile As String
        With dlgFileOpen
            .Filter = "All Files (*.*)|*.*|MARS Files (*.mrs)|*.mrs"
            .FilterIndex = 2
            If CurProj.Directory = "" Then
                .InitialDirectory = MarsDir
            Else
                .InitialDirectory = CurProj.Directory
            End If
            .ShowReadOnly = False
            .CheckFileExists = True
            .CheckPathExists = True
            .CheckPathExists = True
            .ShowDialog()

            Cursor = System.Windows.Forms.Cursors.WaitCursor
            CurProj = Nothing
            CurProj = New Project

            ClearScreenData()

            With CurProj
                CurVessel = .Vessel
                CurField = .WellSites

                .GetDirNFileName(dlgFileOpen.FileName)
                Text = " MARS Console - " & .Title & " (" & .FileName & ")"

                If .ImportData(dlgFileOpen.FileName) Then
                    CurVessel.Name = .Vessel.Name
                    lblVessel.Text = .Vessel.Name
                    .Saved = True
                    If Not .VesselParticular Then
                        MsgBox("MARS is unable to load vessel particulars.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, msgTitle)
                        '   Unload Me
                    End If
                Else
                    MsgBox(msgInputWarning, MsgBoxStyle.OkOnly, msgTitle)
                    .Saved = False
                End If
            End With
            FileOpen_Renamed = False

            NumLines = CurVessel.MoorSystem.MoorLineCount
            UpdateMenu()
            LoadData(True)

            CatFile = MarsDir & CatenaryFile
            Call CurVessel.MoorSystem.GenCatenaryFile(CatFile)

            Cursor = System.Windows.Forms.Cursors.Default
        End With

        Exit Sub

ErrHandler:
        '   User pressed Cancel button
        UpdateMenu()
        Cursor = System.Windows.Forms.Cursors.Default

    End Sub

    Public Sub mnuFileSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileSave.Click
        SaveLC()

        Dim FileOpen_Renamed As Boolean
        Dim FileNum As Integer

        On Error GoTo ErrorHandler

        FileOpen_Renamed = False
        FileNum = FreeFile()

        With CurProj
            If .FileName = "" Then
                mnuFileSaveAs_Click(mnuFileSaveAs, New System.EventArgs())
            Else
                FileOpen(FileNum, .Directory & .FileName, OpenMode.Output)
                FileOpen_Renamed = True
                Cursor = System.Windows.Forms.Cursors.WaitCursor
                If Not .ExportData(FileNum) Then
                End If
                Cursor = System.Windows.Forms.Cursors.Default
                FileClose(FileNum)
                FileOpen_Renamed = False
                .Saved = True
            End If
        End With

        Exit Sub

ErrorHandler:
        FileClose(FileNum)
        Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub

    End Sub

    Public Sub mnuFileSaveAs_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileSaveAs.Click
        SaveLC()
        Dim FileOpen_Renamed As Boolean

        '   should the user cancel the dialog box, exit
        On Error GoTo ErrHandler

        FileOpen_Renamed = False

        '   get file name
        Dim fnum As Integer
        With dlgFileSave
            .Filter = "All Files (*.*)|*.*|MARS Files (*.mrs)|*.mrs"
            .FilterIndex = 2
            If CurProj.Directory = "" Then
                .InitialDirectory = MarsDir
            Else
                .InitialDirectory = CurProj.Directory
            End If
            If CurProj.FileName <> "" Then
                .FileName = CurProj.FileName
            Else
                .FileName = CurProj.Title & ".mrs"
            End If
            .OverwritePrompt = True
            .ShowDialog()

            '       save data
            fnum = FreeFile()
            FileOpen(fnum, .FileName, OpenMode.Output)
            FileOpen_Renamed = True

            Cursor = System.Windows.Forms.Cursors.WaitCursor
            With CurProj
                If .FileName <> "" Then
                    Defaults.AddPreFile(.Directory & .FileName)
                    UpdateFileMenu()
                End If

                .GetDirNFileName(dlgFileOpen.FileName)
                Text = " MARS Console - " & .Title & " (" & .FileName & ")"

                If Not .ExportData(fnum) Then
                End If
                Cursor = System.Windows.Forms.Cursors.Default
                .Saved = True
            End With
            FileClose(fnum)
            FileOpen_Renamed = False
        End With

        Exit Sub

ErrHandler:
        '   User pressed Cancel button
        If FileOpen_Renamed Then FileClose(fnum)
        Cursor = System.Windows.Forms.Cursors.Default
        If Err.Number <> 32755 Then ' don't report cancel dialog error
            MsgBox("Error " & Err.Description & ",  Source: " & Err.Source, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Error")
        End If
    End Sub

    Public Sub mnuFilePrintSetup_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFilePrintSetup.Click

        dlgFilePrint.ShowDialog()

    End Sub

    Public Sub mnuFilePre_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFilePre.Click
        Dim Index As Short = mnuFilePre.GetIndex(eventSender)

        Dim InputFile As String
        Dim FileOpen_Renamed As Boolean

        '   should the user cancel the dialog box, exit
        On Error GoTo ErrHandler

        FileOpen_Renamed = False

        '   save the current project
        If CurProj.Saved = False Then
            If CurProj.Title = "" Then
                msgNotSavedWarning = "Do you want to save changes made to " & " the current project ?"
            Else
                msgNotSavedWarning = "Do you want to save changes made to " & CurProj.Title & " ?"
            End If

            Select Case MsgBox(msgNotSavedWarning, MsgBoxStyle.YesNoCancel, msgTitle)
                Case MsgBoxResult.Yes
                    '           save
                    mnuFileSave_Click(mnuFileSave, New System.EventArgs())
                    If CurProj.FileName <> "" Then Defaults.AddPreFile(CurProj.Directory & CurProj.FileName)
                Case MsgBoxResult.No
                    If CurProj.FileName <> "" Then Defaults.AddPreFile(CurProj.Directory & CurProj.FileName)
                '           do nothing
                Case MsgBoxResult.Cancel
                    Exit Sub
            End Select
        Else
            With CurProj
                If .FileName <> "" Then Defaults.AddPreFile(.Directory & .FileName)
            End With
        End If

        '   input from file
        '  InputFile = Defaults.PreFile(Index + 1)
        InputFile = VB.Right(mnuFilePre(Index).Text, Len(mnuFilePre(Index).Text) - 3)
        If Len(Dir(InputFile)) <= 0 Then
            MsgBox(msgNoFileWarning, MsgBoxStyle.OkOnly, msgTitle)
            Defaults.DelPreFile(Index + 1)
            UpdateFileMenu()
            Exit Sub
        End If

        Cursor = System.Windows.Forms.Cursors.WaitCursor
        CurProj = Nothing
        CurProj = New Project
        ClearScreenData()

        With CurProj
            .GetDirNFileName(InputFile)
            Text = " MARS Console - " & .Title & " (" & .FileName & ")"

            If .ImportData(InputFile) Then
                .Saved = True
                CurVessel = .Vessel
                CurField = .WellSites
                lblVessel.Text = .Vessel.Name
                CurVessel.Name = .Vessel.Name
                lblVessel.Text = CurVessel.Name
                If Not .VesselParticular Then
                    MsgBox("MARS is unable to load vessel particulars.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, msgTitle)
                    '  Unload Me
                End If
            Else
                MsgBox(msgInputWarning, MsgBoxStyle.OkOnly, msgTitle)
                .Saved = False
            End If
        End With
        Cursor = System.Windows.Forms.Cursors.Default

        NumLines = CurVessel.MoorSystem.MoorLineCount
        UpdateMenu()
        LoadData(True)
        ' force catenary output
        Dim CatFile As String
        CatFile = MarsDir & CatenaryFile
        Call CurVessel.MoorSystem.GenCatenaryFile(CatFile)
        Exit Sub

ErrHandler:
        '   User pressed Cancel button
        Cursor = System.Windows.Forms.Cursors.Default
        MsgBox("mnuFilePre: " & Err.Description & ", Source: " & Err.Source, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Error")
    End Sub

    Public Sub mnuFileExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileExit.Click

        Me.Close()

    End Sub

    Public Sub mnuInpEnviron_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuInpEnviron.Click
        frmEnviron.Show()

    End Sub

    Public Sub mnuInpMoor_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuInpMoor.Click

        btnMooring_Click(btnMooring, New System.EventArgs())

    End Sub

    Public Sub mnuInpProjDes_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuInpProjDes.Click
        frmProjDesc.Show()

    End Sub

    Public Sub mnuMoor_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuMoor.Click

        Me.Enabled = False
        SaveLC()
        frmAnalyses.Show()


    End Sub

    Public Sub mnuAnalysesB_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAnalysesB.Click

        Me.Enabled = False
        frmAnalysesB.Show()


    End Sub

    Public Sub mnuMoorLayoutReport_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuMoorLayoutReport.Click
        ' check if has data
        If CurVessel.MoorSystem.MoorLineCount = 0 Then
            MsgBox("Mooring system is not yet defined.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Error")
            Exit Sub
        End If

        Dim ox As New ExcelReporter
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        ox.ReportMooringLayout(txtClientName.Text, txtLocationName.Text, CurVessel)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Public Sub mnuMove_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuMove.Click
        frmMove.Show()

    End Sub

    Public Sub mnuHelpAbout_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuHelpAbout.Click
        frmAbout.Show()

    End Sub

    Function GetDampingPercent(ByVal WaterDepth As Single) As Object

        ' by experience on chain-rig wire system
        Dim Val2, Val1, BaseVal As Single

        If WaterDepth < 1000 Then
            GetDampingPercent = 10
            Exit Function
        ElseIf WaterDepth >= 1000 And WaterDepth < 2000 Then
            Val1 = 1000
            Val2 = 2000
            BaseVal = 10
        ElseIf WaterDepth >= 2000 And WaterDepth < 3000 Then
            Val1 = 2000
            Val2 = 3000
            BaseVal = 20
        ElseIf WaterDepth >= 3000 And WaterDepth < 4000 Then
            Val1 = 3000
            Val2 = 4000
            BaseVal = 30
        ElseIf WaterDepth >= 4000 And WaterDepth < 5000 Then
            Val1 = 4000
            Val2 = 5000
            BaseVal = 40
        ElseIf WaterDepth >= 5000 Then
            GetDampingPercent = 50
            Exit Function
        End If

        GetDampingPercent = (WaterDepth - Val1) / (Val2 - Val1) * 10 + BaseVal
    End Function

    Private Sub SetLblSummary()

        LDLabels(0) = "Line" & vbCrLf & "No."
        LDLabels(1) = "Top" & vbCrLf & "Tension"
        LDLabels(2) = "Line FOS"
        LDLabels(3) = "Payout"
        LDLabels(4) = "Top Angle" & vbCrLf & "(deg)"
        LDLabels(5) = "Local" & vbCrLf & "Spread" & vbCrLf & "Angle" & vbCrLf & "(deg)"
        LDLabels(6) = "Grounded" & vbCrLf & "Length"
        LDLabels(7) = "Anchor" & vbCrLf & "Tension"
        LDLabels(8) = "Anchor" & vbCrLf & "FOS"
        LDLabels(9) = "Scope"
        LDLabels(10) = "Connected"
        LDLabels(11) = "Winch" & vbCrLf & "Functional"

        If IsMetricUnit Then

            LDLabels(1) = LDLabels(1) & vbCrLf & "(KN)"
            LDLabels(3) = LDLabels(3) & vbCrLf & "(m)"
            LDLabels(6) = LDLabels(6) & vbCrLf & "(m)"
            LDLabels(7) = LDLabels(7) & vbCrLf & "(KN)"
            LDLabels(9) = LDLabels(9) & vbCrLf & "(m)"
        Else
            LDLabels(1) = LDLabels(1) & vbCrLf & "(kips)"
            LDLabels(3) = LDLabels(3) & vbCrLf & "(ft)"
            LDLabels(6) = LDLabels(6) & vbCrLf & "(ft)"
            LDLabels(7) = LDLabels(7) & vbCrLf & "(kips)"
            LDLabels(9) = LDLabels(9) & vbCrLf & "(ft)"
        End If
    End Sub

    Private Sub ClearScreenData()
        grdLD.Rows.Clear()
        grdEL.Rows.Clear()

        txtWell(0).Text = "0"
        txtWell(1).Text = "0"
        txtWell(2).Text = "0"
        txtVslSt(0).Text = "0"
        txtVslSt(1).Text = "0"
        txtVslSt(4).Text = "0"
        txtVslSt(5).Text = "0"
        txtVslSt(2).Text = "0"
        txtVslSt(3).Text = "0"
        cboWells.Items.Clear()
        cboWells.Items.Add("Default")
        cboWells.SelectedIndex = 0
    End Sub

    Private Sub InitSummary()

        Dim c As Short

        CheckingGrid = True

        Call SetLblSummary()

        With grdLD
            '.TopLeftHeaderCell.Value = "Line" & vbCrLf & "No."
            .Rows.Clear()
            .RowCount = Max(10, NumLines)
            .ColumnCount = UBound(LDLabels) + 1
            For c = 0 To .ColumnCount - 1
                .Columns(c).FillWeight = 100 / .ColumnCount
                .Columns(c).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Columns(c).HeaderText = LDLabels(c)
            Next

        End With

        '' ---------  Initialize Mooring Forces Grid ---------------

        SetLblMoorForce()

        With grdEL
            .TopLeftHeaderCell.Value = "Direction"
            .ColumnCount = 4
            .RowCount = 1
            For c = 0 To .ColumnCount - 1
                .Columns(c).FillWeight = 100 / .ColumnCount
                .Columns(c).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Columns(c).HeaderText = ELLabels(c)
            Next c

            .Rows(0).Cells(0).Value = "Mooring Forces"
        End With

        CheckingGrid = False

    End Sub


    Private Sub SetLblMoorForce()
        ELLabels(0) = "Direction"
        If IsMetricUnit Then
            ELLabels(1) = "x (E) (KN)"
            ELLabels(2) = "y (N) (KN)"
            ELLabels(3) = "Yaw (KN-m)"
        Else
            ELLabels(1) = "x (E) (kips)"
            ELLabels(2) = "y (N) (kips)"
            ELLabels(3) = "Yaw (Kips-ft)"
        End If
        ELLabels(4) = "Mooring Forces"
    End Sub

    ' update menu

    Private Sub UpdateMenu()

        With CurVessel.MoorSystem
            If .MoorLineCount < 1 Then
                btnUpdate.Enabled = False
                '      btnOnBoard.Enabled = False
                'mnuOnBoard.Enabled = False
                mnuMoor.Enabled = False
                mnuTransient.Enabled = False
                mnuAnalysesB.Enabled = False
                mnuMove.Enabled = False
                '    btnAnalyze.Enabled = False
            Else
                btnUpdate.Enabled = True
                mnuMoor.Enabled = True
                mnuAnalysesB.Enabled = True
                mnuTransient.Enabled = True
                mnuMove.Enabled = True
                '       btnAnalyze.Enabled = True
                If .MoorLineCount < 2 Then
                    mnuMove.Enabled = False
                Else
                    mnuMove.Enabled = True
                End If
            End If
        End With

    End Sub

    Private Sub UpdateFileMenu()

        Dim i, NumPreFiles As Short

        NumPreFiles = Defaults.CountPreFile

        If NumPreFiles > 0 Then mnuLinePreFile.Visible = True
        For i = 1 To NumPreFiles Step 1
            With mnuFilePre(i - 1)
                .Text = i & ". " & Defaults.PreFile(i)
                .Enabled = True
                .Visible = True
            End With
        Next i
        For i = NumPreFiles + 1 To MaxNumPreFiles Step 1
            With mnuFilePre(i - 1)
                .Text = ""
                .Enabled = False
                .Visible = False
            End With
        Next i

    End Sub

    Public Sub LoadData(Optional ByRef FirstTime As Boolean = False)

        If IsMetricUnit Then
            LFactor = 0.3048 ' ft -> m
            FrcFactor = 4.448222 ' kips -> KN
        Else
            LFactor = 1
            FrcFactor = 1
        End If

        Dim c, r, i As Short
        Dim NumWells As Short
        Dim Distance, Bearing As Single

        Dim ShipCurLoc As ShipGlobal

        Dim CurLine As MoorLine

        NumWells = CurField.Count
        If FirstTime Then FirstLoad = True
        With CurField
            cboWells.Items.Clear()
            For i = 1 To NumWells
                cboWells.Items.Add(.Item(i).NameID)
            Next i
            cboWells.SelectedIndex = .CurWellNo - 1
            With .Item(.CurWellNo)
                txtWell(0).Text = Format(.Xg * LFactor, "0.0")
                txtWell(1).Text = Format(.Yg * LFactor, "0.0")
                If .Depth > 0# Then CurVessel.WaterDepth = .Depth
                txtWell(2).Text = Format(CurVessel.WaterDepth * LFactor, "0.0")
            End With
        End With

        With CurVessel
            With .ShipCurGlob
                txtVslSt(0).Text = Format(.Xg * LFactor, "0.00")
                txtVslSt(1).Text = Format(.Yg * LFactor, "0.00")
                txtVslSt(4).Text = Format(RadTo360(.Heading), "0.00")
            End With

            txtVslSt(5).Text = Format(.ShipDraft * LFactor, "0.00")

            Coord2Bear(.ShipCurGlob, Distance, Bearing)
            txtVslSt(2).Text = Format(Distance * LFactor, "0.0")
            txtVslSt(3).Text = Format(Bearing * Radians2Degrees, "0.00")

            .MoorSystem.MoorForce(FMGlob, (CurVessel.ShipCurGlob))
        End With

        Dim HorForce As Single ' output
        Dim Scope As Single

        InitSummary()
        With grdLD
            For r = 0 To NumLines - 1
                CurLine = CurVessel.MoorSystem.MoorLines(r + 1)
                .Rows(r).HeaderCell.Value = (r + 1).ToString
                If CurLine.Connected Then
                    .Rows(r).Cells(0).Value = (r + 1)
                    Scope = CurLine.ScopeByVesselLocation((CurVessel.ShipCurGlob))

                    .Rows(r).Cells(1).Value = Format(CurLine.TensionByScopePOL(Scope, HorForce, (CurLine.Payout)) * 0.001 * FrcFactor, "0.00")

                    .Rows(r).Cells(2).Value = Format(CurLine.FOS, "0.00")
                    .Rows(r).Cells(3).Value = Format(CurLine.Payout * LFactor, "0.0")

                    .Rows(r).Cells(4).Value = Format(CurLine.FLAngle * Radians2Degrees, "0.00")

                    .Rows(r).Cells(5).Value = Format(CurLine.SprdAngle(ShipCurLoc) * Radians2Degrees, "0.00")
                    '                .Text = Format(Dir * Radians2Degrees, "0.00")

                    .Rows(r).Cells(6).Value = Format(CurLine.GrdLen * LFactor, "0.0")

                    .Rows(r).Cells(7).Value = Format(CurLine.AnchPull * 0.001 * FrcFactor, "0.0")

                    .Rows(r).Cells(8).Value = CurLine.AnchorFOS
                    .Rows(r).Cells(9).Value = Format(Scope * LFactor, "0.0")

                    If CurLine.Connected Then
                        .Rows(r).Cells(10).Value = "Yes"
                    Else
                        .Rows(r).Cells(10).Value = "No"
                    End If

                    If CurLine.Winched Then
                        .Rows(r).Cells(11).Value = "Yes"
                    Else
                        .Rows(r).Cells(11).Value = "No"
                    End If

                End If

            Next r
        End With

        With grdEL

            .Rows(0).Cells(1).Value = Format(FMGlob.Fx * 0.001 * FrcFactor, "0.0")

            .Rows(0).Cells(2).Value = Format(FMGlob.Fy * 0.001 * FrcFactor, "0.0")

            .Rows(0).Cells(3).Value = Format(FMGlob.MYaw * 0.001 * FrcFactor * LFactor, "0.0")
        End With

        '   btnOnBoard.Enabled = True
        'mnuOnBoard.Enabled = True

        ' Update Catenary file
        CurVessel.MoorSystem.GenCatenaryFile(MarsDir & CatenaryFile)

        If PlotShow Then frmPlot.UpdatePicture()

    End Sub

    Public Sub SaveLC()

        If IsMetricUnit Then
            LFactor = 0.3048 ' ft -> m
            FrcFactor = 4.448222 ' kips -> KN
        Else
            LFactor = 1
            FrcFactor = 1
        End If

        Dim Distance, Bearing As Single

        With CurVessel
            If optInputSystem(0).Checked Then
                With .ShipCurGlob
                    .Xg = CDbl(Format(CDbl(CheckData(CStr(Val(txtVslSt(0).Text)),  , True)) / LFactor, "0.00"))
                    .Yg = CDbl(Format(CDbl(CheckData(CStr(Val(txtVslSt(1).Text)),  , True)) / LFactor, "0.00"))
                    .Heading = CDbl(CheckData(CStr(Val(txtVslSt(4).Text)),  , True)) * Degrees2Radians
                End With
                Coord2Bear(.ShipCurGlob, Distance, Bearing)
                txtVslSt(2).Text = Format(Distance * LFactor, "0.0")
                txtVslSt(3).Text = Format(Bearing * Radians2Degrees, "0.00")
            Else
                Distance = CDbl(CheckData(CStr(Val(txtVslSt(2).Text) / LFactor),  , True))
                Bearing = CDbl(CheckData(CStr(Val(txtVslSt(3).Text)),  , True)) * Degrees2Radians
                Bear2Coord(.ShipCurGlob, Distance, Bearing)
                With .ShipCurGlob
                    txtVslSt(0).Text = Format(.Xg * LFactor, "0.0")
                    txtVslSt(1).Text = Format(.Yg * LFactor, "0.0")
                    .Heading = CDbl(CheckData(txtVslSt(4).Text,  , True)) * Degrees2Radians
                End With
            End If
            .ShipDraft = CDbl(CheckData(CStr(Val(txtVslSt(5).Text) / LFactor),  , True))
        End With

    End Sub

    Private Function OnBoardData() As Boolean

        Dim r As Short
        Dim FileOpen_Renamed As Boolean
        Dim CurLine As MoorLine

        OnBoardData = False

        On Error GoTo ErrorHandler
        FileOpen_Renamed = False
        FileOpen(FileNumRes, MarsDir & OnBoardFile, OpenMode.Output)
        FileOpen_Renamed = True

        For r = 1 To NumLines
            With CurVessel.MoorSystem.MoorLines(r)
                If .Connected Then
                    PrintLine(FileNumRes, r, .Payout, .TopTension * 0.001, .FLAngle * Radians2Degrees, .SprdAngle((CurVessel.ShipCurGlob)) * Radians2Degrees)
                Else
                    PrintLine(FileNumRes, r, 0#, 0#, 0#, 0#)
                End If
            End With
        Next r

        FileClose(FileNumRes)
        FileOpen_Renamed = False
        OnBoardData = True
        Exit Function

ErrorHandler:
        If FileOpen_Renamed Then FileClose(FileNumRes)
        Exit Function

    End Function

    Private Sub Coord2Bear(ByRef ShipGlob As ShipGlobal, ByRef Distance As Single, ByRef Bearing As Single)

        Dim dx, dy As Single

        With CurField
            dx = ShipGlob.Xg - .Item(.CurWellNo).Xg
            dy = ShipGlob.Yg - .Item(.CurWellNo).Yg
        End With

        Distance = System.Math.Sqrt(dx ^ 2 + dy ^ 2)
        Bearing = Atan(dy, dx)

    End Sub

    Private Sub Bear2Coord(ByRef ShipGlob As ShipGlobal, ByRef Distance As Single, ByRef Bearing As Single)

        Dim dx, dy As Single

        dx = Distance * System.Math.Sin(Bearing)
        dy = Distance * System.Math.Cos(Bearing)

        With CurField
            ShipGlob.Xg = .Item(.CurWellNo).Xg + dx
            ShipGlob.Yg = .Item(.CurWellNo).Yg + dy
        End With

    End Sub

    Public Sub mnuOptions_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuOptions.Click
        VB6.ShowForm(frmOptions, 1, Me)
        RefreshData()
    End Sub

    Public Sub mnuRadarPlot_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRadarPlot.Click
        Call btnPlot_Click(btnPlot.Item(1), New System.EventArgs())
    End Sub

    Public Sub mnuTransient_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuTransient.Click
        VB6.ShowForm(frmTransient, 0, Me) ' can't show non-modal progressbar when modal is displayed
    End Sub

    Private Sub RefreshData()
        RefreshUnitLabels(Me)
        LoadData(False)
    End Sub

    Private Sub grdLD_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdLD.CellDoubleClick
        If e.ColumnIndex = 10 Then
            If grdLD.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = "Yes" Then
                grdLD.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = "No"
                CurVessel.MoorSystem.MoorLines(e.RowIndex + 1).Connected = False

            Else
                grdLD.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = "Yes"
                CurVessel.MoorSystem.MoorLines(e.RowIndex + 1).Connected = True
            End If
        End If


    End Sub
End Class