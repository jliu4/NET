Option Strict Off
Option Explicit On

' Version 8.1
' 2014, Copyright DTCEL, All Rights Reserved
' Jin Liu May 2014

Friend Class frmOtherRiserProperties
    Inherits System.Windows.Forms.Form

    Const APIRP2A_A As Single = 1372700000.0#
    Const APIRP2A_B As Single = 4.38
    Const DEnB_A As Single = 2827900000.0#
    Const DEnB_B As Single = 4.0#
    Private CFStress As Single
    Dim FatCurv As FatigueCurve
    Private CFDensity, CFViscosity As Single



    Private Sub cmdBrowse_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBrowse.Click

        Dim index As Short = cmdBrowse.GetIndex(eventSender)
        Dim dlgFiles As New System.Windows.Forms.OpenFileDialog

        '   should the user cancel the dialog box, exit
        On Error GoTo ErrHandler

        With dlgFiles
            '       set filters to allow selection of all files or just .in
            .Filter = "All Files (*.*)|*.*" '|Input Files (*.in)|*.in"

            '       specify default filter as *.in
            .FilterIndex = 2

            If CurProj.DataDirectory = "" Then
                .InitialDirectory = VIVADIR
            Else
                .InitialDirectory = CurProj.DataDirectory
            End If

            '    .InitDir = CheckPath(txtDirectory, True) 'TBC jjx 07/23
            .CheckFileExists = True
            .CheckPathExists = True 'cdlOFNHideReadOnly +

            '.CancelError = True
            '       display the Open dialog box
            .ShowDialog()

            Select Case index
                Case 0
                    txtModesFile(0).Text = .FileName
                Case 1
                    txtModesFile(1).Text = .FileName
                Case 2
                    txtModesFile(2).Text = .FileName
                Case 3
                    txtModesFile(3).Text = .FileName
            End Select
        End With

        Exit Sub

ErrHandler:
        '   user pressed Cancel button
        Exit Sub

    End Sub


    Private Sub optModes_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optModes.CheckedChanged

        If eventSender.Checked Then
            Dim index As Short = optModes.GetIndex(eventSender)

            Select Case index
                Case 0
                    txtModesFile(0).Enabled = False
                    cmdBrowse(0).Enabled = False
                    txtModesFile(1).Enabled = False
                    cmdBrowse(1).Enabled = False
                    txtModesFile(2).Enabled = False
                    cmdBrowse(2).Enabled = False
                    txtModesFile(3).Enabled = False
                    cmdBrowse(3).Enabled = False

                Case 1
                    txtModesFile(0).Enabled = True
                    cmdBrowse(0).Enabled = True
                    If CurProj.Riser(CurProj.RiserId).FreqFile <> "" Then txtModesFile(0).Text = CurProj.Riser(CurProj.RiserId).FreqFile
                    txtModesFile(1).Enabled = False
                    cmdBrowse(1).Enabled = False
                    txtModesFile(2).Enabled = False
                    cmdBrowse(2).Enabled = False
                    txtModesFile(3).Enabled = False
                    cmdBrowse(3).Enabled = False

                Case 2
                    txtModesFile(1).Enabled = True
                    cmdBrowse(1).Enabled = True
                    If CurProj.Riser(CurProj.RiserId).FreqFile <> "" Then txtModesFile(1).Text = CurProj.Riser(CurProj.RiserId).FreqFile
                    txtModesFile(2).Enabled = True
                    cmdBrowse(2).Enabled = True
                    If CurProj.Riser(CurProj.RiserId).ModeFile <> "" Then txtModesFile(2).Text = CurProj.Riser(CurProj.RiserId).ModeFile
                    txtModesFile(3).Enabled = True
                    cmdBrowse(3).Enabled = True
                    If CurProj.Riser(CurProj.RiserId).CurvFile <> "" Then txtModesFile(3).Text = CurProj.Riser(CurProj.RiserId).CurvFile
                    txtModesFile(0).Enabled = False
                    cmdBrowse(0).Enabled = False
            End Select

        End If

    End Sub


    Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click

        Me.Close()

    End Sub

    Private Sub btnOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnOK.Click


        If optModes(1).Checked = True Then
            If Trim(txtModesFile(0).Text) = "Freq File Name" Or Trim(txtModesFile(0).Text) = "" Then
                MsgBox("Please import a file of frequencies", MsgBoxStyle.OkOnly, "WinVIVA - Importing File Missing")
                txtModesFile(0).Focus()
                Exit Sub
            End If
            CurProj.Riser(CurProj.RiserId).Mode = 1
            CurProj.Riser(CurProj.RiserId).FreqFile = txtModesFile(0).Text
            'FileCopy UserFreqFile, VIVADIR & "freq"
        ElseIf optModes(2).Checked = True Then
            If Trim(txtModesFile(1).Text) = "Freq File Name" Or Trim(txtModesFile(1).Text) = "" Then
                MsgBox("Please import a file of frequencies", MsgBoxStyle.OkOnly, "WinVIVA - Importing File Missing")
                txtModesFile(1).Focus()
                Exit Sub
            End If
            If Trim(txtModesFile(2).Text) = "Modes File Name" Or Trim(txtModesFile(2).Text) = "" Then
                MsgBox("Please import a file of modes", MsgBoxStyle.OkOnly, "WinVIVA - Importing File Missing")
                txtModesFile(1).Focus()
                Exit Sub
            End If
            If Trim(txtModesFile(3).Text) = "Curvature File Name" Or Trim(txtModesFile(3).Text) = "" Then
                MsgBox("Please import a file of curvatures", MsgBoxStyle.OkOnly, "WinVIVA - Importing File Missing")
                txtModesFile(1).Focus()
                Exit Sub
            End If
            CurProj.Riser(CurProj.RiserId).Mode = 2
            CurProj.Riser(CurProj.RiserId).FreqFile = txtModesFile(1).Text
            CurProj.Riser(CurProj.RiserId).ModeFile = txtModesFile(2).Text
            CurProj.Riser(CurProj.RiserId).CurvFile = txtModesFile(3).Text
            'FileCopy UserFreqFile, VIVADIR & "freq" 'default name used by VIVA DOS program
            'FileCopy UserModesFile, VIVADIR & "modes_us"
        Else
            CurProj.Riser(CurProj.RiserId).Mode = 0
            '        CurProj.FreqFile = ""
            '        CurProj.ModeFile = ""
        End If
        '   Check each entry for validity:  is it a number; and is it positive?
        If Not PositiveNumber((txtFatConstA.Text)) Then
            MsgBox("Please enter a valid positive number for the Constant A")
            txtFatConstA.Focus()
            Exit Sub
        ElseIf Not PositiveNumber((txtFatConstB.Text)) Then
            MsgBox("Please enter a valid positive number for the Constant B")
            txtFatConstB.Focus()
            Exit Sub
        ElseIf Not PositiveNumber((txtSCF.Text)) Then
            MsgBox("Please enter a valid positive number for the SCF")
            txtSCF.Focus()
            Exit Sub
        End If

        '   If all numbers are OK, store them in the appropriate object
        CurProj.Riser(CurProj.RiserId).FatigueCurveName = cboFatigueCurves.Text
        CurProj.Riser(CurProj.RiserId).FatigueConstA = CSng(txtFatConstA.Text) * CFStress
        CurProj.Riser(CurProj.RiserId).FatigueConstB = CSng(txtFatConstB.Text)
        CurProj.Riser(CurProj.RiserId).SCF = CSng(txtSCF.Text)

        If btnHiRe.Checked = True Then
            CurProj.Riser(CurProj.RiserId).HiRe = WinVIVA.HiReData.Use
        Else
            CurProj.Riser(CurProj.RiserId).HiRe = WinVIVA.HiReData.NoUse
        End If

        '   remember that the case will need to be saved
        CurProj.Saved = False

        '   unload
        Me.Close()

    End Sub

    Private Sub txtVis_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub cboFatigueCurves_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboFatigueCurves.SelectedIndexChanged

        With FatigueCurves.Item(cboFatigueCurves.SelectedIndex + 1)
            txtFatConstA.Text = CStr(CSng(.A / CFStress))
            txtFatConstB.Text = .B
        End With

    End Sub


    Private Sub frmOtherRiserProperty_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        Select Case CurProj.Riser(CurProj.RiserId).Mode
            Case 0
                optModes(0).Checked = True
                optModes_CheckedChanged(optModes.Item(0), New System.EventArgs())
            Case 1
                optModes(1).Checked = True
                optModes_CheckedChanged(optModes.Item(1), New System.EventArgs())
            Case 2
                optModes(2).Checked = True
                optModes_CheckedChanged(optModes.Item(2), New System.EventArgs())
        End Select

        UpdateTextFields()
        With CurProj
            Text = Text & " : " & CurProj.RiserId & " - " & .Title

            If .Units = VIVAMain.Units.English Then
                CFStress = Lb2N / (In2Ft * Ft2M) ^ 2
                lblUnit.Text = "psi"
            Else
                CFStress = 1000000.0#
                lblUnit.Text = "MPa"
            End If

            LoadFatigueCurves()

            cboFatigueCurves.Text = .Riser(.RiserId).FatigueCurveName
            txtFatConstA.Text = CStr(.Riser(.RiserId).FatigueConstA / CFStress)
            txtFatConstB.Text = CStr(.Riser(.RiserId).FatigueConstB)
            txtSCF.Text = CStr(.Riser(.RiserId).SCF)
            If .Riser(.RiserId).HiRe = VIVAMain.HiReData.Use Then
                btnHiRe.Checked = True
            Else
                btnNoHiRe.Checked = True
            End If

        End With

    End Sub
    Private Sub txtDampingCoef_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDampingCoef.Enter

        '   set the text box properties to select the entire field
        txtDampingCoef.SelectionStart = 0
        txtDampingCoef.SelectionLength = 100

    End Sub

    Private Sub btnAPIConsts_Click()

        txtFatConstA.Text = CStr(APIRP2A_A / CFStress)
        txtFatConstB.Text = CStr(APIRP2A_B)

    End Sub


    Private Sub btnDEnBConsts_Click()

        txtFatConstA.Text = CStr(DEnB_A / CFStress)
        txtFatConstB.Text = CStr(DEnB_B)

    End Sub
    Private Sub txtFatConstB_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFatConstB.Enter

        '   Set the text box properties to select the entire field
        txtFatConstB.SelectionStart = 0
        txtFatConstB.SelectionLength = 100

    End Sub

    Private Sub UpdateTextFields()

        '   load the form fields from curproj
        txtDampingCoef.Text = CStr(CurProj.Riser(CurProj.RiserId).Damping)
    End Sub
    'definition of fatigue curves
    'load them into the combo box
    Private Sub LoadFatigueCurves()

        '   load them into the combo box
        Dim i, NumCurves As Short

        cboFatigueCurves.Items.Clear()
        NumCurves = FatigueCurves.Count()

        For i = 1 To NumCurves
            cboFatigueCurves.Items.Add(FatigueCurves.Item(i).Name)
        Next

    End Sub

End Class