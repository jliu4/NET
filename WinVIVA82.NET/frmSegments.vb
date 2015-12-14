Option Strict Off
Option Explicit On

' Version 8.1
' 2014, Copyright DTCEL, All Rights Reserved
' Jin Liu May 2014

Friend Class frmSegments

    Inherits System.Windows.Forms.Form

    Const LengthCol As Short = 2
    Const TypeCol As Short = 7
    Private riserId As Short
    Private additionalDatabase As Short = 10
    Private ExistingTxt As String
    Private Changed As Boolean
    Private CheckingGrid As Boolean
    Private JEC As Boolean
    Private CurUnits As VIVAMain.Units
    Private CFForce, CFLength, CFDim, CFStress, CFEI As Single
    Private RiserLabels(1, 13) As String
    Private typeSeg As New DataGridViewComboBoxCell()

    Private Sub grdSegments_CellValidating( _
      ByVal sender As Object, _
      ByVal e As System.Windows.Forms. _
      DataGridViewCellValidatingEventArgs) _
      Handles grdSegments.CellValidating

        grdSegments.Rows(e.RowIndex).ErrorText = ""

        If grdSegments.Rows(e.RowIndex).IsNewRow Then Return

        If Not (e.ColumnIndex = 7) Then 'Or e.ColumnIndex = 12) Then
            If Not (IsNumeric(e.FormattedValue)) And Not String.IsNullOrEmpty(e.FormattedValue) Then
                e.Cancel = True
                grdSegments.Rows(e.RowIndex).ErrorText = _
                   "It must be a numeric value."
            End If
        End If

    End Sub

    Private Sub grdSegments_CellValueChanged(ByVal sender As Object, _
    ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) _
    Handles grdSegments.CellValueChanged
        If Not CheckingGrid Then

            If e.ColumnIndex = 2 Or e.ColumnIndex = 3 Or e.ColumnIndex = 11 Then
                UpdateCellEI()
            End If

        End If
    End Sub

    Private Sub frmSegments_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        riserId = CurProj.RiserId

        Text = Text & " : " & riserId & " - " & CurProj.Title
        CheckingGrid = False

        CurUnits = CurProj.Units
        SetLabels()
        IniConvertFactor()

        LoadGridFromProject()

        NumInputForms = NumInputForms + 1
        frmGlobalParms.btnMetric.Enabled = False
        frmGlobalParms.btnEnglish.Enabled = False

    End Sub


    Private Sub frmSegments_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        NumInputForms = NumInputForms - 1
        If NumInputForms = 0 And UnitsChangeable Then
            frmGlobalParms.btnMetric.Enabled = True
            frmGlobalParms.btnEnglish.Enabled = True
        End If

    End Sub


    Private Sub SetLabels()
        '   Metric unit labels
        RiserLabels(VIVAMain.Units.Metric, 0) = "No*"
        RiserLabels(VIVAMain.Units.Metric, 1) = "Number of Joints"
        RiserLabels(VIVAMain.Units.Metric, 2) = "Joint Length (m)"
        RiserLabels(VIVAMain.Units.Metric, 3) = "Main Tube O.D. (mm)"
        RiserLabels(VIVAMain.Units.Metric, 4) = "Wall Thickness (mm)"
        RiserLabels(VIVAMain.Units.Metric, 5) = "Dry Weight* (N/Joint)"
        RiserLabels(VIVAMain.Units.Metric, 6) = "Wet Weight* (N/Joint)"
        RiserLabels(VIVAMain.Units.Metric, 7) = "Buoy. Module Dia. (mm)"
        RiserLabels(VIVAMain.Units.Metric, 8) = "Section Type"
        RiserLabels(VIVAMain.Units.Metric, 9) = "Strakes Height (mm)"
        RiserLabels(VIVAMain.Units.Metric, 10) = "Fairing Thickness (mm)"
        RiserLabels(VIVAMain.Units.Metric, 11) = "Fairing Chord Len. (mm)"
        RiserLabels(VIVAMain.Units.Metric, 12) = "Modulus of Elasticity (MPa)"
        RiserLabels(VIVAMain.Units.Metric, 13) = "Bending Stiffness (MN*m^2)"

        '   English unit labels
        RiserLabels(VIVAMain.Units.English, 0) = "No*"
        RiserLabels(VIVAMain.Units.English, 1) = "Number of Joints"
        RiserLabels(VIVAMain.Units.English, 2) = "Joint Length (ft)"
        RiserLabels(VIVAMain.Units.English, 3) = "Main Tube O.D. (in)"
        RiserLabels(VIVAMain.Units.English, 4) = "Wall Thickness (in)"
        RiserLabels(VIVAMain.Units.English, 5) = "Dry Weight* (lb/Joint)"
        RiserLabels(VIVAMain.Units.English, 6) = "Wet Weight* (lb/Joint)"
        RiserLabels(VIVAMain.Units.English, 7) = "Buoy. Module Dia. (in)"
        RiserLabels(VIVAMain.Units.English, 8) = "Section Type"
        RiserLabels(VIVAMain.Units.English, 9) = "Strakes Height (in)"
        RiserLabels(VIVAMain.Units.English, 10) = "Fairing Thickness (in)"
        RiserLabels(VIVAMain.Units.English, 11) = "Fairing Chord Len. (in)"
        RiserLabels(VIVAMain.Units.English, 12) = "Modulus of Elasticity (psi)"
        RiserLabels(VIVAMain.Units.English, 13) = "Bending Stiffness (lb*in^2)"

    End Sub


    Private Sub IniConvertFactor()

        If CurUnits = VIVAMain.Units.Metric Then
            CFLength = 1.0#
            CFDim = mm2M
            CFForce = 1.0#
            CFStress = 1000000.0#
            CFEI = CFStress * CFLength ^ 2
        Else
            CFLength = Ft2M
            CFDim = In2Ft * Ft2M
            CFForce = Lb2N
            CFStress = Lb2N / (In2Ft * Ft2M) ^ 2
            CFEI = Lb2N * (In2Ft * Ft2M) ^ 2
        End If

    End Sub


    Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click

        Me.Close()

    End Sub


    Private Sub btnOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnOK.Click

        Cursor = System.Windows.Forms.Cursors.WaitCursor
        If UpdateRiserSegments() Then
            CurProj.Saved = False
            Cursor = System.Windows.Forms.Cursors.Default
            Me.Close()
        Else
            Cursor = System.Windows.Forms.Cursors.Default
        End If

    End Sub


    Private Sub grdSegments_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As DataGridViewCellEventArgs) _
        Handles grdSegments.CellClick

        Dim i As Short
        If eventArgs.ColumnIndex = 7 Then
            Dim cell As New DataGridViewComboBoxCell()
            additionalDatabase = VIVACoreFiles.NumFiles - additionalDatabase
            cell.MaxDropDownItems = 4 + additionalDatabase
            cell.FlatStyle = FlatStyle.Flat
            cell.Items.Add("Smooth")
            cell.Items.Add("Vecto")
            cell.Items.Add("Staggard")
            cell.Items.Add("Straked P/D=17")
            i = 0
            While i < additionalDatabase
                i = i + 1
                If IsNothing(VIVACoreFiles.FileNames(9 + i)) Then
                    Exit While

                End If
                cell.Items.Add(VIVACoreFiles.FileNames(9 + i))
            End While
        cell.Value = "Smooth"
        grdSegments(eventArgs.ColumnIndex, eventArgs.RowIndex) = cell
        End If

        If eventArgs.ColumnIndex = 12 Then
            If Me.RdBttnYoungsModules.Checked = True Then
                grdSegments.Columns(eventArgs.ColumnIndex).ReadOnly = True 'TODO Jliu
                Exit Sub
            End If
        End If
    End Sub
    Private Sub grdSegments_LeaveCell(ByVal eventSender As System.Object, ByVal eventArgs As DataGridViewCellEventArgs) _
    Handles grdSegments.CellLeave
        If eventArgs.ColumnIndex = 7 Then
            Dim cell As New DataGridViewTextBoxCell()

            cell.Value = grdSegments(eventArgs.ColumnIndex, eventArgs.RowIndex).Value
            'grdSegments(eventArgs.ColumnIndex, eventArgs.RowIndex) = cell
        End If
    End Sub


    Private Sub SetSegCbo(ByRef SecType As String)

        Dim i As Short

        For i = 1 To typeSeg.Items.Count
            'If SecType = VB6.GetItemString(typeSeg.containsFocus, i - 1) Then Exit For
        Next
        If i > typeSeg.Items.Count Then i = 1
        'typeSeg.SelectedIndex = i - 1

    End Sub

    Private Sub frm_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) _
        Handles MyBase.MouseDown

        If eventArgs.Button = System.Windows.Forms.MouseButtons.Right Then

            ContextMenuStripGridEdit.Show(MousePosition)

        End If

    End Sub

    Public Sub mnuAddRow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAddRow.Click

        AddRow(grdSegments, CheckingGrid)

    End Sub


    Public Sub mnuDeleteRow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDeleteRow.Click

        DeleteRow(grdSegments, CheckingGrid)

    End Sub

    Public Sub paste_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles pasteFromClipboard.Click

        PastefromClipBoardToDataGridView(grdSegments, CheckingGrid, MaxSegments)

    End Sub
    Public Sub copyToExcel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles copyToExcel.Click

        copyToClipBoard(grdSegments)

    End Sub

    Private Function UpdateRiserSegments() As Boolean

        Dim r, c As Short
        Dim NumRows, NumSegments As Short
        Dim NewSeg As Segment

        'avoid triggering the "Leave Cell" event
        CheckingGrid = True
        With CurProj.Riser(riserId)
            If Me.RdBttnYoungsModules.Checked Then
                .CalculateBendingStiffness = True
            ElseIf Me.RdBttnStiffness.Checked Then
                .CalculateBendingStiffness = False
            End If

            If Not (CSng(Me.txtVectoAngle.Text) >= 0 And CSng(Me.txtVectoAngle.Text) < 360) Then
                MsgBox("Please use a value between 0 and 360 degree for the Vecto angle")
                CheckingGrid = False
                UpdateRiserSegments = False
                Exit Function
            End If

            For r = 0 To grdSegments.RowCount - 2
                For c = 0 To grdSegments.ColumnCount - 1
                    If String.IsNullOrEmpty(grdSegments(c, r).Value) Then
                        MsgBox("No blank in all cells")
                        CheckingGrid = False
                        UpdateRiserSegments = False
                        Exit Function
                    End If
                Next
            Next


            .VectoAngle = CSng(Me.txtVectoAngle.Text)
        End With

        With grdSegments
            '       scan for a blank cell in this row; if one is found, bail out
            NumRows = .RowCount
            '       delete the old input
            NumSegments = CurProj.Riser(riserId).Segments.Count
            For r = 1 To NumSegments
                CurProj.Riser(riserId).Segments.Delete(1)
            Next r
            'MsgBox(cboSeg.DataGridView.Rows(0).Cells)

            '       insert the updated input
            If .RowCount > 1 Then
                For r = 0 To .RowCount - 2
                    NewSeg = CurProj.Riser(riserId).Segments.Add
                    NewSeg.NumJoints = CShort(.Rows(r).Cells(0).Value)
                    NewSeg.JointLength = CSng(.Rows(r).Cells(1).Value) * CFLength
                    NewSeg.MainTubeOD = CSng(.Rows(r).Cells(2).Value) * CFDim
                    NewSeg.WallThickness = CSng(.Rows(r).Cells(3).Value) * CFDim
                    NewSeg.DryWeight = CSng(.Rows(r).Cells(4).Value) * CFForce
                    NewSeg.WetWeight = CSng(.Rows(r).Cells(5).Value) * CFForce
                    NewSeg.BuoyModDia = CSng(.Rows(r).Cells(6).Value) * CFDim

                    Select Case (.Rows(r).Cells(7).Value)
                        Case typeSeg.Items(0)
                            NewSeg.SectionType = ICHARValues.SmoothCylinder
                        Case typeSeg.Items(1)
                            NewSeg.SectionType = ICHARValues.VetcoRiser0
                        Case typeSeg.Items(2)
                            NewSeg.SectionType = ICHARValues.StaggardBareBuoyant
                        Case typeSeg.Items(3)
                            NewSeg.SectionType = ICHARValues.TestedStrakes
                        Case typeSeg.Items(4)
                            NewSeg.SectionType = ICHARValues.userAdded11
                        Case typeSeg.Items(5)
                            NewSeg.SectionType = ICHARValues.userAdded12
                        Case typeSeg.Items(6)
                            NewSeg.SectionType = ICHARValues.userAdded13
                        Case typeSeg.Items(7)
                            NewSeg.SectionType = ICHARValues.userAdded14
                        Case typeSeg.Items(8)
                            NewSeg.SectionType = ICHARValues.userAdded15
                        Case typeSeg.Items(9)
                            NewSeg.SectionType = ICHARValues.userAdded16
                        Case typeSeg.Items(10)
                            NewSeg.SectionType = ICHARValues.userAdded17
                        Case typeSeg.Items(11)
                            NewSeg.SectionType = ICHARValues.userAdded18

                    End Select
                    NewSeg.StrakesHeight = CSng(.Rows(r).Cells(8).Value) * CFDim
                    NewSeg.FairThick = CSng(.Rows(r).Cells(9).Value) * CFDim
                    NewSeg.FairChord = CSng(.Rows(r).Cells(10).Value) * CFDim
                    NewSeg.ModulusOfElasticity = CSng(.Rows(r).Cells(11).Value) * CFStress

                    If Me.RdBttnYoungsModules.Checked Then
                        NewSeg.BendingStiffness = NewSeg.EI * CFEI
                    Else
                        NewSeg.BendingStiffness = CSng(.Rows(r).Cells(12).Value) * CFEI

                    End If
                Next r
            End If
        End With

        '   reset the "Leaving Cell" trigger
        CheckingGrid = False
        UpdateRiserSegments = True

    End Function


    Private Sub LoadGridFromProject()

        Dim seg As Segment
        Dim r As Single
        Dim c As Short, i As Short
        'typeSeg.HeaderText = RiserLabels(CurUnits, 7 + 1)
        additionalDatabase = VIVACoreFiles.NumFiles - additionalDatabase

        typeSeg.MaxDropDownItems = 4 + additionalDatabase
        typeSeg.FlatStyle = FlatStyle.Flat
        typeSeg.Items.Add("Smooth")
        typeSeg.Items.Add("Vecto")
        typeSeg.Items.Add("Staggard")
        typeSeg.Items.Add("Straked P/D=17")
        i = 0
        While i < additionalDatabase
            i = i + 1
            If IsNothing(VIVACoreFiles.FileNames(9 + i)) Then
                Exit While

            End If
            typeSeg.Items.Add(VIVACoreFiles.FileNames(9 + i))
        End While
        CheckingGrid = True

        Select Case CurProj.Riser(riserId).CalculateBendingStiffness
            Case True
                Me.RdBttnYoungsModules.Checked = True
            Case False
                Me.RdBttnStiffness.Checked = True
        End Select

        Me.txtVectoAngle.Text = CStr(CurProj.Riser(riserId).VectoAngle)

        If CurProj.Riser(riserId).VectoAngle = -1 Then
            MsgBox("This project file was generated by an old version WinVIVA program." & Chr(13) & _
                "If there is any Vecto riser section, please check the angle.")
            Me.txtVectoAngle.Text = CStr("0")
        End If

        With grdSegments
            .RowCount = CurProj.Riser(riserId).Segments.Count + 1
            .ColumnCount = 13

            For c = 0 To .ColumnCount - 1
                .Columns(c).HeaderText = RiserLabels(CurUnits, c + 1)
                .Columns(c).Width = 50
            Next
            .Columns(4).Width = 61
            .Columns(5).Width = 61
            .Columns(7).Width = 75
            .Columns(12).Width = 95

            r = 0

            For Each seg In CurProj.Riser(riserId).Segments
                If r > MaxSegments Then Exit For
                .Rows(r).Cells(0).Value = CStr(seg.NumJoints)
                .Rows(r).Cells(1).Value = CStr(seg.JointLength / CFLength)
                .Rows(r).Cells(2).Value = CStr(seg.MainTubeOD / CFDim)
                .Rows(r).Cells(3).Value = CStr(seg.WallThickness / CFDim)
                .Rows(r).Cells(4).Value = CStr(seg.DryWeight / CFForce)
                .Rows(r).Cells(5).Value = CStr(seg.WetWeight / CFForce)

                .Rows(r).Cells(6).Value = CStr(seg.BuoyModDia / CFDim)
                .Rows(r).Cells(8).Value = CStr(seg.StrakesHeight / CFDim)

                .Rows(r).Cells(9).Value = CStr(seg.FairThick / CFDim)

                .Rows(r).Cells(10).Value = CStr(seg.FairChord / CFDim)

                .Rows(r).Cells(11).Value = CStr(seg.ModulusOfElasticity / CFStress)

                Select Case CurProj.Riser(riserId).CalculateBendingStiffness
                    Case True
                        .Rows(r).Cells(12).Value = CStr(seg.EI() / CFEI)
                    Case False
                        .Rows(r).Cells(12).Value = CStr(seg.BendingStiffness / CFEI)
                    Case False
                End Select
                r = r + 1
            Next seg

            r = 0
            For Each seg In CurProj.Riser(riserId).Segments
                If r > MaxSegments Then Exit For

                Select Case seg.SectionType
                    Case ICHARValues.SmoothCylinder, ICHARValues.TestedHighRe
                        .Rows(r).Cells(7).Value = typeSeg.Items(0).ToString
                    Case ICHARValues.StaggardBareBuoyant
                        .Rows(r).Cells(7).Value = typeSeg.Items(2).ToString
                    Case ICHARValues.TestedStrakes
                        .Rows(r).Cells(7).Value = typeSeg.Items(3).ToString
                    Case ICHARValues.VetcoRiser0
                        .Rows(r).Cells(7).Value = typeSeg.Items(1).ToString
                        If CurProj.Riser(riserId).VectoAngle = -1 _
                        Then Me.txtVectoAngle.Text = CStr(0)
                    Case ICHARValues.VetcoRiser30
                        .Rows(r).Cells(7).Value = typeSeg.Items(1).ToString
                        If CurProj.Riser(riserId).VectoAngle = -1 _
                        Then Me.txtVectoAngle.Text = CStr(330)
                    Case ICHARValues.VetcoRiser60
                        .Rows(r).Cells(7).Value = typeSeg.Items(1).ToString
                        If CurProj.Riser(riserId).VectoAngle = -1 _
                        Then Me.txtVectoAngle.Text = CStr(300)
                    Case ICHARValues.VetcoRiser90
                        .Rows(r).Cells(7).Value = typeSeg.Items(1).ToString
                        If CurProj.Riser(riserId).VectoAngle = -1 _
                        Then Me.txtVectoAngle.Text = CStr(270)
                    Case ICHARValues.VetcoRiser120
                        .Rows(r).Cells(7).Value = typeSeg.Items(1).ToString
                        If CurProj.Riser(riserId).VectoAngle = -1 _
                        Then Me.txtVectoAngle.Text = CStr(240)
                    Case ICHARValues.VetcoRiser150
                        .Rows(r).Cells(7).Value = typeSeg.Items(1).ToString
                        If CurProj.Riser(riserId).VectoAngle = -1 _
                        Then Me.txtVectoAngle.Text = CStr(210)
                    Case ICHARValues.userAdded11
                        .Rows(r).Cells(7).Value = typeSeg.Items(4).ToString
                    Case ICHARValues.userAdded12
                        .Rows(r).Cells(7).Value = typeSeg.Items(5).ToString
                    Case ICHARValues.userAdded13
                        .Rows(r).Cells(7).Value = typeSeg.Items(6).ToString
                    Case ICHARValues.userAdded14
                        .Rows(r).Cells(7).Value = typeSeg.Items(7).ToString
                    Case ICHARValues.userAdded15
                        .Rows(r).Cells(7).Value = typeSeg.Items(8).ToString
                    Case ICHARValues.userAdded16
                        .Rows(r).Cells(7).Value = typeSeg.Items(9).ToString
                    Case ICHARValues.userAdded17
                        .Rows(r).Cells(7).Value = typeSeg.Items(10).ToString
                    Case ICHARValues.userAdded18
                        .Rows(r).Cells(7).Value = typeSeg.Items(11).ToString
                    Case Else
                        .Rows(r).Cells(7).Value = typeSeg.Items(1).ToString
                End Select
                r = r + 1

            Next seg

        End With
        If (Me.RdBttnYoungsModules.Checked) Then
            grdSegments.Columns(12).ReadOnly = True
        End If
        NumberAllRows(grdSegments)
        CheckingGrid = False

    End Sub



    Private Sub RdBttnYoungsModules_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RdBttnYoungsModules.Click
        grdSegments.Columns(12).ReadOnly = True

    End Sub
    Private Sub RdBttnStiffness_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RdBttnStiffness.Click
        grdSegments.Columns(12).ReadOnly = False

    End Sub


    Private Sub UpdateCellEI()
        Dim r As Integer
        Dim msngOD, msngWallThickness, msngID, msngE As Single

        CheckingGrid = True

        If Me.RdBttnYoungsModules.Checked Then
            With grdSegments
                For r = 0 To .RowCount - 2
                    msngOD = CSng(.Rows(r).Cells(2).Value) * CFDim
                    msngWallThickness = CSng(.Rows(r).Cells(3).Value) * CFDim
                    msngID = msngOD - 2 * msngWallThickness
                    msngE = CSng(.Rows(r).Cells(11).Value) * CFStress
                    .Rows(r).Cells(12).Value = (Format(msngE * (msngOD ^ 4 - msngID ^ 4) * Pi / 64.0 / CFEI, "0.000E-00"))
                Next
            End With
        End If

        CheckingGrid = False

    End Sub

End Class