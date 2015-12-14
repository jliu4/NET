Option Strict Off
Option Explicit On

' Version 8.1
' 2014, Copyright DTCEL, All Rights Reserved
' Jin Liu May 2014

Friend Class frmSupSpring

    Inherits System.Windows.Forms.Form

    Private Changed As Boolean
    Private CheckingGrid As Boolean
    Private JEC As Boolean
    Private ExistingTxt As String
    Private CurUnits As VIVAMain.Units
    Private CFSpringConst, CFLength, CFDamperConst As Single
    Private RiserLabels(1, 3) As String

    Private Sub grdSupSpring_CellValidating( _
      ByVal sender As Object, _
      ByVal e As System.Windows.Forms. _
      DataGridViewCellValidatingEventArgs) _
      Handles grdSupSpring.CellValidating

        grdSupSpring.Rows(e.RowIndex).ErrorText = ""

        If grdSupSpring.Rows(e.RowIndex).IsNewRow Then Return

        If Not IsNumeric(e.FormattedValue) Then
            If Not (IsNumeric(e.FormattedValue)) And Not String.IsNullOrEmpty(e.FormattedValue) Then
                e.Cancel = True
                MsgBox("It must be numeric value")
                grdSupSpring.Rows(e.RowIndex).ErrorText = _
                   "It must be a numeric value."

            End If
        End If
    End Sub

    Private Sub frmSupSpring_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        '   disable the selection of units system in the global parameters form
        NumInputForms = NumInputForms + 1
        frmGlobalParms.btnMetric.Enabled = False
        frmGlobalParms.btnEnglish.Enabled = False

        '   initiate variables
        CheckingGrid = False
        Me.Text = Me.Text & " : " & CurProj.RiserId & " - " & CurProj.Title
        CurUnits = CurProj.Units

        '   load grid
        SetLabels()
        IniConvertFactor()
        LoadGridFromProject()

    End Sub


    Private Sub frmSupSpring_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        '   determine whether to enable the selection of units system
        NumInputForms = NumInputForms - 1

        If NumInputForms = 0 And UnitsChangeable Then
            frmGlobalParms.btnMetric.Enabled = True
            frmGlobalParms.btnEnglish.Enabled = True
        End If

    End Sub


    Private Sub SetLabels()

        '   Metric unit labels
        RiserLabels(VIVAMain.Units.Metric, 0) = "No"
        RiserLabels(VIVAMain.Units.Metric, 1) = "Dist. Along Riser (m)"
        RiserLabels(VIVAMain.Units.Metric, 2) = "Stiffness (N/m)"
        RiserLabels(VIVAMain.Units.Metric, 3) = "Damping (N-sec/m)"

        '   English unit labels
        RiserLabels(VIVAMain.Units.English, 0) = "No"
        RiserLabels(VIVAMain.Units.English, 1) = "Dist. Along Riser (ft)"
        RiserLabels(VIVAMain.Units.English, 2) = "Stiffness (lbs/ft)"
        RiserLabels(VIVAMain.Units.English, 3) = "Damping (lbs-sec/ft)"

    End Sub


    Private Sub IniConvertFactor()

        If CurUnits = VIVAMain.Units.Metric Then
            CFLength = 1.0#
            CFSpringConst = 1.0#
            CFDamperConst = 1.0#
        Else
            CFLength = Ft2M
            CFSpringConst = Lb2N / Ft2M
            CFDamperConst = Lb2N / Ft2M
        End If

    End Sub


    Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click

        Me.Close()

    End Sub


    Private Sub btnOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnOK.Click

        Cursor = System.Windows.Forms.Cursors.WaitCursor

        '   update spring data and unload
        If UpdateSupSpring() Then
            CurProj.Saved = False
            Cursor = System.Windows.Forms.Cursors.Default
            Me.Close()
        Else
            Cursor = System.Windows.Forms.Cursors.Default
        End If

    End Sub

    Public Sub mnuAddRow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAddRow.Click

        AddRow(grdSupSpring, CheckingGrid)

    End Sub

    Public Sub mnuDeleteRow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDeleteRow.Click

        DeleteRow(grdSupSpring, CheckingGrid)

    End Sub

    Public Sub paste_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles pasteFromClipBoard.Click

        PastefromClipBoardToDataGridView(grdSupSpring, CheckingGrid, MaxNumLatSupports)

    End Sub

    Public Sub copyToExcel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CopyToExcel.Click
        copyToClipBoard(grdSupSpring)

    End Sub

    Private Function UpdateSupSpring() As Boolean
        Dim r, c As Short
        Dim NumLatSupports As Short

        For r = 0 To grdSupSpring.RowCount - 2
            For c = 0 To grdSupSpring.ColumnCount - 1
                If String.IsNullOrEmpty(grdSupSpring(c, r).Value) Then
                    MsgBox("No blank in all cells")
                    CheckingGrid = False
                    UpdateSupSpring = False
                    Exit Function
                End If
            Next
        Next

        '   avoid triggering LeaveCell event
        CheckingGrid = True
        With grdSupSpring
            NumLatSupports = CurProj.Riser(CurProj.RiserId).LatSupports.Count
            For r = 1 To NumLatSupports
                CurProj.Riser(CurProj.RiserId).LatSupports.Delete(1)
            Next

            If .RowCount > 1 Then
                For r = 0 To .RowCount - 2
                    If .Rows(r).Cells(0).Value IsNot Nothing And _
                                        .Rows(r).Cells(1).Value IsNot Nothing And _
                                        .Rows(r).Cells(2).Value IsNot Nothing Then
                        CurProj.Riser(CurProj.RiserId).LatSupports.Add(CSng(.Rows(r).Cells(0).Value) * CFLength, _
                                                      CSng(.Rows(r).Cells(1).Value) * CFSpringConst, _
                                                      CSng(.Rows(r).Cells(2).Value) * CFDamperConst)
                    End If
                Next
            End If
        End With
        '   confirm a successful updating
        UpdateSupSpring = True
        '   reset the LeaveCell trigger
        CheckingGrid = False

    End Function

    Private Sub LoadGridFromProject()

        Dim r, c As Short
        Dim LatSupport1 As LatSupport
        ' avoid triggering LeaveCell event

        CheckingGrid = True

        With grdSupSpring

            .RowCount = CurProj.Riser(CurProj.RiserId).LatSupports.Count + 1
            .ColumnCount = 3

            For c = 0 To .ColumnCount - 1
                .Columns(c).HeaderText = RiserLabels(CurUnits, c + 1)
                .Columns(c).Width = 79
            Next c
            .Columns(0).Width = 87
            r = 0

            For Each LatSupport1 In CurProj.Riser(CurProj.RiserId).LatSupports
                If r > MaxNumLatSupports Then Exit For
                .Rows(r).Cells(0).Value = CStr(LatSupport1.Dist / CFLength)
                .Rows(r).Cells(1).Value = CStr(LatSupport1.Stiffness / CFSpringConst)
                .Rows(r).Cells(2).Value = CStr(LatSupport1.Damping / CFDamperConst)

                r = r + 1
            Next LatSupport1

        End With
        NumberAllRows(grdSupSpring)
        CheckingGrid = False
    End Sub

  

End Class