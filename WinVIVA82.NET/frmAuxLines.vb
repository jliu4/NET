Option Strict Off
Option Explicit On

' Version 8.1
' 2014, Copyright DTCEL, All Rights Reserved
' Jin Liu May 2014

Friend Class frmAuxLines

    Inherits System.Windows.Forms.Form

    Private Changed As Boolean
    Private CheckingGrid As Boolean
    Private JEC As Boolean
    Private ExistingTxt As String
    Private CurUnits As VIVAMain.Units
    Private CFDim, CFContent As Single
    Private RiserLabels(1, 3) As String


    Private Sub grdAuxLines_CellValidating( _
          ByVal sender As Object, _
          ByVal e As System.Windows.Forms. _
          DataGridViewCellValidatingEventArgs) _
          Handles grdAuxLines.CellValidating

        grdAuxLines.Rows(e.RowIndex).ErrorText = ""

        If grdAuxLines.Rows(e.RowIndex).IsNewRow Then Return

        If Not IsNumeric(e.FormattedValue) Then
            If Not (IsNumeric(e.FormattedValue)) And Not String.IsNullOrEmpty(e.FormattedValue) Then
                e.Cancel = True
                MsgBox("It must be numeric value")
                grdAuxLines.Rows(e.RowIndex).ErrorText = _
                   "It must be a numeric value."
            End If
        End If
    End Sub

    Private Sub frmAuxLines_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        'disable the selection of units system in the global parameters form
        NumInputForms = NumInputForms + 1
        frmGlobalParms.btnMetric.Enabled = False
        frmGlobalParms.btnEnglish.Enabled = False

        'initiate variables
        CheckingGrid = False
        Me.Text = Me.Text & " : " & CurProj.RiserId & " - " & CurProj.Title
        CurUnits = CurProj.Units

        'load grid
        SetLabels()
        IniConvertFactor()
        LoadGridFromProject()

        'check value of checkbox
        If CurProj.Riser(CurProj.RiserId).AuxLines.AuxInEICalcs Then
            chkUseInEICalc.CheckState = System.Windows.Forms.CheckState.Checked
        Else
            chkUseInEICalc.CheckState = System.Windows.Forms.CheckState.Unchecked
        End If

    End Sub


    Private Sub frmAuxLines_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        'determine whether to enable the selection of units system
        NumInputForms = NumInputForms - 1

        If NumInputForms = 0 And UnitsChangeable Then
            frmGlobalParms.btnMetric.Enabled = True
            frmGlobalParms.btnEnglish.Enabled = True
        End If

    End Sub


    Private Sub SetLabels()

        'Metric unit labels
        RiserLabels(VIVAMain.Units.Metric, 0) = "No"
        RiserLabels(VIVAMain.Units.Metric, 1) = "Outer Dia. (mm)"
        RiserLabels(VIVAMain.Units.Metric, 2) = "Inner Dia. (mm)"
        RiserLabels(VIVAMain.Units.Metric, 3) = "Content Density (kg/m^3)"

        'English unit labels
        RiserLabels(VIVAMain.Units.English, 0) = "No"
        RiserLabels(VIVAMain.Units.English, 1) = "Outer Dia. (in)"
        RiserLabels(VIVAMain.Units.English, 2) = "Inner Dia. (in)"
        RiserLabels(VIVAMain.Units.English, 3) = "Content Density (ppg)"

    End Sub


    Private Sub IniConvertFactor()

        If CurUnits = VIVAMain.Units.Metric Then
            CFDim = mm2M
            CFContent = 1.0#
        Else
            CFDim = In2Ft * Ft2M
            CFContent = Lb2Kg / Gal2M3
        End If

    End Sub


    Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click

        Me.Close()

    End Sub


    Private Sub btnOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnOK.Click

        Cursor = System.Windows.Forms.Cursors.WaitCursor

        'check the value of checkbox and update curproj
        If chkUseInEICalc.CheckState = System.Windows.Forms.CheckState.Checked Then
            If CurProj.Riser(CurProj.RiserId).AuxLines.AuxInEICalcs = False Then
                CurProj.Riser(CurProj.RiserId).AuxLines.AuxInEICalcs = True
                CurProj.Saved = False
            End If
        Else
            If CurProj.Riser(CurProj.RiserId).AuxLines.AuxInEICalcs = True Then
                CurProj.Riser(CurProj.RiserId).AuxLines.AuxInEICalcs = False
                CurProj.Saved = False
            End If
        End If

        'update auxline and unload
        If UpdateAuxLines() Then
            CurProj.Saved = False
            Cursor = System.Windows.Forms.Cursors.Default
            Me.Close()
        Else
            Cursor = System.Windows.Forms.Cursors.Default
        End If

    End Sub


    Public Sub mnuAddRow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAddRow.Click

        AddRow(grdAuxLines, CheckingGrid)
    End Sub

   

    Public Sub mnuDeleteRow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDeleteRow.Click

        DeleteRow(grdAuxLines, CheckingGrid)

    End Sub

    Public Sub paste_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles pasteFromClipboard.Click

        PastefromClipBoardToDataGridView(grdAuxLines, CheckingGrid, MaxAuxLines)

    End Sub

    Public Sub copyToExcel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles copyToExcel.Click

        copyToClipBoard(grdAuxLines)

    End Sub

    Private Function UpdateAuxLines() As Boolean

        Dim r, c As Short
        Dim NumAuxLines As Short
        Dim NewAuxLine As AuxLine

        For r = 0 To grdAuxLines.RowCount - 2
            For c = 0 To grdAuxLines.ColumnCount - 1
                If String.IsNullOrEmpty(grdAuxLines(c, r).Value) Then
                    MsgBox("No blank in all cells")
                    CheckingGrid = False
                    UpdateAuxLines = False
                    Exit Function
                End If
            Next
        Next
        CheckingGrid = True
        With grdAuxLines
    
            'delete the old inputs
            NumAuxLines = CurProj.Riser(CurProj.RiserId).AuxLines.Count

            For r = 1 To NumAuxLines
                CurProj.Riser(CurProj.RiserId).AuxLines.Delete(1)
            Next

            If .RowCount > 1 Then
                For r = 0 To .RowCount - 2
                    If .Rows(r).Cells(0).Value IsNot Nothing And _
                    .Rows(r).Cells(1).Value IsNot Nothing And _
                    .Rows(r).Cells(2).Value IsNot Nothing Then
                        NewAuxLine = CurProj.Riser(CurProj.RiserId).AuxLines.Add
                        NewAuxLine.OD = CSng(.Rows(r).Cells(0).Value) * CFDim
                        NewAuxLine.ID = CSng(.Rows(r).Cells(1).Value) * CFDim
                        NewAuxLine.ContentDensity = CSng(.Rows(r).Cells(2).Value) * CFContent
                    End If
                Next
            End If
          
        End With

        'confirm a successful updating
        UpdateAuxLines = True
        CheckingGrid = False

    End Function


    Private Sub LoadGridFromProject()

        Dim r, c As Short
        Dim AuxLine1 As AuxLine

        CheckingGrid = True

        With grdAuxLines
            .RowCount = CurProj.Riser(CurProj.RiserId).AuxLines.Count + 1
            .ColumnCount = 3

            For c = 0 To .ColumnCount - 1
                .Columns(c).HeaderText = RiserLabels(CurUnits, c + 1)
                .Columns(c).Width = 66
            Next c
            '.Columns(0).Width = 60
            r = 0

            For Each AuxLine1 In CurProj.Riser(CurProj.RiserId).AuxLines
                If r > MaxAuxLines Then Exit For
                .Rows(r).Cells(0).Value = CStr(AuxLine1.OD / CFDim)
                .Rows(r).Cells(1).Value = CStr(AuxLine1.ID / CFDim)
                .Rows(r).Cells(2).Value = CStr(AuxLine1.ContentDensity / CFContent)

                r = r + 1
            Next AuxLine1
            NumberAllRows(grdAuxLines)
        End With

        CheckingGrid = False

    End Sub

End Class