Option Strict Off
Option Explicit On

Imports VB = Microsoft.VisualBasic

Friend Class frmTabulateResults

    Inherits System.Windows.Forms.Form

    Public MaxTabH, NumPoints, MaxTabW As Short

    Public ScrBar As Boolean
    Private ReportTitle As String
    Private ReportOption As Short


    Private Sub frmTabulateResults_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        Text = Text & " : " & CurProj.RiserId & " - " & CurProj.Title
        ' frmMainMenu.Cursor = System.Windows.Forms.Cursors.WaitCursor
        NumPoints = CurProj.NumPoints
        PeriodFatigueReport(CurProj.RiserId)
        '  frmMainMenu.Cursor = System.Windows.Forms.Cursors.Default

    End Sub


    Private Sub PeriodFatigueReport(ByVal riserid As Short)

        Dim r, c As Short
        Dim CFStress, CFLength As Single
        Dim ColFmtStr(0 To 6) As String
        ColFmtStr(0) = "Mode"
        ColFmtStr(1) = "Period(sec)"
        ColFmtStr(2) = "Initial Freq(Hz)"
        ColFmtStr(3) = "Coupled Freq(Hz)"
        ColFmtStr(4) = "Max Amplitude"
        ColFmtStr(5) = "Fatigue Loc"
        ColFmtStr(6) = "Fatigue Life(yr)"
        ReportTitle = "Fatigue Report"

        If CurProj.Units = VIVAMain.Units.English Then
            CFStress = (In2Ft * Ft2M) ^ 2 / Lb2N
            CFLength = 1.0# / Ft2M
            ColFmtStr(4) = ColFmtStr(4) & "(ft)"
            ColFmtStr(5) = ColFmtStr(5) & "(ft)"
        Else
            CFStress = 0.000001
            CFLength = 1.0#
            ColFmtStr(4) = ColFmtStr(4) & "(m)"
            ColFmtStr(5) = ColFmtStr(5) & "(m)"
        End If
        With grdResults
            .RowCount = CurProj.Riser(riserid).NumOutputModes + 2
            .ColumnCount = 7
            For c = 0 To 6
                .Columns(c).HeaderText = ColFmtStr(c)
                .Columns(c).Width = 90
            Next
            .Columns(0).Width = 56
            '       Mode Number
            For r = 1 To CurProj.Riser(riserid).NumOutputModes
                .Rows(r - 1).Cells(0).Value = CStr(CurProj.Riser(riserid).ModeNumber(r))
                .Rows(r - 1).Cells(1).Value = CStr(Format(CurProj.Riser(riserid).FatiguePeriod(CurProj.Riser(riserid).ModeNumber(r)), "0.00E-00"))
                .Rows(r - 1).Cells(2).Value = CStr(Format(CurProj.Riser(riserid).InitialFreqs(CurProj.Riser(riserid).ModeNumber(r)), "0.00E-00"))
                .Rows(r - 1).Cells(3).Value = CStr(Format(CurProj.Riser(riserid).CoupledFreqs(CurProj.Riser(riserid).ModeNumber(r)), "0.00E-00"))
                .Rows(r - 1).Cells(4).Value = CStr(Format(CurProj.Riser(riserid).MaxA(CurProj.Riser(riserid).ModeNumber(r)) * CFLength, "0.000"))

                .Rows(r - 1).Cells(5).Value = CStr(Format(CurProj.Riser(riserid).MaxStressPos(CurProj.Riser(riserid).ModeNumber(r)) * CFLength, "#,##0.00"))
                .Rows(r - 1).Cells(6).Value = CStr(Format(CurProj.Riser(riserid).FatigueLife(CurProj.Riser(riserid).ModeNumber(r)), "0.00E-00"))
            Next
            .Rows(CurProj.Riser(riserid).NumOutputModes).Cells(0).Value = "MM"
            .Rows(CurProj.Riser(riserid).NumOutputModes).Cells(1).Value = "---"
            .Rows(CurProj.Riser(riserid).NumOutputModes).Cells(2).Value = "---"
            .Rows(CurProj.Riser(riserid).NumOutputModes).Cells(3).Value = "---"
            'MsgBox(MaxA(0) & "___" & CFLength)
            .Rows(CurProj.Riser(riserid).NumOutputModes).Cells(4).Value = CStr(Format(CurProj.Riser(riserid).MaxA(0) * CFLength, "0.000"))
            .Rows(CurProj.Riser(riserid).NumOutputModes).Cells(5).Value = CStr(Format(CurProj.Riser(riserid).MaxStressPos(CurProj.Riser(riserid).ModeNumber(0)) * CFLength, "#,##0.00"))
            .Rows(CurProj.Riser(riserid).NumOutputModes).Cells(6).Value = CStr(Format(CurProj.Riser(riserid).FatigueLife(CurProj.Riser(riserid).ModeNumber(0)), "0.00E-00"))
        End With

    End Sub

    Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click

        Me.Close()

    End Sub

    Private Sub save_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) _
    Handles Save.Click
        Dim saveFile As Boolean
        ' frmMainMenu.Cursor = System.Windows.Forms.Cursors.WaitCursor
        If CurProj.nRisers = 1 Then
            If MM.Checked Then saveFile = ImportVIVACoreOutputFilesToExcel(1, 1)
            If AM.Checked Then saveFile = ImportVIVACoreOutputFilesToExcel(2, 1)
            If AC.Checked Then saveFile = ImportVIVACoreOutputFilesToExcel(3, 1)
        Else
            If MM.Checked Then saveFile = ImportVIVACoreOutputFilesToExcel(1, CurProj.RiserId)
            If AM.Checked Then saveFile = ImportVIVACoreOutputFilesToExcel(2, CurProj.RiserId)
            If AC.Checked Then saveFile = ImportVIVACoreOutputFilesToExcel(3, CurProj.RiserId)

        End If
        ' frmMainMenu.Cursor = System.Windows.Forms.Cursors.Default
    End Sub

End Class