Option Strict Off
Option Explicit On
Friend Class frmCatenary
	Inherits System.Windows.Forms.Form

    Private SegLength(MaxNumSeg) As Single
    Private SegTension(MaxNumSeg) As Single
    Private SegAngle(MaxNumSeg) As Single
    Private SegPosition(MaxNumSeg) As Single
    Private JustEnter As Boolean
	Private ExistingTxt As String
	Private LastChanged As Short '1 - topten; 2 - horfrc
	Private InitiateCbo As Boolean

    Private LUnit, FrcUnit As String
    Private LFactor, FrcFactor As Single
    Private cboSegment As New DataGridViewComboBoxCell

    Private cboSegmentColIndex As Short = 2
    Private drawingX As Short
    Private drawingY As Short
    Private curRow As Short
    Private curCol As Short

    ' form load and unload

    Private Sub frmCatenary_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        If IsMetricUnit Then
            LFactor = 0.3048 ' ft -> m
            FrcFactor = 4.448222 ' kips -> KN
            LUnit = " (m)"
            FrcUnit = " (KN)"
        Else
            LFactor = 1
            FrcFactor = 1
            LUnit = " (ft)"
            FrcUnit = " (kips)"
        End If

        Dim r, c As Short

        Dim LengthLabel(4) As String
        Dim DetailLabelC(3) As String
        Dim DetailLabelR(5) As String

        LengthLabel(0) = "Scope" & LUnit
        LengthLabel(1) = "Total Length" & LUnit
        LengthLabel(2) = "Suspended Length" & LUnit
        LengthLabel(3) = "Grounded Length" & LUnit
        LengthLabel(4) = "Stretched Length" & LUnit

        DetailLabelC(1) = "Top"
        DetailLabelC(2) = "Segment"
        DetailLabelC(3) = "Bottom"
        DetailLabelR(1) = "Payout" & LUnit
        DetailLabelR(2) = "Tension" & FrcUnit
        DetailLabelR(3) = "Hor. Force" & FrcUnit
        DetailLabelR(4) = "Angle (deg)"
        DetailLabelR(5) = "Depth at Top" & LUnit
        'set drawing area for the form
        drawingX = Me.Size.Width
        drawingY = grdLength.Location.Y - 10

        With grdLength
            .RowCount = 1
            .ColumnCount = 5
            For c = 0 To .ColumnCount - 1
                .Columns(c).FillWeight = 100 / .ColumnCount
                .Columns(c).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Columns(c).HeaderText = LengthLabel(c)
            Next
        End With

        With grdDetails
            .ColumnCount = 4
            .RowCount = 6

            For c = 0 To .ColumnCount - 1
                .Columns(c).FillWeight = 100 / .ColumnCount
                .Columns(c).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                If c = 2 Then
                    If .Rows(0).Cells(2) IsNot cboSegment Then
                        .Rows(0).Cells(2) = cboSegment
                    End If
                Else
                    .Rows(0).Cells(c).Value = DetailLabelC(c)
                End If
            Next c

            For r = 1 To .RowCount - 1
                .Rows(r).Cells(0).Value = DetailLabelR(r)
            Next
            .Rows(1).Cells(cboSegmentColIndex).Value = Format(SegLength(1) * LFactor, "0.0")
            .Rows(2).Cells(cboSegmentColIndex).Value = Format(SegTension(1) * 0.001 * FrcFactor, "0.00")
            .Rows(3).Cells(cboSegmentColIndex).Value = Format(SegTension(1) * 0.001 * System.Math.Cos(SegAngle(1)) * FrcFactor, "0.00")
            .Rows(4).Cells(cboSegmentColIndex).Value = Format(SegAngle(1) * Radians2Degrees, "0.00")
            .Rows(5).Cells(cboSegmentColIndex).Value = Format(SegPosition(1) * LFactor, "0.0")
        End With
        JustEnter = True
    End Sub

    ' buttons

    Private Sub btnOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnOK.Click

        Me.Close()

    End Sub

    Private Sub btnRefresh_Click()

        If IsMetricUnit Then
            LFactor = 0.3048 ' ft -> m
            FrcFactor = 4.448222 ' kips -> KN
            LUnit = " (m)"
            FrcUnit = " (KN)"
        Else
            LFactor = 1
            FrcFactor = 1
            LUnit = " (ft)"
            FrcUnit = " (kips)"
        End If

        Dim TopTension, HorForce As Single

        Select Case LastChanged
            Case 1
                TopTension = grdDetails.Rows(2).Cells(1).Value * 1000.0# / FrcFactor

                Call frmMoorLines.UpdateCat(TopTension)
                frmMoorLines.ChangeInCat = True
            Case 2
                TopTension = 0#
                HorForce = grdDetails.Rows(3).Cells(1).Value * 1000.0# / FrcFactor

                Call frmMoorLines.UpdateCat(TopTension, HorForce)
                frmMoorLines.ChangeInCat = True
        End Select

    End Sub

    ' combo box
    Private Sub cboLines_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboLines.SelectedIndexChanged

        If InitiateCbo Then Exit Sub
        With frmMoorLines
            .tabMoorLines.SelectedIndex = cboLines.SelectedIndex
            Call .UpdateCat()
        End With

    End Sub

    Private Sub cboSegment_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)

        If IsMetricUnit Then
            LFactor = 0.3048 ' ft -> m
            FrcFactor = 4.448222 ' kips -> KN
            LUnit = " (m)"
            FrcUnit = " (KN)"
        Else
            LFactor = 1
            FrcFactor = 1
            LUnit = " (ft)"
            FrcUnit = " (kips)"
        End If

        Dim ID As Short = DirectCast(sender, ComboBox).SelectedIndex
        With grdDetails
            .Rows(1).Cells(cboSegmentColIndex).Value = Format(SegLength(ID + 1) * LFactor, "0.0")
            .Rows(2).Cells(cboSegmentColIndex).Value = Format(SegTension(ID + 1) * 0.001 * FrcFactor, "0.00")
            .Rows(3).Cells(cboSegmentColIndex).Value = Format(SegTension(ID + 1) * 0.001 * System.Math.Cos(SegAngle(ID + 1)) * FrcFactor, "0.00")
            .Rows(4).Cells(cboSegmentColIndex).Value = Format(SegAngle(ID + 1) * Radians2Degrees, "0.00")
            .Rows(5).Cells(cboSegmentColIndex).Value = Format(SegPosition(ID + 1) * LFactor, "0.0")
        End With
    End Sub

    ' update data

    Public Sub UpdateForm(ByVal NumLines As Short, ByVal CurLine As Short, ByVal NumSegment As Short, ByVal Scope As Single, ByVal lineLength As Single, ByVal SuspendLength As Single, ByVal GroundedLength As Single, ByVal StretchLength As Single, ByVal TopLength As Single, ByVal TopTension As Single, ByVal BottomTension As Single, ByVal TopAngle As Single, ByVal BottomAngle As Single, ByVal TopPosition As Single, ByRef SegmentLength() As Single, ByRef SegmentTension() As Single, ByRef SegmentAngle() As Single, ByRef SegmentPosition() As Single, ByRef CatX() As Single, ByRef CatY() As Single, ByRef Connector() As Short)

        Dim SegMaxTen, i, j, NumPoints As Short
        Dim MaxTen As Single

        Dim Xmin, Xmax, Ymax, Ymin As Single
        Dim Color(3) As Integer
        Dim X(MaxNumSubSeg + 1) As Single
        Dim Y(MaxNumSubSeg + 1) As Single

        If IsMetricUnit Then
            LFactor = 0.3048 ' ft -> m
            FrcFactor = 4.448222 ' kips -> KN
            LUnit = " (m)"
            FrcUnit = " (KN)"
        Else
            LFactor = 1
            FrcFactor = 1
            LUnit = " (ft)"
            FrcUnit = " (kips)"
        End If

        MaxTen = 0#
        For i = 1 To NumSegment
            If i > MaxNumSeg Then Exit For
            SegLength(i) = SegmentLength(i)
            SegTension(i) = SegmentTension(i)
            SegAngle(i) = SegmentAngle(i)
            SegPosition(i) = SegmentPosition(i)

            If SegmentTension(i) > MaxTen Then
                MaxTen = SegmentTension(i)
                SegMaxTen = i
            End If
        Next i

        InitiateCbo = True
        With cboLines
            .Items.Clear()
            For i = 1 To NumLines
                .Items.Add("Line " & i) 'TODO JLIU, it should be the name of line
            Next i
            .SelectedIndex = CurLine - 1
        End With
        InitiateCbo = False

        With cboSegment
            .Items.Clear()
            For i = 1 To NumSegment
                .Items.Add("Segment " & i)
            Next i
            .Value = .Items(SegMaxTen - 1).ToString
        End With

        With grdLength
            .RowCount = 1
            .ColumnCount = 5
            .Rows(0).Cells(0).Value = Format(Scope * LFactor, "0.0")
            .Rows(0).Cells(1).Value = Format(lineLength * LFactor, "0.0")
            .Rows(0).Cells(2).Value = Format(SuspendLength * LFactor, "0.0")
            .Rows(0).Cells(3).Value = Format(GroundedLength * LFactor, "0.0")
            .Rows(0).Cells(4).Value = Format(StretchLength * LFactor, "0.0")
        End With

        With grdDetails
            .ColumnCount = 4
            .RowCount = 6
            .Rows(1).Cells(1).Value = Format(TopLength * LFactor, "0.0")
            .Rows(2).Cells(1).Value = Format(TopTension * 0.001 * FrcFactor, "0.00")
            .Rows(3).Cells(1).Value = Format(TopTension * 0.001 * System.Math.Cos(TopAngle) * FrcFactor, "0.00")
            .Rows(4).Cells(1).Value = Format(TopAngle * Radians2Degrees, "0.00")
            .Rows(5).Cells(1).Value = Format(TopPosition * LFactor, "0.0")
            .Rows(2).Cells(3).Value = Format(BottomTension * 0.001 * FrcFactor, "0.00")
            .Rows(3).Cells(3).Value = Format(BottomTension * 0.001 * System.Math.Cos(BottomAngle) * FrcFactor, "0.00")
            .Rows(4).Cells(3).Value = Format(BottomAngle * Radians2Degrees, "0.00")

            'JLIU TODO should trig the select index call
            With grdDetails
                .Rows(1).Cells(cboSegmentColIndex).Value = Format(SegLength(SegMaxTen) * LFactor, "0.0")
                .Rows(2).Cells(cboSegmentColIndex).Value = Format(SegTension(SegMaxTen) * 0.001 * FrcFactor, "0.00")
                .Rows(3).Cells(cboSegmentColIndex).Value = Format(SegTension(SegMaxTen) * 0.001 * System.Math.Cos(SegAngle(SegMaxTen)) * FrcFactor, "0.00")
                .Rows(4).Cells(cboSegmentColIndex).Value = Format(SegAngle(SegMaxTen) * Radians2Degrees, "0.00")
                .Rows(5).Cells(cboSegmentColIndex).Value = Format(SegPosition(SegMaxTen) * LFactor, "0.0")
            End With

        End With

        Xmax = CatX(1) * LFactor
        Xmin = 0#
        Ymax = Max(CatY(1) * LFactor, 0#)
        Ymin = 0#
        Color(1) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue)
        Color(2) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
        Color(3) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Lime)
        Color(0) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
        'find drawing scales

        'Call drawAxis0(Xmax, Xmin, Ymax, Ymin, "Distance" & LUnit, "Depth" & LUnit, picCatenary, False)

        FileClose(FileNumRes)
        '    Open MarsDir & "catenary.dat" For Output Access Write As #FileNumRes
        FileOpen(FileNumRes, MarsDir & CurLine & "catenary.dat", OpenMode.Output, OpenAccess.Write)
        For i = 1 To NumSegment
            NumPoints = Connector(i) - Connector(i + 1) + 1
            For j = 1 To NumPoints
                X(j) = CatX(Connector(i) - j + 1) * LFactor
                Y(j) = CatY(Connector(i) - j + 1) * LFactor

                If j = NumPoints Then
                    WriteLine(FileNumRes, X(j), Y(j), "<-End Segment " & i)
                Else
                    WriteLine(FileNumRes, X(j), Y(j))
                End If
            Next j
            'Call DrawLine(X, Y, NumPoints, 0, picCatenary)
        Next i

        FileClose(FileNumRes)

    End Sub

    Private Sub grdDetails_EditingControlShowing(ByVal sender As System.Object, ByVal e As DataGridViewEditingControlShowingEventArgs) _
        Handles grdDetails.EditingControlShowing

        curRow = grdDetails.CurrentCell.RowIndex
        curCol = grdDetails.CurrentCell.ColumnIndex

        If curRow = 0 And curCol = 2 Then
            Dim comboBox As ComboBox = TryCast(e.Control, ComboBox)
            AddHandler comboBox.SelectedIndexChanged, AddressOf cboSegment_SelectedIndexChanged
        End If

    End Sub

    Private Sub grdDetails_CellbeginEdit(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdDetails.CellBeginEdit

        Try
            curCol = grdDetails.CurrentCell.ColumnIndex
            curRow = grdDetails.CurrentCell.RowIndex
            If curCol = 1 And (curRow = 2 Or curRow = 3) Then

                If JustEnter Then
                    JustEnter = False
                    'ExistingTxt = grdDetails.CurrentCell.Value
                    VB6.SetCancel(btnOK, False)
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub grdDetails_CellEndEdit(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdDetails.CellEndEdit

        Try
            curCol = grdDetails.CurrentCell.ColumnIndex
            curRow = grdDetails.CurrentCell.RowIndex

            If curCol = 1 And (curRow = 2 Or curRow = 3) Then

                LastChanged = curRow - 1
                btnRefresh_Click()

            End If
        Catch ex As Exception

        End Try
    End Sub

End Class