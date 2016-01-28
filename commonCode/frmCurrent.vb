Option Strict Off
Option Explicit On

Friend Class frmCurrent

    Inherits System.Windows.Forms.Form

    Dim VelUColor As Color = Color.Red
    Dim VelVColor As Color = Color.Lime
    Dim VelAColor As Color = Color.Blue
    Private NumPoints As Short
    Private ExistingTxt As String
    Private Changed As Boolean
    Private CheckingGrid As Boolean
    Private JEC As Boolean 'JEC - just entered cell
    Private Depth(MaxCurrentPair) As Single
    Private CurVelU(MaxCurrentPair) As Single
    Private CurVelV(MaxCurrentPair) As Single
    Private CurVelA(MaxCurrentPair) As Single
    Private Vmax, Dmax, Dmin, Vmin As Single
    Private CFLength, CFVel As Single
    Private CurrentLbls(3) As String
    Private tmpFileNum As Short


    Private Sub frmCurrent_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        Dim i As Short

        CheckingGrid = True
        Text = Text & " - " & CurProj.Title

        If Not IsMetricUnit Then
            VelFactor = 1
            LFactor = 1

            CurrentLbls(0) = "Depth (ft)"
            CurrentLbls(1) = "Vel Y (kn)"
            CurrentLbls(2) = "Vel Z (kn)"
            CurrentLbls(3) = "Vel (kn)"
        Else
            LFactor = 0.3048 ' ft -> m
            VelFactor = 0.5144444 ' knots - > m/s

            CurrentLbls(0) = "Depth (m)"
            CurrentLbls(1) = "Vel Y (m/s)"
            CurrentLbls(2) = "Vel Z (m/s)"
            CurrentLbls(3) = "Vel (m/s)"
        End If

        With grdCurrent
            '       initiate grid
            .RowCount = MaxCurrentPair + 1
            .ColumnCount = 3
            For i = 0 To .ColumnCount - 1
                .Columns(i).HeaderText = CurrentLbls(i)
                .Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Columns(i).FillWeight = 100 / .ColumnCount

            Next

            '       load data
            If CurVessel.EnvLoad.EnvCur.Current.ProfileCount = 0 Then
                .Text = "0"
            Else
                LoadGridFromProject()
                '  UpdatePlotFromGrid()
            End If
        End With

        CheckingGrid = False

    End Sub

    Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click

        Me.Close()

    End Sub

    Private Sub btnOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnOK.Click

        Cursor = System.Windows.Forms.Cursors.WaitCursor
        UpdateCurrentProfile()
        Cursor = System.Windows.Forms.Cursors.Default

        frmEnviron._txtCurr_0.Text = CStr(CurVessel.EnvLoad.EnvCur.Current.Profile(1).Velocity * Ftps2Knots * VelFactor)

        Me.Close()

    End Sub

    Private Sub CellLeave(ByVal eventSender As System.Object, ByVal e As DataGridViewCellEventArgs) _
    Handles grdCurrent.CellValueChanged
        If Not CheckingGrid Then
            With grdCurrent
                'UpdatePlotFromGrid()
                'ReDrawing()
            End With
        End If
    End Sub

    Private Sub RowsRemoved(ByVal eventSender As System.Object, ByVal e As DataGridViewRowsRemovedEventArgs) _
  Handles grdCurrent.RowsRemoved
        ' If Not CheckingGrid Then
        With grdCurrent

            'UpdatePlotFromGrid()
            'ReDrawing()
        End With
        ' End If

    End Sub

    Private Sub grdCurrent_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) _
    Handles MyBase.MouseDown

        If eventArgs.Button = System.Windows.Forms.MouseButtons.Right Then

            ContextMenuStripGridEdit.Show(MousePosition)

        End If

    End Sub


    Public Sub mnuAddRow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAddRow.Click
        'insert a blank row
        AddRow(grdCurrent, CheckingGrid)

    End Sub


    Public Sub mnuDeleteRow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDeleteRow.Click
        DeleteRow(grdCurrent, CheckingGrid)

    End Sub


    Public Sub mnuExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuExit.Click

        Me.Close()

    End Sub


    Private Sub UpdateCurrentProfile()

        Dim r, NumPairs As Short
        Dim DepthText As String
        Dim VelU As String
        Dim VelV As String
        Dim VelA As String

        '   avoid triggering the "Leave Cell" event
        CheckingGrid = True

        '   delete previous input
        NumPairs = CurVessel.EnvLoad.EnvCur.Current.ProfileCount

        '    Debug.Print "NumPairs= " & NumPairs
        For r = NumPairs To 1 Step -1
            CurVessel.EnvLoad.EnvCur.Current.ProfileDelete((1)) 'After deleting first element, second one will become the "first" one
        Next r

        '   insert current input
        With grdCurrent
            For r = 0 To .RowCount - 1

                DepthText = .Rows(r).Cells(0).Value
                VelU = .Rows(r).Cells(1).Value
                VelV = .Rows(r).Cells(2).Value
                VelA = CStr(System.Math.Sqrt(CSng(VelU) ^ 2 + CSng(VelV) ^ 2))

                CurVessel.EnvLoad.EnvCur.Current.ProfileAdd(Depth:=CSng(DepthText), Velocity:=CSng(VelA))

            Next r
        End With

        '   remember that a change has been made
        CurProj.Saved = False
        '   reset the "Leaving Cell" trigger
        CheckingGrid = False

    End Sub

    Private Sub UpdatePlotFromGrid()

        Dim r As Short

        CheckingGrid = True
        NumPoints = 0
        With grdCurrent
            If .RowCount > 1 Then
                For r = 0 To .RowCount - 1

                    Depth(r + 1) = CSng(.Rows(r).Cells(0).Value)

                    CurVelU(r + 1) = CSng(.Rows(r).Cells(1).Value)

                    NumPoints = NumPoints + 1
                    CurVelA(r + 1) = System.Math.Sqrt(CurVelU(r + 1) ^ 2 + CurVelV(r + 1) ^ 2)

                Next
            End If
        End With

        ReDrawing()

        CheckingGrid = False

    End Sub
    Private Sub ReDrawing()
        Dim BoxWidth, BoxHeight, BoxSize As Integer

        BoxWidth = picCurrent.ClientSize.Width
        BoxHeight = picCurrent.ClientSize.Height
        If BoxWidth < BoxHeight Then
            BoxSize = BoxWidth
        Else
            BoxSize = BoxHeight
        End If

        Dim bm As Bitmap = New Bitmap(BoxWidth, BoxHeight)
        Dim gr As Graphics = Graphics.FromImage(bm)
        Dim pen As New System.Drawing.Pen(Color.Black, 1)
        Dim i As Short

        For i = 1 To CurVelA.Length - 1
            'UPGRADE_ISSUE: PictureBox method picGraph.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            gr.DrawLine(pen, CurVelA(i), Depth(i), CurVelA(i + 1), Depth(i + 1))
        Next i
        Dim dest_bounds As New RectangleF(0, 0, BoxWidth, BoxHeight)
        Dim source_bounds As New RectangleF(0, 0, BoxWidth + 1, BoxHeight + 1)
        gr.DrawImage(bm, dest_bounds, source_bounds, GraphicsUnit.Pixel)
        picCurrent.SizeMode = PictureBoxSizeMode.AutoSize
        picCurrent.Image = bm
        pen.Dispose()
        gr.Dispose()

    End Sub

    Private Sub LoadGridFromProject()
        Dim NumPairs As Short
        Dim r As Short

        CheckingGrid = True

        NumPairs = CurVessel.EnvLoad.EnvCur.Current.ProfileCount

        With grdCurrent
            For r = 0 To NumPairs - 1
                .Rows(r).Cells(0).Value = CStr(FormatNumber(CurVessel.EnvLoad.EnvCur.Current.Profile(r + 1).Depth * LFactor, 1))
                .Rows(r).Cells(1).Value = CStr(FormatNumber(CurVessel.EnvLoad.EnvCur.Current.Profile(r + 1).Velocity * Ftps2Knots * VelFactor, 3))
            Next r

        End With

        CheckingGrid = False
    End Sub


    Public Sub mnuFileOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileOpen.Click

        Dim Fn As String

        'should the user cancel the dialog box, exit
        On Error GoTo ErrHandler

        'set filters to allow selection of all files or just .VIV
        dlgCurrentProfileOpen.Filter = "All Files (*.*)|*.*|Current Files (*.cur)|*.cur"
        dlgCurrentProfileSave.Filter = "All Files (*.*)|*.*|Current Files (*.cur)|*.cur"

        'specify default filter as *.CUR
        dlgCurrentProfileOpen.FilterIndex = 2
        dlgCurrentProfileSave.FilterIndex = 2
        dlgCurrentProfileOpen.ShowReadOnly = False

        'display the Open dialog box
        dlgCurrentProfileOpen.ShowDialog()
        dlgCurrentProfileSave.FileName = dlgCurrentProfileOpen.FileName

        'open the file
        tmpFileNum = FreeFile()
        FileOpen(tmpFileNum, dlgCurrentProfileOpen.FileName, OpenMode.Input, OpenAccess.Read)

        'build the object from the input data
        Cursor = System.Windows.Forms.Cursors.WaitCursor
        InputCurrentProfile((tmpFileNum))
        UpdatePlotFromGrid()

        Cursor = System.Windows.Forms.Cursors.Default

        'close the input file and return
        FileClose(tmpFileNum)

        Exit Sub

ErrHandler:
        'user pressed Cancel button
        Exit Sub

    End Sub


    Public Sub mnuFileSaveAs_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileSaveAs.Click

        UpdateCurrentProfile()

        '   should the user cancel the dialog box, exit
        On Error GoTo ErrHandler

        '   set filters to allow selection of all files or just .CUR
        dlgCurrentProfileOpen.Filter = "All Files (*.*)|*.*|Current Files (*.cur)|*.cur"
        dlgCurrentProfileSave.Filter = "All Files (*.*)|*.*|Current Files (*.cur)|*.cur"

        dlgCurrentProfileOpen.FilterIndex = 2
        dlgCurrentProfileSave.FilterIndex = 2
        '   specify the current filename
        dlgCurrentProfileOpen.FileName = "*.cur"
        dlgCurrentProfileSave.FileName = "*.cur"
        dlgCurrentProfileOpen.ShowReadOnly = False
        dlgCurrentProfileSave.OverwritePrompt = True
        '   display the Open dialog box
        dlgCurrentProfileSave.ShowDialog()
        dlgCurrentProfileOpen.FileName = dlgCurrentProfileSave.FileName
        '   open the file for output
        tmpFileNum = FreeFile
        FileOpen(tmpFileNum, dlgCurrentProfileOpen.FileName, OpenMode.Output)

        Cursor = System.Windows.Forms.Cursors.WaitCursor
        '   output
        OutputCurrentProfile((tmpFileNum))
        Cursor = System.Windows.Forms.Cursors.Default
        '   close the output file
        FileClose(tmpFileNum)

        Exit Sub

ErrHandler:
        '   user pressed Cancel button
        Exit Sub

    End Sub

    Private Sub OutputCurrentProfile(ByRef FileNum As Short)

        Dim r, NumPairs As Short
        '   save the form data in the Project object
        With CurVessel.EnvLoad.EnvCur.Current
            NumPairs = .ProfileCount
            '       simple output; profile name, then the depth/current pairs
            For r = 1 To NumPairs
                WriteLine(FileNum, .Profile(r).Depth, .Profile(r).Velocity)
            Next r
        End With

    End Sub

    Private Sub InputCurrentProfile(ByRef FileNum As Short)

        Dim Depth, VelU As Single

        CheckingGrid = True

        Do While Not EOF(FileNum)
            Input(FileNum, Depth)
            Input(FileNum, VelU)
            With grdCurrent

                Dim tmp As String() = {CStr(Depth * LFactor),
                CStr(VelU * Ftps2Knots * VelFactor)}
                .Rows.Add(tmp)

            End With
        Loop

        CheckingGrid = False

    End Sub

End Class