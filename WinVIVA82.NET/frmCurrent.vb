Option Strict Off
Option Explicit On

Friend Class frmCurrent

    Inherits System.Windows.Forms.Form

    Dim VelUColor As Color = Color.Red
    Dim VelVColor As Color = Color.Lime
    Dim VelAColor As Color = Color.Blue
    Dim CurDrawScale As DrawScale
    Dim wmf_file_name As String

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
    Private CurrentLbls(3) As String 'jjx
    Private tmpFileNum As Short

    Private Sub grdCurrent_CellValidating( _
           ByVal sender As Object, _
           ByVal e As System.Windows.Forms. _
           DataGridViewCellValidatingEventArgs) _
          Handles grdCurrent.CellValidating

        grdCurrent.Rows(e.RowIndex).ErrorText = ""

        If grdCurrent.Rows(e.RowIndex).IsNewRow Then Return
        
        If Not IsNumeric(e.FormattedValue) Then
            If Not (IsNumeric(e.FormattedValue)) And Not String.IsNullOrEmpty(e.FormattedValue) Then
                e.Cancel = True
                grdCurrent.Rows(e.RowIndex).ErrorText = _
                   "It must be a numeric value."

            End If
        End If
    End Sub

    Private Sub frmCurrent_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed

        Try
            Kill(wmf_file_name)
        Catch ex As Exception
        End Try

    End Sub


    Private Sub frmCurrent_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        Dim i As Short
        'Dim ColW As Short

        CheckingGrid = True
        Text = Text & " - " & CurProj.Title

        If CurProj.Units = VIVAMain.Units.English Then
            CurrentLbls(0) = "Depth (ft)"
            CurrentLbls(1) = "Vel Y (kn)"
            CurrentLbls(2) = "Vel Z (kn)"
            CurrentLbls(3) = "Vel (kn)"
            CFLength = Ft2M
            CFVel = Kn2MPS
        Else
            CurrentLbls(0) = "Depth (m)"
            CurrentLbls(1) = "Vel Y (m/s)"
            CurrentLbls(2) = "Vel Z (m/s)"
            CurrentLbls(3) = "Vel (m/s)"
            CFLength = 1.0#
            CFVel = 1.0#
        End If

        With grdCurrent
            'initiate grid
            .RowCount = 1
            .ColumnCount = 3
            '.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            For i = 0 To .ColumnCount - 1
                .Columns(i).HeaderText = CurrentLbls(i)
            Next

            'load data
            If CurProj.Water.CurrentProfile.Count = 0 Then
                .Text = "0"
            Else
                LoadGridFromProject()
                txtProfileName.Text = CurProj.Water.CurrentProfile.ProfileName

            End If
        End With

        NumInputForms = NumInputForms + 1
        'frmGlobalParms.Enabled = False
        'frmGlobalParms.btnEnglish.Enabled = False

        'initializing drawing scale
        CurDrawScale.left = 0
        CurDrawScale.top = 0
        CurDrawScale.width = 1
        CurDrawScale.height = 1

        CheckingGrid = False

        Dim path_name As String = Environment.GetEnvironmentVariable("Temp", EnvironmentVariableTarget.User)
        wmf_file_name = RandomFileName(path_name, "tmp")
        'delete file if it exists
        If Len(Dir$(wmf_file_name)) > 0 Then Kill(wmf_file_name)

        UpdatePlotFromGrid()
        ReDrawing()

    End Sub


    Private Sub frmCurrent_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) _
        Handles Me.FormClosed

        'check whether enable units selection
        NumInputForms = NumInputForms - 1

        If NumInputForms = 0 And UnitsChangeable Then
            frmGlobalParms.btnMetric.Enabled = True
            frmGlobalParms.btnEnglish.Enabled = True
        End If

        Try
            Kill(wmf_file_name)
        Catch ex As Exception
        End Try

    End Sub


    Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click

        Me.Close()

    End Sub


    Private Sub btnOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnOK.Click

        Cursor = System.Windows.Forms.Cursors.WaitCursor
        UpdateCurrentProfile()
        Cursor = System.Windows.Forms.Cursors.Default
        Me.Close()

    End Sub
    Private Sub CellLeave(ByVal eventSender As System.Object, ByVal e As DataGridViewCellEventArgs) _
    Handles grdCurrent.CellValueChanged
        If Not CheckingGrid Then
            With grdCurrent
                UpdatePlotFromGrid()
                ReDrawing()
            End With
        End If
    End Sub

    Private Sub RowsRemoved(ByVal eventSender As System.Object, ByVal e As DataGridViewRowsRemovedEventArgs) _
  Handles grdCurrent.RowsRemoved
        ' If Not CheckingGrid Then
        With grdCurrent

            UpdatePlotFromGrid()
            ReDrawing()
        End With
        ' End If

    End Sub

    Private Sub frm_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) _
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
    Public Sub paste_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles pasteFromClipBoard.Click

        PastefromClipBoardToDataGridView(grdCurrent, CheckingGrid, MaxCurrentPair)

    End Sub

    Public Sub copyToExcel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles copyToExcel.Click

        copyToClipBoard(grdCurrent)

    End Sub


    Public Sub mnuExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuExit.Click

        Me.Close()

    End Sub


    Private Sub UpdateCurrentProfile()

        Dim r, c, NumPairs As Short
        Dim DepthText As String
        Dim VelU As String
        Dim VelV As String

        For r = 0 To grdCurrent.RowCount - 2
            For c = 0 To grdCurrent.ColumnCount - 1
                If String.IsNullOrEmpty(grdCurrent(c, r).Value) Then
                    MsgBox("No blank in all cells")
                    CheckingGrid = False

                    Exit Sub
                End If
            Next
        Next
        'avoid triggering the "Leave Cell" event
        CheckingGrid = True

        'reset the name of the current profile
        If txtProfileName.Text = "" Then
            txtProfileName.Text = CurProj.Title & " Current Profile"
        End If

        CurProj.Water.CurrentProfile.ProfileName = txtProfileName.Text

        'delete previous input
        NumPairs = CurProj.Water.CurrentProfile.Count

        'Debug.Print "NumPairs= " & NumPairs
        For r = 1 To NumPairs
            CurProj.Water.CurrentProfile.Delete(1) 'After deleting first element, second one will become the "first" one
        Next

        'insert current input
        With grdCurrent
            For r = 0 To .RowCount - 1
                If .Rows(r).Cells(0).Value IsNot Nothing And _
                .Rows(r).Cells(1).Value IsNot Nothing And _
                .Rows(r).Cells(2).Value IsNot Nothing Then
                    DepthText = .Rows(r).Cells(0).Value
                    VelU = .Rows(r).Cells(1).Value
                    VelV = .Rows(r).Cells(2).Value
                    CurProj.Water.CurrentProfile.Add(CDepth:=CSng(DepthText) * CFLength, CVelU:=CSng(VelU) * CFVel, CVelV:=CSng(VelV) * CFVel)
                End If
            Next
        End With

        'remember that a change has been made
        CurProj.Saved = False
        CheckingGrid = False

    End Sub


    Private Sub UpdatePlotFromGrid()

        Dim r As Short

        CheckingGrid = True
        NumPoints = 0
        With grdCurrent
            If .RowCount > 1 Then
                For r = 0 To .RowCount - 2
                    If .Rows(r).Cells(0).Value IsNot Nothing And _
                       .Rows(r).Cells(1).Value IsNot Nothing And _
                       .Rows(r).Cells(2).Value IsNot Nothing Then
                        Depth(r + 1) = CSng(.Rows(r).Cells(0).Value)

                        CurVelU(r + 1) = CSng(.Rows(r).Cells(1).Value)

                        CurVelV(r + 1) = CSng(.Rows(r).Cells(2).Value)
                        NumPoints = NumPoints + 1
                        CurVelA(r + 1) = System.Math.Sqrt(CurVelU(r + 1) ^ 2 + CurVelV(r + 1) ^ 2)

                    End If
                Next
            End If
        End With

        CheckingGrid = False

    End Sub


    Private Sub LoadGridFromProject()

        Dim Pair As CurrentPair
        Dim r As Short
        CheckingGrid = True
        With grdCurrent
            .RowCount = CurProj.Water.CurrentProfile.Count + 1
            .ColumnCount = 3
            .Columns(0).Width = 74
            .Columns(1).Width = 74
            .Columns(2).Width = 73
            r = 0
            For Each Pair In CurProj.Water.CurrentProfile
                If r > MaxCurrentPair Then Exit For
                .Rows(r).Cells(0).Value = CStr(Pair.Depth / CFLength)
                .Rows(r).Cells(1).Value = CStr(Pair.VelU / CFVel)
                .Rows(r).Cells(2).Value = CStr(Pair.VelV / CFVel)

                r = r + 1
            Next Pair
            .Columns(0).DefaultCellStyle.Format = "#.0"
            
            .Columns(1).DefaultCellStyle.Format = "#.00"
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
        ReDrawing()
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

        'should the user cancel the dialog box, exit
        On Error GoTo ErrHandler

        'set filters to allow selection of all files or just .CUR
        dlgCurrentProfileOpen.Filter = "All Files (*.*)|*.*|Current Files (*.cur)|*.cur"
        dlgCurrentProfileSave.Filter = "All Files (*.*)|*.*|Current Files (*.cur)|*.cur"

        'specify default filter as *.CUR
        dlgCurrentProfileOpen.FilterIndex = 2
        dlgCurrentProfileSave.FilterIndex = 2

        'specify the current filename
        dlgCurrentProfileOpen.FileName = "*.cur"
        dlgCurrentProfileSave.FileName = "*.cur"
        dlgCurrentProfileOpen.ShowReadOnly = False
        dlgCurrentProfileSave.OverwritePrompt = True

        'display the Open dialog box
        dlgCurrentProfileSave.ShowDialog()
        dlgCurrentProfileOpen.FileName = dlgCurrentProfileSave.FileName

        'open the file for output
        tmpFileNum = FreeFile()
        FileOpen(tmpFileNum, dlgCurrentProfileOpen.FileName, OpenMode.Output)

        Cursor = System.Windows.Forms.Cursors.WaitCursor

        'store the fully-qulaified file name
        CurProj.Water.CurrentProfile.CurrentFile = dlgCurrentProfileOpen.FileName

        'output
        OutputCurrentProfile((tmpFileNum))
        Cursor = System.Windows.Forms.Cursors.Default

        'close the output file
        FileClose(tmpFileNum)

        Exit Sub

ErrHandler:
        'user pressed Cancel button
        Exit Sub

    End Sub


    Private Sub OutputCurrentProfile(ByRef FileNum As Short)

        Dim Pair As CurrentPair
        'Dim flgFun As Short

        'save the form data in the Project object
        With CurProj.Water
            'simple output; profile name, then the depth/current pairs
            WriteLine(FileNum, .CurrentProfile.ProfileName)
            WriteLine(FileNum, CurProj.Units)

            For Each Pair In .CurrentProfile
                WriteLine(FileNum, Pair.Depth / CFLength, Pair.VelU / CFVel, Pair.VelV / CFVel)
            Next Pair
        End With

    End Sub


    Private Sub InputCurrentProfile(ByRef FileNum As Short)

        Dim r As Short
        Dim VelU, Depth, VelV As Single
        Dim PCFLength, PCFVel As Single
        Dim PName As String
        Dim ProUnits As VIVAMain.Units

        CheckingGrid = True
        r = 1

        'initialization
        PName = ""

        Input(FileNum, PName)
        txtProfileName.Text = PName
        Input(FileNum, ProUnits)

        If ProUnits = VIVAMain.Units.Metric Then
            PCFLength = 1
            PCFVel = 1
        Else
            PCFLength = Ft2M
            PCFVel = Kn2MPS
        End If

        Do While Not EOF(FileNum)
            Input(FileNum, Depth)
            Input(FileNum, VelU)
            Input(FileNum, VelV)

            With grdCurrent
                If r > MaxCurrentPair Then Exit Do
                Dim tmp As String() = {CStr(Depth * PCFLength / CFLength), _
                   CStr(VelU * PCFVel / CFVel), _
                   CStr(VelV * PCFVel / CFVel)}
                .Rows.Add(tmp)

                r = r + 1
            End With
        Loop

        CheckingGrid = False

    End Sub

    Private Sub ReDrawing()

        Dim wid As Integer = picCurrent.ClientSize.Width
        Dim hgt As Integer = picCurrent.ClientSize.Height

        ' Do nothing if we have no size.
        ' This happens, for example, if the form is minimized.
        If wid < 1 Or hgt < 1 Then Exit Sub

        Dim bm As Bitmap = New Bitmap(wid, hgt)
        Dim me_gr As Graphics = Me.CreateGraphics
        Dim me_hdc As IntPtr = me_gr.GetHdc

        Dim bounds As New RectangleF(0, 0, wid, hgt)
        Dim mf As New Imaging.Metafile(wmf_file_name, me_hdc, bounds, Imaging.MetafileFrameUnit.Pixel)
        me_gr.ReleaseHdc(me_hdc)

        Dim gr As Graphics = Graphics.FromImage(mf)
        gr.PageUnit = GraphicsUnit.Pixel
        gr.Clear(Color.White)

        If NumPoints > 0 Then
            Call MaxAndMin(CurVelA, Vmax, Vmin, NumPoints)
            Call MaxAndMin(Depth, Dmax, Dmin, NumPoints)

            'picCurrent.CreateGraphics().Clear(picCurrent.BackColor)
            'picCurrent.Refresh()
            'picCurrent.BackColor = Color.White

            'find drawing scales
            Dim scl As DrawScale
            Call DrawScaleF(Vmax, Vmin, Dmax, Dmin, picCurrent, scl)

            'DrawAxisNew() is copied from vivarray
            DrawAxisNew(gr, scl, Vmax, Vmin, Dmax, Dmin, CurrentLbls(3), CurrentLbls(0), "0.00", "0", False, False, 3, "Vy", VelUColor, _
                "Vz", VelVColor, "Vtotal", VelAColor)
            DrawLine(gr, scl, CurVelU, Depth, NumPoints, VelUColor)
            DrawLine(gr, scl, CurVelV, Depth, NumPoints, VelVColor)
            DrawLine(gr, scl, CurVelA, Depth, NumPoints, VelAColor)
        End If

        gr.Dispose()
        gr = Graphics.FromImage(bm)
        Dim dest_bounds As New RectangleF(0, 0, wid, hgt)
        Dim source_bounds As New RectangleF(0, 0, wid + 1, hgt + 1)
        gr.DrawImage(mf, bounds, source_bounds, GraphicsUnit.Pixel)
        picCurrent.SizeMode = PictureBoxSizeMode.AutoSize
        picCurrent.Image = bm
        gr.Dispose()

        ClipboardMetafileHelper.PutEnhMetafileOnClipboard(Me.Handle, mf)
        mf.Dispose()

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

End Class