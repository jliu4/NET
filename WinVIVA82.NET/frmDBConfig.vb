Public Class frmDBConfig
    Private bInitializing As Boolean = True
    Private oTmpVIVACoreFiles As VIVADBFiles = New VIVADBFiles
    Private srcFile(10), desFile(10) As String


    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        oTmpVIVACoreFiles = VIVACoreFiles

        PreDBs_Init()
        UserDBs_Init()

        bInitializing = False

        ' Add any initialization after the InitializeComponent() call.

    End Sub


    Private Sub PreDBs_Init()
        Dim i As Integer
        Dim tmpRow(5) As String

        With dvgPreDBs
            .ColumnCount = 5
            .Columns(0).HeaderText = "Filename"
            .Columns(1).HeaderText = "Number of Frequencies"
            .Columns(2).HeaderText = "Number of Cd"
            .Columns(3).HeaderText = "Riser Segment Type"
            .Columns(4).HeaderText = "Segment Type ID"
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            'Set column width
            .Columns(1).Width = 160
            .Columns(2).Width = 160
            .Columns(3).Width = 160
            .Columns(4).Width = 160
        End With

        For i = 0 To 9
            With oTmpVIVACoreFiles
                tmpRow(0) = .FileNames(i)
                tmpRow(1) = CStr(.NumFrq(i))
                tmpRow(2) = CStr(.NumAmp(i))
                tmpRow(3) = .TypeNames(i)
                tmpRow(4) = CStr(.TypeID(i))
                'If i = 0 Then
                'tmpRow(1) = ""
                'tmpRow(2) = ""
                'End If
                If i = 8 Then
                    tmpRow(4) = CStr(.TypeID(0)) & " (iHire =2)"
                End If
            End With

            With dvgPreDBs
                .Rows.Add(tmpRow)
            End With
        Next

    End Sub


    Private Sub UserDBs_Init()
        Dim i As Integer
        Dim clm As DataGridViewColumn = New DataGridViewCheckBoxColumn
        Dim oCell As DataGridViewCheckBoxCell = New DataGridViewCheckBoxCell
        Dim Checked As Boolean
        Dim nUserDBs As Integer
        Dim iFile As Integer

        nUserDBs = oTmpVIVACoreFiles.NumFiles - 10

        With dvgUserDBs
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            .RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True
            .ColumnCount = 1
            .Columns.Add(clm)
            .ColumnCount = 6
            .RowCount = 10
            .Columns(0).HeaderText = "Filename"
            .Columns(1).HeaderText = "Included"
            .Columns(2).HeaderText = "Source Hydrodynamic Database"
            .Columns(3).HeaderText = "Number of Frequencies"
            .Columns(4).HeaderText = "Number of Cd"
            .Columns(5).HeaderText = "Segment Type ID"
            '.Columns(0).ReadOnly = True
            .Columns(5).ReadOnly = True

            .Columns(3).ValueType = GetType(Integer)
            .Columns(4).ValueType = GetType(Integer)

            For i = 0 To 9
                iFile = i + 10
                'dvgUserDBs(0, i).Value = "User" & CStr(i + 1) & ".db"
                dvgUserDBs(5, i).Value = CStr(i + 11)
                oCell = TryCast(dvgUserDBs(1, i), DataGridViewCheckBoxCell)
                Checked = oCell.Value
                If i >= nUserDBs Then
                    dvgUserDBs.Rows(i).ReadOnly = True
                    oCell.Value = False
                Else
                    dvgUserDBs.Rows(i).ReadOnly = False
                    oCell.Value = True
                    dvgUserDBs(0, i).Value = oTmpVIVACoreFiles.FileNames(iFile)
                    dvgUserDBs(3, i).Value = oTmpVIVACoreFiles.NumFrq(iFile)
                    dvgUserDBs(4, i).Value = oTmpVIVACoreFiles.NumAmp(iFile)
                End If
            Next

            i = oTmpVIVACoreFiles.NumFiles - 10
            If i <= 10 Then
                .Rows(i).ReadOnly = False
            End If

            For i = 1 To 10
                iFile = i + 9
                desFile(i) = VIVADIR & oTmpVIVACoreFiles.FileNames(iFile)
            Next

            'Set column width
       
            .Columns(0).Width = 100
            .Columns(1).Width = 30
            .Columns(3).Width = 40
            .Columns(4).Width = 40
            .Columns(5).Width = 60
        End With
    End Sub


    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Dispose()
    End Sub


    Private Sub dvgUserDBs_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dvgUserDBs.CellClick
        Dim Col, Row As Integer
        Dim fileDlg As FileDialog
        Dim sFilename As String
        Dim i As Integer
        'Dim iWidth As Integer

        fileDlg = New OpenFileDialog
        fileDlg.InitialDirectory = VIVADIR
        sFilename = VIVADIR.Substring(0, 1)
        ChDrive(sFilename)
        ChDir(VIVADIR)

        With Me.dvgUserDBs
            Col = .CurrentCell.ColumnIndex
            Row = .CurrentCell.RowIndex
            Select Case Col
                Case 1
                    If Not .Rows(Row).ReadOnly Then
                        If .CurrentCell.Value = True Then
                            .CurrentCell.Value = False
                            For i = Row + 1 To 9
                                dvgUserDBs(Col, i).Value = False
                                .Rows(i).ReadOnly = True
                            Next
                            'Update Temperary VIVACorefiles Struture
                            oTmpVIVACoreFiles.NumFiles = 10 + Row
                        Else
                            .CurrentCell.Value = True
                            oTmpVIVACoreFiles.NumFiles = oTmpVIVACoreFiles.NumFiles + 1
                            If Row < 9 Then
                                Row = Row + 1
                                .Rows(Row).ReadOnly = False
                            End If
                        End If
                    End If

                Case 2 'Set source file names of the user defined DBs
                    If Not .Rows(Row).ReadOnly Then
                        If .CurrentCell.Value <> "" Then
                            fileDlg.FileName = .CurrentCell.Value
                        End If
                        If Not .CurrentCell.ReadOnly Then
                            fileDlg.ShowDialog()
                        End If
                        .CurrentCell.Value = fileDlg.FileName

                        If .CurrentCell.Value <> "" And Me.dvgUserDBs(1, Row).Value = False Then
                            Me.dvgUserDBs(1, Row).Value = True
                            .Rows(Row + 1).ReadOnly = False
                            oTmpVIVACoreFiles.NumFiles = oTmpVIVACoreFiles.NumFiles + 1
                        End If
                    End If
            End Select

        End With
    End Sub


    Private Sub dvgUserDBs_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dvgUserDBs.CellValueChanged
        Dim Col, Row As Integer
        Dim iFile As Integer

        If Not bInitializing Then
            With Me.dvgUserDBs
                Col = .CurrentCell.ColumnIndex
                Row = .CurrentCell.RowIndex
                iFile = 10 + Row
                Select Case Col
                    Case 0
                        If Me.dvgUserDBs(1, Row).Value Then
                            oTmpVIVACoreFiles.FileNames(iFile) = Me.dvgUserDBs(0, Row).Value
                            oTmpVIVACoreFiles.NumFrq(iFile) = Me.dvgUserDBs(3, Row).Value
                            oTmpVIVACoreFiles.NumAmp(iFile) = Me.dvgUserDBs(4, Row).Value
                        End If
                    Case 1
                        If Me.dvgUserDBs(Col, Row).Value Then
                            oTmpVIVACoreFiles.FileNames(iFile) = Me.dvgUserDBs(0, Row).Value
                            oTmpVIVACoreFiles.NumFrq(iFile) = Me.dvgUserDBs(3, Row).Value
                            oTmpVIVACoreFiles.NumAmp(iFile) = Me.dvgUserDBs(4, Row).Value
                        End If
                    Case 2
                        srcFile(iFile - 9) = Me.dvgUserDBs(Col, Row).Value
                    Case 3
                        If Me.dvgUserDBs(1, Row).Value Then
                            oTmpVIVACoreFiles.FileNames(iFile) = Me.dvgUserDBs(0, Row).Value
                            oTmpVIVACoreFiles.NumFrq(iFile) = Me.dvgUserDBs(3, Row).Value
                            oTmpVIVACoreFiles.NumAmp(iFile) = Me.dvgUserDBs(4, Row).Value
                        End If
                    Case 4
                        If Me.dvgUserDBs(1, Row).Value Then
                            oTmpVIVACoreFiles.FileNames(iFile) = Me.dvgUserDBs(0, Row).Value
                            oTmpVIVACoreFiles.NumFrq(iFile) = Me.dvgUserDBs(3, Row).Value
                            oTmpVIVACoreFiles.NumAmp(iFile) = Me.dvgUserDBs(4, Row).Value
                        End If
                End Select
            End With
        End If
    End Sub


    Private Sub dvgUserDBs_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dvgUserDBs.DataError
        Dim Col, Row As Integer

        With dvgUserDBs
            Col = .CurrentCell.ColumnIndex
            Row = .CurrentCell.RowIndex

            Select Case Col
                Case 3, 4                 'check if inputs are numbers
                    .CurrentCell = dvgUserDBs(Col, Row)
                    .CurrentCell.Value = ""
                    MsgBox("Please input a valid number.")
                Case Else
                    Exit Sub
            End Select
        End With

    End Sub


    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim i, iNumOfFiles, iFile As Integer
        Dim iFileNum As Integer

        'copy selected database files to VIVA directory USER*.DB
        If MsgBox("User defined database files will be overwritten by the files you selected, do you want continue", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
            'update globale VIVACorefiles Structure

            For i = 1 To 10
                iFile = i + 9
                desFile(i) = VIVADIR & oTmpVIVACoreFiles.FileNames(iFile)
            Next

            VIVACoreFiles = oTmpVIVACoreFiles

            iNumOfFiles = VIVACoreFiles.NumFiles

            'Update no_files.in
            iFileNum = FreeFile()
            FileOpen(iFileNum, VIVADIR & "no_files.in", OpenMode.Output, OpenAccess.Write)
            PrintLine(iFileNum, iNumOfFiles)
            With VIVACoreFiles
                For i = 0 To iNumOfFiles - 1
                    PrintLine(iFileNum, .FileNames(i) & Format(.NumFrq(i), " #0") & Format(.NumAmp(i), " #0"))
                Next i
            End With
            FileClose(iFileNum)

            Try
                For i = 1 To iNumOfFiles - 10 Step 1
                    If srcFile(i) <> "" And srcFile(i) <> desFile(i) Then
                        FileCopy(srcFile(i), desFile(i))
                    End If
                Next
            Catch ex As Exception

            End Try

        Else
            Exit Sub
        End If

        Me.Dispose()

    End Sub

    Private Sub dvgUserDBs_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dvgUserDBs.CellContentClick

    End Sub
End Class