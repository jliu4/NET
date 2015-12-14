Option Strict Off
Option Explicit On

Public Class frmMain
    Private NumProj As Short
    'current proj's data directory is the default data direcotory for next proj
    Private tmpDataDirSave As String
    Private tmpFileNum As Short
    Private dblCurrents(,) As Double
    Private nWd, nCase As Integer

    Private Sub btnFile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFileProject.Click
        'should the user cancel the dialog box, exit
        On Error GoTo ErrHandler
        Dim VerNum As Single
        'add current proj to previoufile menu
        If Trim(CurProj.FQFileName) <> "" Then
            PreviousFiles.AddPreFile(CurProj.FQFileName)
            'UpdateFilePreMenu()
        End If

        'get input file name
        Me.FileDialogOpen.FileName = "*.viv"
        Me.FileDialogOpen.Filter = "WinVIVA Project Files (*.viv)|*.viv|All Files|*.*"
        FileDialogOpen.FilterIndex = 2


        If CurProj.DataDirectory = "" Then
            FileDialogOpen.InitialDirectory = VIVADIR
        Else
            FileDialogOpen.InitialDirectory = CurProj.DataDirectory
        End If

        FileDialogOpen.ShowReadOnly = False
        FileDialogOpen.ShowDialog()

        'input from file
        tmpFileNum = FreeFile()
        FileOpen(tmpFileNum, FileDialogOpen.FileName, OpenMode.Input, OpenAccess.Read)

        Cursor = System.Windows.Forms.Cursors.WaitCursor
        tmpDataDirSave = CurProj.DataDirectory
        CurProj = New Project
        CurProj.BatchProcess = True
        CurProj.DataDirectory = tmpDataDirSave
        verNum = GetVersion(tmpFileNum)

        If VerNum > 8 Then
            InputProject8((tmpFileNum))
        Else
            InputProject(tmpFileNum)
        End If

        'InputProject((tmpFileNum))
        FileClose(tmpFileNum)

        'data directory and filename also have been set in below function
        CurProj.FQFileName = FileDialogOpen.FileName

        Text = "WinVIVA Batch - " & CurProj.Title & " (" & CurProj.FileName & ")"

        With LblProject
            .Text = "Project Location - " & CurProj.DataDirectory & Chr(13) & Chr(13)
            .Text = .Text & "Project File - " & CurProj.FileName & Chr(13) & Chr(13)
            .Text = .Text & "Project Description - " & CurProj.Desc & Chr(13) & Chr(13)
            If CurProj.Units = Units.English Then
                .Text = .Text & "Project Unit - " & "English" & Chr(13) & Chr(13)
            ElseIf CurProj.Units = Units.Metric Then
                .Text = .Text & "Project Unit - " & "Metric" & Chr(13) & Chr(13)
            End If
        End With

        'ClearPlotTableResults()

        'get existing resutls
        'If Len(Dir(CurProj.DataDirectory & VIVAOutputFiles(1))) > 0 Then
        '    HaveResults = ReadVIVACoreOutputFiles()
        'Else
        '    HaveResults = False
        'End If

        'Me.tbcMain.TabPages.Add(Me.tbpgCurrents)
        Me.btnCurrentFile.Enabled = True
        Cursor = System.Windows.Forms.Cursors.Default

ErrHandler:
        'User pressed Cancel button
        Cursor = System.Windows.Forms.Cursors.Default 'jjx
        Exit Sub
    End Sub

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        InitializeVariables()

        'VIVADIR = GetVIVADirFromRegistry()

        NumProj = 1
        CurProj = New Project
        CurProj.BatchProcess = True
        CurProj.Title = "No Project"
        Text = "WinVIVA Batch - " & CurProj.Title
        ReadInWinVivaIni()
        'UpdateFilePreMenu()

        tmpDataDirSave = CurProj.DataDirectory

        Me.tbcMain.TabPages.Remove(Me.tbpgCurrents)
        Me.btnCurrentFile.Enabled = False
        Me.btnBatch.Enabled = False
        Me.btnResults.Enabled = False
        Me.LblProject.Text = "Open a WinVIVA project file."

    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        CurProj = Nothing
        Me.Close()
        Me.Finalize()
        Me.Dispose(True)
    End Sub

    Private Sub btnCurrentFile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCurrentFile.Click
        Dim dblArr() As String
        Dim dblArrCurrs(1, 1) As String
        Dim strLine As String
        Dim i, j As Integer
        Dim retDiag As Object

        On Error GoTo ErrHandler

        nWd = 0
        nCase = 0

        'get input file name
        Me.FileDialogOpen.FileName = "*.csv"
        Me.FileDialogOpen.Filter = "Current Date Files (*.csv)|*.csv|All Files|*.*"
        FileDialogOpen.FilterIndex = 1


        If CurProj.DataDirectory = "" Then
            FileDialogOpen.InitialDirectory = VIVADIR
        Else
            FileDialogOpen.InitialDirectory = CurProj.DataDirectory
        End If

        FileDialogOpen.ShowReadOnly = False
        retDiag = FileDialogOpen.ShowDialog()

        If retDiag <> 2 Then
            If Me.tbcMain.TabCount = 1 Then
                Me.tbcMain.TabPages.Add(Me.tbpgCurrents)
            End If
        End If


        Me.dgvCurrents.RowCount = 0
        Me.dgvCurrents.ColumnCount = 0

        'input from file
        tmpFileNum = FreeFile()
        FileOpen(tmpFileNum, FileDialogOpen.FileName, OpenMode.Input, OpenAccess.Read)

        Cursor = System.Windows.Forms.Cursors.WaitCursor

        Do While Not EOF(tmpFileNum)
            If nWd = 0 Then
                strLine = LineInput(tmpFileNum)
                dblArr = strLine.Split(Chr(44))
                nCase = dblArr.Length - 1
                Me.dgvCurrents.ColumnCount = nCase + 1
                ReDim dblArrCurrs(nCase, nWd)
                i = 0
                Do While i <= nCase
                    If dblArr(i) = "" Then
                        nCase = i - 1
                        Me.dgvCurrents.ColumnCount = nCase + 1
                        ReDim dblArrCurrs(nCase, nWd)
                    Else
                        dblArrCurrs(i, nWd) = dblArr(i)
                        Me.dgvCurrents.Columns(i).HeaderText = "case" & Trim(Str(i))
                        i = i + 1
                    End If
                Loop
                If CurProj.Units = Units.English Then
                    Me.dgvCurrents.Columns(0).HeaderText = "WD (ft)"
                ElseIf CurProj.Units = Units.Metric Then
                    Me.dgvCurrents.Columns(0).HeaderText = "WD (m)"
                End If
                Me.dgvCurrents.Rows.Add(dblArr)
            Else
                strLine = LineInput(tmpFileNum)
                dblArr = strLine.Split(Chr(44))
                ReDim Preserve dblArrCurrs(nCase, nWd)
                For i = 0 To nCase
                    dblArrCurrs(i, nWd) = dblArr(i)
                Next i
                Me.dgvCurrents.Rows.Add(dblArr)
            End If
            nWd += 1
        Loop
        FileClose(tmpFileNum)

        ReDim dblCurrents(nCase, nWd - 1)

        For i = 0 To nCase
            For j = 0 To nWd - 1
                dblCurrents(i, j) = CDbl(dblArrCurrs(i, j))
            Next j
        Next i

        For i = 0 To nWd - 1
            Me.dgvCurrents.Rows(i).HeaderCell.Value = Trim(Str(i + 1))
        Next

        Me.dgvCurrents.RowHeadersWidth = 50
        Me.dgvCurrents.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Me.tbcMain.SelectedTab = Me.tbpgCurrents
        Me.btnBatch.Enabled = True
        Me.btnResults.Enabled = True
        Cursor = System.Windows.Forms.Cursors.Arrow


ErrHandler:
        'User pressed Cancel button
        'MsgBox("somethings is wrong")
        Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub

    End Sub

    Private Sub btnBatch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnBatch.Click
        Dim WD(nWd - 1), U(nWd - 1), V(nWd - 1) As Single
        Dim iCase, iWd As Integer
        Dim dirCase As String
        Dim dirProj As String
        Dim msgText As String
        Dim retMsg As Boolean
        Dim CurDir As Double

        Dim frmR As frmRun = New frmRun

        dirProj = CurProj.DataDirectory

        'On Error GoTo errorHandler

        Try
            Me.Enabled = False
            frmR.Show()

            With CurProj.Water.CurrentProfile
                iWd = 1
                Do While iWd <= .Count And .Item(iWd).VelA = 0
                    iWd = iWd + 1
                Loop

                If iWd > .Count Then
                    CurDir = 0
                ElseIf .Item(iWd).VelA = 0 Then
                    CurDir = 0
                Else
                    CurDir = System.Math.Atan2(.Item(iWd).VelV, .Item(iWd).VelU)
                End If
            End With

            For iWd = 1 To nWd
                WD(iWd - 1) = dblCurrents(0, iWd - 1)
            Next

            For iCase = 1 To nCase
                msgText = "Running Case No " & Trim(CStr(iCase)) & " of " & Trim(CStr(nCase)) & " cases, please wait ..."
                frmR.lblMsg1.Text = msgText
                Application.DoEvents()

                For iWd = 1 To nWd
                    U(iWd - 1) = dblCurrents(iCase, iWd - 1) * System.Math.Cos(CurDir)
                    V(iWd - 1) = dblCurrents(iCase, iWd - 1) * System.Math.Sin(CurDir)
                Next
                UpdateCurrentProfile(nWd, WD, U, V)
                dirCase = "Case" & Trim(CStr(iCase))
                CurProj.DataDirectory = dirProj & "\" & dirCase
                retMsg = DirExists(CurProj.DataDirectory)
                If (retMsg = False) Then
                    MkDir(CurProj.DataDirectory)
                End If
                CurProj.DataDirectory = dirProj & "\" & dirCase & "\"
                Application.DoEvents()

                If Not Case_Run(frmR, iCase, nCase) Then
                    MsgBox("Batch job stopped due to error.")
                    If Not frmR.IsDisposed Then
                        frmR.Dispose()
                        frmR = Nothing
                    End If
                    Me.Enabled = True
                    Me.Focus()
                    Exit Sub
                End If

                Application.DoEvents()

                If frmR.IsDisposed Then
                    frmR = Nothing
                End If

                frmR.prgBrCaseRun.Value = CInt((CSng(iCase) / CSng(nCase)) * 100)
                Application.DoEvents()
            Next iCase

            CurProj.DataDirectory = dirProj
            MsgBox("Batch run finished successfully!")
            If Not frmR Is Nothing Then
                If Not frmR.IsDisposed Then
                    frmR.Dispose()
                    frmR = Nothing
                End If
            End If
            Me.Enabled = True
            Me.Focus()
            Me.btnResults.Enabled = True
            Exit Sub

        Catch ex As Exception
            If Not frmR Is Nothing Then
                If Not frmR.IsDisposed Then
                    frmR.Dispose()
                    frmR = Nothing
                End If
            End If
            CurProj.DataDirectory = dirProj
            MsgBox("Batch job stopped by user ")
            Me.Enabled = True
            Me.Focus()
            Exit Sub

        End Try

        'errorHandler:
    End Sub

    Private Sub UpdateCurrentProfile(ByVal nWaterDepth As Integer, ByVal Depth() As Single, ByVal VelU() As Single, ByVal VelV() As Single)
        Dim CFLength, CFVel As Single
        Dim i, NumPairs As Short

        'current unit is knots
        CFVel = Kn2MPS
        If CurProj.Units = Units.English Then
            CFLength = Ft2M
        Else
            CFLength = 1.0
        End If

        'delete previous input
        NumPairs = CurProj.Water.CurrentProfile.Count

        'Debug.Print "NumPairs= " & NumPairs
        For i = 1 To NumPairs
            CurProj.Water.CurrentProfile.Delete(1) 'After deleting first element, second one will become the "first" one
        Next i

        'insert current input
        For i = 1 To nWaterDepth
            CurProj.Water.CurrentProfile.Add(CDepth:=Depth(i - 1) * CFLength, CVelU:=VelU(i - 1) * CFVel, CVelV:=VelV(i - 1) * CFVel)
        Next

    End Sub

    Function Case_Run(ByRef frmRun As frmRun, ByVal iCase As Integer, ByVal nCase As Integer) As Boolean

        Dim i As Short
        Dim msgText As String

        'On Error GoTo errorHandler

        Try
            Case_Run = True

            Cursor = System.Windows.Forms.Cursors.WaitCursor
            RemovePreviousVIVAOutputs()
            CopyImportedFiles()

            'check for files we need, and bail out if they aren't found
            'ChDir(VIVADIR)

            If Not VIVAFilesExist() Then
                msgText = "Cannot execute VIVA programs," & Chr(13) & "Batch job will be stopped!"
                MsgBox(msgText, MsgBoxStyle.OkOnly)
                frmRun.Dispose()
                Cursor = System.Windows.Forms.Cursors.Default
                Case_Run = False
                Exit Function
            End If

            'write the model in the format needed by VIVACORE
            msgText = "Preparing VIVA data..."
            If Not WriteDOSVIVAInput() Then
                msgText = "Error in data preparation," & Chr(13) & "Batch job will be stopped!"
                MsgBox(msgText, MsgBoxStyle.OkOnly)
                frmRun.Dispose()
                Cursor = System.Windows.Forms.Cursors.Default
                Case_Run = False
                Exit Function
            End If

            'loop through the programs that make up VIVA
            'backup current output files to another directory, if so requested

            'CurProj.BakFileDirectory = txtDirectory.Text
            'BackupOutputs


            'On Error Resume Next
            ChDrive(VIVADIR)
            ChDir(VIVADIR)

            For i = 1 To NumVIVADOSPrograms
                'update the status for the user
                msgText = "Executing VIVA Program:  " & VIVADOSPrograms(CurProj.nRisers, i)
                frmRun.lblMsg2.Text = msgText
                Application.DoEvents()

                'this is the actual execution step
                ExecCmd(VIVADIR & VIVADOSPrograms(CurProj.nRisers, i), True)
            Next i

            'see if anything went wrong
            If Not VIVACoreOutputsExist() Then
                'if so, warn the user
                HaveResults = False
            Else
                'rename the output files using the project name
                'CopyVIVACoreOutputs
                MoveVIVAOutputs()

                'read the information computed
                'HaveResults = ReadVIVACoreOutputFiles()
                'tell the user the good news
            End If

            Case_Run = True
            Cursor = System.Windows.Forms.Cursors.Default
            Exit Function

            'If HaveResults Then
            '    msgText = "Execution is completed"
            '    Cursor = System.Windows.Forms.Cursors.Default
            '    Case_Run = True
            'Else
            '    msgText = "Abnormal Completion of VIVA Programs"
            '    frmRun.lblMsg2.Text = msgText
            '    Cursor = System.Windows.Forms.Cursors.Default
            '    Case_Run = False
            'End If

        Catch ex As Exception
            Case_Run = False
            Cursor = System.Windows.Forms.Cursors.Default
            'Me.Enabled = True
            'Me.Focus()
            Exit Function
        End Try

        'errorHandler:
    End Function

    Function DirExists(ByVal DirName As String) As Boolean
        On Error GoTo ErrorHandler
        ' test the directory attribute
        DirExists = GetAttr(DirName) And vbDirectory
ErrorHandler:
        ' if an error occurs, this function returns False
    End Function

    Private Sub btnResults_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnResults.Click
        Dim oxApp As Excel.Application, oxBook As Excel.Workbook
        Dim iCase As Long
        Dim NumNodes As Long
        Dim oxSheetSMLife, oxSheetMMLife, oxSheetSMDisp, oxSheetMMDisp As Excel.Worksheet
        Dim oxSheet As Excel.Worksheet
        Dim dblElmLngth As Double
        Dim FirstRow, FirstCol As Long
        Dim i As Long
        Dim ir, ix As Short
        ' Dim outputFiles(2, 4) As String

        Dim strCase, strFN As String, FileNum As Integer
        Dim frmR As frmRun = New frmRun, msgText As String, strBuff As String, nMode, nModeMin, iMode As Long
        Dim lngbuff, dblbuff, fatLifeMin, fatLifeBuff As Double
        'outputFiles(0, 0) = "fat1.out"
        'outputFiles(0, 1) = "fat-mono.out"
        'outputFiles(0, 2) = "fat-multi.out"
        'outputFiles(0, 3) = "out.out"
        'outputFiles(0, 4) = "out_mm.out"

        Try
            frmR.Show()
            msgText = "Initializing "
            frmR.lblMsg1.Text = msgText
            frmR.Text = "Reading batch results"
            Application.DoEvents()

            Me.Enabled = False
            Me.UseWaitCursor = True
            FirstRow = 10
            FirstCol = 1

            oxApp = New Excel.Application
            For ir = 1 To CurProj.nRisers
                ix = ir * CurProj.nRisers - ir
                oxBook = oxApp.Workbooks.Add
                oxBook.Activate()
                oxApp.ScreenUpdating = False
                oxApp.Calculation = Excel.XlCalculation.xlCalculationManual

                oxSheetMMDisp = oxBook.Worksheets.Add
                oxSheetMMDisp.Name = ("Disp_MM")
                oxSheetSMDisp = oxBook.Worksheets.Add
                oxSheetSMDisp.Name = ("Disp_SM")
                oxSheetMMLife = oxBook.Worksheets.Add
                oxSheetMMLife.Name = ("Fatigue_MM")
                oxSheetSMLife = oxBook.Worksheets.Add
                oxSheetSMLife.Name = ("Fatigue_SM")

                'oxBook.Worksheets("sheet1").delete()
                'oxBook.Worksheets("sheet2").delete()
                'oxBook.Worksheets("sheet3").delete()

                NumNodes = CurProj.NumPoints
                dblElmLngth = CurProj.Riser(ir).TotalLength / (NumNodes - 1)

                'format the sheets
                oxSheetSMLife.Cells(1, 1) = "Single mode fatigue life (Year)"
                oxSheetSMLife.Cells(6, 1) = "*Representing modes with least fatigue life"
                oxSheetMMLife.Cells(1, 1) = "Multi mode fatigue life (Year)"
                oxSheetSMDisp.Cells(1, 1) = "Single mode maximum displacement (meter)"
                oxSheetSMDisp.Cells(6, 1) = "*Representing modes with least fatigue life"
                oxSheetMMDisp.Cells(1, 1) = "Multi mode displcacement (meter)"

                For Each oxSheet In oxBook.Worksheets
                    With oxSheet.Cells(1, 1)
                        .font.size = "10"
                        .font.bold = True
                    End With
                    With oxSheet.Cells(6, 1)
                        .font.size = "8"
                    End With
                    With oxSheet.Cells(2, 1)
                        .value = "Project file - " & CurProj.FileName
                        .font.size = "10"
                        '.font.bold = True
                    End With
                    With oxSheet.Cells(3, 1)
                        .value = "Project description - " & CurProj.Desc
                        .font.size = "10"
                        .wraptext = False
                        '.font.bold = True
                    End With
                    With oxSheet.Cells(4, 1)
                        .value = "Riser Length - " & CurProj.Riser(ir).TotalLength & " (meter)"
                        .font.size = "10"
                        '.font.bold = True
                    End With
                    With oxSheet.Cells(5, 1)
                        .value = "Total current profiles - " & nCase
                        .font.size = "10"
                        '.font.bold = True
                    End With
                    With oxSheet.Cells(FirstRow - 2, 1)
                        .value = "WD"
                        .font.bold = True
                        .HorizontalAlignment = Excel.Constants.xlRight
                    End With
                    With oxSheet.Cells(FirstRow - 1, 1)
                        .value = "(meter)"
                        .font.bold = True
                        .HorizontalAlignment = Excel.Constants.xlRight
                    End With
                    For i = 1 To nCase
                        oxSheet.Cells(FirstRow - 2, i + 1) = "Case" & Trim(CStr(i))
                        oxSheet.Cells(FirstRow - 2, i + 1).font.bold = True
                        oxSheet.Cells(FirstRow - 2, i + 1).HorizontalAlignment = Excel.Constants.xlRight
                    Next

                    If frmR.IsDisposed Then
                        frmR = Nothing
                    End If

                    For i = 1 To NumNodes
                        oxSheet.Cells(i + (FirstRow - 1), FirstCol) = (i - 1) * dblElmLngth
                        oxSheet.Cells(i + (FirstRow - 1), FirstCol).numberformat = "0.00"
                        oxSheet.Cells(i + (FirstRow - 1), FirstCol).font.size = 8
                        frmR.prgBrCaseRun.Value = CInt((CSng(i) / CSng(NumNodes)) * 100)
                        Application.DoEvents()
                    Next
                Next

                frmR.prgBrCaseRun.Value = 0
                Application.DoEvents()
                For iCase = 1 To nCase
                    msgText = "Reading results for Case No " & Trim(CStr(iCase)) & " of " & Trim(CStr(nCase)) & " cases, please wait ..."
                    frmR.lblMsg1.Text = msgText
                    Application.DoEvents()

                    'Debug.Print(iCase)
                    strCase = "Case" & Trim(CStr(iCase))

                    strFN = CurProj.DataDirectory & ".\" & strCase & "\" & OutputFiles(ix, fat1_out)
                    FileNum = FreeFile()
                    FileOpen(FileNum, strFN, OpenMode.Input, OpenAccess.Read)
                    strBuff = LineInput(FileNum)
                    strBuff.Split()
                    Input(FileNum, nMode)
                    If nMode = 0 Then
                        FileClose(FileNum)
                        GoTo Line1
                    End If
                    nModeMin = nMode

                    Input(FileNum, lngbuff)
                    Input(FileNum, dblbuff)
                    Input(FileNum, fatLifeMin)
                    Input(FileNum, dblbuff)
                    For iMode = 2 To nMode
                        Input(FileNum, lngbuff)
                        Input(FileNum, dblbuff)
                        Input(FileNum, fatLifeBuff)
                        Input(FileNum, dblbuff)
                        If fatLifeMin > fatLifeBuff Then
                            nModeMin = iMode
                            fatLifeMin = fatLifeBuff
                        End If
                    Next iMode
                    FileClose(FileNum)

                    strFN = CurProj.DataDirectory & strCase & "\" & OutputFiles(ix, fat_mono_out)
                    FileNum = FreeFile()
                    FileOpen(FileNum, strFN, OpenMode.Input, OpenAccess.Read)
                    strBuff = LineInput(FileNum)
                    For iMode = 1 To nModeMin - 1
                        For i = 1 To NumNodes
                            strBuff = LineInput(FileNum)
                        Next i
                    Next iMode
                    For i = 1 To NumNodes
                        Input(FileNum, dblbuff)
                        Input(FileNum, fatLifeBuff)
                        With oxSheetSMLife
                            .Cells(i + (FirstRow - 1), iCase + 1) = fatLifeBuff
                            .Cells(i + (FirstRow - 1), iCase + 1).numberformat = "0.000E+00"
                            .Cells(i + (FirstRow - 1), iCase + 1).font.size = 8
                        End With
                    Next i
                    FileClose(FileNum)

                    strFN = CurProj.DataDirectory & strCase & "\" & OutputFiles(ix, fat_multi_out)
                    FileNum = FreeFile()
                    FileOpen(FileNum, strFN, OpenMode.Input, OpenAccess.Read)
                    For i = 1 To NumNodes
                        Input(FileNum, dblbuff)
                        Input(FileNum, fatLifeBuff)
                        With oxSheetMMLife
                            .Cells(i + (FirstRow - 1), iCase + 1) = fatLifeBuff
                            .Cells(i + (FirstRow - 1), iCase + 1).numberformat = "0.000E+00"
                            .Cells(i + (FirstRow - 1), iCase + 1).font.size = 8
                        End With
                    Next i
                    FileClose(FileNum)

                    strFN = CurProj.DataDirectory & strCase & "\" & OutputFiles(ix, out_out)
                    FileNum = FreeFile()
                    FileOpen(FileNum, strFN, OpenMode.Input, OpenAccess.Read)
                    For iMode = 1 To nModeMin - 1
                        For i = 1 To NumNodes
                            strBuff = LineInput(FileNum)
                        Next i
                    Next iMode
                    For i = 1 To NumNodes
                        Input(FileNum, dblbuff)
                        Input(FileNum, fatLifeBuff)
                        With oxSheetSMDisp
                            .Cells(i + (FirstRow - 1), iCase + 1) = fatLifeBuff
                            .Cells(i + (FirstRow - 1), iCase + 1).numberformat = "0.000E+00"
                            .Cells(i + (FirstRow - 1), iCase + 1).font.size = 8
                        End With
                    Next i
                    FileClose(FileNum)

                    strFN = CurProj.DataDirectory & strCase & "\" & OutputFiles(ix, out_mm_out)
                    FileNum = FreeFile()
                    FileOpen(FileNum, strFN, OpenMode.Input, OpenAccess.Read)
                    For i = 1 To NumNodes
                        Input(FileNum, dblbuff)
                        Input(FileNum, fatLifeBuff)
                        With oxSheetMMDisp
                            .Cells(i + (FirstRow - 1), iCase + 1) = fatLifeBuff
                            .Cells(i + (FirstRow - 1), iCase + 1).numberformat = "0.000E+00"
                            .Cells(i + (FirstRow - 1), iCase + 1).font.size = 8
                        End With
                    Next i
                    FileClose(FileNum)


Line1:
                    Application.DoEvents()
                    If frmR.IsDisposed Then
                        frmR = Nothing
                    End If

                    frmR.prgBrCaseRun.Value = CInt((CSng(iCase) / CSng(nCase)) * 100)
                    Application.DoEvents()
                Next iCase
            Next ir

            If Not frmR Is Nothing Then
                If Not frmR.IsDisposed Then
                    frmR.Dispose()
                    frmR = Nothing
                End If
            End If
            oxApp.Visible = True
            oxApp.ScreenUpdating = True
            oxApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic
            Me.Enabled = True
            Me.UseWaitCursor = False
            Me.Activate()
            oxApp = Nothing
            Exit Sub

        Catch ex As Exception
            If Not frmR Is Nothing Then
                If Not frmR.IsDisposed Then
                    frmR.Dispose()
                    frmR = Nothing
                End If
            End If
            'oxApp.Visible = True
            'oxBook.Close(False)
            'oxApp.Quit()
            'oxApp = Nothing
            Me.Enabled = True
            Me.UseWaitCursor = False
            Me.Activate()
            MsgBox("Action Cancelled!!!")

        End Try

    End Sub
End Class

