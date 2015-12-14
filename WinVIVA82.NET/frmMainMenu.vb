Option Strict Off
Option Explicit On

' Version 8.1
' 2014, Copyright DTCEL, All Rights Reserved
' Jin Liu May 2014

Imports System
Imports System.Diagnostics
Imports System.ComponentModel

Friend Class frmMainMenu

    Inherits System.Windows.Forms.Form

    Private ERROR_FILE_NOT_FOUND As Integer = 2
    Private ERROR_ACCESS_DENIED As Integer = 5
    Private ERROR_NO_ASSOCIATION As Integer = 1155

    Private NumProj As Short
    Private msgNotSavedWarning As String

    'current proj's data directory is the default data direcotory for next proj
    Private tmpDataDirSave As String

    Private tmpFileNum As Short
    Private Const msgTitle As String = " WinVIVA "

    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As _
        String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

    Private Sub frmMainMenu_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        InitializeVariables()
        NumProj = 1
        CurProj = New Project
        
        CurProj.Title = "Project" & NumProj
        Text = "WinVIVA - Main Menu - " & CurProj.Title
        ReadInWinVivaIni()
        UpdateFilePreMenu()
        tmpDataDirSave = CurProj.DataDirectory
        If (CurProj.nRisers = 1) Then 'JLIU TODO
            mnuRiserR.Enabled = False
        Else
            mnuRiserR.Enabled = True

        End If
        If CurProj.nRisers = 1 Then
            mnuResultsR.Enabled = False
        Else
            mnuResultsR.Enabled = True
        End If
        eventArgs.ToString()
    End Sub

    Private Sub frmMainMenu_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        Dim Cancel As Boolean = eventArgs.Cancel
        Dim unloadmode As System.Windows.Forms.CloseReason = eventArgs.CloseReason

        'add current proj to prefile menu
        If Trim(CurProj.FQFileName) <> "" Then
            PreviousFiles.AddPreFile(CurProj.FQFileName)
            UpdateFilePreMenu()
        End If

        OutputWinVivaIni()

        If CurProj.Saved = False Then
            Select Case MsgBox("Do you want to save changes made to " & CurProj.Title & " ?", MsgBoxStyle.YesNoCancel, "WinVIVA - Exit Program")
                Case MsgBoxResult.Yes
                    'save
                    tmpDataDirSave = CurProj.DataDirectory
                    mnuFileSave_Click(mnuFileSave, New System.EventArgs())
                    'Destroy existing case and leave
                    CurProj = Nothing

                Case MsgBoxResult.No
                    'Destroy existing case and leave
                    tmpDataDirSave = CurProj.DataDirectory
                    CurProj = Nothing

                Case MsgBoxResult.Cancel
                    'do nothing
                    Cancel = True
            End Select
        Else
            'Destroy existing case and leave
            CurProj = Nothing
        End If

        eventArgs.Cancel = Cancel

    End Sub


    Public Sub mnuFileNew_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileNew.Click

        If CurProj.Saved = False Then
            Select Case MsgBox("Do you want to save changes made to " & CurProj.Title & " ?", MsgBoxStyle.YesNoCancel, "WinVIVA - Exit Program")
                Case MsgBoxResult.Yes
                    'save
                    mnuFileSave_Click(mnuFileSave, New System.EventArgs())

                    If Trim(CurProj.FQFileName) <> "" Then
                        PreviousFiles.AddPreFile(CurProj.FQFileName)
                        UpdateFilePreMenu()
                    End If

                    'Destroy existing case and start a new one
                    tmpDataDirSave = CurProj.DataDirectory
                    CurProj = New Project
                    CurProj.DataDirectory = tmpDataDirSave
                    NumProj = NumProj + 1
                    CurProj.Title = "Project" & NumProj
                    Text = "WinVIVA - Main Menu - " & CurProj.Title
                    HaveResults = False
                    frmProjDesc.Show()

                Case MsgBoxResult.No
                    If Trim(CurProj.FQFileName) <> "" Then
                        PreviousFiles.AddPreFile(CurProj.FQFileName) 'JLIU TODO
                        UpdateFilePreMenu()
                    End If

                    'Destroy existing case and start a new one
                    tmpDataDirSave = CurProj.DataDirectory
                    CurProj = New Project
                    CurProj.DataDirectory = tmpDataDirSave
                    NumProj = NumProj + 1
                    CurProj.Title = "Project" & NumProj
                    Text = "WinVIVA - Main Menu - " & CurProj.Title
                    HaveResults = False
                    frmProjDesc.Show()

                Case MsgBoxResult.Cancel
                    'do nothing
            End Select
        Else
            If Trim(CurProj.FQFileName) <> "" Then
                PreviousFiles.AddPreFile(CurProj.FQFileName)
                UpdateFilePreMenu()
            End If

            tmpDataDirSave = CurProj.DataDirectory
            CurProj = New Project
            CurProj.DataDirectory = tmpDataDirSave
            NumProj = NumProj + 1
            CurProj.Title = "Project" & NumProj
            Text = "WinVIVA - Main Menu - " & CurProj.Title
            HaveResults = False
            frmProjDesc.Show()
        End If

    End Sub


    Public Sub mnuFileOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileOpen.Click

        'should the user cancel the dialog box, exit
        On Error GoTo ErrHandler
        Dim VerNum As Single
        'save the current project
        If CurProj.Saved = False Then
            Select Case MsgBox("Do you want to save changes made to " & CurProj.Title & " ?", MsgBoxStyle.YesNoCancel, "WinVIVA - Exit Program")
                Case MsgBoxResult.Yes
                    'save
                    mnuFileSave_Click(mnuFileSave, New System.EventArgs())
                    HaveResults = False
                Case MsgBoxResult.No
                    'do nothing
                    HaveResults = False
                Case MsgBoxResult.Cancel
                    Exit Sub
            End Select
        End If

        'add current proj to previoufile menu
        If Trim(CurProj.FQFileName) <> "" Then
            PreviousFiles.AddPreFile(CurProj.FQFileName)
            UpdateFilePreMenu()
        End If

        'get input file name
        FileDialogOpen.Filter = "All Files (*.*)|*.*|VIV Files (*.viv)|*.viv"
        FileDialogSave.Filter = "All Files (*.*)|*.*|VIV Files (*.viv)|*.viv"
        FileDialogOpen.FilterIndex = 2
        FileDialogSave.FilterIndex = 2

        If CurProj.DataDirectory = "" Then
            FileDialogOpen.InitialDirectory = VIVADIR
            FileDialogSave.InitialDirectory = VIVADIR
        Else
            FileDialogOpen.InitialDirectory = CurProj.DataDirectory
            FileDialogSave.InitialDirectory = CurProj.DataDirectory
        End If

        FileDialogOpen.ShowReadOnly = False
        FileDialogOpen.ShowDialog()
        FileDialogSave.FileName = FileDialogOpen.FileName

        'input from file
        tmpFileNum = FreeFile()
        FileOpen(tmpFileNum, FileDialogOpen.FileName, OpenMode.Input, OpenAccess.Read)

        Cursor = System.Windows.Forms.Cursors.WaitCursor
        tmpDataDirSave = CurProj.DataDirectory
        CurProj = New Project
        CurProj.DataDirectory = tmpDataDirSave
        verNum = GetVersion(tmpFileNum)

        If VerNum > 8 Then
            InputProject8((tmpFileNum))
        Else
            InputProject(tmpFileNum)
        End If


        'data directory and filename also have been set in below function
        CurProj.FQFileName = FileDialogOpen.FileName

        Text = "WinVIVA - Main Menu - " & CurProj.Title & " (" & CurProj.FileName & ")"
        FileClose(tmpFileNum)

        ClearPlotTableResults()
        If (CurProj.nRisers = 1) Then 'JLIU TODO
            mnuRiserR.Enabled = False
        Else
            mnuRiserR.Enabled = True

        End If
        If CurProj.nRisers = 1 Then
            mnuResultsR.Enabled = False
        Else
            mnuResultsR.Enabled = True
        End If
        'get existing resutls
        'JLIU TODO, it should move to different place 
        If Len(Dir(CurProj.DataDirectory & OutputFiles(0, out_out))) > 0 Then
            HaveResults = ReadVIVACoreOutputFiles()
        Else
            HaveResults = False
        End If

        Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub

ErrHandler:
        'User pressed Cancel button
        Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub

    End Sub


    Public Sub mnuFileSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileSave.Click

        Dim Fn As String

        Fn = CurProj.FQFileName

        If Fn = "" Then
            mnuFileSaveAs_Click(mnuFileSaveAs, New System.EventArgs())
        Else
            tmpFileNum = FreeFile()
            FileOpen(tmpFileNum, Fn, OpenMode.Output)
            Cursor = System.Windows.Forms.Cursors.WaitCursor
            OutputProject((tmpFileNum))

            Cursor = System.Windows.Forms.Cursors.Default
            FileClose(tmpFileNum)
            CurProj.Saved = True
        End If

    End Sub


    Public Sub mnuFileSaveAs_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileSaveAs.Click

        '   should the user cancel the dialog box, exit
        On Error GoTo ErrHandler

        '   get file name
        FileDialogOpen.Filter = "All Files (*.*)|*.*|viv Files (*.viv)|*.viv"
        FileDialogSave.Filter = "All Files (*.*)|*.*|viv Files (*.viv)|*.viv"
        FileDialogOpen.FilterIndex = 2
        FileDialogSave.FilterIndex = 2

        If CurProj.DataDirectory = "" Then
            FileDialogOpen.InitialDirectory = VIVADIR
            FileDialogSave.InitialDirectory = VIVADIR
        Else
            FileDialogOpen.InitialDirectory = CurProj.DataDirectory
            FileDialogSave.InitialDirectory = CurProj.DataDirectory
        End If

        If CurProj.FileName <> "" Then
            FileDialogOpen.FileName = CurProj.FileName
            FileDialogSave.FileName = CurProj.FileName
        Else
            FileDialogOpen.FileName = CurProj.Title & ".viv"
            FileDialogSave.FileName = CurProj.Title & ".viv"
        End If

        FileDialogSave.OverwritePrompt = True
        FileDialogOpen.ShowReadOnly = False
        FileDialogSave.ShowDialog()
        FileDialogOpen.FileName = FileDialogSave.FileName

        '   save data
        tmpFileNum = FreeFile()
        FileOpen(tmpFileNum, FileDialogOpen.FileName, OpenMode.Output)

        ' add current proj to previoufile menu
        If Trim(CurProj.FQFileName) <> "" Then
            PreviousFiles.AddPreFile(CurProj.FQFileName)
            UpdateFilePreMenu()
        End If

        Cursor = System.Windows.Forms.Cursors.WaitCursor
        CurProj.FQFileName = FileDialogOpen.FileName

        OutputProject((tmpFileNum))

        Cursor = System.Windows.Forms.Cursors.Default
        FileClose(tmpFileNum)

        CurProj.Saved = True
        Text = "WinVIVA - Main Menu - " & CurProj.Title & " (" & CurProj.FileName & ")"
        Exit Sub

ErrHandler:
        '   User pressed Cancel button
        Exit Sub

    End Sub


    Public Sub mnuFilePD_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFilePD.Click

        frmProjDesc.Show()

    End Sub


    Public Sub mnuFilePrintSetup_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFilePrintSetup.Click

        FileDialogPrint.ShowDialog()

    End Sub


    Public Sub mnuFileExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileExit.Click

        Me.Close()

    End Sub


    Public Sub mnuAPGlobal_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAPGlobal.Click

        frmGlobalParms.Show()

    End Sub
    Public Sub mnuRefreshMainMenu_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRefreshMainMenu.Click

        If (CurProj.nRisers = 1) Then 'Or CurProj.sameRiserARiser = True) Then 'JLIU TODO
            Me.mnuRiserR.Enabled = False
        Else
            Me.mnuRiserR.Enabled = True

        End If
        If CurProj.nRisers = 1 Then
            Me.mnuResultsR.Enabled = False
        Else
            Me.mnuResultsR.Enabled = True
        End If
        eventArgs.ToString()

    End Sub


    Public Sub mnuAPCurrentProfiles_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAPCurrentProfiles.Click

        frmCurrent.Show()

    End Sub


    Public Sub mnuFilePre_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFilePre.Click

        Dim index As Short = mnuFilePre.GetIndex(eventSender)
        Dim msgnofilewarning As Object
        Dim Path, InputFile, Fn As String
        Dim verNum As Single
        'should the user cancel the dialog box, exit
        On Error GoTo ErrHandler

        'save the current project
        If CurProj.Saved = False Then

            If CurProj.FileName = "" Then
                msgNotSavedWarning = "Do you want to save changes made to " & " the current project ?"
            Else
                msgNotSavedWarning = "Do you want to save changes made to " & CurProj.FileName & " ?"
            End If

            Select Case MsgBox(msgNotSavedWarning, MsgBoxStyle.YesNoCancel, msgTitle)
                Case MsgBoxResult.Yes
                    'save
                    mnuFileSave_Click(mnuFileSave, New System.EventArgs())
                    HaveResults = False
                Case MsgBoxResult.No
                    'do nothing
                    HaveResults = False
                Case MsgBoxResult.Cancel
                    Exit Sub
            End Select
        End If

        'initializaiton
        Path = ""
        Fn = ""

        'input from file
        InputFile = PreviousFiles.PreFile(index + 1)
        GetDirNFileName(InputFile, Path, Fn)

        If Len(Dir(InputFile)) <= 0 Or Fn = "" Then
            msgnofilewarning = InputFile & " doesn't exist!"
            MsgBox(msgnofilewarning, MsgBoxStyle.OkOnly, msgTitle)
            PreviousFiles.DelPreFile(index + 1)
            UpdateFilePreMenu()
            Exit Sub
        End If

        PreviousFiles.DelPreFile(index + 1)
        If CurProj.FileName <> "" Then PreviousFiles.AddPreFile(CurProj.DataDirectory & CurProj.FileName)

        UpdateFilePreMenu()

        Cursor = System.Windows.Forms.Cursors.WaitCursor
        tmpDataDirSave = CurProj.DataDirectory

        CurProj = New Project
        CurProj.FQFileName = InputFile

        With CurProj
            'GetDirNFileName InputFile
            tmpFileNum = FreeFile()
            FileOpen(tmpFileNum, InputFile, OpenMode.Input, OpenAccess.Read)
            verNum = GetVersion(tmpFileNum)

            If VerNum > 8 Then
                InputProject8((tmpFileNum))
            Else
                InputProject(tmpFileNum)
            End If


            Text = "WinVIVAARY - Main Menu - " & .Title & " (" & .FileName & ")"
            FileClose(tmpFileNum)
            'JLIU TODO should move to different place

            If (CurProj.nRisers = 1) Then 'JLIU TODO
                mnuRiserR.Enabled = False
            Else
                mnuRiserR.Enabled = True

            End If
            If CurProj.nRisers = 1 Then
                mnuResultsR.Enabled = False
            Else
                mnuResultsR.Enabled = True
            End If

            If CurProj.Riser(1).RiserType = VIVAMain.RiserTypes.Rigid Then
                Me.mnuRiserFStaticSolution.Enabled = False
            Else
                Me.mnuRiserFStaticSolution.Enabled = True
            End If
            If CurProj.nRisers = 2 And CurProj.Riser(2).RiserType = VIVAMain.RiserTypes.Rigid Then
                Me.mnuRiserRStaticSolution.Enabled = False
            Else
                Me.mnuRiserRStaticSolution.Enabled = True
            End If
        End With

        ClearPlotTableResults()

        'get existing resutls
        If Len(Dir(CurProj.DataDirectory & OutputFiles(0, summ1_out))) > 0 Then
            HaveResults = ReadVIVACoreOutputFiles()
        Else
            HaveResults = False
        End If

        Cursor = System.Windows.Forms.Cursors.Default

  
        Exit Sub

ErrHandler:
        'User pressed Cancel button
        Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub

    End Sub

    'initiate menu
    Private Sub UpdateFilePreMenu()

        Dim i, NumPreFiles As Short

        NumPreFiles = PreviousFiles.CountPreFile

        If NumPreFiles > 0 Then
            mnuLineFilePre.Visible = True
        Else
            mnuLineFilePre.Visible = False
        End If

        For i = 1 To NumPreFiles Step 1
            With mnuFilePre(i - 1)
                .Text = i & ". " & PreviousFiles.PreFile(i)
                .Enabled = True
                .Visible = True
            End With
        Next i

        For i = NumPreFiles + 1 To MaxNumPreFiles Step 1
            With mnuFilePre(i - 1)
                .Text = ""
                .Enabled = False
                .Visible = False
            End With
        Next i

    End Sub

    Public Sub mnuManual_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuManual.Click

        Dim index As Short = mnuManual.GetIndex(eventSender)
        Dim ReadPdfCmd As String

        ReadPdfCmd = VIVADIR & ManualFile(index)
        Try
            System.Diagnostics.Process.Start(ReadPdfCmd)

        Catch ex As System.ComponentModel.Win32Exception
            Select Case ex.NativeErrorCode
                Case ERROR_FILE_NOT_FOUND
                    MsgBox(ex.Message + ". Check the path.")
                Case ERROR_ACCESS_DENIED
                    MsgBox(ex.Message + ". You do not have permission to print this file.")
                Case ERROR_NO_ASSOCIATION
                    MsgBox(ex.Message + "." + Chr(13) + "The files is " + ReadPdfCmd + ".")
            End Select
        End Try

    End Sub
    Public Sub mnuVIVOfFlexibleStructures_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuVIVOfFlexibleStructures.Click

        Dim ReadPdfCmd As String

        ReadPdfCmd = VIVADIR & "\Reference\" & ReferenceFile(0)
        Try
            System.Diagnostics.Process.Start(ReadPdfCmd)

        Catch ex As System.ComponentModel.Win32Exception
            Select Case ex.NativeErrorCode
                Case ERROR_FILE_NOT_FOUND
                    MsgBox(ex.Message + ". Check the path.")
                Case ERROR_ACCESS_DENIED
                    MsgBox(ex.Message + ". You do not have permission to print this file.")
                Case ERROR_NO_ASSOCIATION
                    MsgBox(ex.Message + "." + Chr(13) + "The files is " + ReadPdfCmd + ".")
            End Select
        End Try

    End Sub
    Public Sub mnuDefinitionOfCoordinateSystems_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFatigueCalculation.Click

        Dim ReadPdfCmd As String

        ReadPdfCmd = VIVADIR & "\Reference\" & ReferenceFile(1)
        Try
            System.Diagnostics.Process.Start(ReadPdfCmd)

        Catch ex As System.ComponentModel.Win32Exception
            Select Case ex.NativeErrorCode
                Case ERROR_FILE_NOT_FOUND
                    MsgBox(ex.Message + ". Check the path.")
                Case ERROR_ACCESS_DENIED
                    MsgBox(ex.Message + ". You do not have permission to print this file.")
                Case ERROR_NO_ASSOCIATION
                    MsgBox(ex.Message + "." + Chr(13) + "The files is " + ReadPdfCmd + ".")
            End Select
        End Try

    End Sub
    Public Sub mnuDefinitionOfReferenceSystems_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuModelingStrakes.Click

        Dim ReadPdfCmd As String

        ReadPdfCmd = VIVADIR & "\Reference\" & ReferenceFile(2)
        Try
            System.Diagnostics.Process.Start(ReadPdfCmd)

        Catch ex As System.ComponentModel.Win32Exception
            Select Case ex.NativeErrorCode
                Case ERROR_FILE_NOT_FOUND
                    MsgBox(ex.Message + ". Check the path.")
                Case ERROR_ACCESS_DENIED
                    MsgBox(ex.Message + ". You do not have permission to print this file.")
                Case ERROR_NO_ASSOCIATION
                    MsgBox(ex.Message + "." + Chr(13) + "The files is " + ReadPdfCmd + ".")
            End Select
        End Try

    End Sub
    Public Sub mnuHowToolStrip_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuOMAE1.Click

        Dim ReadPdfCmd As String

        ReadPdfCmd = VIVADIR & "\Reference\" & ReferenceFile(3)
        Try
            System.Diagnostics.Process.Start(ReadPdfCmd)

        Catch ex As System.ComponentModel.Win32Exception
            Select Case ex.NativeErrorCode
                Case ERROR_FILE_NOT_FOUND
                    MsgBox(ex.Message + ". Check the path.")
                Case ERROR_ACCESS_DENIED
                    MsgBox(ex.Message + ". You do not have permission to print this file.")
                Case ERROR_NO_ASSOCIATION
                    MsgBox(ex.Message + "." + Chr(13) + "The files is " + ReadPdfCmd + ".")
            End Select
        End Try

    End Sub
    Public Sub mnuInlineVIVATool_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuOMAE2.Click

        Dim ReadPdfCmd As String

        ReadPdfCmd = VIVADIR & "\Reference\" & ReferenceFile(4)
        Try
            System.Diagnostics.Process.Start(ReadPdfCmd)

        Catch ex As System.ComponentModel.Win32Exception
            Select Case ex.NativeErrorCode
                Case ERROR_FILE_NOT_FOUND
                    MsgBox(ex.Message + ". Check the path.")
                Case ERROR_ACCESS_DENIED
                    MsgBox(ex.Message + ". You do not have permission to print this file.")
                Case ERROR_NO_ASSOCIATION
                    MsgBox(ex.Message + "." + Chr(13) + "The files is " + ReadPdfCmd + ".")
            End Select
        End Try

    End Sub
    Public Sub mnuModelingStrakesAndFoilSections_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuQuestionAnswers.Click

        Dim ReadPdfCmd As String

        ReadPdfCmd = VIVADIR & "\Reference\" & ReferenceFile(5)
        Try
            System.Diagnostics.Process.Start(ReadPdfCmd)

        Catch ex As System.ComponentModel.Win32Exception
            Select Case ex.NativeErrorCode
                Case ERROR_FILE_NOT_FOUND
                    MsgBox(ex.Message + ". Check the path.")
                Case ERROR_ACCESS_DENIED
                    MsgBox(ex.Message + ". You do not have permission to print this file.")
                Case ERROR_NO_ASSOCIATION
                    MsgBox(ex.Message + "." + Chr(13) + "The files is " + ReadPdfCmd + ".")
            End Select
        End Try

    End Sub
   
    Public Sub mnuVortexInducedVibrationsOfLongCyclindricalStructures_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuVortexInducedVibrationsOfLongCyclindricalStructures.Click

        Dim ReadPdfCmd As String

        ReadPdfCmd = VIVADIR & "\Reference\" & ReferenceFile(6)
        Try
            System.Diagnostics.Process.Start(ReadPdfCmd)

        Catch ex As System.ComponentModel.Win32Exception
            Select Case ex.NativeErrorCode
                Case ERROR_FILE_NOT_FOUND
                    MsgBox(ex.Message + ". Check the path.")
                Case ERROR_ACCESS_DENIED
                    MsgBox(ex.Message + ". You do not have permission to print this file.")
                Case ERROR_NO_ASSOCIATION
                    MsgBox(ex.Message + "." + Chr(13) + "The files is " + ReadPdfCmd + ".")
            End Select
        End Try

    End Sub
    Public Sub mnuVortexInducedVibrationAnalysis_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

        Dim ReadPdfCmd As String

        ReadPdfCmd = VIVADIR & "\Reference\" & ReferenceFile(8)
        Try
            System.Diagnostics.Process.Start(ReadPdfCmd)

        Catch ex As System.ComponentModel.Win32Exception
            Select Case ex.NativeErrorCode
                Case ERROR_FILE_NOT_FOUND
                    MsgBox(ex.Message + ". Check the path.")
                Case ERROR_ACCESS_DENIED
                    MsgBox(ex.Message + ". You do not have permission to print this file.")
                Case ERROR_NO_ASSOCIATION
                    MsgBox(ex.Message + "." + Chr(13) + "The files is " + ReadPdfCmd + ".")
            End Select
        End Try

    End Sub

    Public Sub mnuRiserTypeandBC_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRiserFRiserTypeandBC.Click
        CurProj.RiserId = 1

        If CurProj.Riser(CurProj.RiserId).RiserType = VIVAMain.RiserTypes.Rigid Then 'JLIU TODO
            Me.mnuRiserFStaticSolution.Enabled = False
        Else
            Me.mnuRiserFStaticSolution.Enabled = True
        End If
        eventArgs.ToString()
        frmRiser.Show()

    End Sub

    Public Sub mnuRiserRTypeandBC_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRiserRRiserTypeandBC.Click
        CurProj.RiserId = 2

        If CurProj.Riser(CurProj.RiserId).RiserType = VIVAMain.RiserTypes.Rigid Then 'JLIU TODO
            Me.mnuRiserRStaticSolution.Enabled = False
        Else
            Me.mnuRiserRStaticSolution.Enabled = True
        End If
        eventArgs.ToString()
        frmRiser.Show()

    End Sub

    Public Sub mnuRiserFOtherProperty_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRiserFOtherProperties.Click
        CurProj.RiserId = 1
        frmOtherRiserProperties.Show()

    End Sub

    Public Sub mnuRiserROtherProperty_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRiserROtherProperties.Click
        CurProj.RiserId = 2
        frmOtherRiserProperties.Show()

    End Sub

    Public Sub mnuRiserFSegmentData_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRiserFSegment.Click
        CurProj.RiserId = 1
        frmSegments.Show()

    End Sub

    Public Sub mnuRiserRSegmentData_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRiserRSegment.Click
        CurProj.RiserId = 2
        frmSegments.Show()

    End Sub

    Public Sub mnuRiserFAuxLine_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRiserFAuxline.Click
        CurProj.RiserId = 1
        frmAuxLines.Show()

    End Sub

    Public Sub mnuRiserRAuxLine_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRiserRAuxLine.Click
        CurProj.RiserId = 2
        frmAuxLines.Show()

    End Sub

    Public Sub mnuRiserFSupSpring_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRiserFSupSpring.Click
        CurProj.RiserId = 1
        frmSupSpring.Show()

    End Sub

    Public Sub mnuRiserRSupSpring_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRiserRSupSpring.Click
        CurProj.RiserId = 2
        frmSupSpring.Show()

    End Sub

    Public Sub mnuRiserFStaticSolution_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRiserFStaticSolution.Click
        CurProj.RiserId = 1
        frmStaticSolution.Show()

    End Sub

    Public Sub mnuRiserRStaticSolution_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRiserRStaticSolution.Click
        CurProj.RiserId = 2
        frmStaticSolution.Show()

    End Sub

    Public Sub mnuAPComputation_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAPComputation.Click

        frmComputation.Show()

    End Sub

    Public Sub mnuResultsF_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuResultsF.Click

        CurProj.RiserId = 1

        If Not HaveResults Then
            MsgBox("You must compute or load results before viewing.", MsgBoxStyle.OkOnly, "WinVIVA - No Data To Tabulate")
        Else
            frmTabulateResults.Show()
        End If

    End Sub

    Public Sub mnuResultsR_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuResultsR.Click

        CurProj.RiserId = 2

        If Not HaveResults Then
            MsgBox("You must compute or load results before viewing.", MsgBoxStyle.OkOnly, "WinVIVA - No Data To Tabulate")
        Else
            frmTabulateResults.Show()
        End If

    End Sub


    Public Sub mnuFLViewFile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

        Dim FatFN As String
        Dim i As Short
        i = CurProj.nRisers * CurProj.RiserId - CurProj.RiserId

        FatFN = CurProj.DataDirectory & OutputFiles(i, fat1_out) 'fat1.out

        If Len(Dir(FatFN)) > 0 Then
            frmDosData.Show()
        Else
            MsgBox(FatFN & " was not found or was empty", MsgBoxStyle.OkOnly, "WinVIVA - No Fatigue Data Found")
        End If

    End Sub


    Public Sub mnuHelpAbout_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuHelpAbout.Click
        frmAbout.Show()
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub mnuBatch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuBatch.Click
        System.Diagnostics.Process.Start(VIVADIR & "WinVIVA Batch.exe")
    End Sub

    Private Sub mnuDatabase_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuDatabase.Click
        frmDBConfig.Show()
    End Sub

    Private Sub mnuFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFile.Click

    End Sub

    'Private Sub QuestionsAnswersToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuQuestionsAnswers.Click

    'End Sub
End Class