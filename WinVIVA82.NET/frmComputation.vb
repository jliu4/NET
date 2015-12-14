Option Strict Off
Option Explicit On

' Version 8.1
' 2014, Copyright DTCEL, All Rights Reserved
' Jin Liu May 2014

Friend Class frmComputation

    Inherits System.Windows.Forms.Form


    Private Sub frmComputation_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        Text = Text & " - " & CurProj.Title
        sbrStatus.Text = "Waiting to begin VIVA execution..."

        'check iteration number
        If CurProj.NumIterations < MinIteration Then
            txtIterations.Text = CStr(MinIteration)
        Else
            txtIterations.Text = CStr(CurProj.NumIterations)
        End If

        'check point number
        If CurProj.NumPoints < MinVIVACorePoints Then
            txtNumPoints.Text = CStr(MinVIVACorePoints)
        Else
            txtNumPoints.Text = CStr(CurProj.NumPoints)
        End If

        'dos windows
        If CurProj.ShowDOSBox Then
            optDOSBoxVisible.Checked = True
            optDOSBoxHidden.Checked = False
        Else
            optDOSBoxVisible.Checked = True
            optDOSBoxHidden.Checked = False
        End If

    End Sub


    Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click

        Me.Close()

    End Sub


    Sub btnExecute_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnExecute.Click

        Dim i, j As Short
        On Error GoTo ErrHandler
        Cursor = System.Windows.Forms.Cursors.WaitCursor
        RemovePreviousVIVAOutputs()
        CopyImportedFiles()

        If UpdateProject() Then
            'check for files we need, and bail out if they aren't found
            'ChDir(VIVADIR)

            If Not VIVAFilesExist() Then
                sbrStatus.Text = "Cannot execute VIVA programs"
                Cursor = System.Windows.Forms.Cursors.Default
                Exit Sub
            End If

            'write the model in the format needed by VIVACORE
            sbrStatus.Text = "Preparing VIVA data..."
            If WriteDOSVIVAInput() Then
                sbrStatus.Text = "Ready to run VIVA programs"
            Else
                sbrStatus.Text = "Error in data preparation"
                Cursor = System.Windows.Forms.Cursors.Default
                Exit Sub
            End If

            'On Error Resume Next
            ChDrive(VIVADIR)
            ChDir(VIVADIR)

            For i = 1 To NumVIVADOSPrograms
                'update the status for the user
                sbrStatus.Text = "Executing VIVA Program:  " & VIVADOSPrograms(CurProj.nRisers, i)

                If CurProj.nRisers = 2 And Not My.Computer.FileSystem.FileExists(VIVADIR & VIVADOSPrograms(2, 1)) Then

                    MsgBox("You do not have Vivarray program.")
                    Exit Sub
                End If
                'DoEvents

                'this is the actual execution step
                'MsgBox(VIVADIR & VIVADOSPrograms(CurProj.nRisers, i))
                ExecCmd(VIVADIR & VIVADOSPrograms(CurProj.nRisers, i), (optDOSBoxHidden.Checked))
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
                HaveResults = ReadVIVACoreOutputFiles()
                'tell the user the good news
            End If

            If HaveResults Then
                sbrStatus.Text = "Execution is completed"
            Else
                sbrStatus.Text = "Abnormal Completion of VIVA Programs"
            End If
        Else
            sbrStatus.Text = "Error in data preparation"
        End If

        Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub
ErrHandler:

        MsgBox("Error when excute file. The data may be imcomplete or corrupted", MsgBoxStyle.OkOnly, " WinVIVA - Read Input File Error")
        Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub

    End Sub



    Private Sub btnWriteFiles_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnWriteFiles.Click

        Cursor = System.Windows.Forms.Cursors.WaitCursor

        If UpdateProject() Then
            sbrStatus.Text = "Preparing VIVA data..."

            'write the model in the format needed by VIVA
            'ChDir VIVADIR
            If WriteDOSVIVAInput() Then
                CopyVIVAInputs()
                sbrStatus.Text = "Ready to run VIVA programs"
            Else
                sbrStatus.Text = "Error in data preparation"
            End If
        Else
            sbrStatus.Text = "Error in data preparation"
        End If

        Cursor = System.Windows.Forms.Cursors.Default

    End Sub


    Private Function UpdateProject() As Boolean

        Dim ret As String

        'check iteration number
        If Not IsNumeric(txtIterations.Text) Then
            'no input, set to 30
            MsgBox("You haven't input Interation Number. It is set to " & CStr(MinIteration) & ".", MsgBoxStyle.OkOnly, "WinVIVA -  Iteration Number")
            txtIterations.Text = CStr(MinIteration)
        End If

        If CShort(txtIterations.Text) < MinIteration Then
            'input less than 30, reset it
            MsgBox("Your Interation Number input is increased to " & CStr(MinIteration) & ".", MsgBoxStyle.OkOnly, "WinVIVA -  Iteration Number")
            txtIterations.Text = CStr(MinIteration)
        End If

        If CurProj.NumIterations <> CShort(txtIterations.Text) Then
            'update curproj
            CurProj.NumIterations = CShort(txtIterations.Text)
            CurProj.Saved = False
        End If

        'check point no
        If Not IsNumeric(txtNumPoints.Text) Then
            'no input, set it to 200
            MsgBox("You haven't input Point Number. It is set to " & CStr(MinVIVACorePoints) & ".", MsgBoxStyle.OkOnly, "WinVIVA -  Point Number")
            txtNumPoints.Text = CStr(MinVIVACorePoints)
        End If

        If CShort(txtNumPoints.Text) < MinVIVACorePoints Then
            'input less than 200, reset it
            MsgBox("Your Point Number input is increased to " & CStr(MinVIVACorePoints) & ".", MsgBoxStyle.OkOnly, "WinVIVA -  Point Number")
            txtNumPoints.Text = CStr(MinVIVACorePoints)
        End If

        If CShort(txtNumPoints.Text) > NumVIVACorePoints Then
            'input greater than max, reset it
            MsgBox("Your Point Number input is limited to " & CStr(NumVIVACorePoints) & ".", MsgBoxStyle.OkOnly, "WinVIVA -  Point Number")
            txtNumPoints.Text = CStr(NumVIVACorePoints)
        End If

        If CurProj.NumPoints <> CShort(txtNumPoints.Text) Then
            'update curproj
            CurProj.NumPoints = CShort(txtNumPoints.Text)
            CurProj.Saved = False
        End If

        'always multifreq
        CurProj.FrequencyResponse = VIVAMain.FrequencyResponseValues.MultiFreq

        'dos windows
        If optDOSBoxHidden.Checked Then
            CurProj.ShowDOSBox = False
        Else
            CurProj.ShowDOSBox = True
        End If

        ret = CurProj.Complete

        If ret <> "" Then
            MsgBox(ret, MsgBoxStyle.OkOnly, "WinVIVA - Check Model Data")
            UpdateProject = False
        Else
            UpdateProject = True
        End If

    End Function

End Class