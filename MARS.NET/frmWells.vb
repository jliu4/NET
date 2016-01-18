Option Strict Off
Option Explicit On
Friend Class frmWells
	Inherits System.Windows.Forms.Form
	' frmWells          form to input/edit well site information
	' Version 1.0
	' 2001, Copyright DTCEL, All Rights Reserved
	
	
	' for grid operation
	Private NumWells As Short
	
	Private Changed As Boolean
	Private ExistingTxt As String
	Private CheckingGrid As Boolean
	Private JustEnterCell As Boolean
	Dim LFactor, FrcFactor As Single
	' load form
	
	Private Sub frmWells_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim LUnit, FrcUnit As String
		If IsMetricUnit Then
			LUnit = "(m)"
			FrcUnit = "(KN)"
		Else
			LUnit = "(ft)"
			FrcUnit = "(kips)"
		End If
		
		Dim i As Short

        Dim WSLabels(4) As String
		
		CheckingGrid = True
		
		WSLabels(0) = "No."
		WSLabels(1) = "Name/ID"
		WSLabels(2) = "X (E) " & LUnit
		WSLabels(3) = "Y (N) " & LUnit
		WSLabels(4) = "Depth " & LUnit
		
		NumWells = CurField.Count

        With grdWS
            '       initiate grid
            .RowCount = MaxNumWells
            .ColumnCount = 4
            For i = 0 To .RowCount - 1
                .Rows(i).HeaderCell.Value = (i + 1).ToString
            Next
            For i = 0 To .ColumnCount - 1
                .Columns(i).Width = 70
                .Columns(i).HeaderText = WSLabels(i + 1)
            Next

            '       load data
            If CurField.Count <> 0 Then LoadData()
		End With
		
		CheckingGrid = False
		Changed = False
		
	End Sub
	
	' command buttons
	
	Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click
		
		Me.Close()
		
	End Sub
	
	Private Sub btnOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnOK.Click
		
		Cursor = System.Windows.Forms.Cursors.WaitCursor
		If Not SaveData Then
			Cursor = System.Windows.Forms.Cursors.Default
			Exit Sub
		End If
		Cursor = System.Windows.Forms.Cursors.Default
		CurProj.Saved = False
		Changed = False
		Me.Close()
		
	End Sub
	
	Private Sub frmWells_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
		If Changed Then
			If MsgBox("Data changed. Save?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Warning") = MsgBoxResult.Yes Then
				SaveData()
			End If
		End If
		eventArgs.Cancel = Cancel
	End Sub

    ' grid

    'Private Sub grdWS_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSFlexGridLib.DMSFlexGridEvents_MouseDownEvent)

    'If eventArgs.button = VB6.MouseButtonConstants.RightButton Then
    'UPGRADE_ISSUE: Form method frmWells.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    'PopupMenu(mnuGridEdit)
    ' End If

    ' End Sub

    Public Sub mnuAddRow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAddRow.Click

        AddRow(grdWS, CheckingGrid)

    End Sub

    Public Sub mnuDeleteRow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDeleteRow.Click

        DeleteRow(grdWS, CheckingGrid)

    End Sub

    Private Sub grdWS_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

        ' MSFlexGridEdit(grdWS, txtEdit, System.Windows.Forms.Keys.F10)

    End Sub

    Private Sub grdWS_EnterCell(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

        ExistingTxt = grdWS.Text
        JustEnterCell = True

    End Sub

    '  Private Sub grdWS_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSFlexGridLib.DMSFlexGridEvents_KeyDownEvent)

    ' KeyHandler(grdWS, txtEdit, eventArgs.keyCode, eventArgs.shift, JustEnterCell, ExistingTxt)

    ' End Sub

    Private Sub grdWS_LeaveCell(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

        '   If txtEdit.Visible = True Then
        '  grdWS.Text = txtEdit.Text
        ' txtEdit.Visible = False
        'End If

        If Not CheckingGrid Then
            With grdWS
                'If Trim(.Text) <> ExistingTxt And .ColumnCount > 1 Then
                'If Trim(.Text) <> "" Then .Text = CheckData(.Text)
                '  Changed = True
                'End If
            End With
        End If

    End Sub

    Private Sub txtEdit_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs)
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000

        'EditKeyCode(grdWS, txtEdit, KeyCode, Shift)
        Changed = True
    End Sub

    Private Sub txtEdit_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs)
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        If KeyAscii = Asc(vbCr) Then KeyAscii = 0

        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    ' menus

    Public Sub mnuExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuExit.Click
		
		Me.Close()
		
	End Sub
	
	Public Sub mnuFileOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileOpen.Click

        Dim FileOpen_Renamed As Boolean

        '   should the user cancel the dialog box, exit
        On Error GoTo ErrHandler
		
		FileOpen_Renamed = False

        '   set filters to allow selection of all files or just .WSL
        dlgFilesOpen.Filter = "All Files (*.*)|*.*|Current Files (*.wsl)|*.wsl"
        dlgFilesSave.Filter = "All Files (*.*)|*.*|Current Files (*.wsl)|*.wsl"
		'   specify default filter as *.WSL
		dlgFilesOpen.FilterIndex = 2
		dlgFilesSave.FilterIndex = 2
		dlgFilesOpen.InitialDirectory = CurProj.Directory
		dlgFilesSave.InitialDirectory = CurProj.Directory
        dlgFilesOpen.ShowReadOnly = False
        dlgFilesOpen.CheckFileExists = True
        dlgFilesOpen.CheckPathExists = True
		dlgFilesSave.CheckPathExists = True
		'   display the Open dialog box
		dlgFilesOpen.ShowDialog()
		dlgFilesSave.FileName = dlgFilesOpen.FileName
		'   open the file
		FileOpen(1, dlgFilesOpen.FileName, OpenMode.Input, OpenAccess.Read)
		FileOpen_Renamed = True
		
		'   build the object from the input data
		Cursor = System.Windows.Forms.Cursors.WaitCursor
		CheckingGrid = True
		ImportData((1))
		CheckingGrid = False
		Cursor = System.Windows.Forms.Cursors.Default
		
		'   close the input file and return
		FileClose(1)
		FileOpen_Renamed = False
		
		Exit Sub
		
ErrHandler: 
		'   user pressed Cancel button
		If FileOpen_Renamed Then FileClose(1)
		Cursor = System.Windows.Forms.Cursors.Default
		
	End Sub
	
	Public Sub mnuFileSaveAs_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileSaveAs.Click

        Dim FileOpen_Renamed As Boolean

        '   should the user cancel the dialog box, exit
        On Error GoTo ErrHandler
		
		FileOpen_Renamed = False

        '   set filters to allow selection of all files or just .WSL
        dlgFilesOpen.Filter = "All Files (*.*)|*.*|Current Files (*.wsl)|*.wsl"
        dlgFilesSave.Filter = "All Files (*.*)|*.*|Current Files (*.wsl)|*.wsl"
		'   specify default filter as *.WSL
		dlgFilesOpen.FilterIndex = 2
		dlgFilesSave.FilterIndex = 2
		'   specify the current filename
		dlgFilesOpen.FileName = "*.wsl"
		dlgFilesSave.FileName = "*.wsl"
		dlgFilesOpen.InitialDirectory = CurProj.Directory
		dlgFilesSave.InitialDirectory = CurProj.Directory
        dlgFilesOpen.ShowReadOnly = False
        dlgFilesSave.OverwritePrompt = True
        '   display the Open dialog box
        dlgFilesSave.ShowDialog()
		dlgFilesOpen.FileName = dlgFilesSave.FileName
		'   open the file for output
		Dim fnum As Integer
		fnum = FreeFile
		FileOpen(fnum, dlgFilesOpen.FileName, OpenMode.Output)
		FileOpen_Renamed = True
		
		Cursor = System.Windows.Forms.Cursors.WaitCursor
		'   output
		CheckingGrid = True
		ExportData((fnum))
		CheckingGrid = False
		Cursor = System.Windows.Forms.Cursors.Default
		'   close the output file
		FileClose(fnum)
		FileOpen_Renamed = False
		
		Exit Sub
		
ErrHandler: 
		'   user pressed Cancel button
		If FileOpen_Renamed Then FileClose(fnum)
		Cursor = System.Windows.Forms.Cursors.Default
		
	End Sub

    ' project data operations
    Private Sub LoadData()

        If IsMetricUnit Then
            LFactor = 0.3048 ' ft -> m
            FrcFactor = 4.448222 ' kips -> KN
        Else
            LFactor = 1
            FrcFactor = 1
        End If

        Dim WSD As Well
        Dim r As Short

        With grdWS
            For r = 0 To CurField.Count - 1
                WSD = CurField.Item(r + 1)
                If r > MaxNumWells Then Exit For
                .Rows(r).Cells(0).Value = WSD.NameID
                .Rows(r).Cells(1).Value = VB6.Format(WSD.Xg * LFactor, "0.00")
                .Rows(r).Cells(2).Value = VB6.Format(WSD.Yg * LFactor, "0.00")
                .Rows(r).Cells(3).Value = VB6.Format(WSD.Depth * LFactor, "0.00")


            Next

        End With

    End Sub

    Private Function SaveData() As Boolean
		If IsMetricUnit Then
			LFactor = 0.3048 ' ft -> m
			FrcFactor = 4.448222 ' kips -> KN
		Else
			LFactor = 1
			FrcFactor = 1
		End If
		
		Dim i As Short
        Dim Y, Depth, X As Single
        Dim NameID As String

        'grdWS.RowCount = 0

        '   If Changed Then        '  make sure save when user clicks save no matter what
        '       delete previous input
        CurField.Clear()
		
		'       avoid triggering the "Leave Cell" event
		CheckingGrid = True
		'       insert current input
		With grdWS
            For i = 0 To .RowCount - 1

                If .Rows(i).Cells(0).Value = "" Then Exit For
                NameID = .Rows(i).Cells(0).Value

                Y = CDbl(.Rows(i).Cells(2).Value) / LFactor

                Depth = CDbl(CDbl(.Rows(i).Cells(3).Value) / LFactor) / LFactor
                CurField.Add(NameID, X, Y, Depth)
            Next i
        End With
		'       reset the "Leaving Cell" trigger
		CheckingGrid = False

        SaveData = True
		
	End Function
	
	Private Sub ImportData(ByRef FileNum As Short)
		If IsMetricUnit Then
			LFactor = 0.3048 ' ft -> m
			FrcFactor = 4.448222 ' kips -> KN
		Else
			LFactor = 1
			FrcFactor = 1
		End If
		
		Dim r, c As Short
		Dim X, Depth, Y As Single
		Dim NameID As String
		
		Changed = True

        With grdWS
            r = 0
            Do While Not EOF(FileNum)
                Input(FileNum, NameID)
                Input(FileNum, X)
                Input(FileNum, Y)
                Input(FileNum, Depth)
                If r > MaxNumWells Then Exit Do

                .Rows(r).Cells(0).Value = NameID

                .Rows(r).Cells(1).Value = CStr(X * LFactor)

                .Rows(r).Cells(2).Value = CStr(Y * LFactor)

                .Rows(r).Cells(3).Value = CStr(Depth * LFactor)
                r = r + 1
            Loop
        End With

    End Sub
	
	Private Sub ExportData(ByRef FileNum As Short)
		If IsMetricUnit Then
			LFactor = 0.3048 ' ft -> m
			FrcFactor = 4.448222 ' kips -> KN
		Else
			LFactor = 1
			FrcFactor = 1
		End If
		
		Dim i As Short
		Dim X, Depth, Y As Single
		Dim NameID As String
		
		With grdWS
            For i = 0 To .RowCount - 1

                If .Rows(i).Cells(0).Value = "" Then Exit For
                NameID = .Rows(i).Cells(0).Value

                X = CDbl(.Rows(i).Cells(1).Value) / LFactor


                Y = CDbl(.Rows(i).Cells(2).Value) / LFactor

                Depth = CDbl(.Rows(i).Cells(3).Value) / LFactor
                WriteLine(FileNum, NameID, X, Y, Depth)
            Next i
        End With
		
	End Sub
End Class