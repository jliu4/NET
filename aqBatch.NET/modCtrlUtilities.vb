Option Strict Off
Option Explicit On
Module modCtrlUtilities
	' modCtrlUtilities  user defined control utilities
	' Version 2.0
	' 2001, Copyright DTCEL, All Rights Reserved
	
	' CheckBox
	' MSFlexGrid
	' General
	
	Public Const Knots2Ftps As Double = 1.6878098571
	Public Const Ftps2Knots As Double = 0.5924838013
	
	' ---------------------------------- consttants ----------------
	Public WaterDepth As Double
	Public oMet() As Metocean
	Public IsDamaged As Short
	Public BreakLine() As Short
	Public WorkDir, rWorkDir As String
	Public UDWS As String ' User defined wind spectrum data string
    Public WSPEC As String

    Public NumCases As Short
	Public L1Tmax() As Short
	Public L2Tmax() As Short
	
	'-----------------------------------------------------------------
	'Keyboard constants
	Private Const DotKey As Short = 190
	Private Const EqualKey As Short = 187
	Private Const MinusKey As Short = 189
	Private Const ShiftKey As Short = 16
	
	Private Const Margin As Short = 50

    Public Sub ReportError(ByRef err_Renamed As ErrObject, ByVal Title As String, ByVal msg0 As String)
        Dim msg As String
        msg = msg0 & vbCrLf & Err.Description & vbCrLf & err_Renamed.Source
        MsgBox(msg, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Error")
    End Sub

    Public Function FileRootName(ByVal fname As String) As Object
		Dim pos1, pos2 As Short
		pos1 = InStrRev(fname, "\")
		pos2 = InStrRev(fname, ".")
		If pos1 > 0 And pos2 > 0 Then
			FileRootName = Mid(fname, pos1 + 1, pos2 - pos1 - 1)
		Else
            FileRootName = ""
        End If
	End Function
	
	Public Function FileExtName(ByVal fname As String) As Object
		Dim pos2 As Short
		pos2 = InStr(1, fname, ".")
		If pos2 > 0 Then
			FileExtName = Mid(fname, pos2 + 1)
		Else
            FileExtName = ""
        End If
	End Function
	
	' CheckBox Utilities
	
	Public Sub SetCheckBox(ByRef CheckBox As System.Windows.Forms.CheckBox, ByVal Value As Boolean)
		
		If Value Then
			CheckBox.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			CheckBox.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		
	End Sub
	
	Public Function CheckBoxValue(ByRef CheckBox As System.Windows.Forms.CheckBox) As Boolean
		
		If CheckBox.CheckState = System.Windows.Forms.CheckState.Checked Then
			CheckBoxValue = True
		Else
			CheckBoxValue = False
		End If
		
	End Function



    ' evoke edit box

    Sub MSFlexGridEdit(ByRef MSFlexGrid As System.Windows.Forms.Control, ByRef EditBox As System.Windows.Forms.TextBox, ByRef KeyCode As Short)
		
		
		'   Use the character that was typed
		Select Case KeyCode
			
			'       The F2 key means edit the current text
			Case System.Windows.Forms.Keys.F2
				EditBox.Text = MSFlexGrid.ToString()
				
				'       The F10 key means edit the current text - mouse double click
			Case System.Windows.Forms.Keys.F10
				EditBox.Text = MSFlexGrid.ToString()
				
				'       Anything else means replace the current text
			Case Else
				EditBox.Text = Chr(KeyCode)
		End Select
		SelectText(EditBox)
        '   Show EditBox at the right place
        'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.CellHeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.CellWidth. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.CellTop. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.CellLeft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'EditBox.SetBounds(VB6.TwipsToPixelsX(MSFlexGrid.CellLeft + VB6.PixelsToTwipsX(MSFlexGrid.Left)), VB6.TwipsToPixelsY(MSFlexGrid.CellTop + VB6.PixelsToTwipsY(MSFlexGrid.Top)), VB6.TwipsToPixelsX(MSFlexGrid.CellWidth), VB6.TwipsToPixelsY(MSFlexGrid.CellHeight))
        EditBox.Visible = True
		
		'   And let it work
		EditBox.Focus()
		
	End Sub
	
	Public Sub SelectText(ByRef EditBox As System.Windows.Forms.TextBox)
		EditBox.SelectionStart = 0
		EditBox.SelectionLength = Len(EditBox.Text)
	End Sub
	
	
	' evoke combo boxes
	
	Public Sub MSFlexGridCombo(ByRef MSFlexGrid As System.Windows.Forms.Control, ByRef Combo As System.Windows.Forms.ComboBox, Optional ByRef SetFocus As Boolean = False)

        If SetFocus Then
            Combo.Visible = True
            Combo.Focus()
        End If

    End Sub



    ' insert a row

    Public Sub AddRow(ByRef dataGridView1 As DataGridView, ByRef CheckingGrid As Boolean, Optional ByRef MaxRows As Short = 0)

        Dim CurRow As Short

        CheckingGrid = True
        With dataGridView1
            If .Rows.Count = 1 Then
                CurRow = .Rows.Add
            Else
                CurRow = .CurrentCell.RowIndex
            End If
            Try
                .Rows.Insert(CurRow, New String() {})
            Catch ex As Exception
            End Try


        End With
        NumberAllRows(dataGridView1)
        CheckingGrid = False

    End Sub

    ' delete a row

    Public Sub DeleteRow(ByRef dataGridView1 As DataGridView, ByRef CheckingGrid As Boolean, Optional ByRef MinRows As Short = 0)
        CheckingGrid = True
        With dataGridView1
            Dim CurRow As Short = .CurrentCell.RowIndex
            If CurRow < .RowCount - 1 Then
                .Rows.RemoveAt(CurRow)

            End If

        End With
        NumberAllRows(dataGridView1)

        CheckingGrid = False

    End Sub

    Public Sub copyToClipBoard(ByRef dataGridView1 As DataGridView)

        'dataGridView1.SelectAll()
        Dim dataObj As DataObject = dataGridView1.GetClipboardContent()

        Clipboard.SetDataObject(dataObj, True)
        dataGridView1.ClearSelection()

    End Sub

    Public Sub PastefromClipBoardToDataGridView(ByRef dataGridView1 As DataGridView, ByRef CheckingGrid As Boolean, Optional ByRef MaxRows As Short = 0)

        Dim rowSplitter As Char() = {vbCr, vbLf}
        Dim columnSplitter As Char() = {vbTab}
        Dim maxAllowableRows As Short
        CheckingGrid = True
        'get the text from clipboard

        Dim dataInClipboard As IDataObject = Clipboard.GetDataObject()
        Dim stringInClipboard As String = CStr(dataInClipboard.GetData(DataFormats.Text))
        If stringInClipboard Is Nothing Then Exit Sub
        'split it into lines
        Dim rowsInClipboard As String() = stringInClipboard.Split(rowSplitter, StringSplitOptions.RemoveEmptyEntries)

        'get the row and column of selected cell in grid
        Dim r As Integer = dataGridView1.SelectedCells(0).RowIndex
        Dim c As Integer = dataGridView1.SelectedCells(0).ColumnIndex

        'add rows into grid to fit clipboard lines
        maxAllowableRows = r + rowsInClipboard.Length
        If MaxRows > 0 And r + rowsInClipboard.Length > MaxRows Then
            maxAllowableRows = MaxRows
        End If
        'copy part of columns, rows can not be added.
        If c > 0 And maxAllowableRows > dataGridView1.Rows.Count - 1 Then
            maxAllowableRows = dataGridView1.Rows.Count - 1
        End If
        If (dataGridView1.Rows.Count - 1 < maxAllowableRows) Then
            dataGridView1.Rows.Add(maxAllowableRows - dataGridView1.Rows.Count + 1)
        End If
        ' loop through the lines, split them into cells and place the values in the corresponding cell.
        Dim iRow As Integer = 0
        While iRow < maxAllowableRows - r
            'split row into cell values
            Dim valuesInRow As String() = rowsInClipboard(iRow).Split(columnSplitter)
            'cycle through cell values
            Dim iCol As Integer = 0
            While iCol < valuesInRow.Length
                'assign cell value, only if it within columns of the grid
                If (dataGridView1.ColumnCount - 1 >= c + iCol) Then
                    'if empty, default to 1
                    'If Not IsNumeric(valuesInRow(iCol)) Then
                    'valuesInRow(iCol) = 1
                    'End If
                    dataGridView1.Rows(r + iRow).Cells(c + iCol).Value = valuesInRow(iCol)
                End If

                iCol += 1
            End While
            iRow += 1
        End While
        CheckingGrid = True
    End Sub
    Public Sub NumberAllRows(ByRef dataGridView1 As DataGridView)
        ' Add row headers.
        For i As Integer = 0 To dataGridView1.Rows.Count - 1
            dataGridView1.Rows(i).HeaderCell.Value = (i + 1).ToString()
        Next i
    End Sub




    ' General

    Public Function CheckData(ByVal Entry As String, Optional ByVal Zero As String = "0", Optional ByVal NoWarning As Boolean = False) As String
		
		Dim msg As String
		
		Entry = Trim(Entry)
		
		If Entry = "" Then Entry = Zero
		Do While Not IsNumeric(Entry)
			If NoWarning Then
				Entry = Zero
			Else
				msg = "'" & Entry & "' is not a valid numberic input. Please re-enter here:"
				Entry = InputBox(msg)
				If Entry = "" Then Entry = Zero
			End If
		Loop 
		
		CheckData = Entry
		
	End Function
End Module