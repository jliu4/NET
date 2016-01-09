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
	
	'UPGRADE_NOTE: err was upgraded to err_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub ReportError(ByRef err_Renamed As ErrObject, ByVal Title As String, ByVal msg0 As String)
		Dim msg As String
		msg = msg0 & vbCrLf & Err.Description & vbCrLf & err_Renamed.Source
		MsgBox(msg, MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Error")
	End Sub
	
	Public Function FileRootName(ByVal fname As String) As Object
		Dim pos1, pos2 As Short
		pos1 = InStrRev(fname, "\")
		pos2 = InStrRev(fname, ".")
		If pos1 > 0 And pos2 > 0 Then
			FileRootName = Mid(fname, pos1 + 1, pos2 - pos1 - 1)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object FileRootName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FileRootName = ""
		End If
	End Function
	
	Public Function FileExtName(ByVal fname As String) As Object
		Dim pos2 As Short
		pos2 = InStr(1, fname, ".")
		If pos2 > 0 Then
			FileExtName = Mid(fname, pos2 + 1)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object FileExtName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
	
	' MSFlexGrid Utilities
	
	' key event when in the edit box
	
	Public Sub EditKeyCode(ByRef MSFlexGrid As System.Windows.Forms.Control, ByRef EditBox As System.Windows.Forms.TextBox, ByRef KeyCode As Short, ByRef Shift As Short)
		
		' Standard edit control processing
		Select Case KeyCode
			
			Case System.Windows.Forms.Keys.Escape ' ESC: cancel changes, return focus to MSFlexGrid.
				EditBox.Visible = False
				MSFlexGrid.Focus()
				
			Case System.Windows.Forms.Keys.Return ' ENTER: accept changes, return focus to MSFlexGrid.
				'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MSFlexGrid = EditBox.Text
				EditBox.Visible = False
				MSFlexGrid.Focus()
				
			Case System.Windows.Forms.Keys.Up ' Up: accept changes, move up
				MSFlexGrid.Focus()
				System.Windows.Forms.Application.DoEvents()
				'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.FixedRows. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.Row. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If MSFlexGrid.Row > MSFlexGrid.FixedRows Then
					'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.Row. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					MSFlexGrid.Row = MSFlexGrid.Row - 1
				End If
				
			Case System.Windows.Forms.Keys.Down ' Down: accept changes, move down
				MSFlexGrid.Focus()
				System.Windows.Forms.Application.DoEvents()
				'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.Rows. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.Row. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If MSFlexGrid.Row < MSFlexGrid.Rows - 1 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.Row. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					MSFlexGrid.Row = MSFlexGrid.Row + 1
				End If
				
		End Select
		
	End Sub
	
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
		EditBox.SetBounds(VB6.TwipsToPixelsX(MSFlexGrid.CellLeft + VB6.PixelsToTwipsX(MSFlexGrid.Left)), VB6.TwipsToPixelsY(MSFlexGrid.CellTop + VB6.PixelsToTwipsY(MSFlexGrid.Top)), VB6.TwipsToPixelsX(MSFlexGrid.CellWidth), VB6.TwipsToPixelsY(MSFlexGrid.CellHeight))
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
		
		'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.CellWidth. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.CellTop. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.CellLeft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Combo.SetBounds(VB6.TwipsToPixelsX(MSFlexGrid.CellLeft + VB6.PixelsToTwipsX(MSFlexGrid.Left)), VB6.TwipsToPixelsY(MSFlexGrid.CellTop + VB6.PixelsToTwipsY(MSFlexGrid.Top)), VB6.TwipsToPixelsX(MSFlexGrid.CellWidth), 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y Or Windows.Forms.BoundsSpecified.Width)
		If SetFocus Then
			Combo.Visible = True
			Combo.Focus()
		End If
		
	End Sub
	
	' key handler in editing
	
	Public Sub KeyHandler(ByRef MSFlexGrid As System.Windows.Forms.Control, ByRef EditBox As System.Windows.Forms.Control, ByRef KeyCode As Short, ByRef Shift As Short, ByRef JustEnterCell As Boolean, ByRef ExistingTxt As String, Optional ByRef MaxRowsFixed As Boolean = True)
		
		Dim StrLen, CurCol As Short
		
		With MSFlexGrid
			Select Case KeyCode
				
				Case System.Windows.Forms.Keys.Return, System.Windows.Forms.Keys.Down, System.Windows.Forms.Keys.Up
					'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.Rows. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.Row. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .Row < .Rows - 1 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.Row. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.Row = .Row + 1
					Else
						If Not MaxRowsFixed Then
							'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.Rows. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							.Rows = .Rows + 1
							'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.Row. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.Rows. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							.Row = .Rows - 1
							'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.Row. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.RowHeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							.RowHeight(.Row) = .RowHeight(.Row - 1)
							'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.FixedCols. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If .FixedCols > 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								CurCol = .Col
								'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								.Col = 0
								'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.CellAlignment. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
								'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.FixedRows. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.Rows. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								.Text = CStr(.Rows - .FixedRows)
								'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								.Col = CurCol
							End If
							.Refresh()
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.Row. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.Rows. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							.Row = .Rows - 1
						End If
					End If
					
				Case System.Windows.Forms.Keys.Tab
					If Shift = 1 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.FixedCols. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If .Col > .FixedCols - 1 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							.Col = .Col - 1
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.Cols. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							.Col = .Cols - 1
						End If
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.Cols. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If .Col < .Cols - 1 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							.Col = .Col + 1
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object MSFlexGrid.FixedCols. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							.Col = .FixedCols
						End If
					End If
					
				Case System.Windows.Forms.Keys.Back
					StrLen = Len(MSFlexGrid.Text) - 1
					If StrLen < 0 Then StrLen = 0
					.Text = Left(.Text, StrLen)
					JustEnterCell = False
					
				Case System.Windows.Forms.Keys.Delete
					.Text = ""
					
				Case System.Windows.Forms.Keys.Escape
					.Text = ExistingTxt
					
				Case System.Windows.Forms.Keys.F2
					MSFlexGridEdit(MSFlexGrid, EditBox, KeyCode)
					
				Case System.Windows.Forms.Keys.Add, EqualKey, System.Windows.Forms.Keys.Subtract, MinusKey, System.Windows.Forms.Keys.Decimal, DotKey
					If Shift = 0 Or (Shift = 1 And KeyCode = EqualKey) Then
						Select Case KeyCode
							Case System.Windows.Forms.Keys.Add, EqualKey
								KeyCode = 43 '"+"
							Case System.Windows.Forms.Keys.Subtract, MinusKey
								KeyCode = 45 '"-"
							Case System.Windows.Forms.Keys.Decimal, DotKey
								KeyCode = 46 '"."
						End Select
						If JustEnterCell = True Then
							.Text = Chr(KeyCode)
							JustEnterCell = False
						Else
							.Text = .Text & Chr(KeyCode)
						End If
					End If
					
				Case System.Windows.Forms.Keys.NumPad0 To System.Windows.Forms.Keys.NumPad9
					If Shift = 0 Then
						If JustEnterCell = True Then
							.Text = Chr(KeyCode - &H30s)
							JustEnterCell = False
						Else
							.Text = MSFlexGrid.Text & Chr(KeyCode - &H30s)
						End If
					End If
					
				Case System.Windows.Forms.Keys.D0 To System.Windows.Forms.Keys.D9, System.Windows.Forms.Keys.Space
					If Shift = 0 Or Shift = 1 Then
						If JustEnterCell = True Then
							.Text = Chr(KeyCode)
							JustEnterCell = False
						Else
							.Text = .Text & Chr(KeyCode)
						End If
					End If
					
				Case System.Windows.Forms.Keys.A To System.Windows.Forms.Keys.Z
					If Shift = 0 Then KeyCode = KeyCode + 32
					
					If JustEnterCell = True Then
						.Text = Chr(KeyCode)
						JustEnterCell = False
					Else
						.Text = .Text & Chr(KeyCode)
					End If
			End Select
		End With
		
	End Sub
	
	' insert a row
	
	Public Sub AddRow(ByRef MSFlexGrid As AxMSFlexGridLib.AxMSFlexGrid, ByRef CheckingGrid As Boolean, Optional ByRef MaxRowsFixed As Boolean = True, Optional ByRef CurRow As Short = 0, Optional ByRef CurCol As Short = 0)
		
		Dim R, c As Short
		Dim Entry As String
		
		CheckingGrid = True
		With MSFlexGrid
			.Redraw = False
			
			'       Remember the current position
			'        If IsMissing(CurRow) Then CurRow = .Row
			'        If IsMissing(CurCol) Then CurCol = .Col
			CurRow = .Row
			CurCol = .Col
			If CurRow < .FixedRows Then CurRow = .FixedRows
			If CurCol < .FixedCols Then CurCol = .FixedCols
			
			If Not MaxRowsFixed Then
				.Rows = .Rows + 1
				.Row = .Rows - 1
				.set_RowHeight(.Row, .get_RowHeight(.Row - 1))
				
				If .FixedCols > 0 Then
					.Col = 0
					.Row = .Rows - 1
					.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
					.Text = CStr(.Rows - .FixedRows)
				End If
			End If
			
			'       Do each column in the grid
			For c = .FixedCols To .Cols - 1
				.Col = c
				'           Move upward from the bottom of the table, shifting as you go
				For R = .Rows - 2 To CurRow Step -1
					.Row = R
					Entry = .Text
					.Text = ""
					.Row = R + 1
					.Text = Entry
				Next R
			Next c
			
			.Row = CurRow
			.Col = CurCol
			.Redraw = True
		End With
		CheckingGrid = False
		
	End Sub
	
	' delete a row
	
	Public Sub DeleteRow(ByRef MSFlexGrid As AxMSFlexGridLib.AxMSFlexGrid, ByRef CheckingGrid As Boolean, Optional ByRef MinRowsFixed As Boolean = True, Optional ByRef CurRow As Short = 0, Optional ByRef CurCol As Short = 0)
		
		Dim R, c As Short
		Dim Entry As String
		
		CheckingGrid = True
		With MSFlexGrid
			.Redraw = False
			
			'       Remember the current position
			'        If IsMissing(CurRow) Then CurRow = .Row
			'        If IsMissing(CurCol) Then CurCol = .Col
			CurRow = .Row
			CurCol = .Col
			If CurRow < .FixedRows Then CurRow = .FixedRows
			If CurCol < .FixedCols Then CurCol = .FixedCols
			
			'       Do each column in the grid
			For c = .FixedCols To .Cols - 1
				.Col = c
				'           Move upward from the bottom of the table, shifting as you go
				For R = CurRow To .Rows - 1
					If R = .Rows - 1 Then
						.Row = R
						.Text = ""
					Else
						.Row = R + 1
						Entry = .Text
						.Row = R
						.Text = Entry
					End If
				Next R
			Next c
			
			If Not MinRowsFixed And .Rows > .FixedRows + 1 Then .Rows = .Rows - 1
			
			'       Make the row that "replaced" the deleted one the selection
			.Row = CurRow
			.Col = CurCol
			.Redraw = True
		End With
		CheckingGrid = False
		
	End Sub
	
	Public Sub SetupLineGrid(ByRef grd As AxMSFlexGridLib.AxMSFlexGrid, ByRef RHead As Object, ByRef CHead As Object, ByRef LeftLimit As Single, ByRef RightLimit As Single, ByRef TopLimit As Single, ByRef BottomLimit As Single, ByRef NumDataRows As Short, ByRef SysControl As System.Windows.Forms.Control, ByRef ScaleCols As Boolean, ByRef TopRowScale As Single, ByRef frm As System.Windows.Forms.Form)
		Dim Hmax, w, H, CWidth As Single
		Dim i, NumCols As Short
		Dim RowHeaders As Boolean
		
		NumCols = UBound(CHead) - LBound(CHead) + 1
		If UBound(RHead) = 0 Then
			RowHeaders = False
		Else
			RowHeaders = True
		End If
		
		' Populate the line data grid
		
		With grd
			
			.Cols = NumCols
			.ScrollBars = MSFlexGridLib.ScrollBarsSettings.flexScrollBarVertical
			.WordWrap = True
			.Row = 0
			
			' If the captions on the first row WILL wrap, adjust the row height to allow it
			
			If TopRowScale = 0 Then TopRowScale = 1#
			.set_RowHeight(0, .get_RowHeight(0) * TopRowScale)
			
			' Fill in the column headings
			
			For i = 0 To NumCols - 1
				.Col = i
				.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
				'UPGRADE_WARNING: Couldn't resolve default property of object CHead(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Text = CHead(i)
				'            .CellFontBold = True
			Next i
			
			' Next, populate the row headings
			
			.Rows = NumDataRows + 1
			.Col = 0
			For i = 1 To NumDataRows
				.Row = i
				.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
				If RowHeaders Then
					'UPGRADE_WARNING: Couldn't resolve default property of object RHead(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Text = RHead(i)
					'                .CellFontBold = True
				End If
			Next i
			
			' Adjust the size of the grid to match what we've done
			
		End With
		
		Call ResizeGrid(grd, LeftLimit, RightLimit, TopLimit, BottomLimit, NumDataRows, SysControl, ScaleCols, frm)
		
	End Sub
	
	Public Sub ResizeGrid(ByRef grd As AxMSFlexGridLib.AxMSFlexGrid, ByRef LeftLimit As Single, ByRef RightLimit As Single, ByRef TopLimit As Single, ByRef BottomLimit As Single, ByRef NumDataRows As Short, ByRef SysControl As System.Windows.Forms.Control, ByRef ScaleCols As Boolean, ByRef frm As System.Windows.Forms.Form)
		Dim Hmax, w, H, CWidth As Single
		Dim i, NumCols As Short
		
		NumCols = grd.Cols
		w = 0#
		
		' Populate the line data grid
		
		With grd
			
			' The captions on the first row WILL wrap; adjust the row height to allow this
			
			H = .get_RowHeight(0)
			
			' If desired, compute a dynamic column width based on the size of the form
			' containing the grid, and the number of columns in the grid
			
			If ScaleCols Then
				CWidth = (RightLimit - LeftLimit - 1000#) / NumCols
			End If
			
			' Fill in the column headings
			
			For i = 0 To NumCols - 1
				' Did we scale the columns?  If so, set their width here
				If ScaleCols And CWidth > 0 Then .set_ColWidth(i, CWidth)
				w = w + .get_ColWidth(i) ' Keep track of the total width
			Next i
			
			' If we're NOT scaling the columns, see if we need a horizontal scrollbar
			If (Not ScaleCols) And (w > RightLimit - LeftLimit) Then
				.ScrollBars = MSFlexGridLib.ScrollBarsSettings.flexScrollBarHorizontal
			End If
			
			' Next, measure the row headings
			
			.Rows = NumDataRows + 1
			.Col = 0
			For i = 1 To NumDataRows
				H = H + .get_RowHeight(i) ' Keep track of the height
			Next i
			
			' Adjust the size of the grid to match what we've done.  The additive values
			' in the calculation of the width and height make sure there's no slight
			' truncation of the last row or column
			
			.Height = VB6.TwipsToPixelsY(H + 75)
			
			' Calculate the maximum grid height, given the space allocated on the form
			
			Hmax = BottomLimit - TopLimit - 2 * Margin
			
			' If we've exceeded this allocation, VB will add a vertical scrollbar;
			' adjust the width of the grid by the width of the scrollbar to keep from
			' overlaying the last column
			
			'UPGRADE_WARNING: Control property grd.Parent was upgraded to grd.FindForm which has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
			If .FindForm.WindowState <> System.Windows.Forms.FormWindowState.Minimized Then
				If VB6.PixelsToTwipsY(.Height) > Hmax Then
					.Height = VB6.TwipsToPixelsY(Hmax)
					'UPGRADE_WARNING: Couldn't resolve default property of object SysControl.ScrollBarSize. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Width = VB6.TwipsToPixelsX(w + SysControl.ScrollBarSize + 100)
				Else
					.Width = VB6.TwipsToPixelsX(w + 75)
				End If
			End If
			
			' The grid is located between "LeftLimit" and "RightLimit"; center it there
			
			.Left = VB6.TwipsToPixelsX((LeftLimit + RightLimit - VB6.PixelsToTwipsX(.Width)) / 2#)
			
			' Locate the grid vertically
			
			.Top = VB6.TwipsToPixelsY(Margin + TopLimit)
			
		End With
		
		
	End Sub
	
	' Calculate index for use with TextArray property.
	Function TAIndex(ByRef Row As Short, ByRef Col As Short, ByRef grd As AxMSFlexGridLib.AxMSFlexGrid) As Integer
		TAIndex = Row * grd.Cols + Col
	End Function
	
	Public Function HasEmptyFields(ByVal grd As AxMSFlexGridLib.AxMSFlexGrid) As Boolean
		Dim R, c As Short
		HasEmptyFields = True
		With grd
			For R = .FixedRows To .Rows - 1
				For c = .FixedCols To .Cols - 1
					.Row = R
					.Col = c
					If Not .CellBackColor.equals(System.Drawing.SystemColors.Control) Then
						If .Text = "" Then
							MsgBox("Must complete all fields.", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Input Error")
							Exit Function
						End If
					End If
				Next c
			Next R
		End With
		HasEmptyFields = False
	End Function
	
	Public Sub ColorRow(ByRef grd As AxMSFlexGridLib.AxMSFlexGrid, ByRef R As Short, ByRef Color As Short)
		Dim c As Short
		
		With grd
			.Row = R
			For c = 0 To grd.Cols - 1
				.Col = c
				.CellForeColor = System.Drawing.ColorTranslator.FromOle(QBColor(Color))
			Next c
		End With
		
	End Sub
	
	Public Sub ColorCol(ByRef grd As AxMSFlexGridLib.AxMSFlexGrid, ByRef c As Short, ByRef Color As Short)
		Dim R As Short
		
		With grd
			.Col = c
			For R = 0 To grd.Rows - 1
				.Row = R
				.CellForeColor = System.Drawing.ColorTranslator.FromOle(QBColor(Color))
			Next R
		End With
		
	End Sub
	
	Public Sub RemoveColor(ByRef grd As AxMSFlexGridLib.AxMSFlexGrid)
		Dim R, c As Short
		
		With grd
			For R = 0 To .Rows - 1
				.Row = R
				For c = 0 To .Cols - 1
					.Col = c
					.CellBackColor = System.Drawing.ColorTranslator.FromOle(0) 'White
				Next c
			Next R
		End With
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