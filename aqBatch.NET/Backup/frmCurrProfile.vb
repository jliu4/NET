Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmCurrProfile
	Inherits System.Windows.Forms.Form
	Dim CaseNo As Short
	Dim NumPoints As Short
	Dim JustEnterCell As Boolean
	Dim ExistingText As String
	
	Private Sub CopyCase(ByVal toCaseNo As Short, ByVal fromCaseNo As Short, Optional ByRef blnCopyAll As Boolean = False)
		Dim i, k As Short
		
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If blnCopyAll = False Or IsNothing(blnCopyAll) Then
			With oMet(toCaseNo)
				.Current.Profile.Clear()
				For i = 1 To oMet(fromCaseNo).Current.Profile.Count
					Call .Current.Profile.Add(oMet(fromCaseNo).Current.Profile.Item(i).Depth.Value, oMet(fromCaseNo).Current.Profile.Item(i).Velocity.Value)
				Next i
			End With
		Else ' copy to all
			For k = 1 To NumCases
				If k <> fromCaseNo Then
					With oMet(k)
						.Current.Profile.Clear()
						For i = 1 To oMet(fromCaseNo).Current.Profile.Count
							Call .Current.Profile.Add(oMet(fromCaseNo).Current.Profile.Item(i).Depth.Value, oMet(fromCaseNo).Current.Profile.Item(i).Velocity.Value)
						Next i
					End With
				End If
			Next k
			frmMain.LoadGrid()
			Me.Close()
		End If
		
	End Sub
	
	Private Sub bntOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles bntOK.Click
		' check empty fields
		If HasEmptyFields(grdProfile) Then Exit Sub
		' must include seabed point
		
		' save data to objects
		Dim R As Short
		oMet(CaseNo).Current.Profile.Clear()
		For R = 1 To grdProfile.Rows - 1
			If Not IsNumeric(grdProfile.get_TextMatrix(R, 0)) Or Not IsNumeric(grdProfile.get_TextMatrix(R, 1)) Then
				MsgBox("Please check your data")
				Exit Sub
			End If
			Call oMet(CaseNo).Current.Profile.Add(CSng(grdProfile.get_TextMatrix(R, 0)), CDbl(grdProfile.get_TextMatrix(R, 1)) * Knots2Ftps)
		Next R
		
		'    For r = 1 To oMet(CaseNo).Current.Profile.Count
		'        Debug.Print oMet(CaseNo).Current.Profile(r).Depth
		'        Debug.Print oMet(CaseNo).Current.Profile(r).Velocity
		'    Next r
		
		With frmMain.grdMatrix
			'   .TextMatrix(frmMain.grdMatrix.Row, 9) = grdProfile.TextMatrix(1, 1)
			.set_TextMatrix(frmMain.grdMatrix.Row, frmMain.grdMatrix.Col, txtNumProfilePts.Text)
		End With
		Me.Close()
	End Sub
	
	Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click
		Me.Close()
	End Sub
	
	Private Sub btnCopy_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCopy.Click
		If Len(txtSameAsCase.Text) = 0 Then
			MsgBox("Must specify case number to copy from.")
			Exit Sub
		End If
		Call CopyCase(CShort(VB.Right(lblCaseNo.Text, Len(lblCaseNo.Text) - 4)), CShort(txtSameAsCase.Text))
		'UPGRADE_WARNING: Form event frmCurrProfile.Activated has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		frmCurrProfile_Activated(Me, New System.EventArgs())
	End Sub
	
	Private Sub btnCopyAll_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCopyAll.Click
		If Len(txtSameAsCase.Text) = 0 Then
			MsgBox("Must specify case number to copy from.")
			Exit Sub
		End If
		Call CopyCase(CShort(VB.Right(lblCaseNo.Text, Len(lblCaseNo.Text) - 4)), CShort(txtSameAsCase.Text), True)
	End Sub
	
	'UPGRADE_WARNING: Form event frmCurrProfile.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmCurrProfile_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		Dim i As Short
		If Len(lblCaseNo.Text) = 0 Then Exit Sub
		CaseNo = CShort(Mid(lblCaseNo.Text, InStr(1, lblCaseNo.Text, " ")))
		txtNumProfilePts.Text = CStr(oMet(CaseNo).Current.Profile.Count)
		With grdProfile
			.set_ColWidth(0, VB6.PixelsToTwipsX(.Width) / .Cols * 0.9)
			.set_ColWidth(1, .get_ColWidth(0))
			.set_TextMatrix(0, 0, "Water Depth (ft)")
			.set_TextMatrix(0, 1, "Velocity (knots)")
			' load current profile data from object
			For i = 1 To oMet(CaseNo).Current.Profile.Count
				.set_TextMatrix(.FixedRows + i - 1, 0, oMet(CaseNo).Current.Profile.Item(i).Depth.Value)
				.set_TextMatrix(.FixedRows + i - 1, 1, VB6.Format(oMet(CaseNo).Current.Profile.Item(i).Velocity.Value * Ftps2Knots, "0.000"))
			Next i
		End With
		
	End Sub
	
	
	
	Private Sub grdProfile_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles grdProfile.ClickEvent
		MSFlexGridEdit(grdProfile, txtEdit, System.Windows.Forms.Keys.F10)
	End Sub
	
	Private Sub grdProfile_EnterCell(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles grdProfile.EnterCell
		Dim ExistingTxt As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object ExistingTxt. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ExistingTxt = grdProfile.Text
		JustEnterCell = True
	End Sub
	
	Private Sub grdProfile_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSFlexGridLib.DMSFlexGridEvents_KeyDownEvent) Handles grdProfile.KeyDownEvent
		KeyHandler(grdProfile, txtEdit, eventArgs.KeyCode, eventArgs.Shift, JustEnterCell, ExistingText)
	End Sub
	
	Private Sub grdProfile_LeaveCell(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles grdProfile.LeaveCell
		Dim CheckingGrid As Object
		Dim ExistingTxt As Object
		Dim Changed As Object
		If txtEdit.Visible = True Then
			grdProfile.Text = txtEdit.Text
			txtEdit.Visible = False
		End If
		
		If Not CheckingGrid Then
			With grdProfile
				'UPGRADE_WARNING: Couldn't resolve default property of object ExistingTxt. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Trim(.Text) <> ExistingTxt And .Col > 1 Then
					If Trim(.Text) <> "" Then .Text = CheckData(.Text)
					'UPGRADE_WARNING: Couldn't resolve default property of object Changed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Changed = True
				End If
			End With
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtNumProfilePts.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtNumProfilePts_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNumProfilePts.TextChanged
		If Val(txtNumProfilePts.Text) <= 0 Then txtNumProfilePts.Text = "1"
		NumPoints = Val(txtNumProfilePts.Text)
		grdProfile.Rows = NumPoints + grdProfile.FixedRows
	End Sub
End Class