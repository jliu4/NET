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
        'If HasEmptyFields(grdProfile) Then Exit Sub
        ' must include seabed point

        ' save data to objects
        Dim R As Short
		oMet(CaseNo).Current.Profile.Clear()
        For R = 1 To grdProfile.RowCount - 1

            Call oMet(CaseNo).Current.Profile.Add(CSng(grdProfile.Rows(R).Cells(0).Value), CDbl(grdProfile.Rows(R).Cells(1).Value) * Knots2Ftps)
        Next R

        With frmMain.grdMatrix
            '   .TextMatrix(frmMain.grdMatrix.Row, 9) = grdProfile.TextMatrix(1, 1)
            ' .set_TextMatrix(frmMain.grdMatrix.Row, frmMain.grdMatrix.Col, txtNumProfilePts.Text)
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
        frmCurrProfile_Activated(Me, New System.EventArgs())
    End Sub
	
	Private Sub btnCopyAll_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCopyAll.Click
		If Len(txtSameAsCase.Text) = 0 Then
			MsgBox("Must specify case number to copy from.")
			Exit Sub
		End If
		Call CopyCase(CShort(VB.Right(lblCaseNo.Text, Len(lblCaseNo.Text) - 4)), CShort(txtSameAsCase.Text), True)
	End Sub

    Private Sub frmCurrProfile_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        Dim i As Short
        On Error GoTo ErrHandler

        If Len(lblCaseNo.Text) = 0 Then Exit Sub
        CaseNo = CShort(lblCaseNo.Text)
        txtNumProfilePts.Text = CStr(oMet(CaseNo).Current.Profile.Count)
        With grdProfile
            .ColumnCount = 2
            .RowCount = oMet(CaseNo).Current.Profile.Count
            .Columns(0).FillWeight = 100 / .ColumnCount
            .Columns(1).FillWeight = 100 / .ColumnCount
            .Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            .Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            .Columns(0).HeaderText = "Water Depth (ft)"
            .Columns(1).HeaderText = "Velocity (knots)"
            ' load current profile data from object
            For i = 1 To oMet(CaseNo).Current.Profile.Count
                .Rows(i - 1).Cells(0).Value = oMet(CaseNo).Current.Profile.Item(i).Depth.Value
                .Rows(i - 1).Cells(1).Value = CStr(Format(Convert.ToDouble(oMet(CaseNo).Current.Profile.Item(i).Velocity.Value) * Ftps2Knots, "0.000"))

            Next i
        End With
        Exit Sub
ErrHandler:
        'User pressed Cancel button
        Cursor = System.Windows.Forms.Cursors.Default
        If Len(Err.Description) > 0 Then
            MsgBox(Err.Description)
        End If
        Exit Sub
    End Sub


    Private Sub txtNumProfilePts_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNumProfilePts.TextChanged
        If Val(txtNumProfilePts.Text) <= 0 Then txtNumProfilePts.Text = "1"
        NumPoints = Val(txtNumProfilePts.Text)
        grdProfile.RowCount = NumPoints
    End Sub
End Class