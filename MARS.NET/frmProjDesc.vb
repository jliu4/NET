Option Strict Off
Option Explicit On
Friend Class frmProjDesc
	Inherits System.Windows.Forms.Form
	
	Private Sub btnSetMarsDir_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnSetMarsDir.Click
		lblMarsDir.Text = Dir1.Path
	End Sub
	
	Private Sub Dir1_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Dir1.Change
		File1.Path = Dir1.Path
	End Sub
	
	Private Sub Drive1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Drive1.SelectedIndexChanged
		Dir1.Path = Drive1.Drive
	End Sub
	
	Private Sub File1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles File1.SelectedIndexChanged
        txtVesselName.Text = UCase(getNoExtFileName(File1.FileName))
    End Sub

    Private Sub File1_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles File1.DoubleClick
        txtVesselName.Text = UCase(getNoExtFileName(File1.FileName))
        btnOK_Click(btnOK, New System.EventArgs())
    End Sub

    Private Sub frmProjDesc_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		With CurProj
			txtProjName.Text = .Title
			txtProjDesc.Text = .Description
			txtVesselName.Text = .Vessel.Name
			If Len(.Directory) = 0 Then .Directory = MarsDir
			Drive1.Drive = .Directory
			Dir1.Path = .Directory
			lblMarsDir.Text = .Directory
		End With
		
	End Sub
	
	Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click
		
		Me.Close()
		
	End Sub
	
	Private Sub btnOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnOK.Click
		
		Dim sPath As String
		With CurProj
			.Title = txtProjName.Text
			.Description = txtProjDesc.Text
			.Vessel.Name = txtVesselName.Text
			.Directory = lblMarsDir.Text
			MarsDir = lblMarsDir.Text
			
			frmMain.lblVessel.Text = CurVessel.Name
			
			.Saved = False
#If DTCEL Then
			'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression DTCEL did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
			frmMain.Caption = " MARS Console - " & .Title
#Else
			frmMain.Text = " MARS for " & .Vessel.Name & " - " & .Title
#End If
			
			sPath = Dir1.Path
			
			If Not .VesselParticular(sPath) Then
				MsgBox("MARS is unable to load vessel particulars.", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, msgTitle)
			End If
		End With
		Me.Close()
		
	End Sub
End Class