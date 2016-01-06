Option Strict Off
Option Explicit On
Friend Class frmProgress
	Inherits System.Windows.Forms.Form
	' frmProgress       form to control progress of programs
	' Version 1.01
	' 2001, Copyright DTCEL, All Rights Reserved
	
	
	' properties
	' Progress          progress percentage
	
	Public WriteOnly Property Progress() As Short
		Set(ByVal Value As Short)
			prgProgress.Value = Value
			lblProgress.Text = Value & "% completed"
		End Set
	End Property
	
	Public WriteOnly Property ProgressBatch() As Short
		Set(ByVal Value As Short)
			If Not prgProgressBatch.Visible Then
				prgProgressBatch.Visible = True
				lblBatchStage.Visible = True
				lblBatchPercent.Visible = True
			End If
			prgProgressBatch.Value = Value
			lblBatchPercent.Text = Value & "% completed"
		End Set
	End Property
	
	Private Sub RefreshData()
		
	End Sub
	
	' load/unload form
	
	Private Sub frmProgress_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		CurrentStage.Text = "Initiating the program..."
		Progress = 0
		
		prgProgressBatch.Visible = False
		lblBatchStage.Visible = False
		lblBatchPercent.Visible = False
		
	End Sub
	
	' command buttons
	
	Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click
		
		Cancelled = True
		Me.Close()
		
	End Sub
	
	Private Sub frmProgress_LostFocus(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.LostFocus
		If Me.Visible Then Me.Activate()
	End Sub
	
	Private Sub frmProgress_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim cancel As Boolean = eventArgs.Cancel
		Dim unloadmode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
		Cancelled = True
		eventArgs.Cancel = cancel
	End Sub
End Class