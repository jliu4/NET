Option Strict Off
Option Explicit On
Friend Class frmOptions
	Inherits System.Windows.Forms.Form
	
	Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click
		Me.Close()
	End Sub
	
	Private Sub btnOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnOK.Click
		If optUnit(0).Checked Then
			IsMetricUnit = True
		Else
			IsMetricUnit = False
		End If
		Me.Close()
	End Sub
	
	Private Sub frmOptions_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		If IsMetricUnit Then
			optUnit(0).Checked = True
		Else
			optUnit(1).Checked = True
		End If
	End Sub
End Class