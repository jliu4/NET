Option Strict Off
Option Explicit On
Friend Class frmOptions
    Inherits System.Windows.Forms.Form

    Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnOK.Click
        If btnMetric.Checked Then
            IsMetricUnit = True
        Else
            IsMetricUnit = False
        End If
        Me.Close()
    End Sub

    Private Sub frmOptions_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        If IsMetricUnit Then
            btnMetric.Checked = True
        Else
            btnEnglish.Checked = True
        End If

    End Sub

    ' Private Sub btnEnglish_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEnglish.Click
    ' If sender.Checked Then
    '          IsMetricUnit = False
    ' End If
    '  End Sub

    '  Private Sub btnMetric_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMetric.Click
    ' If sender.Checked Then
    '         IsMetricUnit = True
    ' End If
    '  End Sub
End Class