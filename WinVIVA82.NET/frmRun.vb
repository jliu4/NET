Public Class frmRun

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        Me.Text = "Running Batch Cases ..."
        Me.lblMsg1.Text = ""
        Me.lblMsg2.Text = ""

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.prgBrCaseRun.Dispose()
            Me.Dispose()
        Catch ex As Exception
        End Try
    End Sub
End Class