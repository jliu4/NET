Option Strict Off
Option Explicit On

Friend Class frmProjDesc

    Inherits System.Windows.Forms.Form


    Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click

        Me.Close()

    End Sub


    Private Sub btnOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnOK.Click

        CurProj.Title = txtProjName.Text
        frmMainMenu.Text = "WinVIVA - Main Menu - " & CurProj.Title
        CurProj.Desc = txtProjDesc.Text
        Me.Close()

    End Sub


    Private Sub frmProjDesc_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        txtProjName.Text = CurProj.Title
        txtProjDesc.Text = CurProj.Desc

    End Sub


    Private Sub frmProjDesc_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        CurProj.Title = txtProjName.Text
        CurProj.Desc = txtProjDesc.Text

    End Sub
End Class