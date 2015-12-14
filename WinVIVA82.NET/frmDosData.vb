Option Strict Off
Option Explicit On

' Version 8.1
' 2014, Copyright DTCEL, All Rights Reserved
' Jin Liu May 2014

Friend Class frmDosData

    Inherits System.Windows.Forms.Form

    Private NumLines As Short


    Private Sub frmDosData_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        NumLines = ReadDOSText("FAT")
        LoadText()

        lblFileName.Text = "File: " & CurProj.DataDirectory & OutputFiles(0, 18)

    End Sub


    Private Sub frmDosData_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize

        Dim MinW As Object
        Dim MinH As Short

        MinW = VB6.PixelsToTwipsX(btnCancel.Width) * 2 + VB6.PixelsToTwipsX(lblFileName.Width) + 480
        MinH = VB6.PixelsToTwipsY(btnCancel.Height) + 2480
        If VB6.PixelsToTwipsX(Width) < MinW Then Width = VB6.TwipsToPixelsX(MinW)
        If VB6.PixelsToTwipsY(Height) < MinH Then Height = VB6.TwipsToPixelsY(MinH)

        txtDOSOut.Top = VB6.TwipsToPixelsY(40)
        txtDOSOut.Left = VB6.TwipsToPixelsX(180)
        txtDOSOut.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Width) - 480)
        txtDOSOut.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Height) - 900 - VB6.PixelsToTwipsY(btnCancel.Height))
        btnCancel.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(txtDOSOut.Height) + 120)
        btnCancel.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Width) - 300 - VB6.PixelsToTwipsX(btnCancel.Width))
        btnPrint.Top = btnCancel.Top
        btnPrint.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(btnCancel.Left) - 120 - VB6.PixelsToTwipsX(btnPrint.Width))
        lblFileName.Top = btnCancel.Top
        lblFileName.Left = txtDOSOut.Left
        lblFileName.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(txtDOSOut.Width) - VB6.PixelsToTwipsX(btnCancel.Width) * 2 - 120)

    End Sub


    Private Sub LoadText()

        Dim i As Short

        txtDOSOut.Text = ""

        For i = 1 To NumLines
            txtDOSOut.Text = txtDOSOut.Text & CurProj.Riser(1).DOSOutText(i) & vbCrLf
        Next i

    End Sub


    Private Sub btnPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnPrint.Click

        Dim i As Short
        'Dim Printer As Printing.PrintDocument

        For i = 1 To NumLines
            'UPGRADE_ISSUE: Printer object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
            'UPGRADE_ISSUE: Printer method Printer.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
            'Printer.Print(DOSOutText(i))
        Next i

        'UPGRADE_ISSUE: Printer method Printer.EndDoc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.EndDoc()

    End Sub


    Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click

        Me.Close()

    End Sub


    Public Sub mnuPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPrint.Click

        btnPrint_Click(btnPrint, New System.EventArgs())

    End Sub


    Public Sub mnuClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuClose.Click

        Me.Close()

    End Sub

End Class