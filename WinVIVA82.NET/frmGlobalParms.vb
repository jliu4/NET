Option Strict Off
Option Explicit On

' Version 8.1
' 2014, Copyright DTCEL, All Rights Reserved
' Jin Liu May 2014

Friend Class frmGlobalParms
    Inherits System.Windows.Forms.Form

    Private CFDensity, CFViscosity, CFLength As Single


  
    Private Sub frmGlobalParms_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        Text = Text & " : - " & CurProj.Title
        UnitsChangeable = True
        If CurProj.nRisers = 1 Then
            btnOneRiser.Checked = True
            TwoRisers.Enabled = False
        Else
            btnTwoRisers.Checked = True
            TwoRisers.Enabled = True
       
            If CurProj.sameRiserARiser Then
                chkbtn2.CheckState = System.Windows.Forms.CheckState.Checked
            Else
                chkbtn2.CheckState = System.Windows.Forms.CheckState.Unchecked
            End If
        End If
        '   load the form fields from curproj
        If CurProj.Units = VIVAMain.Units.English Then
            btnEnglish.Checked = True
            lblWaterDensityUnits.Text = "lb/ft^3"
            lblVisUnits.Text = "ft^2/s"
            lblDistance3.Text = "ft"
            lblDistance4.Text = "ft"
            lblDistance1.Text = "ft"
            lblDistance2.Text = "ft"
            CFViscosity = Ft2M ^ 2
            CFDensity = Lb2Kg / Ft2M ^ 3
            CFLength = Ft2M
        Else
            btnMetric.Checked = True
            lblWaterDensityUnits.Text = "kg/m^3"
            lblVisUnits.Text = "m^2/s"
            lblDistance3.Text = "m"
            lblDistance4.Text = "m"
            lblDistance1.Text = "m"
            lblDistance2.Text = "m"
            CFViscosity = 1.0#
            CFDensity = 1.0#
            CFLength = 1.0#
        End If

        If CurProj.DampingMethod = VIVAMain.DampingMethod.ConstantDamping Then
            cDamping.Checked = True
        Else
            mDamping.Checked = True
        End If

        If NumInputForms = 0 Then
            btnMetric.Enabled = True
            btnEnglish.Enabled = True
        Else
            btnMetric.Enabled = False
            btnEnglish.Enabled = False
        End If

        UpdateTextFields()

    End Sub

    Private Sub frmGlobalParms_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        UnitsChangeable = False

    End Sub

    Private Sub btnOneRiser_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnOneRiser.Click
        CurProj.nRisers = 1
        TwoRisers.Enabled = False

    End Sub

    Private Sub btnTwoRisers_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnTwoRisers.Click
        CurProj.nRisers = 2
        CurProj.sameRiserARiser = True

        TwoRisers.Enabled = True

        If CurProj.sameRiserARiser Then
            chkbtn2.CheckState = System.Windows.Forms.CheckState.Checked
        Else
            chkbtn2.CheckState = System.Windows.Forms.CheckState.Unchecked
        End If

    End Sub

    Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click

        Me.Close()

    End Sub

    Private Sub btnOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnOK.Click

        '   check each entry for validity:  is it a number; and is it positive?
        If Not PositiveNumber((txtWaterDensity.Text)) Then
            MsgBox("Please enter a valid positive number for the Water Density")
            txtWaterDensity.Focus()
            Exit Sub
        ElseIf Not PositiveNumber((txtVis.Text)) Then
            MsgBox("Please enter a valid positive number for the Water Viscosity")
            txtVis.Focus()
            Exit Sub
        ElseIf CurProj.nRisers = 2 And (CSng(txtRiserRY.Text) = 0 And CSng(txtRiserRZ.Text) = 0) Then
            MsgBox("A distance needed for two Risers")
            txtRiserRY.Focus()
            Exit Sub

        End If

        If cDamping.Checked = True Then
            CurProj.DampingMethod = VIVAMain.DampingMethod.ConstantDamping
        Else
            CurProj.DampingMethod = VIVAMain.DampingMethod.ModalDamping
        End If

        '   if all numbers are OK, store them in the appropriate object properties
        CurProj.Water.Density = CSng(txtWaterDensity.Text) * CFDensity

        CurProj.Water.Viscosity = CSng(txtVis.Text) * CFViscosity

        If btnOneRiser.Checked = True Then
            CurProj.nRisers = 1
            TwoRisers.Enabled = False
        Else
            CurProj.nRisers = 2
            If chkbtn2.CheckState = System.Windows.Forms.CheckState.Checked Then
                If CurProj.sameRiserARiser = False Then
                    CurProj.sameRiserARiser = True
                    CurProj.CopyRiser()
                End If
            Else
                If CurProj.sameRiserARiser = True Then
                    CurProj.sameRiserARiser = False

                End If

            End If
            TwoRisers.Enabled = True
            CurProj.RiserFLocZ = CSng(txtRiserFZ.Text) * CFLength
            CurProj.RiserFLocY = CSng(txtRiserFY.Text) * CFLength
            CurProj.RiserRLocZ = CSng(txtRiserRZ.Text) * CFLength
            CurProj.RiserRLocY = CSng(txtRiserRY.Text) * CFLength
        End If

        '   see if the units have been changed
        If btnEnglish.Checked = True And CurProj.Units = VIVAMain.Units.Metric Or btnMetric.Checked = True And CurProj.Units = VIVAMain.Units.English Then
            CurProj.ConvertUnits()
        End If

        '   remember that the case will need to be saved
        CurProj.Saved = False
        If (CurProj.nRisers = 1) Then 'Or CurProj.sameRiserARiser = True) Then 'JLIU TODO
            frmMainMenu.mnuRiserR.Enabled = False
        Else
            frmMainMenu.mnuRiserR.Enabled = True

        End If
        If CurProj.nRisers = 1 Then
            frmMainMenu.mnuResultsR.Enabled = False
        Else
            frmMainMenu.mnuResultsR.Enabled = True
        End If
        eventArgs.ToString()
        '   unload
        Me.Close()

    End Sub


    Private Sub UpdateTextFields()

        '   load the form fields from curproj
        txtWaterDensity.Text = CStr(CurProj.Water.Density / CFDensity)

        txtVis.Text = CStr(CurProj.Water.Viscosity / CFViscosity)
        txtRiserFY.Text = CStr(CurProj.RiserFLocY / CFLength)
        txtRiserFZ.Text = CStr(CurProj.RiserFLocZ / CFLength)
        txtRiserRY.Text = CStr(CurProj.RiserRLocY / CFLength)
        txtRiserRZ.Text = CStr(CurProj.RiserRLocZ / CFLength)
    End Sub

    Private Sub txtVis_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtVis.Enter

        '   set the text box properties to select the entire field
        txtVis.SelectionStart = 0
        txtVis.SelectionLength = 100

    End Sub


    Private Sub txtWaterDensity_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWaterDensity.Enter

        '   set the text box properties to select the entire field
        txtWaterDensity.SelectionStart = 0
        txtWaterDensity.SelectionLength = 100

    End Sub

    Private Sub btnEnglish_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEnglish.Click

        If sender.Checked Then
            If CurProj.Units = VIVAMain.Units.Metric Then
                CurProj.ConvertUnits()
                lblWaterDensityUnits.Text = "lb/ft^3"
                lblVisUnits.Text = "ft^2/s"
                lblDistance3.Text = "ft"
                lblDistance4.Text = "ft"
                lblDistance1.Text = "ft"
                lblDistance2.Text = "ft"
                CFDensity = Lb2Kg / Ft2M ^ 3
                CFViscosity = Ft2M ^ 2
                CFLength = 1.0# / Ft2M
                UpdateTextFields()
            End If
        End If

    End Sub

    Private Sub btnMetric_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMetric.Click

        If sender.Checked Then
            If CurProj.Units = VIVAMain.Units.English Then
                CurProj.ConvertUnits()
                lblWaterDensityUnits.Text = "kg/m^3"
                lblVisUnits.Text = "m^2/s"
                lblDistance3.Text = "m"
                lblDistance4.Text = "m"
                lblDistance1.Text = "m"
                lblDistance2.Text = "m"
                CFDensity = 1.0#
                CFViscosity = 1.0#
                CFLength = 1.0#
                UpdateTextFields()
            End If
        End If

    End Sub

    'Private Sub txtRiserRY_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRiserRY.TextChanged

    'End Sub

End Class