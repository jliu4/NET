Option Strict Off
Option Explicit On

' Version 8.1
' 2014, Copyright DTCEL, All Rights Reserved
' Jin Liu May 2014

Friend Class frmRiser

    Inherits System.Windows.Forms.Form

    Private CurUnits As VIVAMain.Units
    Private CFDim, CFLength, CFForce As Single
    Private CFContent, CFStiff As Single
    Private riserId As Short



    Private Sub frmRiser_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        riserId = CurProj.RiserId

        Text = Text & " : " & riserId & " - " & CurProj.Title
        CurUnits = CurProj.Units
        IniConvertFactor()
        SetLabels()
        IniConvertFactor()

        '   load data
        With CurProj.Riser(riserId)
            '       riser type
            btnRiserType(.RiserType).Checked = True

            '       riser
            txtTopTension.Text = CStr(.TopTension / CFForce)
            txtTopLocation.Text = CStr(.TopLocation / CFLength)

            '       mud
            txtMudWeight.Text = CStr(.ContentDensity / CFContent)
            txtMudHeight.Text = CStr(.BuoyancyHead / CFLength)

            '       boundary conditions
            btnUpperBC(.UpperBC).Checked = True
            btnLowerBC(.LowerBC).Checked = True
            '       spring constants
            txtStiffness(0).Text = CStr(.UpperStiffK / CFStiff)
            txtStiffness(1).Text = CStr(.LowerStiffK / CFStiff)

            '       bottom package frame
            '          If .LowerBC = VIVAMain.BoundaryConditions.Free Then
            'EnableBottomPackageFrame(True)
            '        Else
            '        EnableBottomPackageFrame(False)
            '       End If


        End With

        '   disable unit selection
        NumInputForms = NumInputForms + 1
        frmGlobalParms.btnMetric.Enabled = False
        frmGlobalParms.btnEnglish.Enabled = False
        SetLabels()
    End Sub


    Private Sub frmRiser_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        '   check whether to enable unit selection
        NumInputForms = NumInputForms - 1

        If NumInputForms = 0 And UnitsChangeable Then
            frmGlobalParms.btnMetric.Enabled = True
            frmGlobalParms.btnEnglish.Enabled = True
        End If

    End Sub


    Private Sub SetLabels()
        On Error GoTo ErrHandler

        '   set the labels in accordance with the units chosen
        If CurUnits = VIVAMain.Units.English Then
            lblRiserUnit(0).Text = "lb"
            lblRiserUnit(1).Text = "ft"
            lblMudUnit(0).Text = "ppg"

            lblFJUnit(0).Text = "ft-lb/deg"
            lblFJUnit(1).Text = "ft-lb/deg"
            'lblBPUnit(0).Text = "ft"
            'lblBPUnit(1).Text = "lb"
            'lblBPUnit(2).Text = "lb"
            'lblBPUnit(3).Text = "in"

            'For i = 4 To 7
            '     lblBPUnit(i).Text = lblBPUnit(i - 4).Text
            'Next
        Else
            lblRiserUnit(0).Text = "N"
            lblRiserUnit(1).Text = "m"
            lblMudUnit(0).Text = "kg/m^3"

            lblFJUnit(0).Text = "N-m/deg"
            lblFJUnit(1).Text = "N-m/deg"
            'lblBPUnit(0).Text = "m"
            'lblBPUnit(1).Text = "N"
            'lblBPUnit(2).Text = "N"
            'lblBPUnit(3).Text = "mm"

            'For i = 4 To 7
            'lblBPUnit(i).Text = lblBPUnit(i - 4).Text
            'Next
        End If
ErrHandler:
        'User pressed Cancel button
        Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub
    End Sub


    Private Sub IniConvertFactor()

        If CurUnits = VIVAMain.Units.Metric Then
            CFLength = 1.0#
            CFDim = mm2M
            CFForce = 1.0#
            CFContent = 1.0#
            CFStiff = 1.0# / Deg2Rad
        Else
            CFLength = Ft2M
            CFDim = In2Ft * Ft2M
            CFForce = Lb2N
            CFContent = Lb2Kg / Gal2M3
            CFStiff = Lb2N * Ft2M / Deg2Rad
        End If

    End Sub


    Private Sub btnUpperBC_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnUpperBC.CheckedChanged

        If eventSender.Checked Then
            Dim index As Short = btnUpperBC.GetIndex(eventSender)

            Select Case index
                '       pinned or fixed end
                Case 0, 1
                    '            txtStiffness(0).Text = "0"
                    txtStiffness(0).Enabled = False
                    lblStiffness(0).Enabled = False
                    lblFJUnit(0).Enabled = False
                    txtTopTension.Enabled = True
                    lblTopTension.Enabled = True
                    lblRiserUnit(0).Enabled = True

                    '       spring end
                Case 2
                    txtStiffness(0).Enabled = True
                    lblStiffness(0).Enabled = True
                    lblFJUnit(0).Enabled = True
                    txtTopTension.Enabled = True
                    lblTopTension.Enabled = True
                    lblRiserUnit(0).Enabled = True
                    '       free end

                Case 3
                    '            txtStiffness(0).Text = "0"
                    txtStiffness(0).Enabled = False
                    lblStiffness(0).Enabled = False
                    lblFJUnit(0).Enabled = False
                    txtTopTension.Text = "0"
                    txtTopTension.Enabled = False
                    lblTopTension.Enabled = False
                    lblRiserUnit(0).Enabled = False
                    If btnLowerBC(3).Checked = True Then
                        MsgBox("Only one end of the riser may be free!", MsgBoxStyle.OkOnly, "WinVIVA - Riser Boundary Condition Error")
                    End If
            End Select
        End If

    End Sub


    Private Sub btnLowerBC_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnLowerBC.CheckedChanged

        If eventSender.Checked Then
            Dim index As Short = btnLowerBC.GetIndex(eventSender)

            Select Case index
                '       pinned or fixed end
                Case 0, 1
                    '            txtStiffness(1).Text = "0"
                    txtStiffness(1).Enabled = False
                    lblStiffness(1).Enabled = False
                    lblFJUnit(1).Enabled = False
                    'EnableBottomPackageFrame(False)

                    '       spring end
                Case 2
                    txtStiffness(1).Enabled = True
                    lblStiffness(1).Enabled = True
                    lblFJUnit(1).Enabled = True
                    ' EnableBottomPackageFrame(False)

                    '       free end
                Case 3
                    '            txtStiffness(1).Text = "0"
                    txtStiffness(1).Enabled = False
                    lblStiffness(1).Enabled = False
                    lblFJUnit(1).Enabled = False
                    If btnUpperBC(3).Checked = True Then
                        MsgBox("Only one end of the riser may be free!", MsgBoxStyle.OkOnly, "WinVIVA - Riser Boundary Condition Error")
                    End If
                    'EnableBottomPackageFrame(True)
            End Select
        End If

    End Sub


    Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click

        Me.Close()

    End Sub


    Private Sub btnOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnOK.Click

        Dim index As Short

        '   check for the obvious mistake
        If btnUpperBC(3).Checked And btnLowerBC(3).Checked Then
            MsgBox("Only one end of the riser may be free!", MsgBoxStyle.OkOnly, "WinVIVA - Boundary Condition Error")
            Exit Sub
        End If

        With CurProj.Riser(riserId)
            '       riser type
            For index = 0 To NumRiserTypes - 1
                If btnRiserType(index).Checked = True Then
                    .RiserType = CRst(index)
                    Exit For
                End If
            Next index

            '       reset main menu
            If .RiserType = VIVAMain.RiserTypes.Rigid Then
                frmMainMenu.mnuRiserFStaticSolution.Enabled = False
            Else
                frmMainMenu.mnuRiserFStaticSolution.Enabled = True
            End If

            '       boundary condition
            For index = 0 To NumBoundaryConditions - 1
                If btnUpperBC(index).Checked = True Then
                    .UpperBC = CBcs(index)
                    Exit For
                End If
            Next index

            For index = 0 To NumBoundaryConditions - 1
                If btnLowerBC(index).Checked = True Then
                    .LowerBC = CBcs(index)
                    Exit For
                End If
            Next index

            '       update curproj
            If IsNumeric(txtTopTension.Text) Then
                .TopTension = CSng(txtTopTension.Text) * CFForce
            Else
                .TopTension = 0.0#
            End If

            If IsNumeric(txtTopLocation.Text) Then
                .TopLocation = CSng(txtTopLocation.Text) * CFLength
            Else
                .TopTension = 0.0#
            End If

            If IsNumeric(txtMudWeight.Text) Then
                .ContentDensity = CSng(txtMudWeight.Text) * CFContent
            Else
                .ContentDensity = 0.0#
            End If

            If IsNumeric(txtMudHeight.Text) Then
                .BuoyancyHead = CSng(txtMudHeight.Text) * CFLength
            Else
                .BuoyancyHead = 0.0#
            End If

            If IsNumeric(txtStiffness(0).Text) Then
                .UpperStiffK = CSng(txtStiffness(0).Text) * CFStiff
            Else
                .UpperStiffK = 0.0#
            End If

            If IsNumeric(txtStiffness(1).Text) Then
                .LowerStiffK = CSng(txtStiffness(1).Text) * CFStiff
            Else
                .LowerStiffK = 0.0#
            End If

        End With
        '   remember that the case will need to be saved
        CurProj.Saved = False

        '   unload
        Me.Close()

    End Sub


    ' Private Sub EnableBottomPackageFrame(ByRef Flag As Boolean)

    'Dim Con As System.Windows.Forms.Control
    '
    '       For Each Con In Me.Controls
    '          If Con.Parent.Name = "fraBottomPackage" Then Con.Enabled = Flag
    '     Next Con
    '
    '   End Sub

End Class