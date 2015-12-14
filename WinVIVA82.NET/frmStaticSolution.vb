Option Strict Off
Option Explicit On

' Version 8.1
' 2014, Copyright DTCEL, All Rights Reserved
' Jin Liu May 2014

Friend Class frmStaticSolution

    Inherits System.Windows.Forms.Form

    Const DistanceCol As Short = 0
    Const TensionCol As Short = 1
    Const InAngleCol As Short = 2
    'Const OutAngleCol As Short = 3
    Const DepthCol As Short = 3
    Const YdistCol As Short = 4
    Const ZdistCol As Short = 5

    Const VelUCol As Short = 6
    Const VelVCol As Short = 7

    Const TensionChangeCol As Short = 8
    Const NumColumns As Short = 9
    Const IniPoints As Short = 100
    Private riserId As Short
    Private ExistingTxt As String
    Private Changed As Boolean
    Private CheckingGrid As Boolean
    Private JEC As Boolean
    Private CurUnits As VIVAMain.Units
    Private CFAngle, CFLength, CFForce As Single
    Private CFVel, CFChFr As Single
    Private ColumnLabels(1, NumColumns) As String
    Private tmpFileNum As Short

    Private Sub frmStaticSolution_CellValidating( _
          ByVal sender As Object, _
          ByVal e As System.Windows.Forms. _
          DataGridViewCellValidatingEventArgs) _
          Handles grdStaticSolution.CellValidating

        grdStaticSolution.Rows(e.RowIndex).ErrorText = ""

        If grdStaticSolution.Rows(e.RowIndex).IsNewRow Then Return

        If Not IsNumeric(e.FormattedValue) Then
            If Not (IsNumeric(e.FormattedValue)) And Not String.IsNullOrEmpty(e.FormattedValue) Then
                e.Cancel = True
                MsgBox("It must be numeric value")
                grdStaticSolution.Rows(e.RowIndex).ErrorText = _
                   "It must be a numeric value."

            End If
        End If
    End Sub
    Private Sub frmStaticSolution_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        riserId = CurProj.RiserId
        Text = Text & " : " & riserId & " - " & CurProj.Title
        CheckingGrid = False

        CurUnits = CurProj.Units

        SetLabels()
        IniConvertFactor()
        LoadGridFromProject()

        NumInputForms = NumInputForms + 1
        frmGlobalParms.btnMetric.Enabled = False
        frmGlobalParms.btnEnglish.Enabled = False

    End Sub


    Private Sub frmStaticSolution_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        NumInputForms = NumInputForms - 1

        If NumInputForms = 0 And UnitsChangeable Then
            frmGlobalParms.btnMetric.Enabled = True
            frmGlobalParms.btnEnglish.Enabled = True
        End If

    End Sub


    Private Sub SetLabels()

        ' Metric unit labels
        ColumnLabels(VIVAMain.Units.Metric, 0) = "No*"
        ColumnLabels(VIVAMain.Units.Metric, DistanceCol + 1) = "Distance Along Riser* (m)"
        ColumnLabels(VIVAMain.Units.Metric, TensionCol + 1) = "Axial Static Tension (N)"
        ColumnLabels(VIVAMain.Units.Metric, InAngleCol + 1) = "In-Plane Riser Angle* (deg)"
        ' ColumnLabels(VIVAMain.Units.Metric, OutAngleCol + 1) = "Out-Plane Riser Angle* (deg)"
        ColumnLabels(VIVAMain.Units.Metric, DepthCol + 1) = "Depth (m)"
        ColumnLabels(VIVAMain.Units.Metric, YdistCol + 1) = "Position Y (m)"
        ColumnLabels(VIVAMain.Units.Metric, ZdistCol + 1) = "Position Z (m)"
        ColumnLabels(VIVAMain.Units.Metric, VelUCol + 1) = "Current Vy (m/s)"
        ColumnLabels(VIVAMain.Units.Metric, VelVCol + 1) = "Current Vz (m/s)"
       

        ColumnLabels(VIVAMain.Units.Metric, TensionChangeCol + 1) = "Change in Static Tension (N/m)"

        ' English unit labels
        ColumnLabels(VIVAMain.Units.English, 0) = "No*"
        ColumnLabels(VIVAMain.Units.English, DistanceCol + 1) = "Distance Along Riser* (ft)"
        ColumnLabels(VIVAMain.Units.English, TensionCol + 1) = "Axial Static Tension (lb)"
        ColumnLabels(VIVAMain.Units.English, InAngleCol + 1) = "In-Plane Riser Angle* (deg)"
        ' ColumnLabels(VIVAMain.Units.English, OutAngleCol + 1) = "Out-Plane Riser Angle* (deg)"
        ColumnLabels(VIVAMain.Units.English, DepthCol + 1) = "Depth (ft)"
        ColumnLabels(VIVAMain.Units.English, YdistCol + 1) = "Position Y (ft)"
        ColumnLabels(VIVAMain.Units.English, ZdistCol + 1) = "Position Z (ft)"
        ColumnLabels(VIVAMain.Units.English, VelUCol + 1) = "Current Vy (ft/s)"
        ColumnLabels(VIVAMain.Units.English, VelVCol + 1) = "Current Vz (ft/s)"

        ColumnLabels(VIVAMain.Units.English, TensionChangeCol + 1) = "Change in Static Tension (lb/ft)"


    End Sub


    Private Sub IniConvertFactor()

        If CurUnits = VIVAMain.Units.Metric Then
            CFLength = 1.0#
            CFAngle = Deg2Rad
            CFForce = 1.0#
            CFVel = 1.0#
            CFChFr = 1.0#
        Else
            CFLength = Ft2M
            CFAngle = Deg2Rad
            CFForce = Lb2N
            CFVel = Kn2MPS
            CFChFr = Lb2N / Ft2M
        End If

    End Sub


    Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click

        Me.Close()

    End Sub


    Private Sub btnOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnOK.Click

        Cursor = System.Windows.Forms.Cursors.WaitCursor

        If UpdateStaticSolution() Then
            CurProj.Saved = False
            Cursor = System.Windows.Forms.Cursors.Default
            Me.Close()
        Else
            Cursor = System.Windows.Forms.Cursors.Default
        End If

    End Sub

    Public Sub mnuAddRow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAddRow.Click

        addRow(grdStaticSolution, CheckingGrid, inipoints+1)

    End Sub


    Public Sub mnuDeleteRow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDeleteRow.Click

        DeleteRow(grdStaticSolution, CheckingGrid, IniPoints + 1)

    End Sub
    Public Sub paste_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles pasteFromClipboard.Click

        PastefromClipBoardToDataGridView(grdStaticSolution, CheckingGrid, MaxSSPoints)

    End Sub

    Public Sub copyToExcel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles copyToExcel.Click

        copyToClipBoard(grdStaticSolution)

    End Sub


    Public Sub mnuFileClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileClose.Click

        btnCancel_Click(btnCancel, New System.EventArgs())

    End Sub


    Public Sub mnuOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuOpen.Click

        Dim Fn As String

        '   should the user cancel the dialog box, exit
        On Error GoTo ErrHandler

        '   set filters to allow selection of all files or just VIVO
        dlgImportVIVOOpen.Filter = "All Files (*.*)|*.*" '|VIVO|VIVO"

        '   specify default filter as VIVO
        dlgImportVIVOOpen.FilterIndex = 2

        '   display the Open dialog box
        dlgImportVIVOOpen.ShowReadOnly = False
        dlgImportVIVOOpen.InitialDirectory = CurProj.DataDirectory
        dlgImportVIVOOpen.ShowDialog()

        '   open the file
        tmpFileNum = FreeFile()
        FileOpen(tmpFileNum, dlgImportVIVOOpen.FileName, OpenMode.Input, OpenAccess.Read)

        '   build the object from the input data
        Cursor = System.Windows.Forms.Cursors.WaitCursor
        ImportVIVO((tmpFileNum))

        '    LoadGridFromProject
        Cursor = System.Windows.Forms.Cursors.Default

        '   close the input file and return
        FileClose(tmpFileNum)

        Exit Sub

ErrHandler:
        '   user pressed Cancel button
        Exit Sub

    End Sub


    Private Function UpdateStaticSolution() As Boolean

        Dim r, c As Short
        Dim NumPoints As Short
        Dim NewPoint As StaticSolutionPoint
        For r = 0 To grdStaticSolution.RowCount - 2
            For c = 0 To grdStaticSolution.ColumnCount - 1
                If String.IsNullOrEmpty(grdStaticSolution(c, r).Value) Then
                    MsgBox("No blank in all cells")
                    CheckingGrid = False
                    UpdateStaticSolution = False
                    Exit Function
                End If
            Next
        Next
        '   avoid triggering the "Leave Cell" event
        CheckingGrid = True

        With grdStaticSolution
           
            '       delete the old input
            NumPoints = CurProj.Riser(riserId).StaticSolution.Count

            For r = 1 To NumPoints
                CurProj.Riser(riserId).StaticSolution.Delete(1)
            Next r

            '       insert updated input
            If .RowCount > 1 Then
                For r = 0 To .RowCount - 2
                    NewPoint = CurProj.Riser(riserId).StaticSolution.Add
                    NewPoint.Distance = CSng(.Rows(r).Cells(DistanceCol).Value) * CFLength
                    If CSng(.Rows(r).Cells(TensionCol).Value) < 0.0# Then
                        MsgBox("Riser is in compression. Re-enter the data", MsgBoxStyle.OkOnly, "WinVIVA - Axial Tension")
                        UpdateStaticSolution = False
                        CheckingGrid = False
                        Exit Function
                    End If
                    NewPoint.AxialStaticTension = CSng(.Rows(r).Cells(TensionCol).Value) * CFForce

                    NewPoint.InPlaneAngle = CSng(.Rows(r).Cells(InAngleCol).Value) * CFAngle

                    '  NewPoint.OutPlaneAngle = CSng(.Rows(r).Cells(OutAngleCol).Value) * CFAngle

                    NewPoint.Depth = CSng(.Rows(r).Cells(DepthCol).Value) * CFLength
                    NewPoint.Ydist = CSng(.Rows(r).Cells(YdistCol).Value) * CFLength
                    NewPoint.Zdist = CSng(.Rows(r).Cells(ZdistCol).Value) * CFLength

                    NewPoint.UCurrentVelocity = CSng(.Rows(r).Cells(VelUCol).Value) * CFVel

                    NewPoint.VCurrentVelocity = CSng(.Rows(r).Cells(VelVCol).Value) * CFVel

                    NewPoint.ChangeInStaticTension = CSng(.Rows(r).Cells(TensionChangeCol).Value) * CFChFr
                Next
            End If

        End With

        '   reset the "Leaving Cell" trigger
        CheckingGrid = False
        UpdateStaticSolution = True

    End Function


    Private Sub LoadGridFromProject()

        Dim Point As StaticSolutionPoint
        Dim NumSSPoints As Short
        Dim r, c As Short
        Dim ydist As Single
        Dim delta, distance, Ygap, Zgap As Single


        CheckingGrid = True
        Ygap = 0
        Zgap = 0
        If riserId = 2 And CurProj.sameRiserARiser Then
            Ygap = (CurProj.RiserRLocY - CurProj.RiserFLocY) * CFLength

            Zgap = (CurProj.RiserRLocZ - CurProj.RiserFLocZ) * CFLength
        End If
        NumSSPoints = CurProj.Riser(riserId).StaticSolution.Count
        If riserId = 0 And CurProj.Riser(riserId).StaticSolution.Item(1).Ydist = 0 Then
            'this part only to initialize ydist 
            ydist = 0

            For r = NumSSPoints - 1 To 1 Step -1

                Point = CurProj.Riser(riserId).StaticSolution.Item(r + 1)
                distance = Point.Distance
                Point = CurProj.Riser(riserId).StaticSolution.Item(r)
                delta = Point.Distance - distance

                ydist = ydist + delta * System.Math.Cos(Point.InPlaneAngle)
                Point.Ydist = ydist
            Next
        End If
        If NumSSPoints > MaxSSPoints Then NumSSPoints = MaxSSPoints
        If NumSSPoints < IniPoints Then NumSSPoints = IniPoints
        'StaticSolution, calculate default y position by s and inplane angle
      
        With grdStaticSolution

            .RowCount = CurProj.Riser(riserId).StaticSolution.Count + 1

            .ColumnCount = NumColumns
            For c = 0 To .ColumnCount - 1
                .Columns(c).HeaderText = ColumnLabels(CurUnits, c + 1)
                .Columns(c).Width = 70
            Next
            r = 0
            For Each Point In CurProj.Riser(riserId).StaticSolution
                If r > MaxSSPoints Then Exit For
                .Rows(r).Cells(DistanceCol).Value = CStr(Point.Distance / CFLength)
                .Rows(r).Cells(TensionCol).Value = CStr(Point.AxialStaticTension / CFForce)
                .Rows(r).Cells(InAngleCol).Value = CStr(Point.InPlaneAngle / CFAngle)
                '.Rows(r).Cells(OutAngleCol).Value = CStr(Point.OutPlaneAngle / CFAngle)
                .Rows(r).Cells(DepthCol).Value = CStr(Point.Depth / CFLength)

                .Rows(r).Cells(YdistCol).Value = CStr((Point.Ydist + Ygap) / CFLength)
                .Rows(r).Cells(ZdistCol).Value = CStr((Point.Zdist + Zgap) / CFLength)
            
                .Rows(r).Cells(VelUCol).Value = CStr(Point.UCurrentVelocity / CFVel)

                .Rows(r).Cells(VelVCol).Value = CStr(Point.VCurrentVelocity / CFVel)

                .Rows(r).Cells(TensionChangeCol).Value = VB6.Format(Point.ChangeInStaticTension / CFChFr, "0.000")

                r = r + 1

            Next Point
        End With
        NumberAllRows(grdStaticSolution)
        CheckingGrid = False

    End Sub

 

    Private Sub btnEstimateTension_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnEstimateTension.Click

        Cursor = System.Windows.Forms.Cursors.WaitCursor
        ChangingRate(TensionCol, TensionChangeCol)
        Cursor = System.Windows.Forms.Cursors.Default

    End Sub


    Private Sub btnCurrentVel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCurrentVel.Click
        Dim r As Short
        Dim Depth As Single

        Cursor = System.Windows.Forms.Cursors.WaitCursor
        CheckingGrid = True

        With grdStaticSolution
            .RowCount = CurProj.Riser(riserId).StaticSolution.Count

            For r = 0 To .RowCount - 2
                Depth = .Rows(r).Cells(DepthCol - 1).Value
                .Rows(r).Cells(VelUCol).Value = CStr(CurProj.Water.CurrentProfile.VofD(Depth * CFLength, 1) / CFVel)
                .Rows(r).Cells(VelVCol).Value = CStr(CurProj.Water.CurrentProfile.VofD(Depth * CFLength, 2) / CFVel)
            Next

        End With

        CheckingGrid = False
        Cursor = System.Windows.Forms.Cursors.Default

    End Sub


    Private Sub ChangingRate(ByVal InputCol As Short, ByVal OutputCol As Short)

        Dim A1, DeltaA1, DeltaA2, A2 As Single
        Dim S1, DeltaS1, DeltaS2, S2 As Single
        Dim r As Short
        Dim CR As Single
        Dim DataEnd As Boolean

        CheckingGrid = True

        With grdStaticSolution
            
            If .RowCount < 2 Then
                MsgBox("There must be at least two points to compute curvature!", MsgBoxStyle.OkOnly, "WinVIVA - Data for SCR") '
                'Or .Rows(0).Cells(InputCol).Value = "" Or .Rows(0).Cells(DistanceCol).Value = "" _
                '               Or .Rows(1).Cells(InputCol).Value = "" Or .Rows(1).Cells(DistanceCol).Value = "" Then
                '              MsgBox("There must be at least two points to compute curvature!", MsgBoxStyle.OkOnly, "WinVIVA - Data for SCR")
                Exit Sub
            End If

            A1 = CSng(.Rows(0).Cells(InputCol).Value)
            S1 = CSng(.Rows(0).Cells(DistanceCol).Value)
            A2 = CSng(.Rows(1).Cells(InputCol).Value)
            S2 = CSng(.Rows(1).Cells(DistanceCol).Value)
            DeltaA1 = A2 - A1
            DeltaS1 = S2 - S1

            If DeltaS1 = 0 Then
                MsgBox("You may not have two points at the same distance along the riser", MsgBoxStyle.OkOnly, "WinVIVA - Data for SCR")
                Exit Sub
            End If

            CR = DeltaA1 / DeltaS1
            If OutputCol = TensionChangeCol Then CR = -CR
            .Rows(0).Cells(OutputCol).Value = VB6.Format(CR, "0.000")

            r = 0
            DeltaA2 = DeltaA1
            DeltaS2 = DeltaS1
            DataEnd = False

            Do While Not DataEnd
                DeltaA1 = DeltaA2
                DeltaS1 = DeltaS2
                A1 = A2
                S1 = S2
                r = r + 1

                If r = .RowCount - 2 Then
                    DataEnd = True
                Else

                    If .Rows(r + 1).Cells(InputCol).Value = "" Then
                        DataEnd = True
                    Else
                        A2 = CSng(.Rows(r + 1).Cells(InputCol).Value)
                    End If

                    If .Rows(r + 1).Cells(DistanceCol).Value = "" Then
                        DataEnd = True
                    Else
                        S2 = CSng(.Rows(r + 1).Cells(DistanceCol).Value)
                    End If
                End If

                If DataEnd Then
                    CR = DeltaA1 / DeltaS1
                Else
                    DeltaA2 = A2 - A1
                    DeltaS2 = S2 - S1

                    If System.Math.Abs(DeltaS2) <= 0.001 Then
                        If System.Math.Abs(DeltaS1) <= 0.001 Then
                            MsgBox("You may not have three points at the same distance" & " along the riser", MsgBoxStyle.OkOnly, "WinVIVA - Data for SCR")
                            Exit Sub
                        Else
                            CR = DeltaA1 / DeltaS1
                        End If
                    Else
                        If System.Math.Abs(DeltaS1) <= 0.001 Then
                            CR = DeltaA2 / DeltaS2
                        Else
                            CR = (DeltaA1 / DeltaS1 * DeltaS2 + DeltaA2 / DeltaS2 * DeltaS1) / (DeltaS1 + DeltaS2)
                        End If
                    End If
                End If

                If OutputCol = TensionChangeCol Then CR = -CR
                .Rows(r).Cells(OutputCol).Value = VB6.Format(CR, "0.000")
            Loop
        End With

        CheckingGrid = False

    End Sub


    ' import user prepared static solution data to the interface
    ' the  format is different from vivo-n
    ' 
    Private Sub ImportVIVO(ByRef FileNum As Short)

        Dim NumSSPoints, IniPointsNum As Short
        Dim r As Short
        'Dim dum As Single
        Dim Angl, Dist, Tens, Depth As Single

        Dim PCFLength, PCFForce As Single
        Dim PUnits As VIVAMain.Units

        CheckingGrid = True

        Input(FileNum, NumSSPoints)
        Input(FileNum, PUnits)
        If NumSSPoints > MaxSSPoints Then NumSSPoints = MaxSSPoints

        If NumSSPoints < IniPoints Then
            IniPointsNum = IniPoints
        Else
            IniPointsNum = NumSSPoints
        End If

        If PUnits = VIVAMain.Units.Metric Then
            PCFLength = 1
            PCFForce = 1
        Else
            PCFLength = Ft2M
            PCFForce = Lb2N
        End If

        With grdStaticSolution
           
            For r = 1 To NumSSPoints

                Input(FileNum, Dist)
                Input(FileNum, Tens)
                Input(FileNum, Angl)
                Input(FileNum, Depth)
                .Rows(r).Cells(DistanceCol).Value = CStr(Dist * PCFLength / CFLength)

                .Rows(r).Cells(TensionCol).Value = CStr(Tens * PCFForce / CFForce)

                .Rows(r).Cells(InAngleCol).Value = CStr(Angl)
                '.Rows(r).Cells(OutAngleCol).Value = ""

                .Rows(r).Cells(DepthCol).Value = CStr(Depth * PCFLength / CFLength)

                .Rows(r).Cells(YdistCol).Value = ""

                .Rows(r).Cells(ZdistCol).Value = ""

                .Rows(r).Cells(VelUCol).Value = ""
                .Rows(r).Cells(VelVCol).Value = ""
                .Rows(r).Cells(TensionChangeCol).Value = ""

            Next

        End With

        CheckingGrid = False

    End Sub

End Class