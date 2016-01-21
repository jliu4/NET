Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module modPlot
	'Option Explicit
	
	Private Const NumVesselPoints As Short = 16
	Private Const NumArrowPoints As Short = 8
	Private Const NumNorthLetterPoints As Short = 4
	Private Const NumNorthArrowPoints As Short = 7

    Private Const MarginPercent As Single = 0.2
    Private Const TimerInterval As Single = 0.25
    Private Const Fuzz As Single = 0.1

    Private VesselDiagram(NumVesselPoints, 2) As Single
    Private CompassArrow(NumArrowPoints, 2) As Single
    Private NorthArrow(NumNorthArrowPoints, 2) As Single
    Private NorthLetter(NumNorthLetterPoints, 2) As Single
    Private CurrentX As Single, CurrentY As Single

    Public Sub ResizePicture(ByRef pic As System.Windows.Forms.PictureBox, ByRef LeftLimit As Single, ByRef RightLimit As Single, ByRef TopLimit As Single, ByRef BottomLimit As Single)
        Dim Margin As Single

        If pic.FindForm.WindowState <> System.Windows.Forms.FormWindowState.Minimized Then

            Margin = 400.0#

            With pic
                .Left = VB6.TwipsToPixelsX(LeftLimit + Margin)
                .Width = VB6.TwipsToPixelsX(RightLimit - LeftLimit - Margin)
                .Top = VB6.TwipsToPixelsY(TopLimit + Margin)
                .Height = VB6.TwipsToPixelsY(BottomLimit - TopLimit - Margin)
            End With

        End If

    End Sub

    ' make rig nodes
    Public Sub MakeVesselPoints(ByRef ScaleFactor As Single)

        Dim i As Short

        VesselDiagram(1, mX) = 175.5
        VesselDiagram(2, mX) = 195.2
        VesselDiagram(3, mX) = VesselDiagram(1, mX)
        VesselDiagram(4, mX) = 146.0#
        VesselDiagram(5, mX) = VesselDiagram(4, mX)
        VesselDiagram(6, mX) = VesselDiagram(1, mX)
        VesselDiagram(7, mX) = VesselDiagram(2, mX)
        VesselDiagram(8, mX) = VesselDiagram(1, mX)
        VesselDiagram(9, mX) = -VesselDiagram(1, mX)
        VesselDiagram(10, mX) = -VesselDiagram(2, mX)
        VesselDiagram(11, mX) = -VesselDiagram(1, mX)
        VesselDiagram(12, mX) = -VesselDiagram(4, mX)
        VesselDiagram(13, mX) = -VesselDiagram(4, mX)
        VesselDiagram(14, mX) = -VesselDiagram(1, mX)
        VesselDiagram(15, mX) = -VesselDiagram(2, mX)
        VesselDiagram(16, mX) = -VesselDiagram(1, mX)

        VesselDiagram(1, mY_Renamed) = 110.0#
        VesselDiagram(2, mY_Renamed) = 90.3
        VesselDiagram(3, mY_Renamed) = 70.6
        VesselDiagram(4, mY_Renamed) = VesselDiagram(3, mY_Renamed)
        VesselDiagram(5, mY_Renamed) = -VesselDiagram(3, mY_Renamed)
        VesselDiagram(6, mY_Renamed) = -VesselDiagram(3, mY_Renamed)
        VesselDiagram(7, mY_Renamed) = -VesselDiagram(2, mY_Renamed)
        VesselDiagram(8, mY_Renamed) = -VesselDiagram(1, mY_Renamed)
        VesselDiagram(9, mY_Renamed) = -VesselDiagram(1, mY_Renamed)
        VesselDiagram(10, mY_Renamed) = -VesselDiagram(2, mY_Renamed)
        VesselDiagram(11, mY_Renamed) = -VesselDiagram(3, mY_Renamed)
        VesselDiagram(12, mY_Renamed) = -VesselDiagram(3, mY_Renamed)
        VesselDiagram(13, mY_Renamed) = VesselDiagram(3, mY_Renamed)
        VesselDiagram(14, mY_Renamed) = VesselDiagram(3, mY_Renamed)
        VesselDiagram(15, mY_Renamed) = VesselDiagram(2, mY_Renamed)
        VesselDiagram(16, mY_Renamed) = VesselDiagram(1, mY_Renamed)

        If ScaleFactor <> 1.0# Then
            For i = 1 To NumVesselPoints
                VesselDiagram(i, mX) = ScaleFactor * VesselDiagram(i, mX)
                VesselDiagram(i, mY_Renamed) = ScaleFactor * VesselDiagram(i, mY_Renamed)
            Next i
        End If

    End Sub

    ' generate compass arrow points
    Private Sub MakeCompassArrowPoints(ByRef ScaleFactor As Single)

        ' input
        '   ScaleFactor -- scale factor of the arrow

        Dim i As Short

        CompassArrow(1, mX) = -10.0#
        CompassArrow(2, mX) = -4.0#
        CompassArrow(3, mX) = 10.0#
        CompassArrow(4, mX) = 2.0#
        CompassArrow(5, mX) = 2.0#
        CompassArrow(6, mX) = 10.0#
        CompassArrow(7, mX) = -4.0#
        CompassArrow(8, mX) = -10.0#

        CompassArrow(1, mY_Renamed) = 3.0#
        CompassArrow(2, mY_Renamed) = 0#
        CompassArrow(3, mY_Renamed) = 0#
        CompassArrow(4, mY_Renamed) = 3.0#
        CompassArrow(5, mY_Renamed) = -3.0#
        CompassArrow(6, mY_Renamed) = 0#
        CompassArrow(7, mY_Renamed) = 0#
        CompassArrow(8, mY_Renamed) = -3.0#

        If ScaleFactor <> 1.0# Then
            For i = 1 To NumArrowPoints
                CompassArrow(i, mX) = ScaleFactor * CompassArrow(i, mX)
                CompassArrow(i, mY_Renamed) = ScaleFactor * CompassArrow(i, mY_Renamed)
            Next i
        End If

    End Sub

    ' generate north pointer points
    Public Sub MakeNorthPoints(ByRef ArrowScaleFactor As Single, ByRef LetterScaleFactor As Single)

        Dim i As Short

        NorthArrow(1, mX) = -2.0#
        NorthArrow(2, mX) = 0#
        NorthArrow(3, mX) = 0#

        NorthArrow(1, mY_Renamed) = -8.0#
        NorthArrow(2, mY_Renamed) = -10.0#
        NorthArrow(3, mY_Renamed) = 10.0#

        If ArrowScaleFactor <> 1.0# Then
            For i = 1 To NumNorthArrowPoints
                NorthArrow(i, mX) = ArrowScaleFactor * NorthArrow(i, mX)
                NorthArrow(i, mY_Renamed) = ArrowScaleFactor * NorthArrow(i, mY_Renamed)
            Next i
        End If

        NorthLetter(1, mX) = -5.0#
        NorthLetter(2, mX) = -5.0#
        NorthLetter(3, mX) = 5.0#
        NorthLetter(4, mX) = 5.0#

        NorthLetter(1, mY_Renamed) = 5.0#
        NorthLetter(2, mY_Renamed) = -5.0#
        NorthLetter(3, mY_Renamed) = 5.0#
        NorthLetter(4, mY_Renamed) = -5.0#

        If LetterScaleFactor <> 1.0# Then
            For i = 1 To NumNorthLetterPoints
                NorthLetter(i, mX) = LetterScaleFactor * NorthLetter(i, mX)
                NorthLetter(i, mY_Renamed) = LetterScaleFactor * NorthLetter(i, mY_Renamed)
            Next i
        End If

    End Sub

    ' draw the basic vessel shape
    Public Sub DrawVessel(ByRef X0 As Single, ByRef Y0 As Single, ByRef Head0 As Single, ByRef ScaleFactor As Single, ByRef gr As Graphics, ByRef pen As Pen, Optional ByRef LineColor As Short = 0)

        Dim i As Short

        Call MakeVesselPoints(ScaleFactor)
        Call RotatePoints(VesselDiagram, -Head0)

        For i = 1 To NumVesselPoints - 1
            gr.DrawLine(pen, X0 + VesselDiagram(i, mX), Y0 + VesselDiagram(i, mY_Renamed), X0 + VesselDiagram(i + 1, mX), Y0 + VesselDiagram(i + 1, mY_Renamed))
        Next i
        gr.DrawLine(pen, X0 + VesselDiagram(NumVesselPoints, mX), Y0 + VesselDiagram(NumVesselPoints, mY_Renamed), X0 + VesselDiagram(1, mX), Y0 + VesselDiagram(1, mY_Renamed))
    End Sub

    ' draw compass
    Public Sub DrawCompassArrow(ByRef X0 As Single, ByRef Y0 As Single, ByRef ScaleFactor As Single, ByRef angle As Single, ByRef gr As Graphics, ByRef pen As Pen, ByRef Color1 As Short)
        Dim i As Object

        Call MakeCompassArrowPoints(ScaleFactor)
        Call RotatePoints(CompassArrow, -angle)

        'pen.Color = Color.FromArgb(QBColor(Color1))
        pen.Width = 1

        For i = 1 To NumArrowPoints - 1

            gr.DrawLine(pen, X0 + CompassArrow(i, mX), Y0 + CompassArrow(i, mY_Renamed), X0 + CompassArrow(i + 1, mX), Y0 + CompassArrow(i + 1, mY_Renamed))
        Next i

    End Sub

    ' plot environment directions
    Public Sub DrawEnvPlot(ByVal WindDir As Single, ByVal WaveDir As Single, ByVal CurrDir As Single, ByVal VslHead As Single, ByRef pic As System.Windows.Forms.PictureBox)
        ' Input
        '   WindDir -- wind force direction (local)
        '   WaveDir -- wind force direction (local)
        '   CurrDir -- current force direction (local)
        '   VslHead -- vessel heading (clockwise from North)
        '   picEnviron -- reference to output picture box
        ' Output
        '   picEnviron -- output picture


        Dim Margin, ArrowSpacing As Single
        Dim Theight, Twidth, ALWidth As Single

        Dim BoxHeight, BoxWidth, BoxSize As Integer
        Dim Cx, Px, Py, Cy As Single
        Dim Ymax, Xmax, Xmin, Ymin As Single
        Dim VScale, Vmax, Amax, Ascale As Single
        Dim Rwave, Rcc, Rwind, Rcurrent As Single

        BoxWidth = pic.ClientSize.Width
        BoxHeight = pic.ClientSize.Height
        If BoxWidth < BoxHeight Then
            BoxSize = BoxWidth
        Else
            BoxSize = BoxHeight
        End If

        Dim bm As Bitmap = New Bitmap(BoxWidth, BoxHeight)
        Dim gr As Graphics = Graphics.FromImage(bm)
        Dim pen As New System.Drawing.Pen(Color.Black, 1)

        '   Center the graphics in the picture box
        Px = BoxWidth / 2.0#
        Py = BoxHeight / 2.0#

        '   Adjust the force directions to E-N system
        VslHead = PI / 2 - VslHead
        WindDir = VslHead + WindDir
        WaveDir = VslHead + WaveDir
        CurrDir = VslHead + CurrDir

        '   Find the limits of the vessel, for use in sizing the circle that will
        '   encompass it
        Call MakeVesselPoints(1.0#)
        Call ExtremePoints(VesselDiagram, Xmax, Xmin, Ymax, Ymin)

        If Xmax > Ymax Then
            Vmax = Xmax
        Else
            Vmax = Ymax
        End If

        VScale = 0.15 * BoxSize / Vmax

        '   Draw the vessel itself at the center
        Call DrawVessel(Px, Py, VslHead, VScale, gr, pen)


        '       Draw the circle around the vessel, leaving some room
        Rcc = VScale * Vmax * 1.2
        Margin = Rcc / 10
        gr.DrawEllipse(pen, Px - Rcc, Py - Rcc, 2 * Rcc, 2 * Rcc)
        Twidth = gr.MeasureString("N", pic.Font).Width
        Theight = gr.MeasureString("N", pic.Font).Height
        '       Annotate the compass points on the circle
        pic.Font = VB6.FontChangeSize(pic.Font, 8)
        pic.Font = VB6.FontChangeBold(pic.Font, False)
        gr.DrawString("N", pic.Font, pen.Brush, Px - Twidth / 2, Py - Rcc - Margin - Theight)
        gr.DrawString("S", pic.Font, pen.Brush, Px - Twidth / 2, Py + Rcc + Margin)
        gr.DrawString("W", pic.Font, pen.Brush, Px - Rcc - Margin - Twidth, Py - Theight / 2)
        gr.DrawString("E", pic.Font, pen.Brush, Px + Rcc + Margin, Py - Theight / 2)

        '   Find the maximum dimension of the arrow
        Call MakeCompassArrowPoints(1.0#)
        Call ExtremePoints(CompassArrow, Xmax, Xmin, Ymax, Ymin)

        If Xmax > Ymax Then
            Amax = Xmax
        Else
            Amax = Ymax
        End If

        Ascale = 0.2 * VScale * Vmax / Amax
        ArrowSpacing = Amax * Ascale / 1.0#

        '   Compute the radii for the three arrows representing wind, wave and current
        Rwind = Rcc + ArrowSpacing + Theight + Margin
        Rwave = Rwind + ArrowSpacing
        Rcurrent = Rwave + ArrowSpacing

        '   For each arrow, compute the position of its center, then draw it
        Cx = Px + Rwind * System.Math.Cos(WindDir)
        Cy = Py - Rwind * System.Math.Sin(WindDir)

        pen = New System.Drawing.Pen(Color.Blue, 0.5)

        Call DrawCompassArrow(Cx, Cy, Ascale, WindDir, gr, pen, 9)

        Cx = Px + Rwave * System.Math.Cos(WaveDir)
        Cy = Py - Rwave * System.Math.Sin(WaveDir)

        pen = New System.Drawing.Pen(Color.Green, 0.5)

        Call DrawCompassArrow(Cx, Cy, Ascale, WaveDir, gr, pen, 2)

        Cx = Px + Rcurrent * System.Math.Cos(CurrDir)
        Cy = Py - Rcurrent * System.Math.Sin(CurrDir)
        pen = New System.Drawing.Pen(Color.Red, 0.5)
        Call DrawCompassArrow(Cx, Cy, Ascale, CurrDir, gr, pen, 12)

        '   Annotate the plot
        pic.Font = VB6.FontChangeSize(pic.Font, 7)

        pic.Font = VB6.FontChangeBold(pic.Font, False)

        Twidth = gr.MeasureString("Wind  Wave  Current", pic.Font).Width
        Theight = gr.MeasureString("Wind  Wave  Current", pic.Font).Height

        ALWidth = (BoxWidth - Twidth - 4 * Margin) / 3 - 2 * Margin

        Px = 3 * Margin
        Py = Py + Rcurrent + Amax * Ascale + Theight / 2
        'pen.Color = Color.FromArgb(QBColor(9))
        pen = New System.Drawing.Pen(Color.Blue, 0.5)
        gr.DrawLine(pen, Px, Py, Px + ALWidth, Py)
        Px = Px + ALWidth + Margin
        Py = Py - Theight / 2
        pen = New System.Drawing.Pen(Color.Black, 0.5)
        gr.DrawString("Wind", pic.Font, pen.Brush, Px, Py)

        Px = Px + gr.MeasureString("Wind", pic.Font).Width + 3 * Margin
        Py = Py + Theight / 2
        'pen.Color = Color.FromArgb(QBColor(2))
        pen = New System.Drawing.Pen(Color.Green, 0.5)
        gr.DrawLine(pen, Px, Py, Px + ALWidth, Py)
        Px = Px + ALWidth + Margin
        Py = Py - Theight / 2
        pen = New System.Drawing.Pen(Color.Black, 0.5)
        gr.DrawString("Wave", pic.Font, pen.Brush, Px, Py)

        Px = Px + gr.MeasureString("Wave", pic.Font).Width + 3 * Margin
        Py = Py + Theight / 2
        'pen.Color = Color.FromArgb(QBColor(12))
        pen = New System.Drawing.Pen(Color.Red, 0.5)
        gr.DrawLine(pen, Px, Py, Px + ALWidth, Py)
        Px = Px + ALWidth + Margin
        Py = Py - Theight / 2
        pen = New System.Drawing.Pen(Color.Black, 0.5)
        gr.DrawString("Current", pic.Font, pen.Brush, Px, Py)
        Dim dest_bounds As New RectangleF(0, 0, BoxWidth, BoxHeight)
        Dim source_bounds As New RectangleF(0, 0, BoxWidth + 1, BoxHeight + 1)
        gr.DrawImage(bm, dest_bounds, source_bounds, GraphicsUnit.Pixel)
        pic.SizeMode = PictureBoxSizeMode.AutoSize
        pic.Image = bm
        pen.Dispose()
        gr.Dispose()
    End Sub

    ' rotate plot points
    Private Sub RotatePoints(ByRef x As Single(,), ByRef angle As Single)

        ' input
        '   X -- point coordinates
        '   Angle -- rotating angle
        ' output
        '   X -- point coordinates after rotating

        Dim i As Short
        Dim Ax, c, S, Ay As Single

        c = System.Math.Cos(angle)
        S = System.Math.Sin(angle)

        For i = 1 To UBound(x)
            Ax = c * x(i, mX) - S * x(i, mY_Renamed)
            Ay = S * x(i, mX) + c * x(i, mY_Renamed)
            x(i, mX) = Ax
            x(i, mY_Renamed) = Ay
        Next i

    End Sub

    Public Sub ExtremePoints(ByRef GraphicArray As Single(,), ByRef Xmax As Single, ByRef Xmin As Single, ByRef Ymax As Single, ByRef Ymin As Single)

        Dim i As Short

        Xmax = 1.0E-38
        Ymax = Xmax
        Xmin = 1.0E+38
        Ymin = Xmin

        For i = 1 To UBound(GraphicArray)
            If GraphicArray(i, mX) > Xmax Then Xmax = GraphicArray(i, mX)
            If GraphicArray(i, mX) < Xmin Then Xmin = GraphicArray(i, mX)
            If GraphicArray(i, mY_Renamed) > Ymax Then Ymax = GraphicArray(i, mY_Renamed)
            If GraphicArray(i, mY_Renamed) < Ymin Then Ymin = GraphicArray(i, mY_Renamed)
        Next i

    End Sub
    Public Sub DrawTransientPlot(ByVal X0 As Single, ByVal Y0 As Single, ByRef NumPlotPoints As Short, ByRef Xv As Object, ByRef Yv As Object, ByRef Headv As Object, ByVal MaxTransX As Single, ByVal MaxTransY As Single, ByRef pic As System.Windows.Forms.PictureBox)
        Dim ZWD As Single
        Dim i As Short

        Dim Px, Py As Single
        Dim tX, Twidth, Theight, tY As Single
        Dim Ymax, Xmax, Xmin, Ymin As Single
        Dim BoxScale, Vmax, BoxSize, VScale As Single
        Dim HalfSquare, Dist As Single
        Dim BoxWidth, BoxHeight As Integer
        Dim Rsc(20) As Single
        Dim StartTime As Single

        ZWD = CurVessel.WaterDepth

        ' How big is the picture box we're using?  Find the smaller of the width and
        ' height, and we'll use it to compute the plot scale a little later on...

        BoxWidth = pic.ClientSize.Width
        BoxHeight = pic.ClientSize.Height
        If BoxWidth < BoxHeight Then
            BoxSize = BoxWidth
        Else
            BoxSize = BoxHeight
        End If

        Px = BoxWidth / 2.0#
        Py = BoxHeight / 2.0#
        HalfSquare = 0.01 * BoxSize

        Dim bm As Bitmap = New Bitmap(BoxWidth, BoxHeight)
        Dim gr As Graphics = Graphics.FromImage(bm)
        Dim pen As New System.Drawing.Pen(Color.Black, 1)

        'gr.DrawEllipse(pen, Px, Py, 10, 10)
        'draw well as Color.AliceBlue square
        gr.DrawLine(pen, Px + HalfSquare, Py + HalfSquare, Px - HalfSquare, Py + HalfSquare)
        gr.DrawLine(pen, Px - HalfSquare, Py + HalfSquare, Px - HalfSquare, Py - HalfSquare)
        gr.DrawLine(pen, Px - HalfSquare, Py - HalfSquare, Px + HalfSquare, Py - HalfSquare)
        gr.DrawLine(pen, Px + HalfSquare, Py - HalfSquare, Px + HalfSquare, Py + HalfSquare)

        'compute the distance through which the vessel has moved
        Dist = System.Math.Sqrt((MaxTransX - X0) ^ 2 + (MaxTransY - Y0) ^ 2)

        ' For very small offsets, this is confusing; make sure we don't make the
        ' plot scale too big

        If Dist < 0.02 * ZWD Then Dist = 0.02 * ZWD

        Call MakeVesselPoints(1.0#)
        Call ExtremePoints(VesselDiagram, Xmax, Xmin, Ymax, Ymin)
        If Xmax > Ymax Then
            Vmax = Xmax
        Else
            Vmax = Ymax
        End If
        VScale = 0.05 * BoxSize / Vmax
        ' And use it and the maximum distance to compute a scaling factor for the plot
        BoxScale = 0.45 * BoxSize / (Dist + Vmax / 2.0#)

        ' Draw a series of circles indicating fractions of the water depth
        ' While we're at it, annotate each



        '.Font = VB6.FontChangeSize(.Font, 4)
        '.Font = VB6.FontChangeBold(.Font, False)
        'UPGRADE_ISSUE: PictureBox method picPG.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Twidth = .TextWidth("99%")
        'UPGRADE_ISSUE: PictureBox method picPG.TextHeight was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Theight = .TextHeight("99%")

        For i = 1 To 20
            'UPGRADE_WARNING: Couldn't resolve default property of object ZWD. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If Dist > 0.4 * ZWD Then
                If i <= 10 Then
                    Rsc(i) = 0.05 * i
                Else
                    Rsc(i) = Rsc(i - 1) + 0.1
                End If
                'UPGRADE_WARNING: Couldn't resolve default property of object ZWD. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf Dist > 0.2 * ZWD Then
                If i <= 10 Then
                    Rsc(i) = 0.02 * i
                Else
                    Rsc(i) = Rsc(i - 1) + 0.05
                End If
                'UPGRADE_WARNING: Couldn't resolve default property of object ZWD. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf Dist > 0.1 * ZWD Then
                If i <= 10 Then
                    Rsc(i) = 0.01 * i
                Else
                    Rsc(i) = Rsc(i - 1) + 0.02
                End If
            Else
                If i <= 4 Then
                    Rsc(i) = 0.005 * i
                Else
                    Rsc(i) = Rsc(i - 1) + 0.01
                End If
            End If
            If Rsc(i) = 0.01 Then
                pen.Color = Color.FromArgb(QBColor(Rnd() * 12))
                gr.DrawEllipse(pen, Px - Rsc(i) * ZWD * BoxScale, Py - Rsc(i) * ZWD * BoxScale, 2 * Rsc(i) * ZWD * BoxScale, 2 * Rsc(i) * ZWD * BoxScale)
            ElseIf Rsc(i) = 0.02 Then
                pen.Color = Color.FromArgb(QBColor(Rnd() * 9))
                gr.DrawEllipse(pen, Px - Rsc(i) * ZWD * BoxScale, Py - Rsc(i) * ZWD * BoxScale, 2 * Rsc(i) * ZWD * BoxScale, 2 * Rsc(i) * ZWD * BoxScale)

            Else
                gr.DrawEllipse(pen, Px - Rsc(i) * ZWD * BoxScale, Py - Rsc(i) * ZWD * BoxScale, 2 * Rsc(i) * ZWD * BoxScale, 2 * Rsc(i) * ZWD * BoxScale)
            End If
            tX = Px + Twidth * 0.2
            tY = Py + Rsc(i) * ZWD * BoxScale - Theight
            'UPGRADE_ISSUE: PictureBox property picPG.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            ' .CurrentX = tX
            'UPGRADE_ISSUE: PictureBox property picPG.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.CurrentY = tY
            'UPGRADE_ISSUE: PictureBox method picPG.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            'picPG.Print(Format(Rsc(i), "#0.0%"))

        Next i
        pic.Font = VB6.FontChangeSize(pic.Font, 15)
        pic.Font = VB6.FontChangeBold(pic.Font, True)

        gr.DrawString("N", pic.Font, pen.Brush, Px, 0)
        gr.DrawLine(pen, Px, 0, Px, BoxHeight)
        gr.DrawLine(pen, 0, Py, BoxWidth, Py)
        Headv(1) = -Headv(1) + PI / 2.0#
        pen.Width = 3
        Call DrawVessel(Px + BoxScale * (Xv(1) - X0), Py - BoxScale * (Yv(1) - Y0), (Headv(1)), VScale, gr, pen)
        pen.Width = 1
        StartTime = VB.Timer()
        For i = 2 To NumPlotPoints - 1
            Do While VB.Timer() < StartTime + TimerInterval
            Loop
            Headv(i) = -Headv(i) + PI / 2.0# '  heading in radius
            Call DrawVessel(Px + BoxScale * (Xv(i) - X0), Py - BoxScale * (Yv(i) - Y0), (Headv(i)), VScale, gr, pen)
            frmTransient.Refresh()
            StartTime = VB.Timer()
        Next i
        pen.Width = 3 '  draw black color for final position
        pen.Color = Color.Red
        Headv(NumPlotPoints) = -Headv(NumPlotPoints) + PI / 2.0#
        Call DrawVessel(Px + BoxScale * (Xv(NumPlotPoints) - X0), Py - BoxScale * (Yv(NumPlotPoints) - Y0), (Headv(NumPlotPoints)), VScale, gr, pen)
        Dim dest_bounds As New RectangleF(0, 0, BoxWidth, BoxHeight)
        Dim source_bounds As New RectangleF(0, 0, BoxWidth + 1, BoxHeight + 1)
        gr.DrawImage(bm, dest_bounds, source_bounds, GraphicsUnit.Pixel)
        pic.SizeMode = PictureBoxSizeMode.AutoSize
        pic.Image = bm
        pen.Dispose()
        gr.Dispose()

    End Sub

End Module