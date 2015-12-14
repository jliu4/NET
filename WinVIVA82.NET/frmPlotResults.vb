Option Strict Off
Option Explicit On
Friend Class frmPlotResults
	Inherits System.Windows.Forms.Form
	
	'Const WaveRealColor As Integer = 2
	'Const WaveImagColor As Integer = 12
	'Const WaveMagColor As Integer = 9
	'Const PhaseAngleColor As Integer = 5
	'Const CurveColor As Integer = 9

    Dim WaveRealColor As Color = Color.Lime
    Dim WaveImagColor As Color = Color.Yellow
    Dim WaveMagColor As Color = Color.Red
    Dim PhaseAngleColor As Color = Color.Blue
    Dim CurveColor As Color = Color.Blue
    Dim wmf_file_name As String

	Private ModeNum, PlotOption, NumPoints As Short
	Private CurveUnits As VIVAMain.Units
	Private PrintPlot As Boolean
    Private DispDist(NumVIVACorePoints) As Single
	Private Ymin, Ymax As Single
    Private YLab As String

    Private Sub frmPlotResults_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed

        Try
            Kill(wmf_file_name)
        Catch ex As Exception
        End Try

        Me.Dispose()

    End Sub


    Private Sub frmPlotResults_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        Dim i As Short

        'generate wmf file name
        Dim path_name As String = Environment.GetEnvironmentVariable("Temp", EnvironmentVariableTarget.User)
        wmf_file_name = RandomFileName(path_name, "tmp")
        'delete file if it exists
        If Len(Dir$(wmf_file_name)) > 0 Then Kill(wmf_file_name)

        Text = Text & " - " & CurProj.Title

        If NumOutputModes = 0 Then
            cmbMode.Items.Add("")
        Else
            For i = 1 To NumOutputModes
                cmbMode.Items.Add(CStr(ModeNumber(i)))
            Next i
            cmbMode.Items.Add("MM")
        End If

        cmbMode.SelectedIndex = 0
        lblMode.Enabled = False
        cmbMode.Enabled = False

        PlotOption = tabPlot.SelectedItem.Index - 1
        picResults.SetBounds(tabPlot.ClientLeft, tabPlot.ClientTop, tabPlot.ClientWidth, tabPlot.ClientHeight)

        PrintPlot = False
        'do not need mode shape at this moment
        tabPlot.Tabs.Remove(8)

        PlotCurve()

    End Sub


    Private Sub PlotCurve()

        Dim i As Short
        Dim XLab As String
        Dim XValues(NumVIVACorePoints) As Single
        Dim ConversionFactor As Single

        CurveUnits = CurProj.Units

        If CurveUnits = VIVAMain.Units.English Then
            YLab = "Distance (ft)"
            ConversionFactor = 1.0# / Ft2M
        Else
            YLab = "Distance (m)"
            ConversionFactor = 1.0#
        End If

        NumPoints = CurProj.NumPoints

        For i = 1 To NumPoints
            'DispDist(i) = Distance(i) * ConversionFactor
        Next i

        Call MaxAndMin(DispDist, Ymax, Ymin, NumPoints)

        Select Case PlotOption
            Case 0
                If CurveUnits = VIVAMain.Units.English Then
                    ConversionFactor = 1.0# / Kn2MPS
                    XLab = "Current Velocity (knts)"
                Else
                    ConversionFactor = 1.0#
                    XLab = "Current Velocity (m/s)"
                End If
                For i = 1 To NumPoints
                    ' XValues(i) = VelA(i) * ConversionFactor
                Next i
                Call Plot1DArray(XValues, XLab)

            Case 1
                ' Call Plot1DArray(ReynoldsNumber, "Reynolds Number")

            Case 2
                If CurveUnits = VIVAMain.Units.English Then
                    ConversionFactor = 1.0# / Ft2M
                    Call Plot2DArrayColumn(Amplitude, "Amplitude (ft)", ConversionFactor)
                Else
                    Call Plot2DArrayColumn(Amplitude, "Amplitude (m)", 1.0#)
                End If
                '            Call Plot2DArrayColumn(Amplitude, "Amplitude", 1# / CFLength) 'AoverD, #1
            Case 3
                Call Plot2DArrayColumn(DragCoef, "Drag Coefficient", 1.0#)

            Case 4
                If CurveUnits = VIVAMain.Units.English Then
                    ConversionFactor = 1.0# / (Lb2N * Ft2M)
                    Call Plot2DArrayColumn(BndMoment, "Bending Moment (lb-ft)", ConversionFactor)
                Else
                    Call Plot2DArrayColumn(BndMoment, "Bending Moment (N-m)", 1.0#)
                End If

            Case 5
                If CurveUnits = VIVAMain.Units.English Then
                    ConversionFactor = (In2Ft * Ft2M) ^ 2 / Lb2N
                    Call Plot2DArrayColumn(Stress, "Bending Stress (psi)", ConversionFactor)
                Else
                    ConversionFactor = 0.000001
                    Call Plot2DArrayColumn(Stress, "Bending Stress (MPa)", ConversionFactor)
                End If

            Case 6
                If CurveUnits = VIVAMain.Units.English Then
                    ConversionFactor = 1000000.0
                    Call Plot2DArrayColumn(Strain, "Bending Strain (10E-6)", ConversionFactor)
                Else
                    ConversionFactor = 1000000.0
                    Call Plot2DArrayColumn(Strain, "Bending Strain (10E-6)", ConversionFactor)
                End If

            Case 7
                Call PlotModeShapes()
        End Select

    End Sub


    Private Sub PlotModeShapes()

        Dim i As Short
        Dim XLab As String
        Dim Xmax, Xmin As Single
        Dim XmaxTmp1, XmaxTmp, XminTmp, XminTmp1 As Single
        Dim OfstX, ScaleX, ScaleY, OfstY As Single
        Dim XValues(NumVIVACorePoints) As Object
        Dim XValues1(NumVIVACorePoints) As Object
        Dim XValues2(NumVIVACorePoints) As Object
        Dim ConversionFactor As Single

        If cmbMode.Text = "" Or cmbMode.Text = "MM" Then
            ModeNum = 0
        Else
            ModeNum = CShort(cmbMode.Text)
        End If

        If CurveUnits = VIVAMain.Units.English Then
            ConversionFactor = 1.0# / Ft2M
            If ModeNum = 0 Then
                XLab = "Mode Shapes (ft)"
            Else
                XLab = "Mode Shapes (ft), Phase (rad)"
            End If
        Else
            ConversionFactor = 1.0#
            If ModeNum = 0 Then
                XLab = "Mode Shapes (m)"
            Else
                XLab = "Mode Shapes (m), Phase (rad)"
            End If
        End If

        '   Use the wave magnitude to determine the axis maximum
        For i = 1 To NumPoints
            If NumOutputModes = 0 Then
                XValues(i) = 0.0#
                If ModeNum <> 0 Then
                    XValues1(i) = 0.0#
                    XValues2(i) = 0.0#
                End If
            Else
                'XValues(i) = WaveMag(i, ModeNum) * ConversionFactor
                If ModeNum <> 0 Then
                    ' XValues1(i) = WaveReal(i, ModeNum) * ConversionFactor
                    ' XValues2(i) = WaveImag(i, ModeNum) * ConversionFactor
                End If
            End If
        Next i

        Call MaxAndMin(XValues, Xmax, Xmin, NumPoints)

        '   single frequency mode
        If ModeNum <> 0 Then
            '       real and imagine part for min
            Call MaxAndMin(XValues1, XmaxTmp, XminTmp, NumPoints)
            Call MaxAndMin(XValues2, XmaxTmp1, XminTmp1, NumPoints)
            If XminTmp < Xmin Then Xmin = XminTmp
            If XminTmp1 < Xmin Then Xmin = XminTmp1

            '       phase angle for its maximum
            For i = 1 To NumPoints
                If NumOutputModes = 0 Then
                    XValues(i) = 0.0#
                Else
                    'XValues(i) = PhaseAngle(i, ModeNum)
                End If
            Next i
            Call MaxAndMin(XValues, XmaxTmp, XminTmp, NumPoints)
            If XmaxTmp > Xmax Then Xmax = XmaxTmp
        End If

        Dim scl As DrawScale
        DrawScaleF(Xmax, Xmin, Ymax, 0, Me.picResults, scl)
        scl.Left += (0 - Xmin) * scl.Width

        Dim wid, hgt As Integer
        wid = picResults.ClientSize.Width
        hgt = picResults.ClientSize.Height

        Dim bm As Bitmap = New Bitmap(wid, hgt)
        Dim gr As Graphics = Graphics.FromImage(bm)

        '   axes and legends
        If PrintPlot Then
            XLab = XLab & "    [Mode: " & cmbMode.Text & "]"
            PrintAxis(Xmax, Xmin, Ymax, 0, XLab, YLab, ScaleX, ScaleY, OfstX, OfstY)
            PrintModeShapesLegend(Xmax, Xmin, Ymax, Ymin, ScaleX, ScaleY, OfstX, OfstY)
        Else
            DrawAxis(gr, scl, Xmax, Xmin, Ymax, 0, XLab, YLab)
            PlotModeShapesLegend(gr, scl, Xmax, Xmin, Ymax, Ymin)
        End If

        '   Plot the four curves
        '   single frequency mode only
        If ModeNum <> 0 Then
            Call Plot2DLine(gr, scl, WaveReal, WaveRealColor, ModeNum, ScaleX, ScaleY, OfstX, OfstY, ConversionFactor)
            'Call Plot2DLine(WaveImag, WaveImagColor, ModeNum, ScaleX, ScaleY, OfstX, OfstY, ConversionFactor)
            'Call Plot2DLine(PhaseAngle, PhaseAngleColor, ModeNum, ScaleX, ScaleY, OfstX, OfstY, 1.0#)
        End If

        '   include multi-mode
        'Call Plot2DLine(WaveMag, WaveMagColor, ModeNum, ScaleX, ScaleY, OfstX, OfstY, ConversionFactor)

        picResults.Image = bm
        gr.Dispose()

    End Sub


    Private Sub Plot1DArray(ByRef A1D As Object, ByRef XLab As String)

        Dim Xmax, Xmin As Single
        Dim OfstX, ScaleX, ScaleY, OfstY As Single

        Call MaxAndMin(A1D, Xmax, Xmin, NumPoints)

        If Xmin > 0 Then Xmin = 0

        Dim scl As DrawScale
        DrawScaleF(Xmax, Xmin, Ymax, 0, Me.picResults, scl)

        Dim wid, hgt As Integer
        wid = picResults.ClientSize.Width
        hgt = picResults.ClientSize.Height

        Dim me_gr As Graphics = Me.CreateGraphics
        Dim me_hdc As IntPtr = me_gr.GetHdc

        Dim bounds As New RectangleF(0, 0, wid, hgt)
        Dim mf As New Imaging.Metafile(wmf_file_name, me_hdc, bounds, Imaging.MetafileFrameUnit.Pixel)
        me_gr.ReleaseHdc(me_hdc)

        Dim gr As Graphics = Graphics.FromImage(mf)
        gr.PageUnit = GraphicsUnit.Pixel
        gr.Clear(Color.White)

        If PrintPlot Then
            Call PrintAxis(Xmax, Xmin, Ymax, 0, XLab, YLab, ScaleX, ScaleY, OfstX, OfstY)
            Call PrintLine_Renamed(A1D, DispDist, NumPoints, CurveColor, ScaleX, ScaleY, OfstX, OfstY)
        Else
            DrawAxis(gr, scl, Xmax, Xmin, Ymax, 0, XLab, YLab)
            DrawLine(gr, scl, A1D, DispDist, NumPoints, CurveColor)
        End If

        Dim bm As Bitmap = New Bitmap(wid, hgt)
        gr.Dispose()
        gr = Graphics.FromImage(bm)
        Dim dest_bounds As New RectangleF(0, 0, wid, hgt)
        Dim source_bounds As New RectangleF(0, 0, wid + 1, hgt + 1)
        gr.DrawImage(mf, bounds, source_bounds, GraphicsUnit.Pixel)
        picResults.SizeMode = PictureBoxSizeMode.AutoSize
        picResults.Image = bm
        gr.Dispose()

        'ClipboardMetafileHelper.PutEnhMetafileOnClipboard(Me.Handle, mf)
        mf.Dispose()

    End Sub


    Private Sub Plot2DArrayColumn(ByRef A2D As Object, ByRef XLab As String, ByRef ConversionFactor As Single)

        Dim i As Short
        Dim XValues(NumVIVACorePoints) As Single
        Dim Xmax, Xmin As Single
        Dim OfstX, ScaleX, ScaleY, OfstY As Single
        'Dim strMode As String

        If cmbMode.Text = "" Or cmbMode.Text = "MM" Then
            ModeNum = 0
        Else
            ModeNum = CShort(cmbMode.Text)
        End If

        For i = 1 To NumPoints
            If NumOutputModes = 0 Then
                XValues(i) = 0.0#
            Else
                XValues(i) = A2D(i, ModeNum) * ConversionFactor
            End If
        Next i

        Call MaxAndMin(XValues, Xmax, Xmin, NumPoints)

        Dim scl As DrawScale
        DrawScaleF(Xmax, Xmin, Ymax, 0, Me.picResults, scl)

        Dim wid, hgt As Integer
        wid = picResults.ClientSize.Width
        hgt = picResults.ClientSize.Height

        Dim me_gr As Graphics = Me.CreateGraphics
        Dim me_hdc As IntPtr = me_gr.GetHdc

        Dim bounds As New RectangleF(0, 0, wid, hgt)
        Dim mf As New Imaging.Metafile(wmf_file_name, me_hdc, bounds, Imaging.MetafileFrameUnit.Pixel)
        me_gr.ReleaseHdc(me_hdc)

        Dim gr As Graphics = Graphics.FromImage(mf)
        gr.PageUnit = GraphicsUnit.Pixel
        gr.Clear(Color.White)

        If PrintPlot Then
            XLab = XLab & "    [Mode: " & cmbMode.Text & "]"
            Call PrintAxis(Xmax, Xmin, Ymax, 0, XLab, YLab, ScaleX, ScaleY, OfstX, OfstY)
            Call PrintLine_Renamed(XValues, DispDist, NumPoints, CurveColor, ScaleX, ScaleY, OfstX, OfstY)
        Else
            DrawAxis(gr, scl, Xmax, Xmin, Ymax, 0, XLab, YLab)
            Call DrawAxis(gr, scl, Xmax, Xmin, Ymax, 0, XLab, YLab)
            Call DrawLine(gr, scl, XValues, DispDist, NumPoints, CurveColor)
        End If

        gr.Dispose()

        Dim bm As Bitmap = New Bitmap(wid, hgt)
        gr = Graphics.FromImage(bm)
        Dim dest_bounds As New RectangleF(0, 0, wid, hgt)
        Dim source_bounds As New RectangleF(0, 0, wid + 1, hgt + 1)
        gr.DrawImage(mf, bounds, source_bounds, GraphicsUnit.Pixel)
        picResults.SizeMode = PictureBoxSizeMode.AutoSize
        picResults.Image = bm
        gr.Dispose()

        'ClipboardMetafileHelper.PutEnhMetafileOnClipboard(Me.Handle, mf)
        mf.Dispose()

    End Sub
	

    Private Sub Plot2DLine(ByVal gr As Graphics, ByVal s As DrawScale, ByRef A2D As Object, ByRef LineColor As Color, ByRef ModeNum As Short, _
        ByRef ScaleX As Single, ByRef ScaleY As Single, ByRef OfstX As Single, ByRef OfstY As Single, ByRef ConversionFactor As Single)

        Dim i As Short
        Dim XValues(NumVIVACorePoints) As Single

        For i = 1 To NumPoints
            If NumOutputModes = 0 Then
                XValues(i) = 0.0#
            Else
                XValues(i) = A2D(i, ModeNum) * ConversionFactor
            End If
        Next i

        If PrintPlot Then
            Select Case LineColor
                Case WaveRealColor
                    'UPGRADE_ISSUE: Constant vbDashDotDot was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
                    'UPGRADE_ISSUE: Printer property Printer.DrawStyle was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
                    'Printer.DrawStyle = vbDashDotDot
                Case WaveImagColor
                    'UPGRADE_ISSUE: Constant vbDot was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
                    'UPGRADE_ISSUE: Printer property Printer.DrawStyle was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
                    'Printer.DrawStyle = vbDot
                Case WaveMagColor
                    'UPGRADE_ISSUE: Constant vbSolid was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
                    'UPGRADE_ISSUE: Printer property Printer.DrawStyle was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
                    'Printer.DrawStyle = vbSolid
                Case PhaseAngleColor
                    'UPGRADE_ISSUE: Constant vbDash was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
                    'UPGRADE_ISSUE: Printer property Printer.DrawStyle was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
                    'Printer.DrawStyle = vbDash
            End Select

            Call PrintLine_Renamed(XValues, DispDist, NumPoints, LineColor, ScaleX, ScaleY, OfstX, OfstY)
            'UPGRADE_ISSUE: Constant vbSolid was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
            'UPGRADE_ISSUE: Printer property Printer.DrawStyle was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
            'Printer.DrawStyle = vbSolid
        Else
            Call DrawLine(gr, s, XValues, DispDist, NumPoints, LineColor)
        End If

    End Sub
	

    Private Sub PlotModeShapesLegend(ByVal gr As Graphics, ByVal s As DrawScale, ByRef Xmax As Single, ByRef Xmin As Single, _
        ByRef Ymax As Single, ByRef Ymin As Single)

        Dim BoxLength, LineLength As Single     'BoxHeight, 
        Dim LegendY, LegendX, LineY As Single
        Dim TextLength As Single                'PieceLength, SpaceLength, 
        Dim MagLab, RealLab, ImagLab, PALab As String

        gr.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit

        RealLab = "       Real "
        ImagLab = "  Imaginary "
        MagLab = "  Magnitude "
        PALab = "  Phase Angle "

        'scale
        Xmin *= s.Width
        Xmax *= s.Width
        Ymin *= s.Height
        Ymax *= s.Height

        Dim grPen As Pen = New Pen(Color.Black, 2)

        'location and size
        Dim strF As SizeF
        Dim f As Font = New Font("Airal", 8, FontStyle.Bold)
        strF = gr.MeasureString(YLab, f)

        BoxLength = Xmax - Xmin - strF.Width / 2

        strF = gr.MeasureString(RealLab & ImagLab & MagLab & PALab & " ", f)
        TextLength = strF.Width
        LineLength = (BoxLength - TextLength) / 4

        strF = gr.MeasureString(YLab, f)
        LegendX = Xmin + strF.Width / 2
        LegendY = Ymax + strF.Height * 1.1

        'Set the coordinates and start the legend
        Dim CurrentX As Single
        CurrentX = LegendX

        Dim CurrentY As Single
        CurrentY = LegendY

        gr.DrawString(RealLab, f, grPen.Brush, CurrentX, CurrentY)

        'Update the "cursor"
        strF = gr.MeasureString(RealLab, f)
        LegendX = LegendX + strF.Width
        LineY = LegendY + strF.Height / 2

        'Draw the line
        grPen.Color = WaveRealColor
        gr.DrawLine(grPen, LegendX + CInt(LineLength * 0.05), LineY, LegendX + LineLength, LineY)

        'Repeat for each of the other curves
        LegendX = LegendX + LineLength

        CurrentX = LegendX
        CurrentY = LegendY

        gr.DrawString(ImagLab, f, Brushes.Black, CurrentX, CurrentY)

        strF = gr.MeasureString(ImagLab, f)
        LegendX = LegendX + strF.Width

        grPen.Color = WaveImagColor
        gr.DrawLine(grPen, LegendX + CInt(LineLength * 0.05), LineY, LegendX + LineLength, LineY)

        'QBColor (WaveImagColor)

        LegendX = LegendX + LineLength

        CurrentX = LegendX
        CurrentY = LegendY

        gr.DrawString(MagLab, f, Brushes.Black, CurrentX, CurrentY)

        strF = gr.MeasureString(MagLab, f)
        LegendX = LegendX + strF.Width

        'UPGRADE_ISSUE: PictureBox method picResults.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        grPen.Color = WaveMagColor
        gr.DrawLine(grPen, LegendX + CInt(LineLength * 0.05), LineY, LegendX + LineLength, LineY)

        'QBColor (WaveMagColor)

        LegendX = LegendX + LineLength

        CurrentX = LegendX
        CurrentY = LegendY

        gr.DrawString(PALab, f, Brushes.Black, CurrentX, CurrentY)

        strF = gr.MeasureString(PALab, f)
        LegendX = LegendX + strF.Width

        grPen.Color = PhaseAngleColor
        gr.DrawLine(grPen, LegendX + CInt(LineLength * 0.05), LineY, LegendX + LineLength, LineY)

        'QBColor (PhaseAngleColor)

    End Sub
	

    Private Sub PrintModeShapesLegend(ByRef Xmax As Single, ByRef Xmin As Single, ByRef Ymax As Single, ByRef Ymin As Single, ByRef ScaleX As Single, ByRef ScaleY As Single, ByRef OfstX As Single, ByRef OfstY As Single)

        'Const CYLabW As Single = 1.05
        Dim BoxLength, LineLength As Single         'BoxHeight, 
        Dim LegendX As Single                       'LegendY, , LineY
        Dim TextLength As Single                    'PieceLength, SpaceLength, 
        Dim MagLab, RealLab, ImagLab, PALab As String

        RealLab = "       Real "
        ImagLab = "  Imaginary "
        MagLab = "  Magnitude "
        PALab = "  Phase Angle "

        '   location and size

        'UPGRADE_ISSUE: Printer method Printer.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'BoxLength = ScaleX * (Xmax - Xmin) - Printer.TextWidth(YLab) * CYLabW / 2

        'UPGRADE_ISSUE: Printer method Printer.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'TextLength = Printer.TextWidth(RealLab & ImagLab & MagLab & PALab & " ")

        LineLength = (BoxLength - TextLength) / 4

        'UPGRADE_ISSUE: Printer method Printer.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'LegendX = ScaleX * Xmin + OfstX + Printer.TextWidth(YLab) * CYLabW / 2

        'UPGRADE_ISSUE: Printer method Printer.TextHeight was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'LegendY = ScaleY * Ymax + OfstY + Printer.TextHeight(YLab) * 3.0#

        '   Set the coordinates and start the legend
        'UPGRADE_ISSUE: Printer property Printer.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.CurrentX = LegendX

        'UPGRADE_ISSUE: Printer property Printer.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.CurrentY = LegendY

        'UPGRADE_ISSUE: Printer object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        'UPGRADE_ISSUE: Printer method Printer.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.Print(RealLab)

        '   Update the "cursor"
        'UPGRADE_ISSUE: Printer method Printer.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'LegendX = LegendX + Printer.TextWidth(RealLab)

        'UPGRADE_ISSUE: Printer method Printer.TextHeight was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'LineY = LegendY + Printer.TextHeight(RealLab) / 2

        '   Draw the line
        'UPGRADE_ISSUE: Constant vbDashDotDot was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
        'UPGRADE_ISSUE: Printer property Printer.DrawStyle was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.DrawStyle = vbDashDotDot

        'UPGRADE_ISSUE: Printer method Printer.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.Line((LegendX, LineY) - (LegendX + LineLength, LineY), QBColor(WaveRealColor))

        '   Repeat for each of the other curves
        LegendX = LegendX + LineLength

        'UPGRADE_ISSUE: Printer property Printer.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.CurrentX = LegendX

        'UPGRADE_ISSUE: Printer property Printer.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.CurrentY = LegendY

        'UPGRADE_ISSUE: Printer object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        'UPGRADE_ISSUE: Printer method Printer.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.Print(ImagLab)

        'UPGRADE_ISSUE: Printer method Printer.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'LegendX = LegendX + Printer.TextWidth(ImagLab)

        'UPGRADE_ISSUE: Constant vbDot was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
        'UPGRADE_ISSUE: Printer property Printer.DrawStyle was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.DrawStyle = vbDot

        'UPGRADE_ISSUE: Printer method Printer.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.Line((LegendX, LineY) - (LegendX + LineLength, LineY), QBColor(WaveImagColor))

        LegendX = LegendX + LineLength

        'UPGRADE_ISSUE: Printer property Printer.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.CurrentX = LegendX

        'UPGRADE_ISSUE: Printer property Printer.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.CurrentY = LegendY

        'UPGRADE_ISSUE: Printer object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        'UPGRADE_ISSUE: Printer method Printer.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.Print(MagLab)

        'UPGRADE_ISSUE: Printer method Printer.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'LegendX = LegendX + Printer.TextWidth(MagLab)

        'UPGRADE_ISSUE: Constant vbSolid was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
        'UPGRADE_ISSUE: Printer property Printer.DrawStyle was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.DrawStyle = vbSolid

        'UPGRADE_ISSUE: Printer method Printer.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.Line((LegendX, LineY) - (LegendX + LineLength, LineY), QBColor(WaveMagColor))

        LegendX = LegendX + LineLength

        'UPGRADE_ISSUE: Printer property Printer.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.CurrentX = LegendX

        'UPGRADE_ISSUE: Printer property Printer.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.CurrentY = LegendY

        'UPGRADE_ISSUE: Printer object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        'UPGRADE_ISSUE: Printer method Printer.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.Print(PALab)

        'UPGRADE_ISSUE: Printer method Printer.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'LegendX = LegendX + Printer.TextWidth(PALab)

        'UPGRADE_ISSUE: Constant vbDash was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
        'UPGRADE_ISSUE: Printer property Printer.DrawStyle was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.DrawStyle = vbDash

        'UPGRADE_ISSUE: Printer method Printer.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.Line((LegendX, LineY) - (LegendX + LineLength, LineY), QBColor(PhaseAngleColor))

        'UPGRADE_ISSUE: Constant vbSolid was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
        'UPGRADE_ISSUE: Printer property Printer.DrawStyle was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.DrawStyle = vbSolid

    End Sub


    Private Sub cmbMode_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbMode.SelectedIndexChanged

        PlotCurve()

    End Sub


	Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click
		
		Me.Close()
		
	End Sub
	

    Private Sub btnPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnPrint.Click

        PrintDocResults = New Printing.PrintDocument
        dlgPrint.Document = PrintDocResults
        If dlgPrint.ShowDialog = Windows.Forms.DialogResult.OK Then
            dlgPrint.Document.Print()
        End If

    End Sub
	

    Public Sub mnuExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuExit.Click

        Me.Close()

    End Sub
	

    Public Sub mnuPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPlotPrint.Click

        btnPrint_Click(btnPrint, New System.EventArgs())

    End Sub
	

    Private Sub tabPlot_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tabPlot.ClickEvent

        Dim CurIndex As Short

        PlotOption = tabPlot.SelectedItem.Index - 1
        CurIndex = cmbMode.SelectedIndex

        If PlotOption < 2 Then
            lblMode.Enabled = False
            cmbMode.Enabled = False
        Else
            lblMode.Enabled = True
            cmbMode.Enabled = True
            '        If PlotOption = 3 Or PlotOption = 6 Then
            '            If cmbMode.ListCount > NumOutputModes And NumOutputModes > 0 Then
            '                cmbMode.RemoveItem (NumOutputModes)
            '                If CurIndex = NumOutputModes Then cmbMode.ListIndex = 0
            '            End If
            '        Else
            '            If cmbMode.ListCount = NumOutputModes And NumOutputModes > 0 Then _
            ''                cmbMode.AddItem "MM", NumOutputModes
            '        End If
        End If

        PlotCurve()

    End Sub

    Private Sub mnuPlotCopy_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuPlotCopy.Click

        Dim mf As Imaging.Metafile = New Imaging.Metafile(wmf_file_name)
        ClipboardMetafileHelper.PutEnhMetafileOnClipboard(Me.Handle, mf)
        mf.Dispose()

    End Sub

    Private Sub PrintDocResults_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) _
        Handles PrintDocResults.PrintPage

        Dim mf As Imaging.Metafile = New Imaging.Metafile(wmf_file_name)
        e.Graphics.DrawImage(mf, e.MarginBounds)
        e.HasMorePages = False

    End Sub
End Class