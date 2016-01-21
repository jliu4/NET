Option Strict Off
Option Explicit On

Module modPlot
	
    Const MarginPercent As Single = 0.1
    Const Fuzz As Single = 0.1

    Public Structure DrawScale
        Dim Left, Top, Width, Height As Single
    End Structure


    Public Function Divisions(ByVal min As Single, ByVal max As Single) As Short
        ' function called by FormatPlot to space out grid lines
        ' and set max and min on axes

        Dim Num As Short
        Dim divisor, factor, inc, multiplier As Single

        If max = min Then
            inc = 1.0#
            If Not (max = 0.0#) Then
                divisor = 2.0#
                Do While max / inc < 1.0#
                    inc = inc / divisor
                    If divisor = 5.0# Then
                        divisor = 2.0#
                    Else
                        divisor = 5.0#
                    End If
                Loop
            End If
            max = max + inc
            min = min - inc
            Divisions = 2
        Else
            inc = max - min
            factor = 1.0#
            divisor = 2
            Do While inc > 10.0#
                inc = inc / divisor
                factor = factor * divisor
                If divisor = 5.0# Then
                    divisor = 2.0#
                Else
                    divisor = 5.0#
                End If
            Loop
            multiplier = 2.0#
            Do While inc < 2.0#
                inc = inc * multiplier
                factor = factor / multiplier
                If multiplier = 5.0# Then
                    multiplier = 2.0#
                Else
                    multiplier = 5.0#
                End If
            Loop

            Num = CShort(min / factor)
            If Num * factor > min Then
                Num = Num - 1
            End If
            min = Num * factor
            Num = CShort(max / factor)
            If Num * factor < max Then
                Num = Num + 1
            End If
            max = Num * factor
            Divisions = (max - min) / factor
        End If

    End Function

	
   


    Public Sub DrawAxis(ByVal Graph As Graphics, ByVal sc As DrawScale, ByVal Xmax As Single, ByVal Xmin As Single, ByVal Ymax As Single, _
        ByVal Ymin As Single, ByVal XLabel As String, ByVal YLabel As String)

        Dim DeltaX, DeltaY As Single
        Dim LabelX, LabelY As Single
        Dim NumTix As Short
        Dim TicInt, TicLoc As Single
        Dim TicLab As String
        Dim TicY As Single                      'TicW, 
        Dim i As Short
        Dim blnTicAt0 As Boolean

        Dim grPen As Pen = New Pen(Color.Black, 2)

        '   Coordinate system is the original:  Upper-Left-Hand corner is (0,0);
        '   Positive X is to the right, Positive Y is down

        '   Establish the scale
        Xmax *= sc.Width
        Xmin *= sc.Width
        Ymax *= sc.Height
        Ymin *= sc.Height
        DeltaX = Xmax - Xmin
        If DeltaX = 0 Then DeltaX = 0.01
        DeltaY = Ymax - Ymin
        If DeltaY = 0 Then DeltaY = 0.01

        '   Establish the origin
        Graph.ResetTransform()
        Graph.TranslateTransform(sc.Left, sc.Top)
        Graph.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit

        '   Draw the axes
        Graph.DrawLine(grPen, Xmin, Ymin, Xmax, Ymin)
        Graph.DrawLine(grPen, Xmin, Ymin, Xmin, Ymax)

        '   Label them
        '   X-label goes above the axis
        LabelX = Xmin + DeltaX / 2
        LabelY = Ymin

        Dim CurrentX As Single
        Dim strF As SizeF
        Dim f As New System.Drawing.Font("Arial", 8, FontStyle.Bold)

        strF = Graph.MeasureString(XLabel, f)
        CurrentX = LabelX - strF.Width / 2

        Dim CurrentY As Single
        CurrentY = LabelY - strF.Height * 2.4

        Graph.DrawString(XLabel, f, grPen.Brush, CurrentX, CurrentY)

        '   Y-label goes at the bottom of the y-axis
        LabelX = Xmin
        LabelY = Ymax

        strF = Graph.MeasureString(YLabel, f)
        CurrentX = LabelX - strF.Width / 2
        CurrentY = LabelY + strF.Height * 1.1

        Graph.DrawString(YLabel, f, grPen.Brush, CurrentX, CurrentY)

        '   reset parameters
        grPen.Width = 1
        grPen.DashStyle = Drawing2D.DashStyle.Dash
        f = New Font(f, FontStyle.Regular)

        '   Establish tick marks and label them; X-axis first
        strF = Graph.MeasureString("0.123456789", f)
        TicY = Ymin - strF.Height

        NumTix = Divisions(Xmin, Xmax)
        TicInt = DeltaX / NumTix

        '   Close "the box"
        blnTicAt0 = False

        For i = 0 To NumTix - 1
            TicLoc = i * TicInt + Xmin
            If System.Math.Abs(TicLoc) * 100 < DeltaX Then TicLoc = 0
            If TicLoc = 0 Then blnTicAt0 = True
            If TicLoc > Xmax Then TicLoc = Xmax

            Graph.DrawLine(grPen, TicLoc, Ymin, TicLoc, Ymax)

            If TicLoc <> 0 And (System.Math.Abs(TicLoc / sc.Width) > 1000.0# Or System.Math.Abs(TicLoc / sc.Width) < 0.01) Then
                TicLab = Format(TicLoc / sc.Width, "Scientific")
            Else
                TicLab = Format(TicLoc / sc.Width, "##0.00")
            End If

            strF = Graph.MeasureString(TicLab, f)
            CurrentX = TicLoc - strF.Width / 2
            CurrentY = TicY

            Graph.DrawString(TicLab, f, grPen.Brush, CurrentX, CurrentY)
        Next i

        'wider solid pen
        grPen.DashStyle = Drawing2D.DashStyle.Solid
        grPen.Width = 2

        For i = NumTix To NumTix
            TicLoc = i * TicInt + Xmin
            If System.Math.Abs(TicLoc) * 100 < DeltaX Then TicLoc = 0
            If TicLoc = 0 Then blnTicAt0 = True
            If TicLoc > Xmax Then TicLoc = Xmax

            Graph.DrawLine(grPen, TicLoc, Ymin, TicLoc, Ymax)

            If TicLoc <> 0 And (System.Math.Abs(TicLoc / sc.Width) > 1000.0# Or System.Math.Abs(TicLoc / sc.Width) < 0.01) Then
                TicLab = Format(TicLoc / sc.Width, "Scientific")
            Else
                TicLab = Format(TicLoc / sc.Width, "##0.00")
            End If

            strF = Graph.MeasureString(TicLab, f)
            CurrentX = TicLoc - strF.Width / 2
            CurrentY = TicY

            Graph.DrawString(TicLab, f, grPen.Brush, CurrentX, CurrentY)
        Next i

        If Not blnTicAt0 Then
            TicLoc = 0

            Graph.DrawLine(grPen, TicLoc, Ymin, TicLoc, Ymax)

            TicLab = "0.00"
            strF = Graph.MeasureString(TicLab, f)
            CurrentX = TicLoc - strF.Width / 2
            CurrentY = TicY
            Graph.DrawString(TicLab, f, grPen.Brush, CurrentX, CurrentY)
        End If

        '   Y-axis
        strF = Graph.MeasureString("0.123456789", f)
        TicY = strF.Height / 2

        NumTix = Divisions(Ymin, Ymax)
        TicInt = DeltaY / NumTix

        'dash line
        grPen.Width = 1
        grPen.DashStyle = Drawing2D.DashStyle.Dash
        For i = 0 To NumTix - 1
            TicLoc = i * TicInt + Ymin
            If TicLoc > Ymax Then TicLoc = Ymax

            Graph.DrawLine(grPen, Xmin, TicLoc, Xmax, TicLoc)

            TicLab = Format(TicLoc / sc.Height, "####")

            strF = Graph.MeasureString(TicLab, f)
            CurrentX = Xmin - strF.Width * 1.1
            CurrentY = TicLoc - TicY

            Graph.DrawString(TicLab, f, grPen.Brush, CurrentX, CurrentY)
        Next i

        'solid line
        grPen.Width = 2
        grPen.DashStyle = Drawing2D.DashStyle.Solid
        For i = NumTix To NumTix
            TicLoc = i * TicInt + Ymin
            If TicLoc > Ymax Then TicLoc = Ymax

            Graph.DrawLine(grPen, Xmin, TicLoc, Xmax, TicLoc)

            TicLab = Format(TicLoc / sc.Height, "####")

            strF = Graph.MeasureString(TicLab, f)
            CurrentX = Xmin - strF.Width * 1.1
            CurrentY = TicLoc - TicY

            Graph.DrawString(TicLab, f, grPen.Brush, CurrentX, CurrentY)
        Next i

    End Sub


    Public Sub DrawAxisNew(ByVal Graph As Graphics, ByVal sc As DrawScale, ByVal Xmax As Single, ByVal Xmin As Single, ByVal Ymax As Single, _
        ByVal Ymin As Single, ByVal XLabel As String, ByVal YLabel As String, ByVal XTicFormat As String, ByVal YTicFormat As String, _
        ByVal ReverseY As Boolean, ByVal XLogScale As Boolean, ByVal NumLegend As Short, ByVal ParamArray Legend() As Object)

        Dim DeltaX, DeltaY As Single
        Dim LabelX, LabelY As Single
        Dim NumTix As Short
        Dim TicInt, TicLoc As Single
        Dim TicLab As String
        Dim TicY As Single                  'TicW, 

        Dim i As Short
        Dim TL As Single
        Dim blnTicAt0 As Boolean

        Dim BoxLength, LineLength As Single         'BoxHeight, 
        Dim LegendY, LegendX, LineY As Single
        Dim TextLength As Single                    'PieceLength, SpaceLength, 

        With Graph

            'Coordinate system is the original:  Upper-Left-Hand corner is (0,0);
            'Positive X is to the right, Positive Y is down
            'Initialize (in case this is a replot)

            Dim picPen As New Pen(Brushes.Black, 2)

            'picPen.Width = 2.0
            'picPen.LineJoin = Drawing2D.LineJoin.Bevel

            'Scale the axes
            Xmin *= sc.Width
            Xmax *= sc.Width
            Ymin *= sc.Height
            Ymax *= sc.Height
            DeltaX = Xmax - Xmin
            DeltaY = Ymax - Ymin

            'Draw the axes
            'Dim Graphic As Graphics = picGraph.CreateGraphics()

            .TranslateTransform(sc.Left, sc.Top)

            .DrawLine(picPen, Xmin, Ymin, Xmax, Ymin)
            .DrawLine(picPen, Xmin, Ymin, Xmin, Ymax)

            'Label them
            'X-label goes above the axis
            LabelX = Xmin + DeltaX / 2.0
            LabelY = Ymin

            Dim CurrentX As Single
            Dim strF As SizeF
            Dim f As New System.Drawing.Font("Arial", 8)

            strF = .MeasureString(XLabel, f)
            CurrentX = LabelX - strF.Width / 2
            f = New System.Drawing.Font(f, FontStyle.Bold)

            Dim CurrentY As Single
            CurrentY = LabelY - strF.Height * 2.4

            .TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
            .DrawString(XLabel, f, picPen.Brush, CurrentX, CurrentY)

            'Y-label goes at the bottom of the y-axis
            LabelX = Xmin
            LabelY = Ymax

            strF = .MeasureString(LabelY, f)

            CurrentX = LabelX - strF.Width / 1.2
            CurrentY = LabelY + strF.Height * 1.1

            .DrawString(YLabel, f, picPen.Brush, CurrentX, CurrentY)

            'legend
            If NumLegend > 0 Then

                strF = .MeasureString(YLabel, f)
                LegendX = LabelX + strF.Width / 2.0#
                LegendY = LabelY + strF.Height * 1.1

                TextLength = 0
                f = New System.Drawing.Font(f, FontStyle.Regular)

                For i = 1 To NumLegend Step 1
                    .MeasureString(Legend((i - 1) * 2), f)
                    TL = strF.Width

                    If TL > TextLength Then TextLength = TL
                Next i

                'location and size
                BoxLength = Xmax - LegendX
                LineLength = (BoxLength / NumLegend - TextLength) / 1.2
                LegendX = LegendX + TextLength * 0.2

                strF = .MeasureString(Legend(0), f)
                LineY = LegendY + strF.Height / 2

                For i = 1 To NumLegend Step 1
                    'Draw the line

                    picPen.Color = Legend((i - 1) * 2 + 1)
                    .DrawLine(picPen, LegendX, LineY, LegendX + LineLength, LineY)

                    'Update the "cursor"
                    LegendX = LegendX + LineLength * 1.2

                    CurrentX = LegendX
                    CurrentY = LegendY

                    'print the label
                    .DrawString(Legend((i - 1) * 2), f, picPen.Brush, CurrentX, CurrentY)

                    'Update the "cursor"
                    LegendX = LegendX + TextLength * 1.2
                Next i
            End If

            'reset parameters
            picPen.Width = 1
            picPen.DashStyle = Drawing2D.DashStyle.Dash
            picPen.Color = Color.Black

            'Establish tick marks and label them; X-axis first
            strF = .MeasureString("0.123456789", f)
            TicY = Ymin - strF.Height

            NumTix = Divisions(Xmin, Xmax)
            TicInt = DeltaX / NumTix

            'Close "the box"
            blnTicAt0 = False

            For i = 0 To NumTix
                If i = NumTix Then
                    picPen.Width = 2
                    picPen.DashStyle = Drawing2D.DashStyle.Solid
                End If

                TicLoc = i * TicInt + Xmin
                If System.Math.Abs(TicLoc) * 100 < DeltaX Then TicLoc = 0
                If TicLoc = 0 Then blnTicAt0 = True
                If TicLoc > Xmax Then TicLoc = Xmax

                .DrawLine(picPen, TicLoc, Ymin, TicLoc, Ymax)

                If XLogScale Then
                    TicLab = Format(10 ^ (TicLoc / sc.Width), XTicFormat)
                Else
                    TicLab = Format(TicLoc / sc.Width, XTicFormat)
                End If

                strF = .MeasureString(TicLab, f)
                CurrentX = TicLoc - strF.Width / 2
                CurrentY = TicY

                .DrawString(TicLab, f, picPen.Brush, CurrentX, CurrentY)

            Next i

            If Not blnTicAt0 Then
                TicLoc = 0

                .DrawLine(picPen, TicLoc, Ymin, TicLoc, Ymax)

                TicLab = Format(0.0#, "0")

                strF = .MeasureString(TicLab, f)
                CurrentX = TicLoc - strF.Width / 2
                CurrentY = TicY

                'picGraph.Print(TicLab)
                .DrawString(TicLab, f, picPen.Brush, CurrentX, CurrentY)

            End If

            'Y-axis

            .MeasureString("0.123456789", f)
            TicY = strF.Height / 2

            NumTix = Divisions(Ymin, Ymax)
            TicInt = DeltaY / NumTix

            picPen.Width = 1
            picPen.DashStyle = Drawing2D.DashStyle.Dash

            For i = 0 To NumTix
                If i = NumTix Then
                    picPen.Width = 2
                    picPen.DashStyle = Drawing2D.DashStyle.Solid
                End If

                TicLoc = i * TicInt + Ymin
                If TicLoc > Ymax Then TicLoc = Ymax

                .DrawLine(picPen, Xmin, TicLoc, Xmax, TicLoc)

                If ReverseY Then
                    TicLab = Format((Ymax - TicLoc + Ymin) / sc.Height, YTicFormat)
                Else
                    TicLab = Format(TicLoc / sc.Height, YTicFormat)
                End If

                strF = .MeasureString(TicLab, f)
                CurrentX = Xmin - strF.Width * 1.1
                CurrentY = TicLoc - TicY

                .DrawString(TicLab, f, picPen.Brush, CurrentX, CurrentY)

            Next i

            picPen.Dispose()

        End With

    End Sub


    Public Sub MaxAndMin(ByRef X As Object, ByRef Xmax As Object, ByRef Xmin As Object, ByRef NumPoints As Short)

        'Dim LastPoint, Base As Short
        Dim i As Short

        '   Determine maxima and minima
        Xmax = -3.402823E+38
        Xmin = 3.402823E+38

        If NumPoints = 1 Then
            '       one point, bracket around it a bit to make a plot have some width
            Xmax = X(1) * (1 + Fuzz)
            Xmin = X(1) * (1 - Fuzz)

            If Xmax < 0.0# Then Xmax = 0.0#
            If Xmin > 0.0# Then Xmin = 0.0#
        Else
            '       If there are several points in the arrays (usual case)
            For i = 1 To NumPoints
                If X(i) > Xmax Then Xmax = X(i)
                If X(i) < Xmin Then Xmin = X(i)
            Next i

            '       If the minimum and maximum are the same, bracket as in the one-point case
            If Xmax <= 0.0# Then
                Xmax = 0.0#
            ElseIf Xmax < 0.1 Then
                Xmax = 0.1
            ElseIf Xmax < 0.5 Then
                Xmax = CInt(Xmax * 10.0# + 0.5) / 10.0#
            Else
                Xmax = CInt(Xmax + 0.5)
            End If

            If Xmin >= 0.0# Then
                Xmin = 0.0#
            ElseIf Xmin > -0.1 Then
                Xmin = -0.1
            ElseIf Xmin > -0.5 Then
                Xmin = CInt(Xmin * 10.0# - 0.5) / 10.0#
            Else
                Xmin = CInt(Xmin - 0.5)
            End If
            If Xmax = Xmin Then
                Xmax = Xmax * (1 + Fuzz)
                Xmin = Xmin * (1 - Fuzz)
                If Xmax = 0.0# Then
                    Xmax = 1.0#
                    Xmin = -1.0#
                End If
            End If
        End If

    End Sub


    Public Sub DrawLine(ByVal gr As Graphics, ByVal s As DrawScale, ByVal X As Object, ByVal Y As Object, ByVal NumPoints As Short, _
        ByVal LineColor As Color)

        Dim i As Short
        'Dim DeltaX, DeltaY As Single

        ''Establish the scale
        'DeltaX = Xmax - Xmin
        'If DeltaX = 0 Then DeltaX = 0.01
        'DeltaY = Ymax - Ymin
        'If DeltaY = 0 Then DeltaY = 0.01

        'Dim ScaleWidth As Single
        'ScaleWidth = (1 - 2 * MarginPercent) * picGraph.Width / DeltaX

        'Dim ScaleHeight As Single
        'ScaleHeight = (1 - 2 * MarginPercent) * picGraph.Height / DeltaY

        'Dim ScaleLeft As Single
        'ScaleLeft = picGraph.Width * MarginPercent

        'Dim ScaleTop As Single
        'ScaleTop = picGraph.Height * MarginPercent

        'Dim Graphic As Graphics = picGraph.CreateGraphics()

        'shift origin
        gr.ResetTransform()
        gr.TranslateTransform(s.Left, s.Top)
        gr.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias

        Dim picPen As New Pen(LineColor, 2)

        'picPen.Width = 2.0
        'picPen.Color = LineColor

        'test
        'Xmax *= ScaleWidth
        'Ymax *= ScaleHeight
        'Graphic.DrawLine(picPen, Xmin, Ymin, Xmax, Ymax)

        Dim x1, y1, x2, y2 As Single

        'scale all points
        'For i = 1 To NumPoints
        '    X(i) *= ScaleWidth
        '    Y(i) *= ScaleHeight
        'Next

        If NumPoints = 1 Then
            'If there's only one point, plot it
            'UPGRADE_ISSUE: PictureBox method picGraph.PSet was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            x1 = X(1) * s.Width
            y1 = Y(1) * s.Height
            Dim ptRct As New Rectangle(x1 - 1, y1 - 1, 2, 2)
            gr.DrawEllipse(picPen, ptRct)
            'd(X(1), Y(1))

        Else
            'More than one point is a line

            For i = 1 To NumPoints - 1

                x1 = X(i) * s.Width
                y1 = Y(i) * s.Height
                x2 = X(i + 1) * s.Width
                y2 = Y(i + 1) * s.Height
                gr.DrawLine(picPen, x1, y1, x2, y2)

            Next i

            picPen.Width = 1

        End If

        picPen.Dispose()
        X = Nothing
        Y = Nothing

    End Sub

    Public Sub DrawScaleF(ByVal Xmax As Single, ByVal Xmin As Single, ByVal Ymax As Single, ByVal Ymin As Single, _
        ByVal picGraph As PictureBox, ByRef s As DrawScale)

        Dim DeltaX, DeltaY As Single

        With picGraph

            'Coordinate system is the original:  Upper-Left-Hand corner is (0,0);
            'Positive X is to the right, Positive Y is down
            'Initialize (in case this is a replot)

            'Establish the scale
            DeltaX = Xmax - Xmin
            If DeltaX = 0 Then DeltaX = 0.01
            DeltaY = Ymax - Ymin
            If DeltaY = 0 Then DeltaY = 0.01

            s.Width = (1 - 2 * MarginPercent) * picGraph.ClientSize.Width / DeltaX
            s.Height = (1 - 2 * MarginPercent) * picGraph.ClientSize.Height / DeltaY

            'Establish the origin
            s.Left = picGraph.ClientSize.Width * MarginPercent
            s.Top = picGraph.ClientSize.Height * MarginPercent

        End With

    End Sub

End Module