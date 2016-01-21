Option Strict Off
Option Explicit On

Module modPrintGraphics
	
	Const MarginPercent As Single = 0.2
	Const TitlePercent As Single = 0.08
	Const LtrWidthIn As Single = 8.5
	Const LtrHeightIn As Single = 11#
	

    Public Sub PrintAxis(ByVal Xmax As Single, ByVal Xmin As Single, ByVal Ymax As Single, ByVal Ymin As Single, ByVal XLabel As String, _
    ByVal YLabel As String, ByRef ScaleX As Single, ByRef ScaleY As Single, ByRef OfstX As Single, ByRef OfstY As Single)

        Dim PYmax, PXmax, PXmin, PYmin As Single
        Dim PDeltaX, DeltaX, DeltaY, PDeltaY As Single
        Dim LabelX, LabelY As Single                'MarginX, , MarginY 
        Dim NumTix As Short
        Dim TicInt, TicLoc As Single
        Dim TicLab As String
        'Dim TicW, TicY As Single
        'Dim TicYP, TicWP, TicLocP As Single
        Dim i As Short
        Dim JobTitle As String
        Dim blnTicAt0 As Boolean

        '   Coordinate system is the original:  Upper-Left-Hand corner is (0,0);
        '   Positive X is to the right, Positive Y is down

        '   Set the font
        JobTitle = CurProj.Title

        'UPGRADE_ISSUE: Printer property Printer.FontName was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.FontName = "Arial"

        'UPGRADE_ISSUE: Printer property Printer.FontBold was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.FontBold = True

        '   Set the line weight

        'UPGRADE_ISSUE: Printer property Printer.DrawWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.DrawWidth = 4

        '   Establish the scale

        'UPGRADE_ISSUE: Printer property Printer.ScaleMode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.ScaleMode = 5

        DeltaX = Xmax - Xmin
        If DeltaX = 0 Then DeltaX = 0.01
        DeltaY = Ymax - Ymin
        If DeltaY = 0 Then DeltaY = 0.01
        ScaleX = LtrWidthIn * (1 - 2 * MarginPercent) / DeltaX
        ScaleY = LtrHeightIn * (1 - 2 * MarginPercent) / DeltaY
        OfstX = LtrWidthIn * MarginPercent - ScaleX * Xmin
        OfstY = LtrHeightIn * MarginPercent - ScaleY * Ymin

        '   Scale the variables we'll use again and again
        PXmax = ScaleX * Xmax + OfstX
        PXmin = ScaleX * Xmin + OfstX
        PYmax = ScaleY * Ymax + OfstY
        PYmin = ScaleY * Ymin + OfstY
        PDeltaX = PXmax - PXmin
        PDeltaY = PYmax - PYmin


        '   Draw the axes

        'UPGRADE_ISSUE: Printer method Printer.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.Line((PXmin, PYmin) - (PXmax, PYmin))

        'UPGRADE_ISSUE: Printer method Printer.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.Line((PXmin, PYmin) - (PXmin, PYmax))

        '   Job Title
        If JobTitle <> "" Then

            'UPGRADE_ISSUE: Printer property Printer.FontSize was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
            'Printer.FontSize = 10

            LabelX = PXmin + PDeltaX / 2
            LabelY = PYmin - LtrHeightIn * TitlePercent

            'UPGRADE_ISSUE: Printer method Printer.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
            'UPGRADE_ISSUE: Printer property Printer.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
            'Printer.CurrentX = LabelX - Printer.TextWidth(JobTitle) / 2

            'UPGRADE_ISSUE: Printer method Printer.TextHeight was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
            'UPGRADE_ISSUE: Printer property Printer.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
            'Printer.CurrentY = LabelY - Printer.TextHeight(JobTitle) * 3.0#

            'UPGRADE_ISSUE: Printer object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
            'UPGRADE_ISSUE: Printer method Printer.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
            'Printer.Print(JobTitle)

        End If

        '   Label them;
        '   X-label goes above the axis

        'UPGRADE_ISSUE: Printer property Printer.FontSize was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.FontSize = 8

        LabelX = PXmin + PDeltaX / 2
        LabelY = PYmin

        'UPGRADE_ISSUE: Printer method Printer.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'UPGRADE_ISSUE: Printer property Printer.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.CurrentX = LabelX - Printer.TextWidth(XLabel) / 2

        'UPGRADE_ISSUE: Printer method Printer.TextHeight was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'UPGRADE_ISSUE: Printer property Printer.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.CurrentY = LabelY - Printer.TextHeight(XLabel) * 4.0#

        'UPGRADE_ISSUE: Printer object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        'UPGRADE_ISSUE: Printer method Printer.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.Print(XLabel)

        '   Y-label goes at the bottom of the y-axis
        LabelX = PXmin
        LabelY = PYmax

        'UPGRADE_ISSUE: Printer method Printer.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'UPGRADE_ISSUE: Printer property Printer.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.CurrentX = LabelX - Printer.TextWidth(YLabel) / 2

        'UPGRADE_ISSUE: Printer method Printer.TextHeight was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'UPGRADE_ISSUE: Printer property Printer.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.CurrentY = LabelY + Printer.TextHeight(YLabel) * 2.0#

        'UPGRADE_ISSUE: Printer object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        'UPGRADE_ISSUE: Printer method Printer.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.Print(YLabel)

        '   reset parameters

        'UPGRADE_ISSUE: Printer property Printer.DrawWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.DrawWidth = 1

        'UPGRADE_ISSUE: Printer property Printer.FontBold was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'Printer.FontBold = False

        '   Establish tick marks and label them; X-axis first

        'UPGRADE_ISSUE: Printer method Printer.TextHeight was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'TicY = PYmin - Printer.TextHeight("0.123456789") * 1.1

        NumTix = Divisions(Xmin, Xmax)
        TicInt = DeltaX / NumTix
        blnTicAt0 = False

        For i = 0 To NumTix
            TicLoc = i * TicInt + Xmin
            If System.Math.Abs(TicLoc) * 50 < DeltaX Then TicLoc = 0
            If TicLoc = 0 Then blnTicAt0 = True
            If TicLoc > Xmax Then TicLoc = Xmax

            'UPGRADE_ISSUE: Printer method Printer.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
            'Printer.Line((ScaleX * TicLoc + OfstX, PYmin) - (ScaleX * TicLoc + OfstX, PYmax))

            If TicLoc <> 0 And (System.Math.Abs(TicLoc) > 1000.0# Or System.Math.Abs(TicLoc) < 0.01) Then
                TicLab = Format(TicLoc, "Scientific")
            Else
                TicLab = Format(TicLoc, "#0.00")
            End If

            'UPGRADE_ISSUE: Printer method Printer.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
            'UPGRADE_ISSUE: Printer property Printer.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
            'Printer.CurrentX = ScaleX * TicLoc + OfstX - Printer.TextWidth(TicLab) / 2

            'UPGRADE_ISSUE: Printer property Printer.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
            'Printer.CurrentY = TicY

            'UPGRADE_ISSUE: Printer object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
            'UPGRADE_ISSUE: Printer method Printer.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
            'Printer.Print(TicLab)

        Next i

        If Not blnTicAt0 Then
            TicLoc = 0

            'UPGRADE_ISSUE: Printer method Printer.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
            'Printer.Line((ScaleX * TicLoc + OfstX, PYmin) - (ScaleX * TicLoc + OfstX, PYmax))

            TicLab = "0.00"

            'UPGRADE_ISSUE: Printer method Printer.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
            'UPGRADE_ISSUE: Printer property Printer.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
            'Printer.CurrentX = ScaleX * TicLoc + OfstX - Printer.TextWidth(TicLab) / 2

            'UPGRADE_ISSUE: Printer property Printer.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
            'Printer.CurrentY = TicY

            'UPGRADE_ISSUE: Printer object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
            'UPGRADE_ISSUE: Printer method Printer.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
            'Printer.Print(TicLab)

        End If

        '   Y-axis

        'UPGRADE_ISSUE: Printer method Printer.TextHeight was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'TicYP = Printer.TextHeight("0.123456789") / 2

        NumTix = Divisions(Ymin, Ymax)
        TicInt = DeltaY / NumTix

        For i = 0 To NumTix
            TicLoc = i * TicInt + Ymin

            'UPGRADE_ISSUE: Printer method Printer.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
            'Printer.Line((PXmin, ScaleY * TicLoc + OfstY) - (PXmax, ScaleY * TicLoc + OfstY))

            TicLab = Format(TicLoc, "####")

            'UPGRADE_ISSUE: Printer method Printer.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
            'UPGRADE_ISSUE: Printer property Printer.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
            'Printer.CurrentX = PXmin - Printer.TextWidth(TicLab) * 1.1

            'UPGRADE_ISSUE: Printer property Printer.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
            'Printer.CurrentY = ScaleY * TicLoc + OfstY - TicYP

            'UPGRADE_ISSUE: Printer object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
            'UPGRADE_ISSUE: Printer method Printer.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
            'Printer.Print(TicLab)
        Next i

    End Sub


	'UPGRADE_NOTE: PrintLine was upgraded to PrintLine_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Public Sub PrintLine_Renamed(ByRef X As Object, ByRef Y As Object, ByRef NumPoints As Object, ByRef LineColor As Color, _
    ByRef ScaleX As Single, ByRef ScaleY As Single, ByRef OfstX As Single, ByRef OfstY As Single)

        Dim i As Short

        If NumPoints = 1 Then
            '       If there's only one point, plot it

            'UPGRADE_ISSUE: Printer method Printer.PSet was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
            'Printer.PSet(ScaleX * X(1) + OfstX, ScaleY * Y(1) + OfstY)

        Else
            '   More than one point is a line

            'UPGRADE_ISSUE: Printer property Printer.DrawWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
            'Printer.DrawWidth = 6

            For i = 1 To NumPoints - 1

                'UPGRADE_ISSUE: Printer method Printer.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
                'Printer.Line((ScaleX * X(i) + OfstX, ScaleY * Y(i) + OfstY) - (ScaleX * X(i + 1) + OfstX, ScaleY * Y(i + 1) + OfstY), LineColor)

                ' QBColor (LineColor)
            Next i
        End If
    End Sub
End Module