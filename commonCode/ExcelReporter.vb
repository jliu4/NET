Option Strict Off
Option Explicit On
Friend Class ExcelReporter
	Dim oxApp As Microsoft.Office.Interop.Excel.Application
	Dim oxBooks As Microsoft.Office.Interop.Excel.Workbooks
	Dim oxBook As Microsoft.Office.Interop.Excel.Workbook
	Dim oxTmpBook As Microsoft.Office.Interop.Excel.Workbook

    Public Function GetGroundLength(ByVal CurLine As MoorLine, Optional ByRef TopTension As Single = 0#, Optional ByRef HorForce As Single = 0#) As Single
        GetGroundLength = -1

        Dim i, NumSeg As Short

        Dim Scope, FrcHor As Single

        Dim Length, LenStr As Single

        Dim SegLength(MaxNumSeg) As Single
        Dim SegTension(MaxNumSeg) As Single
        Dim SegAngle(MaxNumSeg) As Single
        Dim SegPosition(MaxNumSeg) As Single
        Dim CatX(MaxNumSubSeg * MaxNumSeg + 1) As Single
        Dim CatY(MaxNumSubSeg * MaxNumSeg + 1) As Single
        Dim Connector(MaxNumSeg + 1) As Short

        With CurLine
            If TopTension = 0# Then
                If HorForce = 0# Then
                    Scope = .DesScope
                Else
                    Scope = .ScopeByFrcHorPOL(HorForce, TopTension, .Payout)
                End If
            Else
                Scope = .ScopeByTopTensionPOL(TopTension, FrcHor, .Payout)
            End If

            Length = .LineLen
            LenStr = .LineLenStr

            NumSeg = .SegmentCount
            If NumSeg > MaxNumSeg Then NumSeg = MaxNumSeg
            For i = 1 To NumSeg Step 1
                With .Segments(i)
                    SegLength(i) = .Length
                    SegTension(i) = .TenUpp
                    SegAngle(i) = .AngUpp
                    If i < NumSeg Then SegPosition(i + 1) = .YLow
                End With
            Next i

            SegLength(1) = .Payout
            SegPosition(1) = .Draft - .FairLead.z
            If Scope > 0 Then
                GetGroundLength = CurLine.GrdLen
            Else
                GetGroundLength = -1
            End If
            If Not .CatenaryPoints(CatX, CatY, Connector) Then
                Debug.Print("Catenary calculation failed")
                GetGroundLength = -1
            End If
        End With

    End Function

    Public Function CopyMoorLine(ByVal CurLine As MoorLine) As MoorLine
        Dim j As Short
        Dim oMoorLine As MoorLine
        oMoorLine = New MoorLine
        ' make a copy of the mooring line object so that calc max uplift angle won't alter original line FOS
        oMoorLine.SegmentClear()
        For j = 1 To CurLine.SegmentCount
            oMoorLine.SegmentAdd((CurLine.Segments(j).SegType), (CurLine.Segments(j).Length), (CurLine.Segments(j).TotalLength), (CurLine.Segments(j).Diameter), (CurLine.Segments(j).BS), (CurLine.Segments(j).E1), (CurLine.Segments(j).E2), (CurLine.Segments(j).UnitDryWeight), (CurLine.Segments(j).UnitWetWeight), (CurLine.Segments(j).Buoy), (CurLine.Segments(j).BuoyLength), (CurLine.Segments(j).FrictionCoef))

        Next j
        For j = 1 To CurLine.SegmentCount
            oMoorLine.Segments(j).TenUpp = CurLine.Segments(j).TenUpp
            oMoorLine.Segments(j).AngUpp = CurLine.Segments(j).AngUpp
            oMoorLine.Segments(j).AngLow = CurLine.Segments(j).AngLow
        Next j
        With oMoorLine
            .moorName = CurLine.moorName
            .BottomSlope = CurLine.BottomSlope
            .Connected = True
            .DesScope = CurLine.DesScope

            .Payout = CurLine.Payout
            .PayoutSur = CurLine.PayoutSur
            .PayoutOpr = CurLine.PayoutOpr
            .PretensionSur = CurLine.PretensionSur
            .PretensionOpr = CurLine.PretensionOpr

            .FairLead.SprdAngle = CurLine.FairLead.SprdAngle
            .FairLead.Xs = CurLine.FairLead.Xs
            .FairLead.Ys = CurLine.FairLead.Ys
            .FairLead.z = CurLine.FairLead.z

            .Anchor.Xg = CurLine.Anchor.Xg
            .Anchor.Yg = CurLine.Anchor.Yg
            .WaterDepth = CurLine.WaterDepth

            .Anchor.HoldCap = CurLine.Anchor.HoldCap
            .Anchor.Model = CurLine.Anchor.Model
            .Anchor.Remark = CurLine.Anchor.Remark
            .WinchCap = CurLine.WinchCap
        End With
        CopyMoorLine = oMoorLine
    End Function


    Function FindMax(ByVal LineNo As Short, ByRef CaseNo As Short) As Object
        Dim oxSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim MaxVal As Single
        Dim i As Short
        MaxVal = -1000000
        CaseNo = -10
        With oxApp.ActiveWorkbook
            For i = 1 To .Sheets.Count
                oxSheet = .Sheets(i)
                If InStr(oxSheet.Name, "Case") > 0 Then
                    If oxSheet.Range("V" & (44 + LineNo))._Default <> "DAMAGED" Then
                        If oxSheet.Range("V" & (44 + LineNo)).Value > MaxVal Then
                            MaxVal = oxSheet.Range("V" & (44 + LineNo)).Value
                            CaseNo = i - 2
                        End If
                    End If
                End If
            Next i
        End With
        FindMax = Format(MaxVal, "0.00")
    End Function

    Function FindMin(ByVal LineNo As Short, ByRef CaseNo As Short) As Object
        Dim oxSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim MinVal As Single
        Dim i As Short
        MinVal = 1000000
        With oxApp.ActiveWorkbook
            For i = 1 To .Sheets.Count
                oxSheet = .Sheets(i)
                If InStr(oxSheet.Name, "Case") > 0 Then
                    If oxSheet.Range("V" & (44 + LineNo))._Default <> "DAMAGED" Then
                        If oxSheet.Range("V" & (44 + LineNo)).Value < MinVal Then
                            MinVal = oxSheet.Range("V" & (44 + LineNo)).Value
                            CaseNo = i - 2
                        End If
                    End If
                End If
            Next i
        End With
        FindMin = Format(MinVal, "0.00")
    End Function

    Function GetDamagedLineNo(ByRef oxSheet As Microsoft.Office.Interop.Excel.Worksheet, ByRef NumLines As Short) As String
        Dim i As Short
        GetDamagedLineNo = "Intact"
        For i = 1 To NumLines
            If InStr(oxSheet.Range("V" & (44 + i)).Formula, "DAMAGED") > 0 Then
                GetDamagedLineNo = CStr(i)
            End If
        Next i
    End Function

    Sub ReportMooringLayout(ByVal sClient As String, ByVal sLocationName As String, ByVal oVessel As Vessel, Optional ByRef MoorSystem As MoorSystem = Nothing, Optional ByRef ShipLoc As ShipGlobal = Nothing)
        Dim NumPoints As Object
        Dim TmpVal2 As Object
        Dim TmpVal1 As Object

        Dim i, j As Short
        Dim dx, dy As Single

        ' open excel with template
        oxBook = oxApp.Workbooks.Add(My.Application.Info.DirectoryPath & "\MooringLayout.xltm")
        Dim CtrlLineNo As Short
        On Error GoTo ErrHandler
        CtrlLineNo = CShort(InputBox("Enter the mooring line number you choose to display its catenary profile.", "Catenary Line No:", "1"))

        ' set User input values
        oxBook.Sheets("User Input").Activate()
        oxApp.Range("ClientName").Value = sClient
        oxApp.Range("VesselName").Value = oVessel.Name
        oxApp.Range("LocationName").Value = sLocationName
        oxApp.Range("WaterDepth").Value = oVessel.WaterDepth

        Dim NumLines As Short
        If MoorSystem Is Nothing Then
            NumLines = oVessel.MoorSystem.MoorLineCount
        Else
            NumLines = MoorSystem.MoorLineCount
        End If
        oxApp.Range("NumLines").Value = NumLines
        oxApp.Range("NumLines").Formula = Format(NumLines, "#0")
        Dim MaxVal, MinVal As Single

        Dim tmpVal(NumLines - 1) As Single

        For i = 1 To NumLines
            If MoorSystem Is Nothing Then
                With oVessel.MoorSystem.MoorLines(i)
                    If Trim(.Segments(1).SegType) = "WIRE" Then
                        tmpVal(i - 1) = .Payout
                    Else
                        tmpVal(i - 1) = 0 ' no wire
                    End If
                End With
            Else
                With MoorSystem.MoorLines(i)
                    If Trim(.Segments(1).SegType) = "WIRE" Then
                        tmpVal(i - 1) = .Payout
                    Else
                        tmpVal(i - 1) = 0 ' no wire
                    End If
                End With
            End If
        Next i

        Call GetArrayElemRange(tmpVal, MaxVal, MinVal)
        oxApp.Range("WirePayoutMin").Formula = Format(MinVal, "#0")
        oxApp.Range("WirePayoutMax").Formula = Format(MaxVal, "#0")

        For i = 1 To NumLines
            If MoorSystem Is Nothing Then
                With oVessel.MoorSystem.MoorLines(i)
                    With .Segments(.SegmentCount)
                        If Trim(.SegType) = "CHAIN" Then
                            tmpVal(i - 1) = .Length
                        Else
                            tmpVal(i - 1) = 0 ' no chain
                        End If
                    End With
                End With
            Else
                With MoorSystem.MoorLines(i)
                    With .Segments(.SegmentCount)
                        If Trim(.SegType) = "CHAIN" Then
                            tmpVal(i - 1) = .Length
                        Else
                            tmpVal(i - 1) = 0 ' no chain
                        End If
                    End With
                End With
            End If
        Next i

        Call GetArrayElemRange(tmpVal, MaxVal, MinVal)
        oxApp.Sheets("User Input").Range("B14").Formula = Format(MinVal, "#0")
        oxApp.Sheets("User Input").Range("C14").Formula = Format(MaxVal, "#0")

        If ShipLoc Is Nothing Then
            oxApp.Sheets("User Input").Range("VesselHdg").Value = Format(RadTo360(oVessel.ShipCurGlob.Heading), "#0.00")
            oxApp.Sheets("User Input").Range("vPosX").Value = Format(oVessel.ShipCurGlob.Xg, "#0.00")
            oxApp.Sheets("User Input").Range("vPosY").Value = Format(oVessel.ShipCurGlob.Yg, "#0.00")
        Else
            oxApp.Sheets("User Input").Range("VesselHdg").Value = Format(RadTo360((ShipLoc.Heading)), "#0.00")
            oxApp.Sheets("User Input").Range("vPosX").Value = Format(ShipLoc.Xg, "#0.00")
            oxApp.Sheets("User Input").Range("vPosY").Value = Format(ShipLoc.Yg, "#0.00")
        End If

        For i = 1 To NumLines
            If MoorSystem Is Nothing Then
                tmpVal(i - 1) = oVessel.MoorSystem.MoorLines(i).TopTension / 1000
            Else
                tmpVal(i - 1) = MoorSystem.MoorLines(i).TopTension / 1000
            End If
        Next i
        Call GetArrayElemRange(tmpVal, MaxVal, MinVal)
        oxApp.Range("pretensionMin").Formula = Format(MinVal, "#0")
        oxApp.Range("pretensionMax").Formula = Format(MaxVal, "#0")
        If System.Math.Abs(MaxVal - MinVal) < 0.5 Then
            oxApp.Range("pretensionMax").Formula = ""
        End If

        oxApp.Visible = True

        If MoorSystem Is Nothing Then
            If oVessel.MoorSystem.MoorLines(CtrlLineNo).Connected Then
                If ShipLoc Is Nothing Then
                    oxApp.Range("CtrlLineScope").Formula = Format(oVessel.MoorSystem.MoorLines(CtrlLineNo).ScopeByVesselLocation((oVessel.ShipCurGlob), False), "#0.00")
                Else
                    oxApp.Range("CtrlLineScope").Formula = Format(oVessel.MoorSystem.MoorLines(CtrlLineNo).ScopeByVesselLocation(ShipLoc, False), "#0.00")
                End If
                oxApp.Range("CtrlLineGrdLen").Formula = Format(oVessel.MoorSystem.MoorLines(CtrlLineNo).GrdLen, "#0.00")
                oxApp.Range("CtrlLineAnchorUpliftAngle").Formula = Format((oVessel.MoorSystem.MoorLines(CtrlLineNo).BtmAngle - oVessel.MoorSystem.MoorLines(CtrlLineNo).BottomSlope) * Radians2Degrees, "#0.0")
                oxApp.Range("CtrlLineAnchorLoad").Formula = Format(oVessel.MoorSystem.MoorLines(CtrlLineNo).AnchPull / 1000, "#0")
            Else
                oxApp.Range("CtrlLineScope").Formula = "Broken"
                oxApp.Range("CtrlLineGrdLen").Formula = "Broken"
                oxApp.Range("CtrlLineAnchorUpliftAngle").Formula = "Broken"
                oxApp.Range("CtrlLineAnchorLoad").Formula = "Broken"
            End If
        Else
            If MoorSystem.MoorLines(CtrlLineNo).Connected Then
                If ShipLoc Is Nothing Then
                    oxApp.Range("CtrlLineScope").Formula = Format(oVessel.MoorSystem.MoorLines(CtrlLineNo).ScopeByVesselLocation((oVessel.ShipCurGlob), False), "#0.00")
                Else
                    oxApp.Range("CtrlLineScope").Formula = Format(oVessel.MoorSystem.MoorLines(CtrlLineNo).ScopeByVesselLocation(ShipLoc, False), "#0.00")
                End If
                oxApp.Range("CtrlLineGrdLen").Formula = Format(MoorSystem.MoorLines(CtrlLineNo).GrdLen, "#0")
                oxApp.Range("CtrlLineAnchorUpliftAngle").Formula = Format((MoorSystem.MoorLines(CtrlLineNo).BtmAngle - MoorSystem.MoorLines(CtrlLineNo).BottomSlope) * Radians2Degrees, "#0.0")
                oxApp.Range("CtrlLineAnchorLoad").Formula = Format(MoorSystem.MoorLines(CtrlLineNo).AnchPull / 1000, "#0")
                oxApp.Range("Pretension").Formula = Format(MoorSystem.MoorLines(CtrlLineNo).PretensionOpr / 1000, "#0")
            Else
                oxApp.Range("CtrlLineScope").Formula = "Broken"
                oxApp.Range("CtrlLineGrdLen").Formula = "Broken"
                oxApp.Range("CtrlLineAnchorUpliftAngle").Formula = "Broken"
                oxApp.Range("CtrlLineAnchorLoad").Formula = "Broken"
                oxApp.Range("Pretension").Formula = "Broken"
            End If
        End If

        Dim FrcUnit As String
        Dim FrcFactor, LFactor As Single

        With oxApp
            If IsMetricUnit Then
                FrcUnit = " KN"
                FrcFactor = 4.448222 ' kips -> KN
                LFactor = 0.3048 ' ft -> m
                .Application.DisplayAlerts = False
                .Sheets("Payout").Delete()
                .Application.DisplayAlerts = True
                .Sheets("Payout Metric").Name = "Payout"
            Else
                FrcUnit = " kips"
                FrcFactor = 1
                LFactor = 1
                .Application.DisplayAlerts = False
                .Sheets("Payout Metric").Delete()
                .Application.DisplayAlerts = True
            End If

            '-----------------------------------------------------------------------------------
            ' enter Fairlead data, Scopes, and Grounded Length
            If MoorSystem Is Nothing Then
                With oVessel.MoorSystem
                    For i = 1 To NumLines
                        oxApp.Sheets("Payout").Range("L" & (13 + i - 1)).Value = Format(i, "0")
                        If ShipLoc Is Nothing Then

                            oxApp.Sheets("Payout").Range("M" & (13 + i - 1)).Value = Format(Val(CStr(.MoorLines(i).ScopeByVesselLocation((oVessel.ShipCurGlob)))) * LFactor, "0.00")
                            oxApp.Sheets("Payout").Range("N" & (13 + i - 1)).Value = Format(.MoorLines(i).SpreadAngleTN(oVessel.ShipCurGlob.Heading) * Radians2Degrees, "0.00")
                        Else
                            oxApp.Sheets("Payout").Range("M" & (13 + i - 1)).Value = Format(Val(CStr(.MoorLines(i).ScopeByVesselLocation(ShipLoc))) * LFactor, "0.00")
                            oxApp.Sheets("Payout").Range("N" & (13 + i - 1)).Value = Format(.MoorLines(i).SpreadAngleTN(ShipLoc.Heading) * Radians2Degrees, "0.00")
                        End If
                        oxApp.Sheets("Payout").Range("O" & (13 + i - 1)).Value = Val(CStr(.MoorLines(i).TouchDownXg)) * LFactor
                        oxApp.Sheets("Payout").Range("P" & (13 + i - 1)).Value = Val(CStr(.MoorLines(i).TouchDownYg)) * LFactor
                        oxApp.Sheets("Payout").Range("Q" & (13 + i - 1)).Value = Val(CStr(.MoorLines(i).Anchor.Xg)) * LFactor
                        oxApp.Sheets("Payout").Range("R" & (13 + i - 1)).Value = Val(CStr(.MoorLines(i).Anchor.Yg)) * LFactor
                        oxApp.Sheets("calculations").Range("O" & (50 + i - 1)).Formula = Val(CStr(.MoorLines(i).TouchDownXg))
                        oxApp.Sheets("calculations").Range("P" & (50 + i - 1)).Formula = Val(CStr(.MoorLines(i).TouchDownYg))
                        oxApp.Sheets("calculations").Range("O" & (30 + i - 1)).Formula = .MoorLines(i).Anchor.Xg
                        oxApp.Sheets("calculations").Range("P" & (30 + i - 1)).Formula = .MoorLines(i).Anchor.Yg
                        oxApp.Sheets("calculations").Range("B" & (10 + i - 1)).Formula = Val(CStr(.MoorLines(i).FairLead.Xs))
                        oxApp.Sheets("calculations").Range("C" & (10 + i - 1)).Formula = Val(CStr(.MoorLines(i).FairLead.Ys))
                    Next i
                End With
            Else
                With MoorSystem
                    For i = 1 To NumLines
                        oxApp.Sheets("Payout").Range("L" & (13 + i - 1)).Value = Format(i, "0")
                        If ShipLoc Is Nothing Then
                            oxApp.Sheets("Payout").Range("M" & (13 + i - 1)).Value = Format(Val(CStr(.MoorLines(i).ScopeByVesselLocation((oVessel.ShipCurGlob)))) * LFactor, "0.00")
                            oxApp.Sheets("Payout").Range("N" & (13 + i - 1)).Value = Format(.MoorLines(i).SpreadAngleTN(oVessel.ShipCurGlob.Heading) * Radians2Degrees, "0.00")
                        Else

                            oxApp.Sheets("Payout").Range("M" & (13 + i - 1)).Value = Format(Val(CStr(.MoorLines(i).ScopeByVesselLocation(ShipLoc))) * LFactor, "0.00")
                            oxApp.Sheets("Payout").Range("N" & (13 + i - 1)).Value = Format(.MoorLines(i).SpreadAngleTN(ShipLoc.Heading) * Radians2Degrees, "0.00")
                        End If
                        oxApp.Sheets("Payout").Range("O" & (13 + i - 1)).Value = Val(CStr(.MoorLines(i).TouchDownXg)) * LFactor
                        oxApp.Sheets("Payout").Range("P" & (13 + i - 1)).Value = Val(CStr(.MoorLines(i).TouchDownYg)) * LFactor
                        oxApp.Sheets("Payout").Range("Q" & (13 + i - 1)).Value = Val(CStr(.MoorLines(i).Anchor.Xg)) * LFactor
                        oxApp.Sheets("Payout").Range("R" & (13 + i - 1)).Value = Val(CStr(.MoorLines(i).Anchor.Yg)) * LFactor
                        oxApp.Sheets("calculations").Range("O" & (50 + i - 1)).Formula = Val(CStr(.MoorLines(i).TouchDownXg))
                        oxApp.Sheets("calculations").Range("P" & (50 + i - 1)).Formula = Val(CStr(.MoorLines(i).TouchDownYg))
                        oxApp.Sheets("calculations").Range("O" & (30 + i - 1)).Formula = Val(CStr(.MoorLines(i).Anchor.Xg))
                        oxApp.Sheets("calculations").Range("P" & (30 + i - 1)).Formula = Val(CStr(.MoorLines(i).Anchor.Yg))
                        oxApp.Sheets("calculations").Range("B" & (10 + i - 1)).Formula = Val(CStr(.MoorLines(i).FairLead.Xs))
                        oxApp.Sheets("calculations").Range("C" & (10 + i - 1)).Formula = Val(CStr(.MoorLines(i).FairLead.Ys))
                    Next i
                End With
            End If


            For i = NumLines + 1 To 16 ' clear rest lines
                oxApp.Sheets("Payout").Range("L" & (13 + NumLines) & ":R28").Formula = ""
            Next i

            .Sheets("Payout").ChartObjects(2).Activate() ' mooring layout chart
            If MoorSystem Is Nothing Then
                For j = 1 To NumLines
                    If oVessel.MoorSystem.MoorLines(j).Connected Then
                        .ActiveChart.SeriesCollection(j).Name = Format(oVessel.MoorSystem.MoorLines(j).TopTension / 1000 * FrcFactor, "#0") & FrcUnit
                        .ActiveChart.SeriesCollection(j).Points(2).ApplyDataLabels(AutoText:=True, LegendKey:=False, ShowSeriesName:=True, ShowCategoryName:=False, ShowValue:=False, ShowPercentage:=False, ShowBubbleSize:=False)

                        If ShipLoc Is Nothing Then
                            dy = oVessel.MoorSystem.MoorLines(j).Anchor.Yg * LFactor - oVessel.MoorSystem.MoorLines(j).FairLead.Yg((oVessel.ShipCurGlob)) * LFactor
                        Else
                            dy = oVessel.MoorSystem.MoorLines(j).Anchor.Yg * LFactor - oVessel.MoorSystem.MoorLines(j).FairLead.Yg(ShipLoc) * LFactor
                        End If
                        If dy > 0 Then
                            .ActiveChart.SeriesCollection(j).Points(2).DataLabel.Position = Microsoft.Office.Interop.Excel.XlDataLabelPosition.xlLabelPositionBelow
                            .ActiveChart.SeriesCollection(j + 16).Points(2).DataLabel.Position = Microsoft.Office.Interop.Excel.XlDataLabelPosition.xlLabelPositionAbove
                        Else
                            .ActiveChart.SeriesCollection(j).Points(2).DataLabel.Position = Microsoft.Office.Interop.Excel.XlDataLabelPosition.xlLabelPositionAbove
                            .ActiveChart.SeriesCollection(j + 16).Points(2).DataLabel.Position = Microsoft.Office.Interop.Excel.XlDataLabelPosition.xlLabelPositionBelow
                        End If

                    Else
                        .ActiveChart.SeriesCollection(j).Border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone
                    End If
                Next j
                For j = 32 To 17 + NumLines Step -1
                    '  remove all other mooring lines (series)
                    .ActiveChart.SeriesCollection(j).Delete()
                Next j
                For j = 16 To NumLines + 1 Step -1
                    '  remove all other mooring lines (series)
                    With .ActiveChart.SeriesCollection(j)
                        .Border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone
                        .MarkerStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
                    End With
                Next j
            Else
                For j = 1 To NumLines
                    If MoorSystem.MoorLines(j).Connected Then
                        .ActiveChart.SeriesCollection(j).Name = Format(MoorSystem.MoorLines(j).TopTension / 1000 * FrcFactor, "#0") & FrcUnit
                        .ActiveChart.SeriesCollection(j).Points(2).ApplyDataLabels(AutoText:=True, LegendKey:=False, ShowSeriesName:=True, ShowCategoryName:=False, ShowValue:=False, ShowPercentage:=False, ShowBubbleSize:=False)

                        If ShipLoc Is Nothing Then
                            dy = oVessel.MoorSystem.MoorLines(j).Anchor.Yg * LFactor - oVessel.MoorSystem.MoorLines(j).FairLead.Yg((oVessel.ShipCurGlob)) * LFactor
                        Else
                            dy = oVessel.MoorSystem.MoorLines(j).Anchor.Yg * LFactor - oVessel.MoorSystem.MoorLines(j).FairLead.Yg(ShipLoc) * LFactor
                        End If
                        If dy > 0 Then
                            .ActiveChart.SeriesCollection(j).Points(2).DataLabel.Position = Microsoft.Office.Interop.Excel.XlDataLabelPosition.xlLabelPositionBelow
                            .ActiveChart.SeriesCollection(j + 16).Points(2).DataLabel.Position = Microsoft.Office.Interop.Excel.XlDataLabelPosition.xlLabelPositionAbove
                        Else
                            .ActiveChart.SeriesCollection(j).Points(2).DataLabel.Position = Microsoft.Office.Interop.Excel.XlDataLabelPosition.xlLabelPositionAbove
                            .ActiveChart.SeriesCollection(j + 16).Points(2).DataLabel.Position = Microsoft.Office.Interop.Excel.XlDataLabelPosition.xlLabelPositionBelow
                        End If
                    Else
                        .ActiveChart.SeriesCollection(j).Border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone
                    End If
                Next j
                For j = 32 To 17 + NumLines Step -1
                    ' first remove all other mooring lines (series)
                    .ActiveChart.SeriesCollection(j).Delete()
                Next j
                For j = 16 To NumLines + 1 Step -1
                    '  remove all other mooring lines (series)
                    With .ActiveChart.SeriesCollection(j)
                        .Border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone
                        .MarkerStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
                    End With
                Next j
            End If

            With .ActiveChart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory)
                .MinimumScaleIsAuto = True
                .MaximumScaleIsAuto = True
                .MinorUnitIsAuto = True
                .MajorUnitIsAuto = True
                .Crosses = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .ReversePlotOrder = False
                .ScaleType = Microsoft.Office.Interop.Excel.XlTrendlineType.xlLinear
                .DisplayUnit = Microsoft.Office.Interop.Excel.Constants.xlNone
                TmpVal1 = .MaximumScale - .MinimumScale
            End With
            With .ActiveChart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue)
                .MinimumScaleIsAuto = True
                .MaximumScaleIsAuto = True
                .MinorUnitIsAuto = True
                .MajorUnitIsAuto = True
                .Crosses = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .ReversePlotOrder = False
                .ScaleType = Microsoft.Office.Interop.Excel.XlTrendlineType.xlLinear
                .DisplayUnit = Microsoft.Office.Interop.Excel.Constants.xlNone
                TmpVal2 = .MaximumScale - .MinimumScale
            End With
            .ActiveChart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory).MaximumScale = .ActiveChart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory).MinimumScale + Max(TmpVal1, TmpVal2)
            .ActiveChart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue).MaximumScale = .ActiveChart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue).MinimumScale + Max(TmpVal1, TmpVal2)
        End With
        '-------------------------------------------------------------------------------------------
        ' paste catenary profile
        If 1 = 1 Then
            Dim NumSegment As Short
            Dim IsConnected As Boolean

            Dim CatX(MaxNumSubSeg * MaxNumSeg + 1) As Single
            Dim CatY(MaxNumSubSeg * MaxNumSeg + 1) As Single
            Dim Connector(MaxNumSeg + 1) As Short

            If MoorSystem Is Nothing Then
                With oVessel.MoorSystem.MoorLines(CtrlLineNo)
                    NumSegment = .SegmentCount
                    IsConnected = .Connected
                    Call .CatenaryPoints(CatX, CatY, Connector)
                End With
            Else
                With MoorSystem.MoorLines(CtrlLineNo)
                    NumSegment = .SegmentCount
                    IsConnected = .Connected
                    Call .CatenaryPoints(CatX, CatY, Connector)
                End With
            End If

            For i = 1 To NumSegment
                NumPoints = Connector(i) - Connector(i + 1) + 1
                For j = 1 To NumPoints
                    oxApp.Sheets("catenary Input").Range("A" & (5 + Connector(i) - j + 1)).Value = CatX(Connector(i) - j + 1)
                    oxApp.Sheets("catenary Input").Range("B" & (5 + Connector(i) - j + 1)).Value = CatY(Connector(i) - j + 1)
                    oxApp.Sheets("catenary Input").Range("C" & (5 + Connector(i) - j + 1)).Value = ""
                    oxApp.Sheets("catenary Input").Range("C" & (5 + Connector(i))).Value = "<<<-Connect"
                Next j
            Next i

            Dim NumSeries As Short
            With oxApp.ActiveWorkbook

                .ActiveSheet.ChartObjects(1).Activate() '  Catenary plot

                With .ActiveChart
                    With .Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue)
                        .MinimumScaleIsAuto = True
                        .MaximumScaleIsAuto = True
                        .MinorUnitIsAuto = True
                        .MajorUnitIsAuto = True
                        .Crosses = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .ReversePlotOrder = True
                        .ScaleType = Microsoft.Office.Interop.Excel.XlTrendlineType.xlLinear
                        .DisplayUnit = Microsoft.Office.Interop.Excel.Constants.xlNone
                    End With

                    ' first remove all series
                    For i = .SeriesCollection.Count To 1 Step -1
                        .SeriesCollection(i).Delete()
                    Next i

                    If MoorSystem Is Nothing Then
                        NumSeries = oVessel.MoorSystem.MoorLines(CtrlLineNo).SegmentCount
                    Else
                        NumSeries = MoorSystem.MoorLines(CtrlLineNo).SegmentCount
                    End If
                    For i = 1 To NumSeries
                        .SeriesCollection.NewSeries()
                    Next i

                    For i = 1 To NumSeries 'JLIU TODO
                        With .SeriesCollection(i)
                            If oxApp.Sheets("catenary Input") IsNot Nothing Then
                                If IsMetricUnit Then ' plot from bottom up on excel  catenary input sheet, i.e. fairlead down
                                    .XValues = oxApp.Sheets("catenary Input").Range("D" & (5 + Connector(i)) & ":D" & (5 + Connector(i + 1)))
                                    .Values = oxApp.Sheets("catenary Input").Range("E" & (5 + Connector(i)) & ":E" & (5 + Connector(i + 1)))
                                Else
                                    'JLIU TODO
                                    .XValues = oxApp.Sheets("catenary Input").Range("A" & (5 + Connector(i)) & ":A" & (5 + Connector(i + 1)))
                                    .Values = oxApp.Sheets("catenary Input").Range("B" & (5 + Connector(i)) & ":B" & (5 + Connector(i + 1)))
                                End If
                            End If
                            With .Border
                                If IsConnected Then
                                    '                   .ColorIndex = xlAutomatic
                                    '                   .Color = vbBlack
                                Else
                                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone
                                    '                            .Color = xlTransparent
                                End If
                                If InStr(oVessel.MoorSystem.MoorLines(CtrlLineNo).Segments(i).SegType, "CHAIN") > 0 Then
                                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
                                Else
                                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                                End If
                                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                            End With
                            .MarkerStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
                            .ApplyDataLabels(AutoText:=False, LegendKey:=False, ShowSeriesName:=False, ShowCategoryName:=False, ShowValue:=False, ShowPercentage:=False, ShowBubbleSize:=False)
                        End With
                    Next i

                    .SeriesCollection(1).Ponits(1).ApplyDataLabels(AutoText:=False, LegendKey:=False, ShowSeriesName:=True, ShowCategoryName:=False, ShowValue:=False, ShowPercentage:=False, ShowBubbleSize:=False)
                    With .SeriesCollection(1).Points(1).DataLabel
                        .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                        .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                        .ReadingOrder = Microsoft.Office.Interop.Excel.Constants.xlContext
                        .Position = Microsoft.Office.Interop.Excel.XlDataLabelPosition.xlLabelPositionAbove
                        .Orientation = Microsoft.Office.Interop.Excel.XlOrientation.xlHorizontal
                        .Characters.Text = "Line " & CtrlLineNo
                        With .Characters.Font
                            .Name = "Arial"
                            .FontStyle = "Bold"
                            .Size = 14
                            .Strikethrough = False
                            .Superscript = False
                            .Subscript = False
                            .OutlineFont = False
                            .Shadow = False
                            .Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleNone
                            .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        End With
                    End With

                End With
                .ActiveSheet.Range("A1").Activate() ' de-select chart
                .ActiveSheet.PageSetup.PrintArea = "$A$1:$R$38"
                With .ActiveSheet.PageSetup
                    .FitToPagesWide = 1
                    .FitToPagesTall = 1
                    .PrintErrors = Microsoft.Office.Interop.Excel.XlPrintErrors.xlPrintErrorsDisplayed
                End With
                '        ' hide template sheets
                ' .Sheets("calculations").Visible = False
                '.Sheets("User Input").Visible = False
                '.Sheets("catenary Input").Visible = False

            End With
        End If
        OutputSummaryTab(oxApp, oVessel, oVessel.WaterDepth, MoorSystem, ShipLoc)
        OutputLinePropertiesTab(oxApp, oVessel, MoorSystem)
        OutMooringAQWATab(oxApp, oVessel, MoorSystem)
        Exit Sub
ErrHandler:
        MsgBox("Error Creating Mooring Layout Report: " & Err.Description, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Error")
    End Sub

    Sub ReportDODO(ByRef CurVessel As Vessel)
        Dim oxSheet As Microsoft.Office.Interop.Excel.Worksheet

        If oxBook Is Nothing Then
            oxBook = oxApp.Workbooks.Add(My.Application.Info.DirectoryPath & "\DODO2.xltx")
            oxBook.Sheets("input").activate()
            oxSheet = DirectCast(oxApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
            oxSheet.Range("$B$1").Value = CurVessel.Name
            oxSheet.Range("$B$2").Value = CurVessel.ShipCurGlob.Xg
            oxSheet.Range("$B$3").Value = CurVessel.ShipCurGlob.Yg
            oxSheet.Range("$B$4").Value = RadTo360(CurVessel.ShipCurGlob.Heading)
            oxSheet.Range("$B$5").Value = CurVessel.ShipDraftOpr
            oxSheet.Range("$B$6").Value = CurVessel.ShipDraftSur
            oxSheet.Range("$B$7").Value = CurVessel.Riser.mass / 1000 * MassFactor
            oxSheet.Range("$B$8").Value = CurVessel.Riser.TopTen / 1000 * FrcFactor
            oxSheet.Range("$B$9").Value = CurVessel.WaterDepth
            oxSheet.Range("$B$10").Value = CurVessel.Riser.LFJDepth
        End If

        oxBook.Sheets("dodo").Copy(After:=oxBook.Sheets(1))
        oxBook.Sheets("dodo (2)").Activate()
        oxBook.Sheets("dodo (2)").Name = CurVessel.EnvLoad.EnvCur.Name
        oxSheet = DirectCast(oxApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
        oxSheet.Range("$B$16").Value = oxBook.Sheets.Count - 5 'JLIU TODO
        oxSheet.Range("$B$1").Value = CurVessel.EnvLoad.EnvCur.Wind.Velocity * Ftps2Knots
        oxSheet.Range("$D$1").Value = CurVessel.EnvLoad.EnvCur.Wind.Heading * Radians2Degrees
        oxSheet.Range("$B$2").Value = CurVessel.EnvLoad.EnvCur.Wave.Height
        oxSheet.Range("$D$2").Value = CurVessel.EnvLoad.EnvCur.Wave.Heading * Radians2Degrees
        oxSheet.Range("$G$2").Value = CurVessel.EnvLoad.EnvCur.Wave.Period

        oxSheet.Range("$B$3").Value = CurVessel.EnvLoad.EnvCur.Current.Profile(1).Velocity * Ftps2Knots
        oxSheet.Range("$D$3").Value = CurVessel.EnvLoad.EnvCur.Current.Heading * Radians2Degrees

        With oxSheet.QueryTables.Add(Connection:="TEXT;" & CurProj.Directory & "appvel.out", Destination:=oxSheet.Range("$N$3"))
            .PreserveFormatting = False
            .RefreshStyle = Microsoft.Office.Interop.Excel.XlCellInsertionMode.xlOverwriteCells 'xlOverwriteCells
            .AdjustColumnWidth = False
            .TextFileParseType = Microsoft.Office.Interop.Excel.XlTextParsingType.xlDelimited
            .TextFileCommaDelimiter = True
            .Refresh(BackgroundQuery:=False)
        End With
        With oxSheet.QueryTables.Add(Connection:="TEXT;" & CurProj.Directory & "offset.out", Destination:=oxSheet.Range("$N$27"))
            .PreserveFormatting = False
            .RefreshStyle = Microsoft.Office.Interop.Excel.XlCellInsertionMode.xlOverwriteCells 'xlOverwriteCells
            .AdjustColumnWidth = False
            .TextFileParseType = Microsoft.Office.Interop.Excel.XlTextParsingType.xlDelimited
            .TextFileCommaDelimiter = True
            .Refresh(BackgroundQuery:=False)

        End With
        oxSheet.Columns("N:X").Font.Name = "Calibri"
        oxSheet.Columns("N:X").Font.Size = 8
        oxSheet.Columns("N:X").Font.Bold = False
        'Return control of Excel to the user.
        oxApp.Visible = True
        oxApp.ScreenUpdating = True
        oxApp.CutCopyMode = False
        oxApp.UserControl = True

    End Sub

    Function CleanupObjects() As Object
        On Error Resume Next ' ignore if window non-exist
        oxBook = Nothing
        oxBooks = Nothing
        oxApp = Nothing
    End Function

    Private Function DeleteWorksheet(ByRef strSheetName As String) As Boolean
        On Error Resume Next

        ExcelGlobal_definst.Application.DisplayAlerts = False
        ExcelGlobal_definst.ActiveWorkbook.Worksheets(strSheetName).Delete()
        ExcelGlobal_definst.Application.DisplayAlerts = True
        ' Return True if no error occurred;
        ' otherwise return False.
        DeleteWorksheet = Not CBool(Err.Number)
    End Function

    Public Sub New()
        MyBase.New()
        oxApp = New Microsoft.Office.Interop.Excel.Application
        oxBooks = oxApp.Workbooks
    End Sub

    Protected Overrides Sub Finalize()
        Call CleanupObjects()
        MyBase.Finalize()
    End Sub

    Sub ReportMooringAnalysisResults(ByVal sTitle As String, ByRef SubTitle As String, ByVal NumCases As Short, ByVal IsDataEnglishUnit As Boolean, ByRef oVessel As Vessel, ByRef MeanPosition As Motion, ByRef InitPos As Motion, ByRef Stiff As Force, ByRef SigLFM As Motion, ByRef SigWFM As Motion, ByRef Ten() As Single, ByRef TenLF() As Single, ByRef TenWF() As Single)
        Dim nrowcop As Object
        Dim nrowd As Object
        Dim NLA As Object
        Dim DamLineCur As Object
        Dim nd As Object
        Dim NL As Object
        Dim StiffRz As Object
        Dim StiffY As Object
        Dim StiffX As Object
        Dim vslzacc As Object
        Dim VslZWF As Object
        Dim VslYWF As Object
        Dim VslXWF As Object
        Dim VslZLF As Object
        Dim VslYLF As Object
        Dim VslXLF As Object
        Dim VslZ As Object
        Dim VslY As Object
        Dim VslX As Object
        Dim WaveRz2 As Object
        Dim WaveY2 As Object
        Dim WaveX2 As Object
        Dim WindRz2 As Object
        Dim WindY2 As Object
        Dim WindX2 As Object
        Dim CurrRz2 As Object
        Dim CurrY2 As Object
        Dim CurrX2 As Object
        Dim WaveRz As Object
        Dim WaveY As Object
        Dim WaveX As Object
        Dim WindRz As Object
        Dim WindY As Object
        Dim WindX As Object
        Dim CurrRz As Object
        Dim CurrY As Object
        Dim CurrX As Object
        Dim ThruRz As Object
        Dim ThruY As Object
        Dim ThruX As Object


        Dim BrokenLineNo(40, NumCases) As Short
        Dim TotalNumBroken(NumCases) As Short

        Dim NameDir As String
        Dim NumLineA, NumLines As Short
        Dim NCase As Short
        Dim CaseNam As String
        Dim NameSheet As String
        Dim Row, Col As Short
        Dim WindVel, CurrVel As Single
        Dim WaveHs As Single
        Dim SwellHs As Single

        Dim j, i, CritSegNo As Short
        Dim IsDamaged As Boolean
        IsDamaged = False
        NumLines = oVessel.MoorSystem.MoorLineCount

        Dim MaxVal, MinVal As Single
        Dim tmpVal(NumLines) As Single
        ' post-processer will be able to handler multiple lines damaged in each case
        With oVessel.MoorSystem
            For i = 1 To NumCases
                TotalNumBroken(i) = 0
                For j = 1 To .MoorLineCount
                    If Not .MoorLines(j).Connected Then
                        IsDamaged = True
                        TotalNumBroken(i) = TotalNumBroken(i) + 1
                        BrokenLineNo(TotalNumBroken(i), i) = j
                    Else
                    End If
                Next j
            Next i
        End With

        NameDir = "Case" ' assume subdir name "case"

        oxBook = oxApp.Workbooks.Add(My.Application.Info.DirectoryPath & "\MARS_MooringResults.xlt")
        ' set intro-input values

        oxBook.Sheets("Intro-Input").Activate()
        oxApp.Range("MainTitle").Formula = sTitle
        oxApp.Range("SubTitle").Formula = SubTitle

        With oVessel
            If System.Math.Abs(.MoorSystem.MoorLines(1).Payout - .MoorSystem.MoorLines(1).PayoutOpr) < 0.5 Then
                For i = 1 To NumLines
                    tmpVal(i) = .MoorSystem.MoorLines(i).PretensionOpr / 1000
                Next i
            Else
                For i = 1 To NumLines
                    tmpVal(i) = .MoorSystem.MoorLines(i).PretensionSur / 1000
                Next i
            End If
            Call GetArrayElemRange(tmpVal, MaxVal, MinVal)
            If MaxVal - MinVal < 5 Then
                oxApp.Range("OperTension").Value = Format((MinVal + MaxVal) / 2, "#0")
            Else
                oxApp.Range("OperTension").Value = Format(MinVal, "#0") & " - " & Format(MaxVal, "#0")
                oxApp.Range("PreTensionKN").Value = Format(MinVal * 4.448222, "#0") & " - " & Format(MaxVal * 4.448222, "#0") ' kips -> KN
            End If
            oxApp.Range("BreakStrength").Formula = Format(.MoorSystem.MoorLines(1).Segments(1).BS / 1000, "#0")
            oxApp.Range("WD").Formula = .WaterDepth
            oxApp.Range("xx").Formula = InitPos.Surge
            oxApp.Range("yy").Formula = InitPos.Sway
            oxApp.Range("zz").Formula = ""
            oxApp.Range("rx").Formula = ""
            oxApp.Range("ry").Formula = ""
            oxApp.Range("rz").Formula = Format(InitPos.Yaw * Radians2Degrees, "0.00")
        End With
        oxApp.Visible = True

        Dim frcH, minFS, maxT As Single
        Dim EndRow As Short
        For NCase = 1 To NumCases
            If Not IsDamaged Then
                CaseNam = CStr(NCase)
            Else
                CaseNam = NCase & "D"
            End If

            NameSheet = NameDir & CaseNam
            CaseNam = NameDir & CaseNam

            '     If IsDamaged Then
            '         NumLineA = NumLines - TotalNumBroken(NCase)
            '     Else
            NumLineA = NumLines
            '     End If

            oxBook.Sheets("MOORING1").Copy(After:=oxBook.Sheets(NCase + 1))
            oxBook.Sheets("MOORING1 (2)").Name = CaseNam
            oxBook.ActiveSheet.Cells(72, 16) = CaseNam

            '--------------------------------------------------------------------------------------

            With oxApp

                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells._Default(6, 3) = (oVessel.EnvLoad.EnvCur.Wind.Heading) * Radians2Degrees
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells._Default(6, 7) = Format(RadTo360(oVessel.EnvLoad.EnvCur.Wave.Heading), "#0.0")
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells._Default(6, 8) = ""
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells._Default(8, 7) = oVessel.EnvLoad.EnvCur.Wave.Period
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells._Default(8, 8) = ""
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells._Default(6, 12) = (oVessel.EnvLoad.EnvCur.Current.Heading) * Radians2Degrees

                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells._Default(25, 6) = ""
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells._Default(25, 7) = ""
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells._Default(25, 8) = (MeanPosition.Yaw) * Radians2Degrees

                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells._Default(26, 6) = ""
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells._Default(26, 7) = ""
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells._Default(26, 8) = (SigLFM.Yaw) * Radians2Degrees

                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells._Default(27, 6) = ""
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells._Default(27, 7) = ""
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells._Default(27, 8) = (SigWFM.Yaw) * Radians2Degrees

                '--------------------------------------------------------------------------
                WindVel = oVessel.EnvLoad.EnvCur.Wind.Velocity
                WaveHs = oVessel.EnvLoad.EnvCur.Wave.Height
                SwellHs = 0
                If oVessel.EnvLoad.EnvCur.Current.ProfileCount > 0 Then
                    CurrVel = oVessel.EnvLoad.EnvCur.Current.Profile(1).Velocity
                Else
                    CurrVel = 0
                End If

                'UPGRADE_WARNING: Couldn't resolve default property of object ThruX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ThruX = 0
                'UPGRADE_WARNING: Couldn't resolve default property of object ThruY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ThruY = 0
                'UPGRADE_WARNING: Couldn't resolve default property of object ThruRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ThruRz = 0

                'UPGRADE_WARNING: Couldn't resolve default property of object CurrX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                CurrX = oVessel.EnvLoad.FCurrGlob.Fx
                'UPGRADE_WARNING: Couldn't resolve default property of object CurrY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                CurrY = oVessel.EnvLoad.FCurrGlob.Fy
                'UPGRADE_WARNING: Couldn't resolve default property of object CurrRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                CurrRz = oVessel.EnvLoad.FCurrGlob.MYaw
                'UPGRADE_WARNING: Couldn't resolve default property of object WindX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                WindX = oVessel.EnvLoad.FWindGlob.Fx
                'UPGRADE_WARNING: Couldn't resolve default property of object WindY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                WindY = oVessel.EnvLoad.FWindGlob.Fy
                'UPGRADE_WARNING: Couldn't resolve default property of object WindRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                WindRz = oVessel.EnvLoad.FWindGlob.MYaw
                'UPGRADE_WARNING: Couldn't resolve default property of object WaveX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                WaveX = oVessel.EnvLoad.FWaveGlob.Fx
                'UPGRADE_WARNING: Couldn't resolve default property of object WaveY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                WaveY = oVessel.EnvLoad.FWaveGlob.Fy
                'UPGRADE_WARNING: Couldn't resolve default property of object WaveRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                WaveRz = oVessel.EnvLoad.FWaveGlob.MYaw

                oVessel.EnvLoad.ShipHead = oVessel.ShipCurGlob.Heading

                'UPGRADE_WARNING: Couldn't resolve default property of object CurrX2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                CurrX2 = oVessel.EnvLoad.FCurrLocl.Fx
                'UPGRADE_WARNING: Couldn't resolve default property of object CurrY2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                CurrY2 = oVessel.EnvLoad.FCurrLocl.Fy
                'UPGRADE_WARNING: Couldn't resolve default property of object CurrRz2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                CurrRz2 = oVessel.EnvLoad.FCurrLocl.MYaw
                'UPGRADE_WARNING: Couldn't resolve default property of object WindX2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                WindX2 = oVessel.EnvLoad.FWindLocl.Fx
                'UPGRADE_WARNING: Couldn't resolve default property of object WindY2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                WindY2 = oVessel.EnvLoad.FWindLocl.Fy
                'UPGRADE_WARNING: Couldn't resolve default property of object WindRz2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                WindRz2 = oVessel.EnvLoad.FWindLocl.MYaw
                'UPGRADE_WARNING: Couldn't resolve default property of object WaveX2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                WaveX2 = oVessel.EnvLoad.FWaveLocl.Fx
                'UPGRADE_WARNING: Couldn't resolve default property of object WaveY2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                WaveY2 = oVessel.EnvLoad.FWaveLocl.Fy
                'UPGRADE_WARNING: Couldn't resolve default property of object WaveRz2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                WaveRz2 = oVessel.EnvLoad.FWaveLocl.MYaw

                'UPGRADE_WARNING: Couldn't resolve default property of object VslX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                VslX = MeanPosition.Surge
                'UPGRADE_WARNING: Couldn't resolve default property of object VslY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                VslY = MeanPosition.Sway
                'UPGRADE_WARNING: Couldn't resolve default property of object VslZ. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                VslZ = ""

                'UPGRADE_WARNING: Couldn't resolve default property of object VslXLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                VslXLF = SigLFM.Surge
                'UPGRADE_WARNING: Couldn't resolve default property of object VslYLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                VslYLF = SigLFM.Sway
                'UPGRADE_WARNING: Couldn't resolve default property of object VslZLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                VslZLF = ""

                'UPGRADE_WARNING: Couldn't resolve default property of object VslXWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                VslXWF = SigWFM.Surge
                'UPGRADE_WARNING: Couldn't resolve default property of object VslYWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                VslYWF = SigWFM.Sway
                'UPGRADE_WARNING: Couldn't resolve default property of object VslZWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                VslZWF = ""
                'UPGRADE_WARNING: Couldn't resolve default property of object vslzacc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                vslzacc = "" ' ignore heave accelleration

                'UPGRADE_WARNING: Couldn't resolve default property of object StiffX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                StiffX = Stiff.Fx
                'UPGRADE_WARNING: Couldn't resolve default property of object StiffY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                StiffY = Stiff.Fy
                'UPGRADE_WARNING: Couldn't resolve default property of object StiffRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                StiffRz = Stiff.MYaw

                If IsDataEnglishUnit Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(7, 3) = WindVel
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(7, 7) = WaveHs
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(7, 8) = SwellHs
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(7, 12) = CurrVel

                    'UPGRADE_WARNING: Couldn't resolve default property of object ThruX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(14, 9) = ThruX
                    'UPGRADE_WARNING: Couldn't resolve default property of object ThruY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(14, 10) = ThruY
                    'UPGRADE_WARNING: Couldn't resolve default property of object ThruRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(14, 11) = ThruRz
                    'UPGRADE_WARNING: Couldn't resolve default property of object CurrX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(16, 9) = CurrX
                    'UPGRADE_WARNING: Couldn't resolve default property of object CurrY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(16, 10) = CurrY
                    'UPGRADE_WARNING: Couldn't resolve default property of object CurrRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(16, 11) = CurrRz
                    'UPGRADE_WARNING: Couldn't resolve default property of object WindX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(17, 9) = WindX
                    'UPGRADE_WARNING: Couldn't resolve default property of object WindY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(17, 10) = WindY
                    'UPGRADE_WARNING: Couldn't resolve default property of object WindRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(17, 11) = WindRz
                    'UPGRADE_WARNING: Couldn't resolve default property of object WaveX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(18, 9) = WaveX
                    'UPGRADE_WARNING: Couldn't resolve default property of object WaveY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(18, 10) = WaveY
                    'UPGRADE_WARNING: Couldn't resolve default property of object WaveRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(18, 11) = WaveRz

                    'UPGRADE_WARNING: Couldn't resolve default property of object CurrX2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(16, 12) = CurrX2
                    'UPGRADE_WARNING: Couldn't resolve default property of object CurrY2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(16, 13) = CurrY2
                    'UPGRADE_WARNING: Couldn't resolve default property of object CurrRz2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(16, 14) = CurrRz2
                    'UPGRADE_WARNING: Couldn't resolve default property of object WindX2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(17, 12) = WindX2
                    'UPGRADE_WARNING: Couldn't resolve default property of object WindY2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(17, 13) = WindY2
                    'UPGRADE_WARNING: Couldn't resolve default property of object WindRz2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(17, 14) = WindRz2
                    'UPGRADE_WARNING: Couldn't resolve default property of object WaveX2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(18, 12) = WaveX2
                    'UPGRADE_WARNING: Couldn't resolve default property of object WaveY2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(18, 13) = WaveY2
                    'UPGRADE_WARNING: Couldn't resolve default property of object WaveRz2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(18, 14) = WaveRz2

                    'UPGRADE_WARNING: Couldn't resolve default property of object VslX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(25, 3) = VslX
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(25, 4) = VslY
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslZ. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(25, 5) = VslZ

                    'UPGRADE_WARNING: Couldn't resolve default property of object VslXLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(26, 3) = VslXLF
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslYLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(26, 4) = VslYLF
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslZLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(26, 5) = VslZLF

                    'UPGRADE_WARNING: Couldn't resolve default property of object VslXWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(27, 3) = VslXWF
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslYWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(27, 4) = VslYWF
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslZWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(27, 5) = VslZWF
                    'UPGRADE_WARNING: Couldn't resolve default property of object vslzacc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(28, 5) = vslzacc

                    'UPGRADE_WARNING: Couldn't resolve default property of object StiffX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(14, 2) = StiffX
                    'UPGRADE_WARNING: Couldn't resolve default property of object StiffY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(15, 2) = StiffY
                    'UPGRADE_WARNING: Couldn't resolve default property of object StiffRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(16, 2) = StiffRz
                Else ' if results are in Metric unit, convert to English to fill in the excel template
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(7, 3) = WindVel / 0.3048
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(7, 7) = WaveHs / 0.3048
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(7, 8) = SwellHs / 0.3048
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(7, 12) = CurrVel / 0.3048

                    ' FORCES  KN  SO * 1000

                    ' Thrusters

                    'UPGRADE_WARNING: Couldn't resolve default property of object ThruX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(14, 9) = 1000 * ThruX / 9.80665 / 0.4536
                    'UPGRADE_WARNING: Couldn't resolve default property of object ThruY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(14, 10) = 1000 * ThruY / 9.80665 / 0.4536
                    'UPGRADE_WARNING: Couldn't resolve default property of object ThruRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(14, 11) = 1000 * ThruRz / 9.80665 / 0.4536 / 0.3048

                    ' Current Forces
                    'UPGRADE_WARNING: Couldn't resolve default property of object CurrX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(16, 9) = 1000 * CurrX / 9.80665 / 0.4536
                    'UPGRADE_WARNING: Couldn't resolve default property of object CurrY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(16, 10) = 1000 * CurrY / 9.80665 / 0.4536
                    'UPGRADE_WARNING: Couldn't resolve default property of object CurrRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(16, 11) = 1000 * CurrRz / 9.80665 / 0.4536 / 0.3048
                    ' Wind Forces
                    'UPGRADE_WARNING: Couldn't resolve default property of object WindX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(17, 9) = 1000 * WindX / 9.80665 / 0.4536
                    'UPGRADE_WARNING: Couldn't resolve default property of object WindY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(17, 10) = 1000 * WindY / 9.80665 / 0.4536
                    'UPGRADE_WARNING: Couldn't resolve default property of object WindRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(17, 11) = 1000 * WindRz / 9.80665 / 0.4536 / 0.3048
                    ' Wave Forces
                    'UPGRADE_WARNING: Couldn't resolve default property of object WaveX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(18, 9) = 1000 * WaveX / 9.80665 / 0.4536
                    'UPGRADE_WARNING: Couldn't resolve default property of object WaveY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(18, 10) = 1000 * WaveY / 9.80665 / 0.4536
                    'UPGRADE_WARNING: Couldn't resolve default property of object WaveRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(18, 11) = 1000 * WaveRz / 9.80665 / 0.4536 / 0.3048

                    ' Motions
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(25, 3) = VslX / 0.3048
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(25, 4) = VslY / 0.3048
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslZ. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(25, 5) = VslZ / 0.3048

                    'UPGRADE_WARNING: Couldn't resolve default property of object VslXLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(26, 3) = VslXLF / 0.3048
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslYLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(26, 4) = VslYLF / 0.3048
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslZLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(26, 5) = VslZLF / 0.3048

                    'UPGRADE_WARNING: Couldn't resolve default property of object VslXWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(27, 3) = VslXWF / 0.3048
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslYWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(27, 4) = VslYWF / 0.3048
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslZWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(27, 5) = VslZWF / 0.3048

                    ' Acceleration
                    'UPGRADE_WARNING: Couldn't resolve default property of object vslzacc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(28, 5) = vslzacc / 0.3048

                    ' Mooring Stiffness
                    'UPGRADE_WARNING: Couldn't resolve default property of object StiffX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(14, 2) = 1000 * StiffX / 9.80665 / 0.4536 * 0.3048
                    'UPGRADE_WARNING: Couldn't resolve default property of object StiffY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(15, 2) = 1000 * StiffY / 9.80665 / 0.4536 * 0.3048
                    'UPGRADE_WARNING: Couldn't resolve default property of object StiffRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells._Default(16, 2) = 1000 * StiffRz / 9.80665 / 0.4536

                End If


                Row = 35
                'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                NL = 0 ' intact line counter
                'UPGRADE_WARNING: Couldn't resolve default property of object nd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                nd = 0 ' damaged line counter
                'UPGRADE_WARNING: Couldn't resolve default property of object DamLineCur. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                DamLineCur = 0 ' current damaged line no.
                For NLA = 1 To NumLineA
                    Row = Row + 1
                    'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    NL = NL + 1
                    If IsDamaged Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object DamLineCur. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object nd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If BrokenLineNo(nd, NCase) > 0 And nd < TotalNumBroken(NCase) And NL > DamLineCur Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object nd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            nd = nd + 1
                            'UPGRADE_WARNING: Couldn't resolve default property of object nd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object DamLineCur. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            DamLineCur = BrokenLineNo(nd, NCase)
                        End If
                    End If
                    'UPGRADE_WARNING: Couldn't resolve default property of object DamLineCur. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If NL = DamLineCur Then
                        For Col = 3 To 5
                            'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .Cells._Default(Row, Col) = 0
                        Next Col
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells._Default(Row + 9, 21) = "Dam"
                    Else
                        If IsDataEnglishUnit Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object NLA. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .Cells._Default(Row, 3) = Ten(NLA)
                            'UPGRADE_WARNING: Couldn't resolve default property of object NLA. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .Cells._Default(Row, 4) = TenLF(NLA)
                            'UPGRADE_WARNING: Couldn't resolve default property of object NLA. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .Cells._Default(Row, 5) = TenWF(NLA)
                        Else
                            'UPGRADE_WARNING: Couldn't resolve default property of object NLA. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .Cells._Default(Row, 3) = Ten(NLA) / 9.80665 / 0.4536 * 1000
                            'UPGRADE_WARNING: Couldn't resolve default property of object NLA. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .Cells._Default(Row, 4) = TenLF(NLA) / 9.80665 / 0.4536 * 1000
                            'UPGRADE_WARNING: Couldn't resolve default property of object NLA. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .Cells._Default(Row, 5) = TenWF(NLA) / 9.80665 / 0.4536 * 1000
                        End If
                    End If

                Next NLA

                If Not IsDamaged Then
                    .Range("k24").Value = "Intact"
                Else
                    .Range("k24").Value = "Damaged"
                End If

                .Range(.Cells._Default(45 + NumLines, 16), .Cells._Default(64, 30)).Delete(Shift:=Microsoft.Office.Interop.Excel.XlDirection.xlUp)

                For j = 1 To NumLines
                    maxT = Max(1.5 * TenLF(j) + TenWF(j), TenLF(j) + 1.86 * TenWF(j)) + Ten(j)
                    If maxT > 0 Then
                        With oVessel.MoorSystem.MoorLines(j)
                            If .ScopeByTopTensionPOL(maxT, frcH, .Payout) > 0 Then
                                minFS = .FOS
                                CritSegNo = oVessel.MoorSystem.MoorLines(j).CriticalSegNo
                            Else
                                minFS = .Segments(1).BS / maxT
                                CritSegNo = 1
                            End If
                        End With
                    End If
                    .Range("W" & (44 + j)).FormulaR1C1 = CritSegNo
                    .Range("V" & (44 + j)).FormulaR1C1 = Format(minFS, "0.00")
                Next j

                If IsDamaged Then ' clear content for the damaged lines
                    For nd = 1 To TotalNumBroken(NCase)
                        'UPGRADE_WARNING: Couldn't resolve default property of object nd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object nrowd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        nrowd = 44 + BrokenLineNo(nd, NCase)
                        'UPGRADE_WARNING: Couldn't resolve default property of object nrowcop. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        nrowcop = 44 + NumLines - 1
                        ' below replaced by delete above
                        '  .Range("Q" & nrowd & ":AD" & nrowcop).Copy
                        '   .Range("Q" & nrowd + 1).PasteSpecial xlPasteValues, xlPasteSpecialOperationNone
                        'UPGRADE_WARNING: Couldn't resolve default property of object nrowd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Range("Q" & nrowd, "U" & nrowd).FormulaR1C1 = "  "
                        'UPGRADE_WARNING: Couldn't resolve default property of object nrowd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Range("V" & nrowd).FormulaR1C1 = "DAMAGED"
                        'UPGRADE_WARNING: Couldn't resolve default property of object nrowd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Range("Y" & nrowd, "AC" & nrowd).FormulaR1C1 = "  "
                        'UPGRADE_WARNING: Couldn't resolve default property of object nrowd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Range("AD" & nrowd).FormulaR1C1 = "DAMAGED"
                    Next nd
                End If


                '      Set Header and Title
                EndRow = 45 + NumLines
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveSheet.PageSetup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                With .ActiveSheet.PageSetup
                    ' set printArea based on result Unit
                    If IsDataEnglishUnit Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveSheet.PageSetup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .PrintArea = "$P$1:$V$" & EndRow
                    Else
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveSheet.PageSetup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .PrintArea = "$X$1:$AD$" & EndRow
                    End If
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveSheet.PageSetup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .PrintTitleRows = ""
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveSheet.PageSetup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .PrintTitleColumns = ""
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveSheet.PageSetup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .LeftHeader = "&D"
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveSheet.PageSetup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .CenterHeader = " "
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveSheet.PageSetup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .RightHeader = "Prepared by DTCEL"
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveSheet.PageSetup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .CenterFooter = "David Tein Consulting Engineers, Ltd." & vbCrLf & "11777 Katy Freeway, Suite 434, Houston, TX 77079, 281-531-0888, FAX 281-531-5888, dtcel@dtcel.com"

                End With

            End With
        Next NCase

        ' hide template sheet
        'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.Sheets().Visible. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        oxApp.ActiveWorkbook.Sheets("MOORING1").Visible = False

        ' cleanup
    End Sub


    Sub FormatSpreadChart(ByRef oChart As Microsoft.Office.Interop.Excel.Chart)
        Dim i As Short

        With oChart
            .ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatterSmooth
        End With

        For i = 1 To 16 ' ungrounded mooring lines
            With oChart.SeriesCollection(i)
                'UPGRADE_WARNING: Couldn't resolve default property of object oChart.SeriesCollection(i).Border. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                With .Border
                    'UPGRADE_WARNING: Couldn't resolve default property of object oChart.SeriesCollection().Border. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .ColorIndex = 1
                    'UPGRADE_WARNING: Couldn't resolve default property of object oChart.SeriesCollection().Border. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                    'UPGRADE_WARNING: Couldn't resolve default property of object oChart.SeriesCollection().Border. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDot
                End With
                'UPGRADE_WARNING: Couldn't resolve default property of object oChart.SeriesCollection().MarkerBackgroundColorIndex. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .MarkerBackgroundColorIndex = 1
                'UPGRADE_WARNING: Couldn't resolve default property of object oChart.SeriesCollection().MarkerForegroundColorIndex. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .MarkerForegroundColorIndex = 1
                'UPGRADE_WARNING: Couldn't resolve default property of object oChart.SeriesCollection().MarkerStyle. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .MarkerStyle = Microsoft.Office.Interop.Excel.Constants.xlSquare
                'UPGRADE_WARNING: Couldn't resolve default property of object oChart.SeriesCollection().Smooth. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Smooth = True
                'UPGRADE_WARNING: Couldn't resolve default property of object oChart.SeriesCollection().MarkerSize. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .MarkerSize = 5
                'UPGRADE_WARNING: Couldn't resolve default property of object oChart.SeriesCollection().Shadow. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Shadow = False
            End With
        Next i

        For i = 17 To 32 '  grounded length
            With oChart.SeriesCollection(i)
                'UPGRADE_WARNING: Couldn't resolve default property of object oChart.SeriesCollection(i).Border. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                With .Border
                    'UPGRADE_WARNING: Couldn't resolve default property of object oChart.SeriesCollection().Border. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .ColorIndex = 1
                    'UPGRADE_WARNING: Couldn't resolve default property of object oChart.SeriesCollection().Border. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                    'UPGRADE_WARNING: Couldn't resolve default property of object oChart.SeriesCollection().Border. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                End With
                'UPGRADE_WARNING: Couldn't resolve default property of object oChart.SeriesCollection().MarkerBackgroundColorIndex. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .MarkerBackgroundColorIndex = 1
                'UPGRADE_WARNING: Couldn't resolve default property of object oChart.SeriesCollection().MarkerForegroundColorIndex. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .MarkerForegroundColorIndex = 1
                'UPGRADE_WARNING: Couldn't resolve default property of object oChart.SeriesCollection().MarkerStyle. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .MarkerStyle = Microsoft.Office.Interop.Excel.Constants.xlCircle
                'UPGRADE_WARNING: Couldn't resolve default property of object oChart.SeriesCollection().Smooth. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Smooth = True
                'UPGRADE_WARNING: Couldn't resolve default property of object oChart.SeriesCollection().MarkerSize. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .MarkerSize = 5
                'UPGRADE_WARNING: Couldn't resolve default property of object oChart.SeriesCollection().Shadow. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Shadow = False
            End With
        Next i

    End Sub

    Sub OutputSummaryTab(ByRef oxApp As Microsoft.Office.Interop.Excel.Application, ByRef oVessel As Vessel, Optional ByVal WD As Single = 0, Optional ByRef MoorSystem As MoorSystem = Nothing, Optional ByRef ShipLoc As ShipGlobal = Nothing)
        Dim Row, i, NumLines As Short
        Dim CurLine As MoorLine
        Dim MoorLine As New MoorLine
        Dim BS, RetVal As Single

        Dim TmpVal2 As Single
        ' send payout summary data to excel
        If MoorSystem Is Nothing Then
            NumLines = oVessel.MoorSystem.MoorLineCount
        Else
            NumLines = MoorSystem.MoorLineCount
        End If
        Dim OutTenh As Single
        With oxApp.Sheets("Summary")
            'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
            If IsNothing(WD) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Range("C2").Formula = oVessel.WaterDepth
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Range("C2").Formula = WD
            End If
            For i = 1 To NumLines
                If MoorSystem Is Nothing Then
                    CurLine = oVessel.MoorSystem.MoorLines(i)
                Else
                    CurLine = MoorSystem.MoorLines(i)
                End If

                ' make a copy of the mooring line object so that calc max uplift angle won't alter original line FOS
                MoorLine = CopyMoorLine(CurLine)

                BS = MoorLine.Segments(1).BS / 1.43 '  use top Segment breaking strength
                TmpVal2 = CurLine.Payout
                RetVal = MoorLine.ScopeByTopTensionPOL(BS, OutTenh, TmpVal2)
                If RetVal < 0 Then
                    'frmMain.Activate()
                    MsgBox("Catenary calculation did not converge for line " & i & ". You have to enter the Max Anchor Uplift Angles manually for the report Summary.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Error")
                End If
                Row = i + 6
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Range("B" & Row).Value = i
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Range("C" & Row).Value = CurLine.WaterDepth
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Range("D" & Row).Value = CurLine.BottomSlope * Radians2Degrees
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Range("E" & Row).Value = CurLine.TopTension / 1000
                If CurLine.Connected Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Range("F" & Row).Value = CurLine.FOS
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Range("L" & Row).Value = CurLine.Payout
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Range("M" & Row).Value = CurLine.GrdLen
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Range("F" & Row).Value = "-"
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Range("L" & Row).Value = "-"
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Range("M" & Row).Value = "-"
                End If
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Range("G" & Row).Value = CurLine.AnchPull / 1000
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Range("H" & Row).Value = CurLine.BtmAngle * Radians2Degrees - CurLine.BottomSlope * Radians2Degrees
                If RetVal < 0 Then ' did not converge
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Range("I" & Row).Value = ""
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Range("I" & Row).Value = MoorLine.BtmAngle * Radians2Degrees - CurLine.BottomSlope * Radians2Degrees
                End If
                '   Debug.Print " MaxAngle= " & MoorLine.BtmAngle * Radians2Degrees & " Actual Angle = " & MoorLine.BtmAngle * Radians2Degrees - CurLine.BottomSlope * Radians2Degrees
                '   Debug.Print "i= " & i & " BS/1.43= " & BS & " Slope= " & CurLine.BottomSlope * Radians2Degrees
                If ShipLoc Is Nothing Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Range("J" & Row).Value = Format(RadTo360(CurLine.SprdAngle((oVessel.ShipCurGlob))), "0.0")
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Range("K" & Row).Value = CurLine.ScopeByVesselLocation((oVessel.ShipCurGlob))
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Range("J" & Row).Value = Format(RadTo360(CurLine.SprdAngle(ShipLoc)), "0.0")
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Range("K" & Row).Value = CurLine.ScopeByVesselLocation(ShipLoc)
                End If
            Next i
            For i = NumLines + 1 To 16
                Row = i + 6
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Range("B" & Row & ":M" & Row).Value = ""
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Range("O" & Row & ":Z" & Row).Value = ""
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Range("O" & Row & ":Z" & Row).Formula = ""
            Next i
            ' setup print area
            If IsMetricUnit Then
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().PageSetup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .PageSetup.PrintArea = "$O$1:$Z$25"
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().PageSetup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .PageSetup.PrintArea = "$B$1:$M$25"
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().PageSetup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape
        End With
    End Sub

    Sub OutputLinePropertiesTab(ByRef oxApp As Microsoft.Office.Interop.Excel.Application, ByRef oVessel As Vessel, Optional ByRef MoorSystem As MoorSystem = Nothing)
        Dim Row, i, NumLines As Short
        Dim CurLine As MoorLine

        Dim j As Short

        ' send payout summary data to excel
        If MoorSystem Is Nothing Then
            NumLines = oVessel.MoorSystem.MoorLineCount
        Else
            NumLines = MoorSystem.MoorLineCount
        End If

        With oxApp.Sheets("Line Properties")
            Row = 6
            For i = 1 To NumLines
                If MoorSystem Is Nothing Then
                    CurLine = oVessel.MoorSystem.MoorLines(i)
                Else
                    CurLine = MoorSystem.MoorLines(i)
                End If
                For j = 1 To CurLine.SegmentCount
                    Row = Row + 1
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Range("B" & Row).Value = i
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Range("C" & Row).Value = j
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Range("D" & Row).Value = CurLine.Segments(j).SegType
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Range("E" & Row).Value = CurLine.Segments(j).TotalLength
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Range("F" & Row).Value = CurLine.Segments(j).Diameter
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Range("G" & Row).Value = CurLine.Segments(j).BS / 1000
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Range("H" & Row).Value = CurLine.Segments(j).E1 / 1000
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Range("I" & Row).Value = CurLine.Segments(j).E2 / 1000
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Range("J" & Row).Value = CurLine.Segments(j).UnitDryWeight
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Range("K" & Row).Value = CurLine.Segments(j).UnitWetWeight
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Range("L" & Row).Value = CurLine.Segments(j).Buoy / 1000
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Range("M" & Row).Value = CurLine.Segments(j).BuoyLength
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Range("N" & Row).Value = CurLine.Segments(j).FrictionCoef
                Next j
            Next i

            ' Row = Row + 1

            ' Draw table bottom border
            'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .Range("B" & CStr(Row) & ":AB" & CStr(Row)).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium

            ' setup print area
            If IsMetricUnit Then
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().PageSetup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .PageSetup.PrintArea = "$O$1:$AB$" & Row + 1
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().PageSetup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .PageSetup.PrintArea = "$B$1:$N$" & Row + 1
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().PageSetup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape

            ' clean up rest rows
            'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .Range("B" & CStr(Row + 1) & ":AB" & CStr(Row + 200)).Clear()
            'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .Range("O" & CStr(Row)).Clear()
        End With

    End Sub
    Sub OutMooringAQWATab(ByRef oxApp As Microsoft.Office.Interop.Excel.Application, ByRef oVessel As Vessel, Optional ByRef MoorSystem As MoorSystem = Nothing)
        Dim i, NumLines As Short
        Dim CurLine As MoorLine

        Dim j As Short

        ' send payout summary data to excel
        If MoorSystem Is Nothing Then
            NumLines = oVessel.MoorSystem.MoorLineCount
        Else
            NumLines = MoorSystem.MoorLineCount
        End If
        Dim WD As Single
        WD = oVessel.WaterDepth

        With oxApp.Sheets("Mooring")
            .Range("A1:J100").clear()

            .cells.font.name = "Courier New"
            .cells.font.size = 9

            Dim CurRow As Integer
            CurRow = 1
            For i = 1 To NumLines
                CurLine = oVessel.MoorSystem.MoorLines(i)
                .Cells(CurRow, 1) = "14COMP"
                .Cells(CurRow, 2) = 3
                .Cells(CurRow, 3) = 20
                .Cells(CurRow, 5) = CurLine.SegmentCount
                .Cells(CurRow, 6) = WD - 30.0#
                .Cells(CurRow, 6).NumberFormat = "0"
                .Cells(CurRow, 7) = WD + 30.0#
                .Cells(CurRow, 7).NumberFormat = "0"
                .Cells(CurRow, 8) = 0#
                .Cells(CurRow, 8).NumberFormat = "0"
                CurRow = CurRow + 1
                For j = CurLine.SegmentCount To 1 Step -1
                    .Cells(CurRow, 1) = "14ECAT"
                    .Cells(CurRow, 6) = CurLine.Segments(j).UnitDryWeight / 32.18
                    .Cells(CurRow, 6).NumberFormat = "0.0000"
                    .Cells(CurRow, 7) = (CurLine.Segments(j).UnitDryWeight - CurLine.Segments(j).UnitWetWeight) / 32.18 / 1.99
                    .Cells(CurRow, 7).NumberFormat = "0.0000"
                    .Cells(CurRow, 8) = 3.1415926 * CurLine.Segments(j).Diameter ^ 2 / 4 * CurLine.Segments(j).E1
                    .Cells(CurRow, 8).NumberFormat = "0.000E+00"
                    .Cells(CurRow, 9) = CurLine.Segments(j).BS
                    .Cells(CurRow, 9).NumberFormat = "0.000E+00"
                    If j = 1 Then
                        .Cells(CurRow, 10) = CurLine.Payout
                    Else
                        .Cells(CurRow, 10) = CurLine.Segments(j).Length
                    End If
                    .Cells(CurRow, 10).NumberFormat = "0.0"
                    CurRow = CurRow + 1

                    .Cells(CurRow, 1) = "14ECAH"
                    .Cells(CurRow, 6) = 1.0#
                    .Cells(CurRow, 6).NumberFormat = "0.0"
                    If Left(CurLine.Segments(j).SegType, 4) = "CHAI" Then
                        .Cells(CurRow, 8) = 2.4
                        .Cells(CurRow, 10) = 0.1
                    Else
                        .Cells(CurRow, 8) = 1.2
                        .Cells(CurRow, 10) = 0.025
                    End If
                    .Cells(CurRow, 8).NumberFormat = "0.0"
                    .Cells(CurRow, 10).NumberFormat = "0.000"
                    .Cells(CurRow, 9) = CurLine.Segments(j).Diameter / 12
                    .Cells(CurRow, 9).NumberFormat = "0.0000"
                    CurRow = CurRow + 1
                Next j

                If i = NumLines Then
                    .Cells(CurRow, 1) = "END14NLID"
                Else
                    .Cells(CurRow, 1) = "14NLID"
                End If
                .Cells(CurRow, 2) = 1
                .Cells(CurRow, 3) = CurLine.Anchor.Node
                .Cells(CurRow, 4) = 0
                .Cells(CurRow, 5) = CurLine.FairLead.Node
                CurRow = CurRow + 1
            Next i

        End With
        With oxApp.Sheets("anchor")
            .Range("A1:F50").clear()

            .cells.font.name = "Courier New"
            .cells.font.size = 9

            Dim CurRow As Integer
            CurRow = 1
            For i = 1 To NumLines
                .Cells(i, 1) = "01" & oVessel.MoorSystem.MoorLines(i).Anchor.Node
                .Cells(i, 4) = oVessel.MoorSystem.MoorLines(i).Anchor.Xg * LFactor
                .Cells(i, 5) = oVessel.MoorSystem.MoorLines(i).Anchor.Yg * LFactor
                .Cells(i, 6) = -oVessel.MoorSystem.MoorLines(i).WaterDepth

            Next
        End With
        With oxApp.Sheets("fairlead")
            .Range("A1:F50").clear()

            .cells.font.name = "Courier New"
            .cells.font.size = 9

            Dim CurRow As Integer
            CurRow = 1
            For i = 1 To NumLines
                .Cells(i, 1) = "01" & oVessel.MoorSystem.MoorLines(i).FairLead.Node
                .Cells(i, 4) = oVessel.MoorSystem.MoorLines(i).FairLead.Xs * LFactor
                .Cells(i, 5) = oVessel.MoorSystem.MoorLines(i).FairLead.Ys * LFactor
                .Cells(i, 6) = -oVessel.MoorSystem.MoorLines(i).FairLead.z * LFactor
            Next
        End With

    End Sub



End Class