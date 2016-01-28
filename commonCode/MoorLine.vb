Option Strict Off
Option Explicit On
Friend Class MoorLine
	Implements System.Collections.IEnumerable
    Public Const MaxNumSeg As Short = 24
    Public Const MaxNumSubSeg As Short = 24
    Private Const MinSubSegLen As Single = 50

    Private Const MaxIter As Short = 2000

    Private Const AngleTol As Single = 0.001
    Private Const DistTol As Single = 0.00005
    Private Const DepthTol As Single = 0.0005
    Private Const ForceTol As Single = 0.5

    Private Const NumSubSeg As Short = 20

    Private Const BuoyHeight As Single = 2.0#

    ' mooring line properties
    ' name:         mooring line name
    ' Segments:     mooring line segments
    ' FairLead:     fairlead location & node (default 950i)
    ' Anchor:       anchor location & node (default 910i)
    ' Payout:       payout of the 1st segment (rig wire)
    ' PayoutSur:    payout at survival draft
    ' PayoutOpr:    payout at operating draft
    ' PretensionSur:pretension at survival draft
    ' PretensionOpr:pretension at operating draft
    ' DesScope:     design scope
    ' Draft:        vessel draft at fairlead
    ' WaterDepth:   waterdepth at anchor point
    ' BottomSlope:  sea bottom slope at anchor point
    ' Connected:    flag mooring line connected
    ' Winched:      flag winch functional
    ' WinchCap:     winch capacity, limiting top tension
    ' IsSpring:     flag whether this line is a spring line (not catenary line)

    ' Scope:        scope of the mooring line
    ' SprdAngle:    spreading angle (0 - 2 pi)
    ' TopTension:   top tension
    ' FLAngle:      fair lead angle
    ' BtmAngle:     anchor uplift angle
    ' GrdLen:       grounded length
    ' LineLen:      total line length
    ' LineLenStr:   total line length after stretching
    ' BuoyDepth:    buoy depth
    ' AnchPull:     anchor pull
    ' FMoorLocl:    horizontal force and it components by mooring line
    ' Tension:      top tension at a particular scope
    ' MaxTension:   max tension along the line
    ' FOS:          safety factor
    ' AnchorFOS:    anchor safety factor on holding capacity

    ' methodes

    ' AnchorLocation:   find anchor location by scope and spreading angle
    ' TensionByScopePOL:calculate tension by scope and payout
    ' ScopeByTensionPOL:calculate scope by tension and payout
    ' POLByScopeTension:calcualte payout by scope and tension
    ' POLByScopeFrcHor: calcualte payout by scope and horizontal force
    ' CatenaryPoints:   generate catenary plot points

    ' internal functions

    ' Catenary:         calculate catenary
    ' CatNoBuoy:        calculate catenary for mooring line w/o buoys
    ' CatWBuoy:         calculate catenary for mooring line with buoys
    ' TouchBtm:         find horizontal tension when kth segment just touches bottem
    Private msngMoorName As String
    Private mcolSegments As Collection
    Private mclsFairLead As FairLead
    Private mclsAnchor As Anchor
    Private mclsFMoorLocl As Force

    Private msngPayout As Single
    Private msngTopTension As Single
    Private msngGrdLen As Single
    Private msngAnchPull As Single

    Private msngPayoutSur As Single
    Private msngPayoutOpr As Single
    Private msngPretensionSur As Single
    Private msngPretensionOpr As Single

    Private msngDesScope As Single
    Private msngScope As Single
    Private msngSprdAngle As Single
    Private msngDraft As Single
    Private msngWaterDepth As Single
    Private msngDepthFL As Single
    Private msngBottomSlope As Single
    Private mblnConnected As Boolean
    Private mblnWinched As Boolean
    Private msngWinchCap As Single

    Private mTouchDownX As Single
    Private mTouchDownY As Single

    Private mblnSpring As Boolean
    Private mblnTBCalc As Boolean
    Private msngFh0k(MaxNumSeg) As Single
    Private msngThU0k(MaxNumSeg, MaxNumSeg) As Single
    Private msngThL0k(MaxNumSeg, MaxNumSeg) As Single

    Private msngTSigLF As Single
    Private msngTSigWF As Single
    Private mintCriticalSegNo As Short

    Public Sub New()
        MyBase.New()
        mcolSegments = New Collection
        mclsFairLead = New FairLead
        mclsAnchor = New Anchor
        mclsFMoorLocl = New Force

        mblnSpring = False
        mblnTBCalc = False
        mblnConnected = True
        mblnWinched = True

        msngWinchCap = 0#
    End Sub

    Public Property moorName() As String
        Get

            moorName = msngMoorName

        End Get
        Set(ByVal Value As String)

            msngMoorName = Value

        End Set
    End Property

    Public Property Payout() As Single
        Get

            Payout = msngPayout

        End Get
        Set(ByVal Value As Single)

            msngPayout = Value
            mblnTBCalc = False

        End Set
    End Property

    Public ReadOnly Property CriticalSegNo() As Short
        Get
            CriticalSegNo = mintCriticalSegNo
        End Get
    End Property


    Public Property TopTension() As Single
        Get

            If mblnConnected And mcolSegments.Count() > 0 Then
                TopTension = mcolSegments.Item(1).TenUpp
                If TopTension = 0# Then TopTension = msngTopTension
            Else
                TopTension = 0#
            End If

        End Get
        Set(ByVal Value As Single)

            msngTopTension = Value

        End Set
    End Property

    Public ReadOnly Property FLAngle() As Single
        Get

            FLAngle = mcolSegments.Item(1).AngUpp

        End Get
    End Property

    Public ReadOnly Property BtmAngle() As Single
        Get

            Dim NumSeg As Short

            NumSeg = mcolSegments.Count()
            BtmAngle = mcolSegments.Item(NumSeg).AngLow

        End Get
    End Property

    Public ReadOnly Property AnchPull() As Single
        Get

            If IsSpring Then '2.2.1
                AnchPull = TopTension
            Else
                AnchPull = msngAnchPull
            End If

        End Get
    End Property

    Public ReadOnly Property GrdLen() As Single
        Get

            If IsSpring Then '2.2.1
                GrdLen = 0#
            Else
                GrdLen = msngGrdLen
            End If

        End Get
    End Property

    Public ReadOnly Property TouchDownXg() As Single
        Get
            TouchDownXg = mTouchDownX
        End Get
    End Property

    Public ReadOnly Property TouchDownYg() As Single
        Get
            TouchDownYg = mTouchDownY
        End Get
    End Property

    Public ReadOnly Property LineLen() As Single
        Get

            Dim Segment_Renamed As Segment

            LineLen = 0#
            If mcolSegments.Count() > 0 Then
                For Each Segment_Renamed In mcolSegments
                    LineLen = LineLen + Segment_Renamed.Length
                Next Segment_Renamed
                LineLen = LineLen - mcolSegments.Item(1).Length + msngPayout
            End If
        End Get
    End Property

    Public ReadOnly Property LineLenStr() As Single
        Get

            Dim Segment_Renamed As Segment

            LineLenStr = 0#
            For Each Segment_Renamed In mcolSegments
                LineLenStr = LineLenStr + Segment_Renamed.LengthStr
            Next Segment_Renamed

        End Get
    End Property

    Public ReadOnly Property BuoyDepth() As Single
        Get

            Dim i As Short
            Dim tmp As Single

            BuoyDepth = 0#
            For i = 1 To mcolSegments.Count()
                With mcolSegments.Item(i)
                    If .Buoy < -0.00001 And BuoyDepth < .YLow Then BuoyDepth = .YLow
                    If .UnitWetWeight < -0.00001 Then
                        If i > 1 Then
                            tmp = mcolSegments.Item(i - 1).YLow
                        Else
                            tmp = 0#
                        End If
                        If BuoyDepth < tmp Then BuoyDepth = tmp
                        If BuoyDepth < .YLow Then BuoyDepth = .YLow
                    End If
                End With
            Next

        End Get
    End Property

    Public ReadOnly Property MaxTension() As Single
        Get

            Dim Segment_Renamed As Segment

            MaxTension = 0#
            For Each Segment_Renamed In mcolSegments
                If Segment_Renamed.TenUpp > MaxTension Then MaxTension = Segment_Renamed.TenUpp
            Next Segment_Renamed

        End Get
    End Property

    Public ReadOnly Property FOS() As Single
        Get

            Dim i As Short

            FOS = 1000000.0#
            For i = 1 To mcolSegments.Count()
                With mcolSegments.Item(i)
                    If .TenUpp > 0# Then
                        If .BS / .TenUpp < FOS Then
                            FOS = .BS / .TenUpp
                            mintCriticalSegNo = i
                        End If
                    End If
                End With
            Next

        End Get
    End Property

    Public ReadOnly Property AnchorFOS() As String
        Get
            If mclsAnchor.HoldCap = 0# Then
                AnchorFOS = "--"
            ElseIf msngAnchPull <= 1 Then
                AnchorFOS = "Large"
            Else
                AnchorFOS = Format(mclsAnchor.HoldCap / msngAnchPull, "0.00")
            End If

        End Get
    End Property


    Public Property PayoutSur() As Single
        Get

            PayoutSur = msngPayoutSur

        End Get
        Set(ByVal Value As Single)

            msngPayoutSur = Value

        End Set
    End Property


    Public Property PayoutOpr() As Single
        Get

            PayoutOpr = msngPayoutOpr

        End Get
        Set(ByVal Value As Single)

            msngPayoutOpr = Value

        End Set
    End Property


    Public Property PretensionSur() As Single
        Get

            PretensionSur = msngPretensionSur

        End Get
        Set(ByVal Value As Single)

            msngPretensionSur = Value

        End Set
    End Property


    Public Property PretensionOpr() As Single
        Get

            PretensionOpr = msngPretensionOpr

        End Get
        Set(ByVal Value As Single)

            msngPretensionOpr = Value

        End Set
    End Property

    Public ReadOnly Property FMoorLocl(ByVal ShipGlob As ShipGlobal, ByVal Moved As Boolean, Optional ByVal ReCalculate As Boolean = True) As Force
        Get

            Dim X, Y As Single
            Dim HorTension As Single

            If Not mblnConnected Then
                With mclsFMoorLocl
                    .Fx = 0#
                    .Fy = 0#
                    .MYaw = 0#
                End With
                msngTopTension = 0#

                FMoorLocl = mclsFMoorLocl
                Exit Property
            End If

            If Not ReCalculate Then
                FMoorLocl = mclsFMoorLocl
                Exit Property
            End If

            If mclsFairLead.NeedUpdate Or mclsAnchor.NeedUpdate Or Moved Then
                msngScope = System.Math.Sqrt((mclsAnchor.Xg - mclsFairLead.Xg(ShipGlob)) ^ 2 + (mclsAnchor.Yg - mclsFairLead.Yg(ShipGlob)) ^ 2)
                X = mclsAnchor.Xg - mclsFairLead.Xg(ShipGlob)
                Y = mclsAnchor.Yg - mclsFairLead.Yg(ShipGlob)
                msngSprdAngle = Atan(Y, X) - ShipGlob.Heading
                If msngSprdAngle < 0# Then msngSprdAngle = 2.0# * PI + msngSprdAngle

                mclsFairLead.NeedUpdate = False
                mclsAnchor.NeedUpdate = False
            End If

            msngTopTension = TensionByScopePOL(msngScope, HorTension, msngPayout)

            With mclsFMoorLocl
                .Fx = HorTension * System.Math.Cos(msngSprdAngle)
                .Fy = -HorTension * System.Math.Sin(msngSprdAngle)
                .MYaw = - .Fx * mclsFairLead.Ys + .Fy * mclsFairLead.Xs
            End With

            FMoorLocl = mclsFMoorLocl

        End Get
    End Property


    Public Property DesScope() As Single
        Get

            DesScope = msngDesScope

        End Get
        Set(ByVal Value As Single)

            msngDesScope = Value
            msngScope = msngDesScope

        End Set
    End Property

    Public ReadOnly Property Scope() As Single
        Get

            Scope = msngScope

        End Get
    End Property

    ' if ship location and orientation are provided, calculate actual spreading angle
    ' based on anchor location
    ' This is the spread angle in vessel north - bow north
    Public ReadOnly Property SprdAngle(ByVal ShipGlob As ShipGlobal, Optional ByVal Moved As Boolean = False) As Single
        Get

            Dim X, Y As Single
            If mclsFairLead.NeedUpdate Or mclsAnchor.NeedUpdate Or Moved Then
                msngScope = System.Math.Sqrt((mclsAnchor.Xg - mclsFairLead.Xg(ShipGlob)) ^ 2 + (mclsAnchor.Yg - mclsFairLead.Yg(ShipGlob)) ^ 2)

                X = mclsAnchor.Xg - mclsFairLead.Xg(ShipGlob)
                Y = mclsAnchor.Yg - mclsFairLead.Yg(ShipGlob)
                msngSprdAngle = Atan(Y, X) - ShipGlob.Heading
                If msngSprdAngle < 0# Then msngSprdAngle = 2.0# * PI + msngSprdAngle
                '    msngSprdAngle = msngSprdAngle - Fix(msngSprdAngle / (PI * 2)) * PI * 2

                mclsFairLead.NeedUpdate = False
                mclsAnchor.NeedUpdate = False
            End If

            SprdAngle = msngSprdAngle
        End Get
    End Property

    Public ReadOnly Property SpreadAngleTN(ByVal VesselHdg As Single) As Single
        Get

            Dim tmpVal As Single

            tmpVal = msngSprdAngle + VesselHdg

            If tmpVal < 0# Then
                SpreadAngleTN = 2.0# * PI + tmpVal
            ElseIf tmpVal > (2.0# * PI) Then
                SpreadAngleTN = tmpVal - 2.0# * PI
            Else
                SpreadAngleTN = tmpVal
            End If

        End Get
    End Property


    Public Property Draft() As Single
        Get

            Draft = msngDraft

        End Get
        Set(ByVal Value As Single)

            msngDraft = Value
            mblnTBCalc = False

        End Set
    End Property


    Public Property WaterDepth() As Single
        Get

            WaterDepth = msngWaterDepth

        End Get
        Set(ByVal Value As Single)

            msngWaterDepth = Value
            mblnTBCalc = False

        End Set
    End Property


    Public Property BottomSlope() As Single
        Get

            BottomSlope = msngBottomSlope

        End Get
        Set(ByVal Value As Single)

            msngBottomSlope = Value
            mblnTBCalc = False

        End Set
    End Property


    Public Property Connected() As Boolean
        Get

            Connected = mblnConnected

        End Get
        Set(ByVal Value As Boolean)

            mblnConnected = Value

        End Set
    End Property


    Public Property Winched() As Boolean
        Get

            Winched = mblnWinched

        End Get
        Set(ByVal Value As Boolean)

            mblnWinched = Value

        End Set
    End Property


    Public Property WinchCap() As Single
        Get

            WinchCap = msngWinchCap

        End Get
        Set(ByVal Value As Single)

            msngWinchCap = Value

        End Set
    End Property

    Public ReadOnly Property IsSpring() As Boolean
        Get

            Dim i, nSp As Short

            With mcolSegments
                nSp = 0
                For i = 1 To .Count()
                    If .Item(i).SegType = "SPRING" Then nSp = nSp + 1
                Next
                If nSp = .Count() Then IsSpring = True
            End With

        End Get
    End Property

    Public ReadOnly Property Segments(ByVal vntIndexKey As Object) As Segment
        Get

            Segments = mcolSegments.Item(vntIndexKey)

        End Get
    End Property

    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        GetEnumerator = mcolSegments.GetEnumerator
    End Function

    Public ReadOnly Property FairLead() As FairLead
        Get

            FairLead = mclsFairLead

        End Get
    End Property

    Public ReadOnly Property Anchor() As Anchor
        Get

            Anchor = mclsAnchor

        End Get
    End Property

    Public Sub SegmentAdd(ByRef SegType As String, ByRef Length As Single, ByRef TotalLength As Single, ByRef Diameter As Single, ByRef BS As Single, ByRef E1 As Single, ByRef E2 As Single, ByRef UnitDryWeight As Single, ByRef UnitWetWeight As Single, ByRef Buoy As Single, ByRef BuoyLength As Single, ByRef FrictionCoef As Single) '2.1.4

        Dim NewSegment As New Segment

        With NewSegment
            .SegType = SegType
            .Length = Length
            .TotalLength = TotalLength
            .Diameter = Diameter
            .BS = BS
            .E1 = E1
            .E2 = E2
            .UnitDryWeight = UnitDryWeight
            .UnitWetWeight = UnitWetWeight
            .Buoy = Buoy
            .BuoyLength = BuoyLength
            .FrictionCoef = FrictionCoef
        End With

        mcolSegments.Add(NewSegment)

        mblnTBCalc = False

    End Sub

    Public Sub SegmentClear()

        Dim NumSegment, i As Short

        NumSegment = mcolSegments.Count()

        For i = NumSegment To 1 Step -1
            mcolSegments.Remove(i)
        Next

        mblnTBCalc = False

    End Sub

    Public Function SegmentCount() As Integer

        SegmentCount = mcolSegments.Count()

    End Function

    Public Sub SegmentDelete(ByRef vntIndexKey As Object)

        mcolSegments.Remove(vntIndexKey)

        mblnTBCalc = False

    End Sub

    ' determine anchor location based on scope, ship location and orientation
    ' and design spreading angle
    Public Sub AnchorLocation(ByRef Scope As Single, ByRef ShipGlob As ShipGlobal)

        Dim Alpha As Single

        Alpha = PI / 2.0# - (mclsFairLead.SprdAngle + ShipGlob.Heading)
        mclsAnchor.Xg = mclsFairLead.Xg(ShipGlob) + Scope * System.Math.Cos(Alpha)
        mclsAnchor.Yg = mclsFairLead.Yg(ShipGlob) + Scope * System.Math.Sin(Alpha)

    End Sub

    Public Sub GetTouchDownLocation(ByRef ShipGlob As ShipGlobal)
        Dim dx, Alpha, dy As Single

        '    Alpha = PI / 2# - (mclsFairLead.SprdAngle + ShipGlob.Heading)
        '    mTouchDownX = mclsFairLead.Xg(ShipGlob) + (msngScope - msngGrdLen) * Cos(Alpha)
        '    mTouchDownY = mclsFairLead.Yg(ShipGlob) + (msngScope - msngGrdLen) * Sin(Alpha)

        dx = mclsAnchor.Xg - mclsFairLead.Xg(ShipGlob)
        dy = mclsAnchor.Yg - mclsFairLead.Yg(ShipGlob)
        If dx = 0 Then
            mTouchDownX = mclsAnchor.Xg
            If dy > 0 Then
                mTouchDownY = mclsAnchor.Yg - msngGrdLen
            Else
                mTouchDownY = mclsAnchor.Yg + msngGrdLen
            End If
        Else
            Alpha = System.Math.Atan(dy / dx)
            If dx > 0 Then
                mTouchDownX = mclsAnchor.Xg - msngGrdLen * System.Math.Abs(System.Math.Cos(Alpha))
            Else
                mTouchDownX = mclsAnchor.Xg + msngGrdLen * System.Math.Abs(System.Math.Cos(Alpha))
            End If
            If dy > 0 Then
                mTouchDownY = mclsAnchor.Yg - msngGrdLen * System.Math.Abs(System.Math.Sin(Alpha))
            Else
                mTouchDownY = mclsAnchor.Yg + msngGrdLen * System.Math.Abs(System.Math.Sin(Alpha))
            End If
        End If

    End Sub

    ' calculate toptension based on scope and payout
    Public Function TensionByScopePOL(ByRef Scope As Single, ByRef FrcHor As Single, ByRef POL As Single) As Single

        ' Input
        '   Scope:      mooring line scope (horizontal distance)
        '   POL:        payout length of the 1st segment
        '   mooring line properties from this class
        ' Output
        '   FrcHor:     applied horizontal force
        '   return:     corresponding top tension (-1 - unsuccessfull calculation)

        Dim i As Short
        Dim ScopeC, t0 As Single
        Dim CatenyCalc As Boolean
        Dim ThTB, MinBS, ThTS As Single
        Dim ScpTS, ScpTB, ScpTol As Single
        Dim FstTB, ThFound, FstTS As Boolean
        Dim Segment_Renamed As Segment
        Dim CalSpring As Boolean

        '   Use a multiple of the breaking strength of the weakest segment of the
        '   mooring line to determine a limit for the iteration to determine the
        '   tension
        '
        '   The limits of the iteration are the most problematic part of the procedure.
        '   While it might be intuitively appealing to limit the tension to the strength
        '   of the weakest segment, the nature of the iterative procedure requires that
        '   more latitude be given during the computation

        mblnTBCalc = False
        ScpTol = Max(DistTol * Scope, 0.1)

        MinBS = 100000000000.0#
        For Each Segment_Renamed In mcolSegments
            With Segment_Renamed
                If .BS < MinBS Then MinBS = .BS
            End With
        Next Segment_Renamed

        ThTB = MinBS * 3.0#
        ThTS = 1.0#
        FrcHor = (ThTB + ThTS) / 2.0#
        t0 = (ThTB + ThTS) / 2.0# '2.2.1
        ThFound = False
        FstTB = True
        FstTS = True

        '   At times, even the relatively generous estimate for a maximum tension
        '   used above may fail.  In order to limit that possibility while preserving
        '   performance, we will try increasing the limits of the iteration as many as
        '   three times (j = 1 to 3) if convergence has not been achieved within a
        '   reasonable number of tries (20 in this case)

        '    DoEvents
        '    For j = 1 To 5
        '        For i = 1 To MaxIter / 5       ' julie remove this to make it faster -- quit at MaxIter if not converge
        Do While i < MaxIter
            If IsSpring Then '2.2.1
                CalSpring = Spring(POL, t0, FrcHor, ScopeC)
                If Not CalSpring Then Exit Do

                If System.Math.Abs(ScopeC - Scope) < ScpTol Then
                    ThFound = True
                    Exit Do
                ElseIf Scope < ScopeC Then
                    If System.Math.Abs(ScopeC - ScpTB) < 0.5 * ScpTol And System.Math.Abs(ScopeC - Scope) < 2.5 * ScpTol Then
                        ThFound = True
                        Exit Do
                    End If

                    ThTB = t0
                    ScpTB = ScopeC
                    If FstTB Then FstTB = False
                    If FstTB Or FstTS Then
                        t0 = (ThTB + ThTS) / 2.0#
                    Else
                        t0 = ThTS + (ThTB - ThTS) * (Scope - ScpTS) / (ScpTB - ScpTS)
                    End If
                Else
                    If System.Math.Abs(ScopeC - ScpTS) < 0.5 * ScpTol And System.Math.Abs(ScopeC - Scope) < 2.5 * ScpTol Then
                        ThFound = True
                        Exit Do
                    End If

                    ThTS = t0
                    ScpTS = ScopeC
                    If FstTS Then FstTS = False
                    If FstTB Or FstTS Then
                        t0 = (ThTB + ThTS) / 2.0#
                    Else
                        t0 = ThTS + (ThTB - ThTS) * (Scope - ScpTS) / (ScpTB - ScpTS)
                    End If
                End If

                If System.Math.Abs(ThTB - ThTS) < ForceTol Then
                    CalSpring = Spring(POL, t0, FrcHor, ScopeC)
                    ThFound = True
                    Exit Do
                End If
            Else
                CatenyCalc = Catenary(POL, FrcHor, t0, ScopeC, msngGrdLen, msngAnchPull)
                If Not CatenyCalc Then Exit Do

                If System.Math.Abs(ScopeC - Scope) < ScpTol Then
                    ThFound = True
                    Exit Do
                ElseIf Scope < ScopeC Then
                    If System.Math.Abs(ScopeC - ScpTB) < 0.5 * ScpTol And System.Math.Abs(ScopeC - Scope) < 2.5 * ScpTol Then
                        ThFound = True
                        Exit Do
                    End If

                    ThTB = FrcHor
                    ScpTB = ScopeC
                    If FstTB Then FstTB = False
                    If FstTB Or FstTS Then
                        FrcHor = (ThTB + ThTS) / 2.0#
                    Else
                        FrcHor = ThTS + (ThTB - ThTS) * (Scope - ScpTS) / (ScpTB - ScpTS)
                    End If
                Else
                    If System.Math.Abs(ScopeC - ScpTS) < 0.5 * ScpTol And System.Math.Abs(ScopeC - Scope) < 2.5 * ScpTol Then
                        ThFound = True
                        Exit Do
                    End If

                    ThTS = FrcHor
                    ScpTS = ScopeC
                    If FstTS Then FstTS = False
                    If FstTB Or FstTS Then
                        FrcHor = (ThTB + ThTS) / 2.0#
                    Else
                        FrcHor = ThTS + (ThTB - ThTS) * (Scope - ScpTS) / (ScpTB - ScpTS)
                    End If
                End If

                If System.Math.Abs(ThTB - ThTS) < ForceTol Then
                    CatenyCalc = Catenary(POL, FrcHor, t0, ScopeC, msngGrdLen, msngAnchPull)
                    ThFound = True
                    Exit Do
                End If
            End If
        Loop
        '        If ThFound Then
        '            Exit For
        '        Else
        '            ThTB = ThTB * 2#
        '            ThTS = ThTS / 2#
        '            FrcHor = (ThTB + ThTS) / 2#
        '        End If
        '    Next j

        '   Did we succeed?  If so, return the computed value; if not, send a value of -1
        '   to indicate the failure (tension should never be negative)
        If ThFound Then
            TensionByScopePOL = t0
        Else
            TensionByScopePOL = -1.0#
        End If

    End Function

    ' if ship location and orientation are provided, calculate actual scope
    ' based on anchor location
    Public Function ScopeByVesselLocation(ByRef ShipGlob As ShipGlobal, Optional ByRef Moved As Boolean = False) As Single

        Dim X, Y As Single

        If mclsFairLead.NeedUpdate Or mclsAnchor.NeedUpdate Or Moved Then
            msngScope = System.Math.Sqrt((mclsAnchor.Xg - mclsFairLead.Xg(ShipGlob)) ^ 2 + (mclsAnchor.Yg - mclsFairLead.Yg(ShipGlob)) ^ 2)

            X = mclsAnchor.Xg - mclsFairLead.Xg(ShipGlob)
            Y = mclsAnchor.Yg - mclsFairLead.Yg(ShipGlob)
            msngSprdAngle = Atan(Y, X) - ShipGlob.Heading
            If msngSprdAngle < 0# Then msngSprdAngle = 2.0# * PI + msngSprdAngle

            mclsFairLead.NeedUpdate = False
            mclsAnchor.NeedUpdate = False
        End If

        ScopeByVesselLocation = msngScope
        Call GetTouchDownLocation(ShipGlob)

    End Function

    Public Function TensionByVesselLocation(ByRef ShipGlob As ShipGlobal) As Single

        Dim Scope, HorTension As Single

        If Not mblnConnected Then
            TensionByVesselLocation = 0#
            Exit Function
        End If

        Scope = System.Math.Sqrt((mclsAnchor.Xg - mclsFairLead.Xg(ShipGlob)) ^ 2 + (mclsAnchor.Yg - mclsFairLead.Yg(ShipGlob)) ^ 2)

        TensionByVesselLocation = TensionByScopePOL(Scope, HorTension, msngPayout)

    End Function

    ' calculate scope based on top tension and payout
    Public Function ScopeByTopTensionPOL(ByRef TopTen As Single, ByRef FrcHor As Single, ByRef POL As Single) As Single

        ' Input
        '   TopTen:     top tension
        '   POL:        payout length of the 1st segment
        '   mooring line properties from this class
        ' Output
        '   FrcHor:     applied horizontal force
        '   return:     corresponding scope

        Dim i As Short
        Dim CatenyCalc As Boolean
        Dim ScopeC, t0 As Single
        Dim T1, T2 As Single
        Dim ThFound As Boolean
        Dim CalSpring As Boolean

        mblnTBCalc = False

        T1 = TopTen
        T2 = 1.0#
        ThFound = False
        i = 0
        Do While i < MaxIter
            If IsSpring Then '2.2.1
                CalSpring = Spring(POL, TopTen, FrcHor, ScopeC)
                If CalSpring Then ThFound = True
                Exit Do
            End If

            i = i + 1
            FrcHor = (T1 + T2) / 2 ' set start trial value of Top Horizontal Tension
            '   Input
            '   FrcHor:     applied horizontal force
            '   POL:        payout length of the 1st segment
            CatenyCalc = Catenary(POL, FrcHor, t0, ScopeC, msngGrdLen, msngAnchPull)
            If Not CatenyCalc Then Exit Do
            '       Debug.Print "t0, FrcHor= " & t0, FrcHor

            If System.Math.Abs(t0 - TopTen) < ForceTol Then
                ThFound = True
                Exit Do
            End If
            If t0 > TopTen Then
                T1 = FrcHor
            Else
                T2 = FrcHor
            End If

            If System.Math.Abs(T1 - T2) < ForceTol Then
                ThFound = False
                Exit Do
            End If
        Loop

        '   Did we succeed?  If so, return the computed value; if not, send a value of -1
        '   to indicate the failure (scope should never be negative)
        If ThFound Then
            ScopeByTopTensionPOL = ScopeC
        Else
            ScopeByTopTensionPOL = -1.0#
            '   Debug.Print "returned -1: Last Found Scope, FrcHor = " & ScopeC, FrcHor
        End If

    End Function

    ' calculate scope based on top tension and payout - replaced by Julie with ScopeByTopTensionPOL
    Public Function ScopeByTensionPOL_A(ByRef TopTen As Single, ByRef FrcHor As Single, ByRef POL As Single) As Single

        ' Input
        '   TopTen:     top tension
        '   POL:        payout length of the 1st segment
        '   mooring line properties from this class
        ' Output
        '   FrcHor:     applied horizontal force
        '   return:     corresponding scope

        Dim i As Short
        Dim CatenyCalc As Boolean
        Dim ScopeC, t0 As Single
        Dim ThTB, ThTS As Single
        Dim T0TB, T0TS As Single
        Dim FstTB, ThFound, FstTS As Boolean

        mblnTBCalc = False

        ThTB = TopTen
        ThTS = 1.0#
        FrcHor = (ThTB + ThTS) / 2.0#
        ThFound = False
        FstTB = True
        FstTS = True

        System.Windows.Forms.Application.DoEvents()
        For i = 1 To MaxIter
            CatenyCalc = Catenary(POL, FrcHor, t0, ScopeC, msngGrdLen, msngAnchPull)
            If Not CatenyCalc Then Exit For

            If System.Math.Abs(t0 - TopTen) < ForceTol Then
                ThFound = True
                Exit For
            ElseIf t0 > TopTen Then
                ThTB = FrcHor
                T0TB = t0
                If FstTB Then FstTB = False
                If FstTB Or FstTS Then
                    FrcHor = (ThTB + ThTS) / 2.0#
                Else
                    FrcHor = ThTS + (ThTB - ThTS) * (TopTen - T0TS) / (T0TB - T0TS)
                End If
            Else
                ThTS = FrcHor
                T0TS = t0
                If FstTS Then FstTS = False
                If FstTB Or FstTS Then
                    FrcHor = (ThTB + ThTS) / 2.0#
                Else
                    FrcHor = ThTS + (ThTB - ThTS) * (TopTen - T0TS) / (T0TB - T0TS)
                End If
            End If

            '       This strategy can get "stuck"; make sure this doesn't happen by looking for
            '       convergence of AlphaTB and AlphaTS
            If System.Math.Abs(ThTB - ThTS) < 2 * ForceTol Then
                CatenyCalc = Catenary(POL, FrcHor, t0, ScopeC, msngGrdLen, msngAnchPull)
                ThFound = True
                Exit For
            End If
        Next i

        '   Did we succeed?  If so, return the computed value; if not, send a value of -1
        '   to indicate the failure (scope should never be negative)
        If ThFound Then
            ScopeByTensionPOL_A = ScopeC
        Else
            ScopeByTensionPOL_A = -1.0#
        End If

    End Function

    ' calculate scope based on horiztonal force and payout
    Public Function ScopeByFrcHorPOL(ByRef FrcHor As Single, ByRef TopTen As Single, ByRef POL As Single) As Single

        ' Input
        '   FrcHor:     applied horizontal force
        '   POL:        payout length of the 1st segment
        '   mooring line properties from this class
        ' Output
        '   TopTen:     top tension
        '   return:     corresponding scope

        Dim CatenyCalc, CalSpring As Boolean
        Dim t0, ScopeC, f0 As Single
        Dim T1, T2 As Single
        Dim ThFound As Boolean
        Dim i As Short

        mblnTBCalc = False

        System.Windows.Forms.Application.DoEvents()

        T1 = FrcHor
        T2 = 1.0#
        ThFound = False
        i = 0
        Do While i < MaxIter
            If Not IsSpring Then '2.2.1
                CatenyCalc = Catenary(POL, FrcHor, t0, ScopeC, msngGrdLen, msngAnchPull)
                If CatenyCalc Then ThFound = True
                Exit Do
            End If

            i = i + 1
            t0 = (T1 + T2) / 2 ' set start trial value of Top Horizontal Tension
            CalSpring = Spring(POL, t0, f0, ScopeC)
            If Not CalSpring Then Exit Do

            If System.Math.Abs(f0 - FrcHor) < ForceTol Then
                ThFound = True
                Exit Do
            End If
            If f0 > FrcHor Then
                T1 = t0
            Else
                T2 = t0
            End If

            If System.Math.Abs(T1 - T2) < ForceTol Then
                ThFound = False
                Exit Do
            End If
        Loop

        If Not ThFound Then
            ScopeByFrcHorPOL = -1.0#
        Else
            TopTen = t0
            ScopeByFrcHorPOL = ScopeC
        End If

    End Function

    ' calculate payout based on scope and top tension
    Public Function PayoutByScopeTopTension(ByRef Scope As Single, ByRef TopTen As Single) As Single

        ' Input
        '   Scope:      scope
        '   TopTen:     top tension
        '   mooring line properties from this class
        ' Output
        '   return:     required payout

        Dim i As Short
        Dim POL, POT1, POT2, POTol As Single
        Dim ScpTol As Single
        Dim ScopeC, FrcHor As Single
        Dim FstTB, POFound, FstTS As Boolean
        Dim CalSpring As Boolean

        mblnTBCalc = False
        ScpTol = Max(DistTol * Scope, 0.1)

        POT1 = 0#
        POT2 = mcolSegments.Item(1).TotalLength '  top segment length

        POTol = Max(DistTol * POT1, 0.1)
        POL = (POT2 + POT1) / 2 ' starting trial value of Payout

        POFound = False
        FstTB = False
        FstTS = True

        Do While i < MaxIter
            POL = (POT2 + POT1) / 2 ' starting trial value of Payout

            If IsSpring Then '2.2.1
                CalSpring = Spring(POL, TopTen, FrcHor, ScopeC)
                If Not CalSpring Then Exit Do
            Else
                ScopeC = ScopeByTopTensionPOL(TopTen, FrcHor, POL) ' Hor Force is output
            End If

            If System.Math.Abs(ScopeC - Scope) < ScpTol Then '  check if scope found
                POFound = True
                Exit Do
            End If
            If ScopeC < 0 Then ' catenary calc failed
                ' i.e.  ScopeC=0
                POT1 = 0
                POT2 = POL / 2
            ElseIf ScopeC < Scope Then
                POT1 = POL ' increase POL
            Else
                POT2 = POL ' decrease POL
            End If
            If System.Math.Abs(POT2 - POT1) < POTol Then
                POFound = False
                Exit Do
            End If
        Loop
        '   Did we succeed?  If so, return the computed value; if not, send a value of -1
        '   to indicate the failure (payout should never be negative)
        If POFound Then
            PayoutByScopeTopTension = POL
        Else
            PayoutByScopeTopTension = -1.0#
            '   Debug.Print "returned -1: Last found POL = " & POL
        End If

    End Function

    ' calculate payout based on scope and top tension- Deleted by Julie Guan replaced by PayoutByScopeTopTension for better convergence
    Public Function POLByScopeTension(ByRef Scope As Single, ByRef TopTen As Single) As Single

        ' Input
        '   Scope:      scope
        '   TopTen:     top tension
        '   mooring line properties from this class
        ' Output
        '   return:     required payout

        Dim i As Short
        Dim POL, POTB, POTS, POTol As Single
        Dim ScpTS, ScpTB, ScpTol As Single
        Dim ScopeC, FrcHor As Single
        Dim FstTB, POFound, FstTS As Boolean

        mblnTBCalc = False
        ScpTol = Max(DistTol * Scope, 0.1)

        POTS = 0#
        POTB = mcolSegments.Item(1).Length '  top segment length

        '    ScpTB = ScopeByTensionPOL(TopTen, FrcHor, POTB)
        '    If ScpTB < Scope + ScpTol Then
        '        POLByScopeTension = -POTB
        '        Exit Function
        '    End If

        POTol = Max(DistTol * POTB, 0.1)
        '    If msngPayout <> 0# And msngPayout <> POTB Then
        '        POL = msngPayout
        '    Else
        POL = (POTS + POTB) / 2 ' starting trial value of Payout
        '    End If
        POFound = False
        FstTB = False
        FstTS = True

        '   DoEvents
        For i = 1 To MaxIter
            ScopeC = ScopeByTensionPOL_A(TopTen, FrcHor, POL) ' Hor Force is output
            If ScopeC < 0# Then
                If FstTB Then
                    POTB = POL
                ElseIf FstTS Then
                    POTS = POL
                Else
                    Exit For
                End If
                POL = (POTB + POTS) / 2

            ElseIf System.Math.Abs(ScopeC - Scope) < ScpTol Then
                POFound = True
                Exit For

            ElseIf Scope < ScopeC Then
                If POL = 0# Then Exit For
                If System.Math.Abs(ScopeC - ScpTB) < 0.5 * ScpTol And System.Math.Abs(ScopeC - Scope) < 2.5 * ScpTol Then
                    POFound = True
                    Exit For
                End If

                POTB = POL
                ScpTB = ScopeC
                If FstTB Then FstTB = False
                If FstTB Or FstTS Then
                    POL = (POTB + POTS) / 2.0#
                Else
                    POL = POTS + (POTB - POTS) * (Scope - ScpTS) / (ScpTB - ScpTS)
                End If
            Else
                If POL > mcolSegments.Item(1).Length Then Exit For
                If System.Math.Abs(ScopeC - ScpTS) < 0.5 * ScpTol And System.Math.Abs(ScopeC - Scope) < 2.5 * ScpTol Then
                    POFound = True
                    Exit For
                End If

                POTS = POL
                ScpTS = ScopeC
                If FstTS Then FstTS = False
                If FstTB Or FstTS Then
                    POL = (POTB + POTS) / 2.0#
                Else
                    POL = POTS + (POTB - POTS) * (Scope - ScpTS) / (ScpTB - ScpTS)
                End If
            End If

            If System.Math.Abs(POTB - POTS) < POTol Then
                ScopeC = ScopeByTensionPOL_A(TopTen, FrcHor, POL)
                POFound = False
                Exit For
            End If
        Next i

        '   Did we succeed?  If so, return the computed value; if not, send a value of -1
        '   to indicate the failure (payout should never be negative)
        If POFound Then
            POLByScopeTension = POL
        Else
            POLByScopeTension = -1.0#
        End If

    End Function

    ' calculate payout by scope and horzontal force
    Public Function POLByScopeFrcHor(ByRef Scope As Single, ByRef FrcHor As Single) As Single

        ' Input
        '   Scope:      scope
        '   FrcHor:     horizontal force
        '   mooring line properties from this class
        ' Output
        '   return:     required scope

        Dim i As Short
        Dim CatenyCalc, CalSpring As Boolean
        Dim POL, POTB, POTS, POTol As Single
        Dim ScpTS, ScpTB, ScpTol As Single
        Dim t0, ScopeC, f0 As Single
        Dim FstTB, POFound, FstTS As Boolean

        mblnTBCalc = False
        ScpTol = Max(DistTol * Scope, 0.1)

        POTS = 0#
        POTB = mcolSegments.Item(1).Length

        If IsSpring Then '2.2.1
            t0 = FrcHor
            CalSpring = Spring(POTB, t0, f0, ScpTB)
        Else
            CatenyCalc = Catenary(POTB, FrcHor, t0, ScpTB, msngGrdLen, msngAnchPull)
        End If
        If ScpTB < Scope Then
            POLByScopeFrcHor = -POTB
            Exit Function
        End If

        POTol = Max(DistTol * POTB, 0.1)
        If msngPayout <> 0# And msngPayout <> POTB Then
            POL = msngPayout
        Else
            POL = (POTS + POTB) * 0.5
        End If

        POFound = False
        FstTB = False
        FstTS = True

        System.Windows.Forms.Application.DoEvents()
        For i = 1 To MaxIter
            ScopeC = ScopeByFrcHorPOL(FrcHor, t0, POL) '2.2.1
            '        CatenyCalc = Catenary(POL, FrcHor, t0, ScopeC, msngGrdLen, msngAnchPull)
            If ScopeC < 0# Then Exit For

            If System.Math.Abs(ScopeC - Scope) < ScpTol Then
                POFound = True
                Exit For
            ElseIf Scope < ScopeC Then
                If POL = 0# Then Exit For
                If System.Math.Abs(ScopeC - ScpTB) < 0.5 * ScpTol And System.Math.Abs(ScopeC - Scope) < 2.5 * ScpTol Then
                    POFound = True
                    Exit For
                End If

                POTB = POL
                ScpTB = ScopeC
                If FstTB Then FstTB = False
                If FstTB Or FstTS Then
                    POL = (POTB + POTS) / 2.0#
                Else
                    POL = POTS + (POTB - POTS) * (Scope - ScpTS) / (ScpTB - ScpTS)
                End If
            Else
                If POL > mcolSegments.Item(1).Length Then Exit For
                If System.Math.Abs(ScopeC - ScpTS) < 0.5 * ScpTol And System.Math.Abs(ScopeC - Scope) < 2.5 * ScpTol Then
                    POFound = True
                    Exit For
                End If

                POTS = POL
                ScpTS = ScopeC
                If FstTS Then FstTS = False
                If FstTB Or FstTS Then
                    POL = (POTB + POTS) / 2.0#
                Else
                    POL = POTS + (POTB - POTS) * (Scope - ScpTS) / (ScpTB - ScpTS)
                End If
            End If

            If System.Math.Abs(POTB - POTS) < POTol Then
                ScopeC = ScopeByFrcHorPOL(FrcHor, t0, POL) '2.2.1
                '            CatenyCalc = Catenary(POL, FrcHor, t0, ScopeC, msngGrdLen, msngAnchPull)
                POFound = False
                Exit For
            End If
        Next i

        '   Did we succeed?  If so, return the computed value; if not, send a value of -1
        '   to indicate the failure (payout should never be negative)
        If POFound Then
            POLByScopeFrcHor = POL
        Else
            POLByScopeFrcHor = -1.0#
        End If

    End Function

    ' generate plot points
    Public Function CatenaryPoints(ByRef CatX() As Single, ByRef CatY() As Single, ByRef Connector() As Short) As Boolean
        ' Connector saves the array index of catX or CatY at which a connector exists

        Dim NumSeg, i, j, NumSubSeg As Short

        Dim TotalLenA, TotalLenB As Single
        Dim SegLenStr, SegLength, SubSegLen As Single
        Dim TopAng, TopHorTen, theta As Single
        Dim SegUnitWgt, TSubSegWgt As Single

        CatenaryPoints = False

        NumSeg = mcolSegments.Count()

        If NumSeg = 0 Then Exit Function

        TotalLenA = 0#
        TotalLenB = 0#
        Connector(NumSeg + 1) = 1

        ' starting from Anchor
        CatX(1) = mcolSegments.Item(NumSeg).XLow
        CatY(1) = mcolSegments.Item(NumSeg).YLow

        For i = NumSeg To 1 Step -1
            If i = 1 Then
                SegLength = msngPayout
            Else
                SegLength = mcolSegments.Item(i).Length
            End If
            SegLenStr = mcolSegments.Item(i).LengthStr

            TotalLenA = TotalLenB + SegLenStr
            If SegLength = 0# Or SegLenStr = 0# Then
                Connector(i) = Connector(i + 1)

            ElseIf IsSpring Then
                Connector(i) = Connector(i + 1) + 1

                If i = 1 Then
                    CatX(Connector(i)) = 0#
                    CatY(Connector(i)) = msngDraft - mclsFairLead.z
                Else
                    CatX(Connector(i)) = mcolSegments.Item(i - 1).XLow
                    CatY(Connector(i)) = mcolSegments.Item(i - 1).YLow
                    '  If CatX(1) - CatX(Connector(i)) < msngGrdLen Then CatY(Connector(i)) = CatY(1)                   ' make sure lowest is on the ground
                End If

            ElseIf TotalLenA < msngGrdLen Then
                If i = 1 Then Exit Function
                Connector(i) = Connector(i + 1) + 1

                CatX(Connector(i)) = mcolSegments.Item(i - 1).XLow
                CatY(Connector(i)) = mcolSegments.Item(i - 1).YLow
                '   If CatX(1) - CatX(Connector(i)) < msngGrdLen Then CatY(Connector(i)) = CatY(1)                   ' make sure lowest is on the ground

            Else
                TopAng = Segments(i).AngUpp
                TopHorTen = Segments(i).TenUpp * System.Math.Cos(TopAng)
                SegUnitWgt = Segments(i).UnitWetWeight * SegLength / SegLenStr

                If TotalLenB < msngGrdLen Then
                    NumSubSeg = (TotalLenA - msngGrdLen) / MinSubSegLen
                    '                NumSubSeg = (TotalLenA + TotalLenB - msngGrdLen) / MinSubSegLen
                    If NumSubSeg > MaxNumSubSeg - 1 Then
                        NumSubSeg = MaxNumSubSeg - 1
                    ElseIf NumSubSeg < 2 Then
                        NumSubSeg = 2
                    End If
                    SubSegLen = (TotalLenA - msngGrdLen) / NumSubSeg
                    '                SubSegLen = (TotalLenA + TotalLenB - msngGrdLen) / NumSubSeg
                    NumSubSeg = NumSubSeg + 1
                Else
                    NumSubSeg = SegLenStr / MinSubSegLen
                    If NumSubSeg > MaxNumSubSeg Then
                        NumSubSeg = MaxNumSubSeg
                    ElseIf NumSubSeg < 2 Then
                        NumSubSeg = 2
                    End If
                    SubSegLen = SegLenStr / NumSubSeg
                End If
                Connector(i) = Connector(i + 1) + NumSubSeg

                If i = 1 Then
                    CatX(Connector(i)) = 0#
                    CatY(Connector(i)) = msngDraft - mclsFairLead.z
                Else
                    CatX(Connector(i)) = mcolSegments.Item(i - 1).XLow
                    CatY(Connector(i)) = mcolSegments.Item(i - 1).YLow
                    '   If CatX(1) - CatX(Connector(i)) < msngGrdLen Then CatY(Connector(i)) = CatY(1)                   ' make sure lowest is on the ground
                End If

                For j = Connector(i) - 1 To Connector(i + 1) + 1 Step -1
                    TSubSegWgt = (Connector(i) - j) * SubSegLen * SegUnitWgt
                    If TopHorTen = 0 Then Exit Function
                    theta = System.Math.Atan(System.Math.Tan(TopAng) - TSubSegWgt / TopHorTen)
                    On Error Resume Next
                    CatX(j) = CatX(Connector(i)) + TopHorTen / SegUnitWgt * System.Math.Log(System.Math.Tan(TopAng / 2.0# + PI / 4.0#) / System.Math.Tan(theta / 2.0# + PI / 4.0#))
                    CatY(j) = CatY(Connector(i)) + TopHorTen / SegUnitWgt * (1.0# / System.Math.Cos(TopAng) - 1.0# / System.Math.Cos(theta))
                    '   If CatX(1) - CatX(j) < msngGrdLen Then CatY(j) = CatY(1)                   ' make sure lowest is on the ground
                Next j
            End If
            TotalLenB = TotalLenA
        Next i

        CatenaryPoints = True

    End Function

    ' calculate forces on spring line   2.2.1
    Private Function Spring(ByRef POL As Single, ByRef TopTen As Single, ByRef FrcHor As Single, ByRef ScopeC As Single) As Boolean

        ' Input
        '   POL:        payout of top segment
        '   TopTen:     top tension
        '   mooring line properties from this class
        ' Output
        '   FrcHor:     horizontal force
        '   ScopeC:     scope
        '   return:     success

        Dim i, NumSeg As Short
        Dim TotalLen, SegLen0 As Single
        Dim tmp2, tmp1, angle As Single
        Dim DepthFL As Single

        Spring = False
        If Not IsSpring Then Exit Function

        NumSeg = mcolSegments.Count()
        DepthFL = mclsFairLead.z - msngDraft + msngWaterDepth

        With mcolSegments
            TotalLen = 0#
            For i = 1 To NumSeg
                With .Item(i)
                    If i = 1 Then
                        SegLen0 = POL
                        If LineLen = 0# Then SegLen0 = 0.01
                    Else
                        'UPGRADE_WARNING: Couldn't resolve default property of object mcolSegments.Item().Length. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        SegLen0 = .Length
                    End If

                    If TopTen < 0# Then TopTen = 0#
                    tmp1 = TopTen / (.E1 * .XArea)
                    tmp2 = tmp1 * 2.0# / (System.Math.Sqrt(1.0# + 4.0# * tmp1 * .E2 / .E1) + 1.0#)
                    .LengthStr = SegLen0 * (1 + tmp2)
                    TotalLen = TotalLen + .LengthStr
                End With
            Next i

            ScopeC = System.Math.Sqrt(TotalLen ^ 2 - DepthFL ^ 2)
            FrcHor = TopTen * ScopeC / TotalLen
            angle = Acos(ScopeC / TotalLen)

            For i = NumSeg To 1
                If i = NumSeg Then
                    .Item(i).XLow = ScopeC
                    .Item(i).YLow = msngWaterDepth
                Else
                    .Item(i).XLow = .Item(i + 1).XLow - .Item(i + 1).LengthStr / TotalLen * ScopeC
                    .Item(i).YLow = .Item(i + 1).YLow - .Item(i + 1).LengthStr / TotalLen * DepthFL
                End If
                .Item(i).AngUpp = angle
                .Item(i).AngLow = angle
                .Item(i).TenUpp = TopTen
            Next i
        End With

        Spring = True

    End Function

    ' calculate the catenary
    Private Function Catenary(ByRef POL As Single, ByRef FrcHor As Single, ByRef TopTen As Single, ByRef ScopeC As Single, ByRef GrdLen As Single, ByRef AnchPull As Single) As Boolean

        ' Input
        '   FrcHor:     applied horizontal force
        '   POL:        payout length of the 1st segment
        '   mooring line properties from this class
        ' Output
        '   TopTen:     corresponding top tension
        '   ScopeC:     calculated scope
        '   GrdLen:     grounded mooring line length
        '   AnchPull:   pull at anchor
        '   return:     flag successful calculation
        ' Output to moorline properties
        '   SegLen(i):  streched segment length
        '   Xi(i):      lower end scope
        '   Yi(i):      lower end depth
        '   ThU(i):     upper end angle of ith segment
        '   ThL(i):     lower end angle of ith segment

        Dim i As Short

        Dim NumSeg, NumBuoy As Short
        Dim SegLen0(MaxNumSeg) As Single
        Dim SegLen(MaxNumSeg) As Single
        Dim SegXA(MaxNumSeg) As Single
        Dim SegWgt(MaxNumSeg) As Single
        Dim SegE1(MaxNumSeg) As Single
        Dim SegE2(MaxNumSeg) As Single
        Dim SegFrc(MaxNumSeg) As Single
        Dim SegBuoy0(MaxNumSeg) As Single
        Dim BuoyLen(MaxNumSeg) As Single

        Dim DepthFL As Single
        Dim Xi(MaxNumSeg) As Single
        Dim Yi(MaxNumSeg) As Single
        Dim ThU(MaxNumSeg) As Single
        Dim ThL(MaxNumSeg) As Single
        Dim TenU(MaxNumSeg) As Single

        NumSeg = mcolSegments.Count()
        DepthFL = mclsFairLead.z - msngDraft + msngWaterDepth

        NumBuoy = 0
        For i = 1 To NumSeg
            With mcolSegments.Item(i)
                If i = 1 Then
                    SegLen0(i) = POL
                    SegWgt(i) = POL / .Length * .TotalWeight
                Else
                    SegLen0(i) = .Length
                    SegWgt(i) = .TotalWeight
                End If
                SegFrc(i) = .FrictionCoef
                SegXA(i) = .XArea
                SegE1(i) = .E1
                SegE2(i) = .E2

                SegBuoy0(i) = .Buoy
                BuoyLen(i) = .BuoyLength
                If System.Math.Abs(SegBuoy0(i)) > 0.00001 Then NumBuoy = NumBuoy + 1
            End With
        Next i

        If NumBuoy = 0 Then
            Catenary = CatNoBuoy(FrcHor, DepthFL, NumSeg, SegLen0, SegLen, SegWgt, SegFrc, SegXA, SegE1, SegE2, TopTen, Xi, Yi, ThU, ThL, TenU, GrdLen, AnchPull)
        Else
            Catenary = CatWBuoy(FrcHor, DepthFL, NumSeg, SegLen0, SegLen, SegWgt, SegFrc, SegXA, SegE1, SegE2, SegBuoy0, BuoyLen, TopTen, Xi, Yi, ThU, ThL, TenU, GrdLen, AnchPull)
        End If
        ScopeC = Xi(NumSeg)

        For i = 1 To NumSeg
            With mcolSegments.Item(i)
                .LengthStr = SegLen(i)
                .XLow = Xi(i)
                .YLow = Yi(i) + msngDraft - mclsFairLead.z
                .AngUpp = ThU(i)
                .AngLow = ThL(i)
                .TenUpp = TenU(i)
            End With
        Next i

    End Function

    ' calculate the catenary (w/o buoys)
    Private Function CatNoBuoy(ByRef FrcHor As Single, ByRef DepthFL As Single, ByRef NumSeg As Short, ByRef SegLen0() As Single, ByRef SegLen() As Single, ByRef SegWgt() As Single, ByRef SegFrc() As Single, ByRef SegXA() As Single, ByRef SegE1() As Single, ByRef SegE2() As Single, ByRef TopTen As Single, ByRef Xi() As Single, ByRef Yi() As Single, ByRef ThU() As Single, ByRef ThL() As Single, ByRef TenU() As Single, ByRef GrdLen As Single, ByRef AnchPull As Single) As Boolean

        ' Input
        '   FrcHor:     applied horizontal force
        '   DepthFL:    fairlead depth (to anchor)
        '   mooring line properties
        ' Output
        '   TopTen:     corresponding top tension
        '   SegLen(i):  stretched segment length
        '   Xi(i):      lower end scope
        '   Yi(i):      lower end depth
        '   ThU(i):     upper end angle of ith segment
        '   ThL(i):     lower end angle of ith segment
        '   TenU(i):    upper end tension
        '   GrdLen:     grounded mooring line length
        '   AnchPull:   pull at anchor
        '   return:     flag successful calculation

        Dim icount, i, j As Short
        Dim AdjT0 As Boolean

        Dim ThUk1, ThUk2 As Single

        Dim TotSus, totlen, SegSus As Single

        Dim Ten, ten0 As Single
        Dim tmp2, tmp1, tmp3 As Single
        Dim T00, T0min, Dt0, T01 As Single

        Dim Dp1, DepthC, Dp0, DpTol As Single
        Dim ScopeC As Single

        Dim ipull As Boolean
        Dim nbot As Short

        '   initialize
        CatNoBuoy = False
        DpTol = Min(DepthTol * msngWaterDepth, 2.0#)

        totlen = 0#
        For i = 1 To NumSeg
            SegLen(i) = SegLen0(i)
            totlen = totlen + SegLen0(i)
        Next i

        '   1st run TouchBtm
        If Not mblnTBCalc Then mblnTBCalc = TouchBtm(DepthFL, NumSeg, SegLen0, SegWgt, SegFrc, SegXA, SegE1, SegE2, msngFh0k, msngThU0k, msngThL0k)
        '   if touchbtm cannot be calculated properly
        If Not mblnTBCalc Then Exit Function

        '   very small FrcHor
        If FrcHor <= 1.0# Then
            TopTen = 0#
            DepthC = 0#
            For i = 1 To NumSeg
                DepthC = DepthC + SegLen(i)
                If DepthC < DepthFL Then
                    TopTen = TopTen + SegWgt(i)
                    Xi(i) = 0#
                    Yi(i) = DepthC
                Else
                    TopTen = TopTen + SegWgt(i) * (DepthFL - DepthC) / SegLen(i)
                    Xi(i) = DepthFL - DepthC
                    Yi(i) = DepthFL
                    nbot = i
                    Exit For
                End If
            Next

            TenU(1) = TopTen
            For i = 2 To NumSeg
                TenU(i) = Max(TenU(i - 1) - SegWgt(i - 1), 0#)
            Next

            For i = nbot + 1 To NumSeg Step 1
                Xi(i) = Xi(i - 1) + SegLen(i)
                Yi(i) = Yi(i - 1)
            Next

            ScopeC = Xi(NumSeg)
            DepthC = Yi(NumSeg)
            AnchPull = 0#

            CatNoBuoy = True
            Exit Function
        End If

        '   Check and mark which segment soft-touches bottom and initiate TopTen

        '   Check COND(1), No vertical anchor pull (IPULL=0).
        '   Check COND(2), Vertical anchor pull not zero (IPULL=1).
        '   NBOT is the segment number of the first segment
        '   that touches the bottom.

        tmp1 = DepthFL / totlen
        '   if water deeper than mooring line length
        If tmp1 > 1.0# Then Exit Function
        TopTen = FrcHor / System.Math.Cos(0.6 * Asin(tmp1) + 0.4 * PI / 2)

        If FrcHor <= msngFh0k(1) Then
            ipull = False
            nbot = 1
            ThUk1 = PI / 2.0#
            ThUk2 = msngThU0k(1, 1)
            TopTen = Max(TopTen, FrcHor / System.Math.Cos(ThUk2))
        ElseIf FrcHor > msngFh0k(NumSeg) Then
            ipull = True
            nbot = NumSeg
            ThUk1 = msngThU0k(1, NumSeg)
            ThUk2 = Asin(tmp1)
            TopTen = Min(TopTen, FrcHor / System.Math.Cos(ThUk1))
        Else
            For i = 1 To NumSeg - 1
                If (FrcHor > msngFh0k(i) And FrcHor <= msngFh0k(i + 1)) Then
                    ipull = False
                    nbot = i + 1
                    ThUk1 = msngThU0k(1, i)
                    ThUk2 = msngThU0k(1, i + 1)
                    If msngFh0k(i) = 0# Then
                        TopTen = Max(TopTen, FrcHor / System.Math.Cos(ThUk2))
                    Else
                        tmp1 = FrcHor / System.Math.Cos(ThUk1)
                        tmp2 = FrcHor / System.Math.Cos(ThUk2)
                        TopTen = tmp1 + (tmp2 - tmp1) * (FrcHor - msngFh0k(i)) / (msngFh0k(i + 1) - msngFh0k(i))
                    End If
                    Exit For
                End If
            Next i
        End If

        Dt0 = 0.05 * TopTen
        If Dt0 < 1000.0# Then Dt0 = 1000.0#

        For icount = 1 To MaxIter * 5
            ThU(1) = Acos(FrcHor / TopTen)

            '       mooring line angles
            '       Loop for suspended segments
            Do
                AdjT0 = False
                For i = 1 To nbot - 1
                    ThL(i) = System.Math.Atan(System.Math.Tan(ThU(i)) - SegWgt(i) / FrcHor)
                    ThU(i + 1) = ThL(i)
                Next i
                If (ThL(nbot - 1) <= msngBottomSlope And SegWgt(nbot - 1) > 0#) Then
                    If Dt0 >= 0# Then
                        TopTen = TopTen + Dt0 * 0.5
                        Dt0 = Dt0 * 1.5
                    Else
                        TopTen = TopTen - Dt0 * 0.5
                        Dt0 = Dt0 * 0.5
                    End If
                    ThU(1) = Acos(FrcHor / TopTen)
                    AdjT0 = True
                End If
            Loop While AdjT0

            '       1st bottomed segment
            If ipull Then
                '       COND(2): Solve for TopTen and SK when vertical pull on anchor
                '                is not zero (IPULL=1).
                ThL(nbot) = System.Math.Atan(System.Math.Tan(ThU(nbot)) - SegWgt(nbot) / FrcHor)
            Else
                '       COND(1): Solve for TopTen and SK when Kth segment soft-touches
                '                bottom, and no vertical pull on anchor (IPULL=0).
                ThL(nbot) = msngBottomSlope
            End If

            '       Line segments on seabed.
            If (nbot < NumSeg) Then
                For i = nbot + 1 To NumSeg
                    ThU(i) = msngBottomSlope
                    ThL(i) = msngBottomSlope
                Next i
            End If

            '       Figure out the actual segment lengths for each and every segment
            '       accounting for line elasticity / stretch.
            '       A 10-point integration was used for stretch calculation.
            '       This part of the code does not change if a portion of the segment
            '       touches the bottom (seabed).
            ten0 = TopTen
            TenU(1) = TopTen
            totlen = 0#
            TotSus = 0#
            If Not ipull Then
                SegSus = FrcHor * SegLen(nbot) / SegWgt(nbot) * (System.Math.Tan(ThU(nbot)) - System.Math.Tan(ThL(nbot)))
                If SegSus > SegLen(nbot) Then SegSus = SegLen(nbot)
            End If
            For i = 1 To NumSeg
                If (SegE1(i) > 0#) Then
                    tmp3 = 0#
                    For j = 1 To NumSubSeg
                        If ipull Then
                            'there is no grounded portion
                            tmp1 = System.Math.Atan(System.Math.Tan(ThU(i)) - SegWgt(i) * j / NumSubSeg / FrcHor)
                            Ten = FrcHor / System.Math.Cos(tmp1)
                        Else
                            'there is grounded portion
                            If i = nbot Then 'this is 1st grounded segment
                                If j * SegLen(i) / NumSubSeg < SegSus Then
                                    'suspended portion
                                    tmp1 = System.Math.Atan(System.Math.Tan(ThU(i)) - SegWgt(i) * j / NumSubSeg / FrcHor)
                                    Ten = FrcHor / System.Math.Cos(tmp1)
                                Else
                                    If (j - 1) * SegLen(i) / NumSubSeg < SegSus Then
                                        '1st grounded portion
                                        tmp1 = ThL(i)
                                        Ten = FrcHor / System.Math.Cos(tmp1) - SegWgt(i) * (j / NumSubSeg - SegSus / SegLen(i)) * (System.Math.Sin(tmp1) + SegFrc(i) * System.Math.Cos(tmp1))
                                    Else
                                        'grounded portion
                                        tmp1 = ThL(i)
                                        Ten = ten0 - SegWgt(i) / NumSubSeg * (System.Math.Sin(tmp1) + SegFrc(i) * System.Math.Cos(tmp1))
                                    End If
                                End If
                            ElseIf i < nbot Then
                                'suspended segment
                                tmp1 = System.Math.Atan(System.Math.Tan(ThU(i)) - SegWgt(i) * j / NumSubSeg / FrcHor)
                                Ten = FrcHor / System.Math.Cos(tmp1)
                            Else
                                'all grounded
                                tmp1 = ThL(i)
                                Ten = ten0 - SegWgt(i) * j / NumSubSeg * (System.Math.Sin(tmp1) + SegFrc(i) * System.Math.Cos(tmp1))
                            End If
                        End If
                        If Ten < 0# Then Ten = 0#
                        tmp1 = 0.5 * (Ten + ten0) / (SegE1(i) * SegXA(i))
                        tmp2 = tmp1 * 2.0# / (System.Math.Sqrt(1.0# + 4.0# * tmp1 * SegE2(i) / SegE1(i)) + 1.0#)
                        tmp3 = tmp3 + SegLen0(i) / NumSubSeg * tmp2
                        ten0 = Ten
                    Next j
                    SegLen(i) = SegLen0(i) + tmp3
                Else
                    SegLen(i) = SegLen0(i)
                    tmp1 = ThL(i)
                    If i = nbot Then
                        If SegLen(i) = 0# Then '2.1.5
                            ten0 = FrcHor / System.Math.Cos(tmp1) '2.1.5
                        Else
                            ten0 = FrcHor / System.Math.Cos(tmp1) - SegWgt(i) * (1.0# - SegSus / SegLen(i)) * (System.Math.Sin(tmp1) + SegFrc(i) * System.Math.Cos(tmp1))
                        End If
                    ElseIf i > nbot Then
                        ten0 = ten0 - SegWgt(i) * (System.Math.Sin(tmp1) + SegFrc(i) * System.Math.Cos(tmp1))
                    Else
                        ten0 = FrcHor / System.Math.Cos(tmp1)
                    End If
                End If
                totlen = totlen + SegLen(i)
                If ipull Then
                    TotSus = totlen
                Else
                    If i < nbot Then
                        TotSus = TotSus + SegLen(i)
                    ElseIf i = nbot Then
                        SegSus = FrcHor * SegLen(nbot) / SegWgt(nbot) * (System.Math.Tan(ThU(nbot)) - System.Math.Tan(ThL(nbot)))
                        If SegSus > SegLen(nbot) Then SegSus = SegLen(nbot)
                        TotSus = TotSus + SegSus
                    End If
                End If
                If i < NumSeg Then TenU(i + 1) = ten0
            Next i
            GrdLen = totlen - TotSus
            AnchPull = ten0

            '       mooring line depth and scope
            ScopeC = 0#
            DepthC = 0#
            T0min = 0#
            For i = 1 To nbot
                If System.Math.Abs(SegWgt(i)) > 0.000001 Then
                    DepthC = DepthC + FrcHor * (1.0# / System.Math.Cos(ThU(i)) - 1.0# / System.Math.Cos(ThL(i))) * SegLen(i) / SegWgt(i)
                    ScopeC = ScopeC + FrcHor * SegLen(i) / SegWgt(i) * System.Math.Log(System.Math.Tan(0.5 * ThU(i) + PI / 4.0#) / System.Math.Tan(0.5 * ThL(i) + PI / 4.0#))
                    If i = nbot Then
                        T0min = T0min + SegSus * SegWgt(i) / SegLen(i)
                    Else
                        T0min = T0min + SegWgt(i)
                    End If
                Else
                    DepthC = DepthC + SegLen(i) * System.Math.Sin(0.5 * (ThU(i) + ThL(i)))
                    ScopeC = ScopeC + SegLen(i) * System.Math.Cos(0.5 * (ThU(i) + ThL(i)))
                End If
                Xi(i) = ScopeC
                Yi(i) = DepthC
            Next i

            If Not ipull Then
                Xi(nbot) = Xi(nbot) + (SegLen(nbot) - SegSus) * System.Math.Cos(msngBottomSlope)
                Yi(nbot) = Yi(nbot) + (SegLen(nbot) - SegSus) * System.Math.Sin(msngBottomSlope)
                For i = nbot + 1 To NumSeg
                    Xi(i) = Xi(i - 1) + SegLen(i) * System.Math.Cos(msngBottomSlope)
                    Yi(i) = Yi(i - 1) + SegLen(i) * System.Math.Sin(msngBottomSlope)
                Next
                DepthC = Yi(NumSeg)
            End If

            '       check match convergence criteria
            Dp1 = DepthC - DepthFL

            If System.Math.Abs(Dp1) < DpTol Or System.Math.Abs(Dt0) < ForceTol Then
                CatNoBuoy = True
                Exit For
            ElseIf ThUk1 = PI / 2.0# And FrcHor < 0.001 * Min(T0min, TopTen) And System.Math.Abs(Dp1) < 10.0# * DpTol Then
                TopTen = T0min - Dp1 * SegWgt(nbot) / SegLen(nbot)
                CatNoBuoy = True
                Exit For
            End If

            If icount = 1 Then
                T01 = TopTen
                If Dp1 > 0# Then
                    Dt0 = -System.Math.Abs(Dt0)
                Else
                    Dt0 = System.Math.Abs(Dt0)
                End If
                TopTen = TopTen + Dt0
                Dp0 = Dp1
            Else
                T00 = T01
                T01 = TopTen
                If Dp1 = Dp0 Then
                    Dt0 = T01 - T00
                Else
                    Dt0 = -Dp1 * (T01 - T00) / (Dp1 - Dp0)
                End If
                If System.Math.Abs(Dt0) > 0.1 * T01 Then Dt0 = 0.1 * System.Math.Sign(Dt0) * T01
                TopTen = T01 + Dt0
                Dp0 = Dp1
            End If

            If (TopTen < 1.001 * FrcHor) Then TopTen = 1.001 * FrcHor

        Next icount

    End Function

    ' calculate the horizontal force when kth segment just touch the bottom
    Private Function TouchBtm(ByRef DepthFL As Single, ByRef NumSeg As Short, ByRef SegLen0() As Single, ByRef SegWgt() As Single, ByRef SegFrc() As Single, ByRef SegXA() As Single, ByRef SegE1() As Single, ByRef SegE2() As Single, ByRef Fh0k() As Single, ByRef ThU0k(,) As Single, ByRef ThL0k(,) As Single) As Boolean

        ' Input
        '   DepthFL:    fairlead depth (to anchor)
        '   mooring line properties from this class
        ' Output
        '   Fh0k(k):    horizontal force when kth segment just touch the bottom
        '   ThU0k(i,k): ith segment upper end angle when kth segment just touch the bottom
        '   ThL0k(i,k): ith segment lower end angle when kth segment just touch the bottom
        '   return:     flag successful calculation

        Dim nstart As Short

        Dim k, i, j, icount As Short

        Dim SegLen(MaxNumSeg) As Single

        Dim TotWgt, totlen As Single
        Dim ThU(MaxNumSeg) As Single
        Dim ThL(MaxNumSeg) As Single

        Dim DepthC, depth0 As Single
        Dim Fh1, Fh0, DFh As Single
        Dim Dp1, Dp0, DpTol As Single
        Dim tmp2, tmp1, tmp3 As Single
        Dim ten0, Ten, theta As Single

        '   initialize
        DpTol = Min(DepthTol * msngWaterDepth, 2.0#)

        TotWgt = 0#
        totlen = 0#
        For i = 1 To NumSeg
            TotWgt = TotWgt + SegWgt(i)
            totlen = totlen + SegLen0(i)

            For j = 1 To NumSeg
                ThU0k(i, j) = msngBottomSlope
                ThL0k(i, j) = msngBottomSlope
            Next j
        Next i

        DepthC = 0#
        For i = 1 To NumSeg
            depth0 = DepthC
            DepthC = DepthC + SegLen0(i)

            '       be able to touch the bottom?
            If (DepthFL >= 0.99 * (depth0 + (totlen - depth0) * System.Math.Sin(msngBottomSlope)) And DepthFL < 0.99 * (DepthC + (totlen - DepthC) * System.Math.Sin(msngBottomSlope))) Then nstart = i
        Next i

        If nstart > 1 Then
            For i = 1 To nstart - 1
                Fh0k(i) = 0#
                For j = 1 To i
                    ThU0k(j, i) = PI / 2.0#
                    ThL0k(j, i) = PI / 2.0#
                Next j
            Next i
        ElseIf nstart < 1 Then
            TouchBtm = False
            Exit Function
        End If

        '   Solve for right Fh0K using iteration that matches total depth
        '   so that the lower end of Kth segment just touches bottom.
        For k = nstart To NumSeg
            TouchBtm = False

            DFh = ((TotWgt / 20.0#) * k * k) / (nstart ^ 2) ' initial increment for f0k
            If (k = 1) Then
                Fh0k(k) = 5.5 * DFh
            Else
                Fh0k(k) = Fh0k(k - 1) + 5.5 * DFh
            End If
            ThL(k) = msngBottomSlope

            For icount = 1 To MaxIter * 5
                '           going backwards and solve for all angle THETA's at joints.
                For j = k To 1 Step -1
                    If (Fh0k(k) <= 0#) Then Fh0k(k) = 0.001
                    ThU(j) = System.Math.Atan(System.Math.Tan(ThL(j)) + SegWgt(j) / Fh0k(k))
                    If (j > 1) Then ThL(j - 1) = ThU(j)
                Next j

                '           Figure out the actual segment lengths for each and every segment
                '           accounting for line elasticity / stretch.
                '           A 10-point integration was used for stretch calculation.
                '           This part of the code does not change if a portion of the segment
                '           touches the bottom (seabed).
                For i = 1 To k
                    If (SegE1(i) > 0#) Then
                        tmp3 = 0#
                        For j = 1 To NumSubSeg
                            theta = System.Math.Atan(System.Math.Tan(ThU(i)) - SegWgt(i) * (j - 0.5) / NumSubSeg / Fh0k(k))
                            Ten = Fh0k(k) / System.Math.Cos(theta)
                            tmp1 = Ten / (SegE1(i) * SegXA(i))
                            tmp2 = tmp1 * 2.0# / (System.Math.Sqrt(1.0# + 4.0# * tmp1 * SegE2(i) / SegE1(i)) + 1.0#)
                            tmp3 = tmp3 + SegLen0(i) / NumSubSeg * tmp2
                        Next j
                        SegLen(i) = SegLen0(i) + tmp3
                    Else
                        SegLen(i) = SegLen0(i)
                    End If
                Next i

                ten0 = Fh0k(k) / System.Math.Cos(msngBottomSlope)
                For i = k + 1 To NumSeg
                    If (SegE1(i) > 0#) Then
                        tmp3 = 0#
                        For j = 1 To NumSubSeg
                            Ten = ten0 - (j - 0.5) / NumSubSeg * SegWgt(i) * (System.Math.Sin(msngBottomSlope) + SegFrc(i) * System.Math.Cos(msngBottomSlope))
                            If Ten < 0# Then Ten = 0#
                            tmp1 = Ten / (SegE1(i) * SegXA(i))
                            tmp2 = tmp1 * 2.0# / (System.Math.Sqrt(1.0# + 4.0# * tmp1 * SegE2(i) / SegE1(i)) + 1.0#)
                            tmp3 = tmp3 + SegLen0(i) / NumSubSeg * tmp2
                        Next j
                        SegLen(i) = SegLen0(i) + tmp3
                        ten0 = ten0 - SegWgt(i) * (System.Math.Sin(msngBottomSlope) + SegFrc(i) * System.Math.Cos(msngBottomSlope))
                    Else
                        SegLen(i) = SegLen0(i)
                    End If
                Next i

                '           Find total depth and compare with given total water depth depthw
                DepthC = 0#
                For i = 1 To k
                    If (System.Math.Abs(SegWgt(i)) > 0.000001) Then
                        DepthC = DepthC + (1.0# / System.Math.Cos(ThU(i)) - 1.0# / System.Math.Cos(ThL(i))) * SegLen(i) / SegWgt(i) * Fh0k(k)
                    Else
                        DepthC = DepthC + SegLen(i) * System.Math.Sin(ThU(i))
                    End If
                Next i
                For i = k + 1 To NumSeg
                    DepthC = DepthC + SegLen(i) * System.Math.Sin(msngBottomSlope)
                Next i

                '           check match convergence criteria
                Dp1 = DepthC - DepthFL

                If System.Math.Abs(Dp1) < DpTol Or System.Math.Abs(DFh) < ForceTol Then
                    TouchBtm = True
                    Exit For
                End If

                If icount = 1 Then
                    Fh1 = Fh0k(k)
                    If Dp1 < 0# Then DFh = -DFh
                    Fh0k(k) = Fh0k(k) + DFh
                    Dp0 = Dp1
                Else
                    Fh0 = Fh1
                    Fh1 = Fh0k(k)
                    If Dp1 = Dp0 Then
                        DFh = -(Fh1 - Fh0)
                    Else
                        DFh = -Dp1 * (Fh1 - Fh0) / (Dp1 - Dp0)
                    End If
                    If System.Math.Abs(DFh) > 0.1 * Fh1 Then DFh = 0.1 * System.Math.Sign(DFh) * Fh1
                    Fh0k(k) = Fh1 + DFh
                    Dp0 = Dp1
                End If
            Next icount

            If Not TouchBtm Then Exit For

            For i = 1 To k
                ThU0k(i, k) = ThU(i)
                ThL0k(i, k) = ThL(i)
            Next i
        Next k

    End Function

    ' calculate the catenary (w buoys)
    Private Function CatWBuoy(ByRef FrcHor As Single, ByRef DepthFL As Single, ByRef NumSeg As Short, ByRef SegLen0() As Single, ByRef SegLen() As Single, ByRef SegWgt() As Single, ByRef SegFrc() As Single, ByRef SegXA() As Single, ByRef SegE1() As Single, ByRef SegE2() As Single, ByRef SegBuoy0() As Single, ByRef BuoyLen() As Single, ByRef TopTen As Single, ByRef Xi() As Single, ByRef Yi() As Single, ByRef ThU() As Single, ByRef ThL() As Single, ByRef TenU() As Single, ByRef GrdLen As Single, ByRef AnchPull As Single) As Boolean

        ' Input
        '   FrcHor:     applied horizontal force
        '   DepthFL:    fairlead depth (to anchor)
        '   mooring line properties
        ' Output
        '   TopTen:     corresponding top tension
        '   SegLen(i):  stretched segment length
        '   Xi(i):      lower end scope
        '   Yi(i):      lower end depth
        '   ThU(i):     upper end angle of ith segment
        '   ThL(i):     lower end angle of ith segment
        '   TenU(i):    upper end tension
        '   GrdLen:     grounded mooring line length
        '   AnchPull:   pull at anchor
        '   return:     flag successful calculation

        Dim j, i, kk As Short
        Dim icount As Short

        Dim SegBuoy(MaxNumSeg) As Single

        Dim nsub(MaxNumSeg) As Short
        Dim SubSegLen(MaxNumSeg) As Single
        Dim SubSeg1(MaxNumSeg) As Single
        Dim SubSeg2(MaxNumSeg) As Single
        Dim SubSeg3(MaxNumSeg) As Single

        Dim xbots(MaxNumSeg) As Single
        Dim ybots(MaxNumSeg) As Single
        Dim xbote(MaxNumSeg) As Single
        Dim ybote(MaxNumSeg) As Single
        Dim Tension(MaxNumSeg, 4) As Single
        Dim ThM(MaxNumSeg) As Single

        Dim ZeroTen As Boolean
        Dim Grounded(MaxNumSeg) As Boolean

        Dim jS1(MaxNumSeg) As Short
        Dim jS2(MaxNumSeg) As Short
        Dim mbuoy(MaxNumSeg) As Short
        Dim mclwt(MaxNumSeg) As Short
        Dim ibup(MaxNumSeg) As Short
        Dim ibdn(MaxNumSeg) As Short
        Dim iwup(MaxNumSeg) As Short
        Dim iwdn(MaxNumSeg) As Short
        Dim ijGrd(MaxNumSeg, 2000) As Short

        Dim Dp, DepthC, DpA, DpTol As Single
        Dim Fact, Ten, TanSlope As Single
        Dim DThU As Single
        Dim idir, idir0 As Short

        Dim itmax1, iback, itmax2 As Short
        Dim DpMaxi, DpMax As Single
        Dim jS2a, jS1a, jtmp As Short

        Dim xtmp, theta, ytmp As Single
        Dim tmp1 As Single
        Dim tmp3, tmp2 As Single

        '   initialize
        CatWBuoy = False
		DpTol = Min(DepthTol * msngWaterDepth, 2)
		TanSlope = System.Math.Tan(msngBottomSlope)
		TopTen = 1.155 * FrcHor
		
		For i = 1 To NumSeg
			SegBuoy(i) = SegBuoy0(i)
			SegLen(i) = SegLen0(i)
			nsub(i) = 2000
			SubSegLen(i) = Max(SegLen(i) / nsub(i), 4)
			nsub(i) = Int(SegLen(i) / SubSegLen(i))
			If nsub(i) = 0 Then Exit Function
			SubSegLen(i) = SegLen(i) / nsub(i)
			SubSeg1(i) = SegLen(i)
			SubSeg2(i) = 0#
			jS1(i) = nsub(i) + 1
			If i = NumSeg Then
				jS2(i) = nsub(i)
				SubSeg3(i) = 0#
			Else
				jS2(i) = 0
				SubSeg3(i) = SegLen(i)
			End If
			
			ibup(i) = 0
			ibdn(i) = 0
			iwup(i) = 0
			iwdn(i) = 0
			
			For j = 1 To nsub(i)
				ijGrd(i, j) = 100
			Next j
		Next i
		
		Xi(0) = 0#
		Yi(0) = 0#
		
		For kk = 1 To MaxIter
			
			'       compute the catenary without line elasticity and ground correction.
			ZeroTen = False
			ThU(1) = Acos(FrcHor / TopTen)
			idir0 = 1
			DThU = 0.02
			For icount = 1 To MaxIter * 5
				If ThU(1) < -PI / 2# + 0.000001 Then
					idir0 = idir
					idir = 1
					DThU = DThU / 2#
					ThU(1) = ThU(1) + idir * DThU
				ElseIf ThU(1) > PI / 2 - 0.000001 Then 
					idir0 = idir
					idir = -1
					DThU = DThU / 2#
					ThU(1) = ThU(1) + idir * DThU
				End If
				TopTen = FrcHor / System.Math.Cos(ThU(1))
				Tension(1, 1) = TopTen
				
				'           Loop for all segments
				DepthC = 0#
				For i = 1 To NumSeg
					If SubSeg2(i) <= 0# Then
						'               no bottom portion, do as usual with full weight.
						ThL(i) = System.Math.Atan(System.Math.Tan(ThU(i)) - SegWgt(i) / (Tension(i, 1) * System.Math.Cos(ThU(i))))
						Tension(i, 4) = Tension(i, 1) * System.Math.Cos(ThU(i)) / System.Math.Cos(ThL(i))
						ThM(i) = (ThU(i) + ThL(i)) * 0.5
						Tension(i, 2) = (2# * Tension(i, 1) + Tension(i, 4)) / 3#
						Tension(i, 3) = (Tension(i, 1) + 2# * Tension(i, 4)) / 3#
						
						If SegWgt(i) = 0 Or System.Math.Cos(ThU(i)) = System.Math.Cos(ThL(i)) Then Exit Function
						
						DepthC = DepthC + Tension(i, 1) * System.Math.Cos(ThU(i)) * SegLen(i) / SegWgt(i) * (1# / System.Math.Cos(ThU(i)) - 1# / System.Math.Cos(ThL(i)))
					Else
						'               bottom portion found, treat as zero weight
						If jS1(i) = 1 Then
							'                   bottom portion reaches upper end
							If jS2(i) = nsub(i) Then
								'                       bottom portion reaches both ends:
								ThM(i) = ThU(i)
								ThL(i) = ThU(i)
								Tension(i, 2) = Tension(i, 1)
								Tension(i, 3) = Tension(i, 2) - SegWgt(i) * SubSeg2(i) / SegLen(i) * (System.Math.Sin(ThM(i)) + SegFrc(i) * System.Math.Cos(ThM(i)))
								If Tension(i, 3) < 0.01 * SegWgt(i) Then
									If Tension(i, 3) <= 0# Then
										ZeroTen = True
										Tension(i, 3) = 0#
									End If
									If SegBuoy(i) > 0.00001 Then Tension(i, 3) = 0.01 * SegWgt(i)
								End If
								Tension(i, 4) = Tension(i, 3)
								
								DepthC = DepthC + SegLen(i) * System.Math.Sin(ThU(i))
							Else
								'                       bottom portion reaches upper end only:
								ThM(i) = ThU(i)
								Tension(i, 2) = Tension(i, 1)
								Tension(i, 3) = Tension(i, 2) - SegWgt(i) * SubSeg2(i) / SegLen(i) * (System.Math.Sin(ThM(i)) + SegFrc(i) * System.Math.Cos(ThM(i)))
								If Tension(i, 3) < 0.01 * SegWgt(i) Then
									If Tension(i, 3) <= 0# Then
										ZeroTen = True
										Tension(i, 3) = 0#
									End If
									If i <> NumSeg Then Tension(i, 3) = 0.01 * SegWgt(i)
								End If
								
								ThL(i) = System.Math.Atan(System.Math.Tan(ThU(i)) - SegWgt(i) * SubSeg3(i) / SegLen(i) / (Tension(i, 3) * System.Math.Cos(ThM(i))))
								Tension(i, 4) = Tension(i, 3) * System.Math.Cos(ThM(i)) / System.Math.Cos(ThL(i))
								
								DepthC = DepthC + Tension(i, 3) * System.Math.Cos(ThM(i)) * SegLen(i) / SegWgt(i) * (1# / System.Math.Cos(ThM(i)) - 1# / System.Math.Cos(ThL(i))) + SubSeg2(i) * System.Math.Sin(ThM(i))
							End If
						Else
							'                   bottom portion away from upper end
							If jS2(i) = nsub(i) Then
								'                       bottom portion reaches lower end only:
								If Tension(i, 1) > 0 Then
									ThM(i) = System.Math.Atan(System.Math.Tan(ThU(i)) - SegWgt(i) * SubSeg1(i) / SegLen(i) / (Tension(i, 1) * System.Math.Cos(ThU(i))))
								End If
								Tension(i, 2) = Tension(i, 1) * System.Math.Cos(ThU(i)) / System.Math.Cos(ThM(i))
								Tension(i, 3) = Tension(i, 2) - SegWgt(i) * SubSeg2(i) / SegLen(i) * (System.Math.Sin(ThM(i)) + SegFrc(i) * System.Math.Cos(ThM(i)))
								If Tension(i, 3) < 0.01 * SegWgt(i) Then
									If Tension(i, 3) <= 0# Then
										ZeroTen = True
										Tension(i, 3) = 0#
									End If
									If SegBuoy(i) > 0.00001 Then Tension(i, 3) = 0.01 * SegWgt(i)
								End If
								Tension(i, 4) = Tension(i, 3)
								ThL(i) = ThM(i)
								
								DepthC = DepthC + Tension(i, 1) * System.Math.Cos(ThU(i)) * SegLen(i) / SegWgt(i) * (1# / System.Math.Cos(ThU(i)) - 1# / System.Math.Cos(ThM(i))) + SubSeg2(i) * System.Math.Sin(ThM(i))
							Else
								'                       bottom portion is in the middle (general 3-section case):
								ThM(i) = System.Math.Atan(System.Math.Tan(ThU(i)) - SegWgt(i) * SubSeg1(i) / SegLen(i) / (Tension(i, 1) * System.Math.Cos(ThU(i))))
								Tension(i, 2) = Tension(i, 1) * System.Math.Cos(ThU(i)) / System.Math.Cos(ThM(i))
								Tension(i, 3) = Tension(i, 2) - SegWgt(i) * SubSeg2(i) / SegLen(i) * (System.Math.Sin(ThM(i)) + SegFrc(i) * System.Math.Cos(ThM(i)))
								If Tension(i, 3) < 0.01 * SegWgt(i) Then
									If Tension(i, 3) <= 0# Then
										ZeroTen = True
										Tension(i, 3) = 0#
									End If
									If i <> NumSeg Then Tension(i, 3) = 0.01 * SegWgt(i)
								End If
								If Tension(i, 3) > 0 Then
									ThL(i) = System.Math.Atan(System.Math.Tan(ThM(i)) - SegWgt(i) * SubSeg3(i) / SegLen(i) / (Tension(i, 3) * System.Math.Cos(ThM(i))))
								End If
								Tension(i, 4) = Tension(i, 3) * System.Math.Cos(ThM(i)) / System.Math.Cos(ThL(i))
								
								DepthC = DepthC + Tension(i, 1) * System.Math.Cos(ThU(i)) * SegLen(i) / SegWgt(i) * (1# / System.Math.Cos(ThU(i)) - 1# / System.Math.Cos(ThM(i))) + Tension(i, 3) * System.Math.Cos(ThM(i)) * SegLen(i) / SegWgt(i) * (1# / System.Math.Cos(ThM(i)) - 1# / System.Math.Cos(ThL(i))) + SubSeg2(i) * System.Math.Sin(ThM(i))
							End If
						End If
					End If
					
					If i < NumSeg Then
						If System.Math.Abs(SegBuoy(i)) > 0.00001 Then
							ThU(i + 1) = System.Math.Atan(System.Math.Tan(ThL(i)) + SegBuoy(i) / (Tension(i, 4) * System.Math.Cos(ThL(i))))
							Tension(i + 1, 1) = Tension(i, 4) * System.Math.Cos(ThL(i)) / System.Math.Cos(ThU(i + 1))
						Else
							ThU(i + 1) = ThL(i)
							Tension(i + 1, 1) = Tension(i, 4)
						End If
					End If
				Next i
				
				DpA = DepthFL - DepthC
				
				'           check match convergence criteria
				'           note: vertical height difference is controlled by top tension.
				If System.Math.Abs(DpA) < DpTol Then Exit For
				'            If Abs(DpA) < DpTol or zeroten Then Exit For
				'            If Abs(Tension(NumSeg, 4) * Cos(NumSeg)) < ForceTol Then Exit For
				
				If DThU < 0.0000000001 Then Exit Function
				If TopTen > 200# * FrcHor Then Exit Function
				
				If DpA < 0# Then
					idir = -1
				Else
					idir = 1
				End If
				If idir * idir0 < 0 Then DThU = DThU / 2#
				idir0 = idir
				
				ThU(1) = ThU(1) + idir * DThU
			Next icount
			
			iback = 0
			
			'       compute node coordinates and scope
			For i = 1 To NumSeg
				If SubSeg2(i) <= 0# Then
					'           no bottom portion, do as usual with full weight.
					Grounded(i) = False
					
					Xi(i) = Xi(i - 1) + Tension(i, 1) * System.Math.Cos(ThU(i)) * SegLen(i) / SegWgt(i) * System.Math.Log(System.Math.Tan(ThU(i) / 2# + PI / 4#) / System.Math.Tan(ThL(i) / 2# + PI / 4#))
					Yi(i) = Yi(i - 1) + Tension(i, 1) * System.Math.Cos(ThU(i)) * SegLen(i) / SegWgt(i) * (1# / System.Math.Cos(ThU(i)) - 1# / System.Math.Cos(ThL(i)))
				Else
					'           bottom portion found, treat as zero weight
					Grounded(i) = True
					
					If jS1(i) = 1 Then
						'               bottom portion reaches upper end
						If jS2(i) = nsub(i) Then
							'                   bottom portion reaches both ends:
							Xi(i) = Xi(i - 1) + SegLen(i) * System.Math.Cos(ThU(i))
							Yi(i) = Yi(i - 1) + SegLen(i) * System.Math.Sin(ThU(i))
						Else
							'                   bottom portion reaches upper end only:
							xbots(i) = Xi(i - 1)
							ybots(i) = Yi(i - 1)
							xbote(i) = xbots(i) + SubSeg2(i) * System.Math.Cos(ThM(i))
							ybote(i) = ybots(i) + SubSeg2(i) * System.Math.Sin(ThM(i))
							Xi(i) = xbote(i) + Tension(i, 3) * System.Math.Cos(ThM(i)) * SegLen(i) / SegWgt(i) * System.Math.Log(System.Math.Tan(ThM(i) / 2# + PI / 4#) / System.Math.Tan(ThL(i) / 2# + PI / 4#))
							Yi(i) = ybote(i) + Tension(i, 3) * System.Math.Cos(ThM(i)) * SegLen(i) / SegWgt(i) * (1# / System.Math.Cos(ThM(i)) - 1# / System.Math.Cos(ThL(i)))
						End If
					Else
						'               bottom portion away from upper end
						If jS2(i) = nsub(i) Then
							'                   bottom portion reaches lower end only:
							xbots(i) = Xi(i - 1) + Tension(i, 1) * System.Math.Cos(ThU(i)) * SegLen(i) / SegWgt(i) * System.Math.Log(System.Math.Tan(ThU(i) / 2# + PI / 4#) / System.Math.Tan(ThM(i) / 2# + PI / 4#))
							ybots(i) = Yi(i - 1) + Tension(i, 1) * System.Math.Cos(ThU(i)) * SegLen(i) / SegWgt(i) * (1# / System.Math.Cos(ThU(i)) - 1# / System.Math.Cos(ThL(i)))
							xbote(i) = xbots(i) + SubSeg2(i) * System.Math.Cos(ThM(i))
							ybote(i) = ybots(i) + SubSeg2(i) * System.Math.Sin(ThM(i))
							Xi(i) = xbote(i)
							Yi(i) = ybote(i)
						Else
							'                   bottom portion is in the middle (general 3-section case):
							xbots(i) = Xi(i - 1) + Tension(i, 1) * System.Math.Cos(ThU(i)) * SegLen(i) / SegWgt(i) * System.Math.Log(System.Math.Tan(ThU(i) / 2# + PI / 4#) / System.Math.Tan(ThM(i) / 2# + PI / 4#))
							ybots(i) = Yi(i - 1) + Tension(i, 1) * System.Math.Cos(ThU(i)) * SegLen(i) / SegWgt(i) * (1# / System.Math.Cos(ThU(i)) - 1# / System.Math.Cos(ThM(i)))
							xbote(i) = xbots(i) + SubSeg2(i) * System.Math.Cos(ThM(i))
							ybote(i) = ybots(i) + SubSeg2(i) * System.Math.Sin(ThM(i))
							Xi(i) = xbote(i) + Tension(i, 3) * System.Math.Cos(ThM(i)) * SegLen(i) / SegWgt(i) * System.Math.Log(System.Math.Tan(ThM(i) / 2# + PI / 4#) / System.Math.Tan(ThL(i) / 2# + PI / 4#))
							Yi(i) = ybote(i) + Tension(i, 3) * System.Math.Cos(ThM(i)) * SegLen(i) / SegWgt(i) * (1# / System.Math.Cos(ThM(i)) - 1# / System.Math.Cos(ThL(i)))
						End If
					End If
				End If
				
				TenU(i) = Tension(i, 1)
			Next i
			
			DpMax = 0#
			For i = 1 To NumSeg
				'           check buoy emergence; adjust buoyancy if emergence found.
				If SegBuoy0(i) > 0.00001 Then
					Dp = mclsFairLead.z - Yi(i) + BuoyLen(i) - DpA
					If SegBuoy(i) > 0.00001 And Dp > DpTol Then
						If kk > 25 Then
							If ibup(i) > 3 Or ibdn(i) > 3 Then
								Fact = 0.003
							Else
								Fact = 0.01
							End If
						Else
							Fact = 0.04
						End If
						SegBuoy(i) = SegBuoy(i) - Fact * SegBuoy0(i)
						
						If SegBuoy(i) < 0# Then SegBuoy(i) = 0#
						
						iback = 1
						If mbuoy(i) <> -1 Then
							mbuoy(i) = -1
							ibdn(i) = ibdn(i) + 1
						End If
					End If
					If SegBuoy(i) < SegBuoy0(i) And Dp < -DpTol Then
						If kk > 25 Then
							If ibup(i) > 3 Or ibdn(i) > 3 Then
								Fact = 0.0015
							Else
								Fact = 0.005
							End If
						Else
							Fact = 0.02
						End If
						SegBuoy(i) = SegBuoy(i) + Fact * SegBuoy0(i)
						
						If SegBuoy(i) > SegBuoy0(i) Then SegBuoy(i) = SegBuoy0(i)
						
						iback = 1
						If mbuoy(i) <> 1 Then
							mbuoy(i) = 1
							ibup(i) = ibup(i) + 1
						End If
					End If
				End If
				
				'           check clump weight; adjust weight if hitting bottom.
				If SegBuoy0(i) < -0.00001 Then
					Dp = DepthFL - Yi(i) - BuoyLen(i) - DpA - (Xi(NumSeg) - Xi(i)) * TanSlope
					If SegBuoy(i) < -0.00001 And Dp < -DpTol Then
						If kk > 25 Then
							If iwup(i) > 3 Or iwdn(i) > 3 Then
								Fact = 0.003
							Else
								Fact = 0.01
							End If
						Else
							Fact = 0.04
						End If
						SegBuoy(i) = SegBuoy(i) - Fact * SegBuoy0(i)
						
						If SegBuoy(i) > 0# Then SegBuoy(i) = 0#
						
						iback = 1
						If (mclwt(i) <> 1) Then
							mclwt(i) = 1
							iwup(i) = iwup(i) + 1
						End If
					End If
					If (SegBuoy(i) > SegBuoy0(i) And Dp > DpTol) Then
						If kk > 25 Then
							If iwup(i) > 3 Or iwdn(i) > 3 Then
								Fact = 0.0015
							Else
								Fact = 0.005
							End If
						Else
							Fact = 0.02
						End If
						SegBuoy(i) = SegBuoy(i) + Fact * SegBuoy0(i)
						
						If SegBuoy(i) < SegBuoy0(i) Then SegBuoy(i) = SegBuoy0(i)
						
						iback = 1
						If (mclwt(i) <> -1) Then
							mclwt(i) = -1
							iwdn(i) = iwdn(i) + 1
						End If
					End If
				End If
				
				'           Mark for line touching bottom
				'           initialize maximum distance below / above seabed.
				jS1a = jS1(i)
				jS2a = jS2(i)
				DpMaxi = 0#
				
				If Not Grounded(i) Then
					tmp1 = DepthFL - Yi(i) - DpA - (Xi(NumSeg) - Xi(i)) * TanSlope
					tmp2 = DepthFL - Yi(i - 1) - DpA - (Xi(NumSeg) - Xi(i - 1)) * TanSlope
					Dp = Min(tmp1, tmp2)
					
					If Dp < -DpTol Then
						Grounded(i) = True
					Else
						tmp3 = (ThU(i) - msngBottomSlope) * (ThL(i) - msngBottomSlope)
						If tmp3 <= 0# Then
							xtmp = Xi(i - 1) + Tension(i, 1) * System.Math.Cos(ThU(i)) * SegLen(i) / SegWgt(i) * System.Math.Log(System.Math.Tan(ThU(i) / 2# + PI / 4#) / System.Math.Tan(msngBottomSlope / 2# + PI / 4#))
							ytmp = Yi(i - 1) + Tension(i, 1) * SegLen(i) / SegWgt(i) * (1# - System.Math.Cos(ThU(i)) / System.Math.Cos(msngBottomSlope))
							Dp = DepthFL - ytmp - DpA - (Xi(NumSeg) - xtmp) * TanSlope
							
							If Dp < -DpTol Then Grounded(i) = True
						End If
					End If
				End If
				
				If Grounded(i) Then
					For j = 1 To nsub(i)
						If (SubSeg2(i) <= 0#) Then
							'               no bottom portion, do as usual with full weight. y-coordinates:
							theta = System.Math.Atan(System.Math.Tan(ThU(i)) - SegWgt(i) * (j - 0.5) / nsub(i) / (Tension(i, 1) * System.Math.Cos(ThU(i))))
							xtmp = Xi(i - 1) + Tension(i, 1) * System.Math.Cos(ThU(i)) * SegLen(i) / SegWgt(i) * System.Math.Log(System.Math.Tan(ThU(i) / 2# + PI / 4#) / System.Math.Tan(theta / 2# + PI / 4#))
							ytmp = Yi(i - 1) + Tension(i, 1) * SegLen(i) / SegWgt(i) * (1# - System.Math.Cos(ThU(i)) / System.Math.Cos(theta))
						Else
							'               bottom portion found, treat as zero weight.
							If jS1(i) = 1 Then
								If jS2(i) = nsub(i) Then
									'                       bottom portion reaches both ends:
									xtmp = Xi(i - 1) + (j - 0.5) / nsub(i) * SegLen(i) * System.Math.Cos(ThU(i))
									ytmp = Yi(i - 1) + (j - 0.5) / nsub(i) * SegLen(i) * System.Math.Sin(ThU(i))
								Else
									'                       bottom portion reaches upper end only:
									If j <= jS2(i) Then
										xtmp = Xi(i - 1) + (j - 0.5) / nsub(i) * SegLen(i) * System.Math.Cos(ThU(i))
										ytmp = Yi(i - 1) + (j - 0.5) / nsub(i) * SegLen(i) * System.Math.Sin(ThU(i))
									Else
										xtmp = Xi(i - 1) + SubSeg2(i) * System.Math.Cos(ThM(i))
										ytmp = Yi(i - 1) + SubSeg2(i) * System.Math.Sin(ThM(i))
										theta = System.Math.Atan(System.Math.Tan(ThM(i)) - SegWgt(i) * (j - jS2(i) - 0.5) / nsub(i) / (Tension(i, 3) * System.Math.Cos(ThM(i))))
										xtmp = xtmp + Tension(i, 3) * System.Math.Cos(ThM(i)) * SegLen(i) / SegWgt(i) * System.Math.Log(System.Math.Tan(ThU(i) / 2# + PI / 4#) / System.Math.Tan(theta / 2# + PI / 4#))
										ytmp = ytmp + Tension(i, 3) * SegLen(i) / SegWgt(i) * (1# - System.Math.Cos(ThM(i)) / System.Math.Cos(theta))
									End If
								End If
							Else
								If jS2(i) = nsub(i) Then
									'                       bottom portion reaches lower end only:
									If j < jS1(i) Then
										If Tension(i, 1) > 0 Then
											theta = System.Math.Atan(System.Math.Tan(ThU(i)) - SegWgt(i) * (j - 0.5) / nsub(i) / (Tension(i, 1) * System.Math.Cos(ThU(i))))
										Else
											theta = PI / 2
										End If
										xtmp = Xi(i - 1) + Tension(i, 1) * System.Math.Cos(ThU(i)) * SegLen(i) / SegWgt(i) * System.Math.Log(System.Math.Tan(ThU(i) / 2# + PI / 4#) / System.Math.Tan(theta / 2# + PI / 4#))
										ytmp = Yi(i - 1) + Tension(i, 1) * SegLen(i) / SegWgt(i) * (1# - System.Math.Cos(ThU(i)) / System.Math.Cos(theta))
									Else
										xtmp = Xi(i - 1) + Tension(i, 1) * System.Math.Cos(ThU(i)) * SegLen(i) / SegWgt(i) * System.Math.Log(System.Math.Tan(ThU(i) / 2# + PI / 4#) / System.Math.Tan(ThM(i) / 2# + PI / 4#))
										ytmp = Yi(i - 1) + Tension(i, 1) * SegLen(i) / SegWgt(i) * (1# - System.Math.Cos(ThU(i)) / System.Math.Cos(ThM(i)))
										xtmp = xtmp + (j - jS1(i) + 0.5) / nsub(i) * SegLen(i) * System.Math.Cos(ThM(i))
										ytmp = ytmp + (j - jS1(i) + 0.5) / nsub(i) * SegLen(i) * System.Math.Sin(ThM(i))
									End If
								Else
									'                       bottom portion is in the middle:
									If j < jS1(i) Then
										'                           1st portion
										theta = System.Math.Atan(System.Math.Tan(ThU(i)) - SegWgt(i) * (j - 0.5) / nsub(i) / (Tension(i, 1) * System.Math.Cos(ThU(i))))
										xtmp = Xi(i - 1) + Tension(i, 1) * System.Math.Cos(ThU(i)) * SegLen(i) / SegWgt(i) * System.Math.Log(System.Math.Tan(ThU(i) / 2# + PI / 4#) / System.Math.Tan(theta / 2# + PI / 4#))
										ytmp = Yi(i - 1) + Tension(i, 1) * SegLen(i) / SegWgt(i) * (1# - System.Math.Cos(ThU(i)) / System.Math.Cos(theta))
									ElseIf j <= jS2(i) Then 
										'                           2nd portion
										xtmp = Xi(i - 1) + Tension(i, 1) * System.Math.Cos(ThU(i)) * SegLen(i) / SegWgt(i) * System.Math.Log(System.Math.Tan(ThU(i) / 2# + PI / 4#) / System.Math.Tan(ThM(i) / 2# + PI / 4#))
										ytmp = Yi(i - 1) + Tension(i, 1) * SegLen(i) / SegWgt(i) * (1# - System.Math.Cos(ThU(i)) / System.Math.Cos(ThM(i)))
										xtmp = xtmp + (j - jS1(i) + 0.5) / nsub(i) * SegLen(i) * System.Math.Cos(ThM(i))
										ytmp = ytmp + (j - jS1(i) + 0.5) / nsub(i) * SegLen(i) * System.Math.Sin(ThM(i))
									Else
										'                           3rd portion
										theta = System.Math.Atan(System.Math.Tan(ThM(i)) - SegWgt(i) * (j - jS2(i) - 0.5) / nsub(i) / (Tension(i, 3) * System.Math.Cos(ThM(i))))
										xtmp = Xi(i - 1) + Tension(i, 1) * System.Math.Cos(ThU(i)) * SegLen(i) / SegWgt(i) * System.Math.Log(System.Math.Tan(ThU(i) / 2# + PI / 4#) / System.Math.Tan(ThM(i) / 2# + PI / 4#)) + SubSeg2(i) * System.Math.Cos(ThM(i)) + Tension(i, 3) * System.Math.Cos(ThM(i)) * SegLen(i) / SegWgt(i) * System.Math.Log(System.Math.Tan(ThM(i) / 2# + PI / 4#) / System.Math.Tan(theta / 2# + PI / 4#))
										ytmp = Yi(i - 1) + Tension(i, 1) * SegLen(i) / SegWgt(i) * (1# - System.Math.Cos(ThU(i)) / System.Math.Cos(ThM(i))) + SubSeg2(i) * System.Math.Sin(ThM(i)) + Tension(i, 3) * SegLen(i) / SegWgt(i) * (1# - System.Math.Cos(ThM(i)) / System.Math.Cos(theta))
									End If
								End If
							End If
						End If
						
						'               check if below sea bottom
						Dp = DepthFL - ytmp - (Xi(NumSeg) - xtmp) * TanSlope - DpA
						If Dp < -DpTol Then
							If System.Math.Abs(Dp) > DpMaxi And ijGrd(i, j) >= 100 Then
								DpMaxi = System.Math.Abs(Dp)
								If DpMaxi > DpMax Then DpMax = DpMaxi
							End If
							If j < jS1a Then jS1a = j
							If j > jS2a Then jS2a = j
						End If
						If Dp > DpTol Then
							If System.Math.Abs(Dp) > DpMaxi And ijGrd(i, j) < 100 Then
								DpMaxi = System.Math.Abs(Dp)
								If DpMaxi > DpMax Then DpMax = DpMaxi
							End If
							If j = jS1a And j < nsub(i) Then jS1a = j + 1
							If j > jS1a And j <= jS2a And i <> NumSeg Then jS2a = j - 1
						End If
					Next j
					
					'             found lowest/highest point to turn weight to zero/full.
					If DpMaxi > DpTol Then
						iback = 2
						
						If SubSeg2(i) <= 0# Then
							If jS1a = jS2a Then
								jS1(i) = jS1a
								jS2(i) = jS2a
							ElseIf jS1a > jS2a Then 
								jS1(i) = nsub(i) + 1
								jS2(i) = 0
							Else
								If i = NumSeg Then
									jtmp = jS1a + nsub(i) * 3
									jS1(i) = jtmp / 4
								Else
									jtmp = jS1a * 5 + jS2a * 3
									jS1(i) = jtmp / 8 + (jtmp / 4) Mod 2
									jtmp = jS1a * 3 + jS2a * 5
									jS2(i) = jtmp / 8 + (jtmp / 4) Mod 2
								End If
							End If
						Else
							If jS1a = jS1(i) And jS2a = jS2(i) Then
								iback = 0#
							ElseIf jS1a > jS2a Then 
								If jS1(i) >= jS2(i) - 1 Then
									jS1(i) = nsub(i) + 1
									jS2(i) = 0
								Else
									If i = NumSeg Then
										jtmp = jS1(i) + nsub(i)
										jS1(i) = jtmp / 2
									Else
										jtmp = jS1(i) * 3 + jS2(i)
										jS1(i) = jtmp / 4 + (jtmp / 2) Mod 2
										jtmp = jS1(i) + jS2(i) * 3
										jS2(i) = jtmp / 4 + (jtmp / 2) Mod 2
									End If
								End If
							Else
								If jS1a < jS1(i) Then
									jtmp = jS1(i) * 3 + jS1a
									jS1(i) = jtmp / 4 + (jtmp / 2) Mod 2
								Else
									'                            jTmp = jS1(i) * 3 + jS1a
									'                            jTmp = jTmp / 4 + (jTmp / 2) Mod 2
									'                            If jTmp > 4 + jS1(i) Then jTmp = 4 + jS1(i)
									'                            jS1(i) = jTmp
									jS1(i) = jS1(i) + 1
								End If
								If jS2a > jS2(i) Then
									jtmp = jS2(i) * 3 + jS2a
									jS2(i) = jtmp / 4 + (jtmp / 2) Mod 2
								Else
									'                            jTmp = jS2(i) * 3 + jS2a
									'                            jTmp = jTmp / 4 + (jTmp / 2) Mod 2
									'                            If jTmp < jS2(i) - 4 Then jTmp = jS2(i) - 4
									'                            jS2(i) = jTmp
									jS2(i) = jS2(i) - 1
								End If
							End If
						End If
						If i = NumSeg Then jS2(i) = nsub(i)
						
						For j = 1 To nsub(i)
							If j < jS1(i) Then
								If ijGrd(i, j) < 100 Then ijGrd(i, j) = ijGrd(i, j) + 101
							ElseIf j <= jS2(i) Then 
								If ijGrd(i, j) >= 100 Then ijGrd(i, j) = ijGrd(i, j) - 99
							Else
								If ijGrd(i, j) < 100 Then ijGrd(i, j) = ijGrd(i, j) + 101
							End If
						Next j
					End If
				End If
				
				'           figure out the actual segment lengths for each and every segment
				'           accounting for line elasticity / stretch.
				If SegE1(i) > 0# Then
					tmp3 = 0#
					For j = 1 To NumSubSeg
						If (SegWgt(i) = 0) Then
							'                   zero weight
							Ten = Tension(i, 1)
						Else
							If SubSeg2(i) <= 0# Then
								'                       no bottom portion
								theta = System.Math.Atan(System.Math.Tan(ThU(i)) - SegWgt(i) * (j - 0.5) / NumSubSeg / (Tension(i, 1) * System.Math.Cos(ThU(i))))
								Ten = Tension(i, 1) * System.Math.Cos(ThU(i)) / System.Math.Cos(theta)
							Else
								'                       bottom portion found
								If jS1(i) = 1 Then
									If jS2(i) = nsub(i) Then
										'                               bottom portion reaches both ends
										Ten = Tension(i, 1) - (j - 0.5) / NumSubSeg * SegWgt(i) * (System.Math.Sin(ThM(i)) + SegFrc(i) * System.Math.Cos(ThM(i)))
									Else
										'                               bottom portion reaches upper end only
										If (j - 0.5) * SegLen(i) / NumSubSeg < SubSeg2(i) Then
											Ten = Tension(i, 1) - (j - 0.5) / NumSubSeg * SegWgt(i) * (System.Math.Sin(ThM(i)) + SegFrc(i) * System.Math.Cos(ThM(i)))
										Else
											theta = System.Math.Atan(System.Math.Tan(ThM(i)) - SegWgt(i) * (j - jS2(i) - 0.5) / NumSubSeg / (Tension(i, 3) * System.Math.Cos(ThM(i))))
											Ten = Tension(i, 3) * System.Math.Cos(ThM(i)) / System.Math.Cos(theta)
										End If
									End If
								Else
									If jS2(i) = nsub(i) Then
										'                               bottom portion reaches lower end only
										If (j - 0.5) * SegLen(i) / NumSubSeg < SubSeg1(i) Then
											If Tension(i, 1) > 0 Then
												theta = System.Math.Atan(System.Math.Tan(ThU(i)) - SegWgt(i) * (j - 0.5) / NumSubSeg / (Tension(i, 1) * System.Math.Cos(ThU(i))))
											Else
												theta = PI / 2
											End If
											Ten = Tension(i, 1) * System.Math.Cos(ThU(i)) / System.Math.Cos(theta)
										Else
											Ten = Tension(i, 2) - SegWgt(i) * (j - jS1(i) + 0.5) / NumSubSeg * (System.Math.Sin(ThM(i)) + SegFrc(i) * System.Math.Cos(ThM(i)))
										End If
									Else
										'                               bottom portion is in the middle
										If (j - 0.5) * SegLen(i) / NumSubSeg < SubSeg1(i) Then
											'                                   1st portion
											theta = System.Math.Atan(System.Math.Tan(ThU(i)) - SegWgt(i) * (j - 0.5) / NumSubSeg / (Tension(i, 1) * System.Math.Cos(ThU(i))))
											Ten = Tension(i, 1) * System.Math.Cos(ThU(i)) / System.Math.Cos(theta)
										ElseIf (j - 0.5) * SegLen(i) / NumSubSeg < SubSeg1(i) + SubSeg2(i) Then 
											'                                   2nd portion
											Ten = Tension(i, 2) - SegWgt(i) * (j - jS1(i) + 0.5) / NumSubSeg * (System.Math.Sin(ThM(i)) + SegFrc(i) * System.Math.Cos(ThM(i)))
										Else
											'                                   3rd portion
											theta = System.Math.Atan(System.Math.Tan(ThM(i)) - SegWgt(i) * (j - jS2(i) - 0.5) / NumSubSeg / (Tension(i, 3) * System.Math.Cos(ThM(i))))
											Ten = Tension(i, 3) * System.Math.Cos(ThM(i)) / System.Math.Cos(theta)
										End If
									End If
								End If
							End If
						End If
						If (Ten < 0#) Then Ten = 0#
						tmp1 = Ten / (SegE1(i) * SegXA(i))
						tmp2 = tmp1 * 2# / (System.Math.Sqrt(1# + 4 * tmp1 * SegE2(i) / SegE1(i)) + 1#)
						tmp3 = tmp3 + (SegLen0(i) / NumSubSeg) * tmp2
					Next j
					
					SegLen(i) = SegLen0(i) + tmp3
					SubSegLen(i) = SegLen(i) / nsub(i)
					
					If jS1(i) > nsub(i) Then
						SubSeg1(i) = SegLen(i)
						SubSeg2(i) = 0#
						If i = NumSeg Then
							SubSeg3(i) = 0#
						Else
							SubSeg3(i) = SegLen(i)
						End If
					Else
						SubSeg1(i) = (jS1(i) - 1) * SubSegLen(i)
						SubSeg2(i) = (jS2(i) - jS1(i) + 1) * SubSegLen(i)
						SubSeg3(i) = SegLen(i) - SubSeg1(i) - SubSeg2(i)
					End If
				End If
			Next i
			
			'       to ensure line stretch/elasticity is accounted for.
			If (kk = 1) Then iback = 2
			
			'       every iteration, check infinitely oscilating situation:
			itmax1 = 0
			itmax2 = 0
			For i = 1 To NumSeg
				jtmp = ibup(i)
				If ibdn(i) > jtmp Then jtmp = ibdn(i)
				If iwup(i) > jtmp Then jtmp = iwup(i)
				If iwdn(i) > jtmp Then jtmp = iwdn(i)
				If jtmp > itmax1 Then itmax1 = jtmp
				For j = 1 To nsub(i)
					jtmp = ijGrd(i, j)
					If jtmp >= 100 Then jtmp = jtmp - 100
					If jtmp > itmax2 Then itmax2 = jtmp
				Next j
			Next i
			If itmax1 > 5 And itmax2 > 5 Or itmax1 > 10 Or itmax2 > 10 Then
				If DpMax < 3# * DpTol Then Exit For
				Exit Function
			End If
			
			If iback = 0 Then Exit For
		Next kk
		
		CatWBuoy = True
		
		TopTen = Tension(1, 1)
		AnchPull = Tension(NumSeg, 4)
		GrdLen = 0#
		For i = 1 To NumSeg
			GrdLen = GrdLen + Max(SubSeg2(i), 0#)
		Next i
		
	End Function
End Class