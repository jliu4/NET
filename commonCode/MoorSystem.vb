Option Strict Off
Option Explicit On
Friend Class MoorSystem
	Implements System.Collections.IEnumerable
	
	' mooring system
	
	' ShipGlob:     current ship location in global system
	' ShipDraft:    ship draft
	
	' WinchCap:     winch capacity, limiting top tension
	' MoorLines:    mooring lines
	' FMoorGlob:    horizontal force and it components in global system
	' FMoorLocl:    horizontal force and it components in local system
	' StiffGlob:    mooring system stiffness in global system
	' StiffLocl:    mooring system stiffness in local system
	
	Private mclsShipGlob As ShipGlobal
    Private msngShipDraft As Single

    Private msngWinchCap As Single
    Private mcolMoorLines As Collection
    Private mclsFMoorGlob As Force
    Private mclsFMoorLocl As Force
    Private mclsStiffGlob As Force
    Private mclsStiffLocl As Force

    Public Sub New()
        MyBase.New()
        mclsShipGlob = New ShipGlobal

        msngWinchCap = 0#

        mcolMoorLines = New Collection
        mclsFMoorGlob = New Force
        mclsFMoorLocl = New Force

        mclsStiffGlob = New Force
        mclsStiffLocl = New Force
    End Sub

    Protected Overrides Sub Finalize()
        mclsShipGlob = Nothing
        mcolMoorLines = Nothing
        mclsFMoorGlob = Nothing
        mclsFMoorLocl = Nothing
        mclsStiffGlob = Nothing
        mclsStiffLocl = Nothing

        MyBase.Finalize()
    End Sub


    Public Property ShipDraft() As Single
        Get

            ShipDraft = msngShipDraft

        End Get
        Set(ByVal Value As Single)

            Dim MoorLine_Renamed As MoorLine

            msngShipDraft = Value

            For Each MoorLine_Renamed In mcolMoorLines
                MoorLine_Renamed.Draft = msngShipDraft
            Next MoorLine_Renamed

        End Set
    End Property

    Public ReadOnly Property ShipGlob() As ShipGlobal
        Get

            ShipGlob = mclsShipGlob

        End Get
    End Property


    Public Property WinchCap() As Single
        Get

            WinchCap = msngWinchCap

        End Get
        Set(ByVal Value As Single)

            msngWinchCap = Value

        End Set
    End Property

    Public ReadOnly Property FMoorLocl(Optional ByVal ReCalculate As Boolean = True) As Force
        Get

            If ReCalculate Then Call TotFMoor(mclsFMoorLocl, mclsShipGlob, True)

            FMoorLocl = mclsFMoorLocl

        End Get
    End Property

    Public ReadOnly Property FMoorGlob(Optional ByVal ReCalculate As Boolean = True) As Force
        Get

            If ReCalculate Then
                Call TotFMoor(mclsFMoorLocl, mclsShipGlob)
                Call FGlobFrmLocl(mclsFMoorGlob, mclsFMoorLocl, (mclsShipGlob.Heading))
            End If

            FMoorGlob = mclsFMoorGlob

        End Get
    End Property

    Public ReadOnly Property StiffLocl(Optional ByVal ReCalculate As Boolean = True) As Force
        Get

            If ReCalculate Then
                Call MoorStiff(mclsStiffLocl, mclsShipGlob, False)
            End If

            StiffLocl = mclsStiffLocl

        End Get
    End Property

    Public ReadOnly Property StiffGlob(Optional ByVal ReCalculate As Boolean = False) As Force
        Get

            If ReCalculate Then Call MoorStiff(mclsStiffGlob, mclsShipGlob)

            StiffGlob = mclsStiffGlob

        End Get
    End Property

    Public ReadOnly Property MoorLines(ByVal vntIndexKey As Object) As MoorLine
        Get
            If vntIndexKey > 0 Then
                MoorLines = mcolMoorLines.Item(vntIndexKey)
            End If

        End Get
    End Property

    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        GetEnumerator = mcolMoorLines.GetEnumerator
    End Function

    Public Sub MoorLineAdd()

        Dim NewMoorLine As New MoorLine

        NewMoorLine.Draft = msngShipDraft
        mcolMoorLines.Add(NewMoorLine)

    End Sub

    Public Function MoorLineCount() As Integer

        MoorLineCount = mcolMoorLines.Count()

    End Function

    Public Sub MoorLineDelete(ByRef vntIndexKey As Object)

        mcolMoorLines.Remove(vntIndexKey)

    End Sub

    Public Function InputML(ByVal FileNum As Short) As Boolean

        Dim i, j As Short
        Dim NumLine, NumSeg As Short
        Dim Payout, Scope, TopTen As Single
        Dim PayoutSur, PayoutOpr As Single
        Dim PretenSur, PretenOpr As Single
        Dim FLY, SprdAng, FLX, FLZ As Single
        Dim AnchorY, AnchorX, WaterDepth As Single
        Dim BtmSlp As String
        Dim AnchModel, AnchRemark As String
        Dim HoldCap As Single
        Dim SegTp As String
        Dim TLg, Lg, dia As Single
        Dim E1, BS, E2 As Single
        Dim Buoy, DryWt, WetWt, BuoyL As Single
        Dim FrCoef As Single
        Dim fairLeadNode, anchorNode As Short

        Dim WinchCap As Single

        Dim NewMoorLine As MoorLine

        InputML = False

        On Error GoTo ErrorHandler

        Input(FileNum, NumLine)

        Do While mcolMoorLines.Count() > 0
            mcolMoorLines.Remove((1))
        Loop

        Dim TmpStr As String

        For i = 1 To NumLine
            NewMoorLine = New MoorLine

            Input(FileNum, Scope)
            Input(FileNum, Payout)
            Input(FileNum, TopTen)
            Input(FileNum, PayoutSur)
            Input(FileNum, PretenSur)
            Input(FileNum, PayoutOpr)
            Input(FileNum, PretenOpr)
            Input(FileNum, WinchCap)
            Input(FileNum, SprdAng)
            Input(FileNum, FLX)
            Input(FileNum, FLY)
            Input(FileNum, FLZ)
            Input(FileNum, fairLeadNode)
            Input(FileNum, AnchorX)
            Input(FileNum, AnchorY)
            Input(FileNum, WaterDepth)
            Input(FileNum, BtmSlp)
            Input(FileNum, anchorNode)
            Input(FileNum, AnchModel)
            Input(FileNum, HoldCap)
            Input(FileNum, AnchRemark)

            If fairLeadNode = 0 Then
                fairLeadNode = 9500 + i
            End If

            If anchorNode = 0 Then
                anchorNode = 9100 + i
            End If

            With NewMoorLine
                .Draft = msngShipDraft
                .DesScope = Scope
                .Payout = Payout
                .TopTension = TopTen
                .FairLead.SprdAngle = SprdAng * Degrees2Radians
                .FairLead.Xs = FLX
                .FairLead.Ys = FLY
                .FairLead.z = FLZ
                .FairLead.Node = fairLeadNode
                .Anchor.Xg = AnchorX
                .Anchor.Yg = AnchorY
                .Anchor.Node = anchorNode
                .WaterDepth = WaterDepth
                .BottomSlope = CSng(Val(BtmSlp)) * Degrees2Radians

                With .Anchor
                    .Model = AnchModel
                    .HoldCap = HoldCap
                    .Remark = AnchRemark
                End With
                .WinchCap = WinchCap

                .PayoutSur = PayoutSur
                .PretensionSur = PretenSur
                .PayoutOpr = PayoutOpr
                .PretensionOpr = PretenOpr
            End With

            Input(FileNum, NumSeg)
            For j = 1 To NumSeg
                Input(FileNum, SegTp)
                Input(FileNum, Lg)
                Input(FileNum, TLg)
                Input(FileNum, dia)
                Input(FileNum, BS)
                Input(FileNum, E1)
                Input(FileNum, E2)
                Input(FileNum, DryWt)
                Input(FileNum, WetWt)
                Input(FileNum, Buoy)
                Input(FileNum, BuoyL)
                Input(FileNum, FrCoef)
                Call NewMoorLine.SegmentAdd(SegTp, Lg, TLg, dia, BS, E1, E2, DryWt, WetWt, Buoy, BuoyL, FrCoef)
            Next j
            mcolMoorLines.Add(NewMoorLine)
        Next i

        InputML = True
        Exit Function
ErrorHandler:
        If Len(Err.Description) > 0 Then MsgBox(Err.Description, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Error")
        InputML = False
    End Function

    Public Function OutputML(ByVal FileNum As Short) As Boolean

        Dim i, j As Short
        Dim Segment_Renamed As Segment
        Dim NumLine As Short

        OutputML = False

        NumLine = mcolMoorLines.Count()
        WriteLine(FileNum, NumLine)

        For i = 1 To NumLine
            With mcolMoorLines.Item(i)
                WriteLine(FileNum, .DesScope, .Payout, .TopTension, .PayoutSur, .PretensionSur, .PayoutOpr, .PretensionOpr, .WinchCap)
                WriteLine(FileNum, .FairLead.SprdAngle * Radians2Degrees, .FairLead.Xs, .FairLead.Ys, .FairLead.z, .Fairlead.NODE)
                WriteLine(FileNum, .Anchor.Xg, .Anchor.Yg, .WaterDepth, .BottomSlope * Radians2Degrees, .Anchor.NODE)
                With .Anchor
                    WriteLine(FileNum, .Model, .HoldCap, .Remark)
                End With

                WriteLine(FileNum, .SegmentCount)
            End With

            For Each Segment_Renamed In mcolMoorLines.Item(i)
                With Segment_Renamed
                    WriteLine(FileNum, .SegType, .Length, .TotalLength, .Diameter, .BS, .E1, .E2, .UnitDryWeight, .UnitWetWeight, .Buoy, .BuoyLength, .FrictionCoef)
                End With
            Next Segment_Renamed
        Next i

        OutputML = True

    End Function

    Public Function GenCatenaryFile(ByVal FileName As String) As Boolean
        ' catenaryfile for plot z from baseline

        Dim k, i, j, L As Short
        Dim Yfl, Xfl, Zfl As Single
        Dim Y, X, z As Single
        Dim CatX(MaxNumSubSeg * MaxNumSeg + 1) As Single
        Dim CatY(MaxNumSubSeg * MaxNumSeg + 1) As Single
        Dim Connector(MaxNumSeg + 1) As Short
        Dim SprdAngle As Single
        Dim NumPoints, NP As Short
        Dim FileOpen_Renamed As Boolean

        On Error GoTo ErrorHandler

        GenCatenaryFile = False
        FileOpen_Renamed = False

        FileClose(FileNumRes)
        FileOpen(FileNumRes, FileName, OpenMode.Output, OpenAccess.Write)
        FileOpen_Renamed = True

        For i = 1 To mcolMoorLines.Count()
            With mcolMoorLines.Item(i)
                If .Connected Then
                    If Not .CatenaryPoints(CatX, CatY, Connector) Then Exit Function
                    SprdAngle = .SprdAngle(mclsShipGlob)

                    WriteLine(FileNumRes, "Line " & i)
                    With .FairLead
                        Xfl = .Xs
                        Yfl = .Ys
                        Zfl = .z
                    End With
                    WriteLine(FileNumRes, Xfl, Yfl, Zfl)

                    For j = 1 To .SegmentCount
                        NumPoints = Connector(j) - Connector(j + 1)
                        If NumPoints > 10 Then
                            NP = 10
                        Else
                            NP = NumPoints
                        End If
                        For k = 1 To NP Step 1
                            L = NumPoints / NP * k
                            X = CatX(Connector(j) - L) * System.Math.Cos(-SprdAngle) + Xfl
                            Y = CatX(Connector(j) - L) * System.Math.Sin(-SprdAngle) + Yfl
                            z = -CatY(Connector(j) - L) + msngShipDraft
                            WriteLine(FileNumRes, X, Y, z)
                        Next k
                    Next j
                End If
            End With
        Next i

        FileClose(FileNumRes)
        FileOpen_Renamed = False
        GenCatenaryFile = True
        Exit Function

ErrorHandler:
        If FileOpen_Renamed Then FileClose(FileNumRes)
        MsgBox("Error: " & Err.Description & ", Source: " & Err.Source, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Error")
    End Function

    Public Sub MoorForce(ByRef FMoor As Force, ByRef ShipLoc As ShipGlobal, Optional ByRef IsGlobal As Boolean = True)

        Dim FMLocl As New Force

        Call TotFMoor(FMLocl, ShipLoc)
        If IsGlobal Then
            Call FGlobFrmLocl(FMoor, FMLocl, (ShipLoc.Heading))
        Else
            With FMoor
                .Fx = FMLocl.Fx
                .Fy = FMLocl.Fy
                .MYaw = FMLocl.MYaw
            End With
        End If


    End Sub

    Public Sub MoorStiff(ByRef Stiff As Force, Optional ByRef ShipLoc As ShipGlobal = Nothing, Optional ByRef IsGlobal As Boolean = True)

        Dim WaterDepth, Alpha As Single
        Dim dy, dx, Drz As Single
        Dim SgnY, SgnX, SgnR As Single
        Dim DFy, DFx, DMz As Single
        Dim ShipMove As New ShipGlobal
        Dim MoorLine_Renamed As MoorLine
        Dim FLocl1 As New Force
        Dim FGlob1 As New Force
        Dim FLocl, FGlob As Force

        If IsNothing(ShipLoc) Then
            ShipLoc = mclsShipGlob
            FLocl = mclsFMoorLocl
            FGlob = mclsFMoorGlob
        Else
            FLocl = New Force
            FGlob = New Force
        End If

        Call TotFMoor(FLocl, ShipLoc)
        Call FGlobFrmLocl(FGlob, FLocl, (ShipLoc.Heading))

        '   determine the moving direction
        If IsGlobal Then
            With FGlob
                SgnX = -System.Math.Sign(.Fx)
                SgnY = -System.Math.Sign(.Fy)
                SgnR = -System.Math.Sign(.MYaw)
            End With
        Else
            With FLocl
                SgnX = -System.Math.Sign(.Fx)
                SgnY = -System.Math.Sign(.Fy)
                SgnR = -System.Math.Sign(.MYaw)
            End With
        End If
        If SgnX = 0# Then SgnX = 1.0#
        If SgnY = 0# Then SgnY = 1.0#
        If SgnR = 0# Then SgnR = 1.0#


        '   determine the movements
        WaterDepth = 0#
        For Each MoorLine_Renamed In mcolMoorLines
            If WaterDepth < MoorLine_Renamed.WaterDepth Then WaterDepth = MoorLine_Renamed.WaterDepth
        Next MoorLine_Renamed

        dx = Max(0.001 * WaterDepth, 10) * SgnX
        dy = Max(0.001 * WaterDepth, 10) * SgnY
        Drz = 0.002 * SgnR

        '   calculate stiffness in x direction (global)
        With ShipMove
            If IsGlobal Then
                .Xg = ShipLoc.Xg + dx
                .Yg = ShipLoc.Yg
            Else
                Alpha = PI / 2.0# - ShipLoc.Heading
                .Xg = ShipLoc.Xg + dx * System.Math.Cos(Alpha)
                .Yg = ShipLoc.Yg + dx * System.Math.Sin(Alpha)
            End If
            .Heading = ShipLoc.Heading
        End With

        Call TotFMoor(FLocl1, ShipMove)
        If IsGlobal Then
            Call FGlobFrmLocl(FGlob1, FLocl1, (ShipMove.Heading))
            DFx = Max(System.Math.Abs(-FGlob1.Fx + FGlob.Fx), 10.0#) * SgnX
        Else
            DFx = Max(System.Math.Abs(-FLocl1.Fx + FLocl.Fx), 10.0#) * SgnX
        End If

        '   calculate stiffness in y direction (global)
        With ShipMove
            If IsGlobal Then
                .Xg = ShipLoc.Xg
                .Yg = ShipLoc.Yg + dy
            Else
                .Xg = ShipLoc.Xg - dy * System.Math.Sin(Alpha)
                .Yg = ShipLoc.Yg + dy * System.Math.Cos(Alpha)
            End If
            .Heading = ShipLoc.Heading
        End With

        Call TotFMoor(FLocl1, ShipMove)
        If IsGlobal Then
            Call FGlobFrmLocl(FGlob1, FLocl1, (ShipMove.Heading))
            DFy = Max(System.Math.Abs(-FGlob1.Fy + FGlob.Fy), 10.0#) * SgnY
        Else
            DFy = Max(System.Math.Abs(-FLocl1.Fy + FLocl.Fy), 10.0#) * SgnY
        End If

        '   calculate stiffness in yaw direction (global)
        With ShipMove
            .Xg = ShipLoc.Xg
            .Yg = ShipLoc.Yg
            .Heading = ShipLoc.Heading - Drz
            ' heading is clock-wise, yaw is counter-clock-wise
        End With

        Call TotFMoor(FLocl1, ShipMove)
        DMz = Max(System.Math.Abs(-FLocl1.MYaw + FLocl.MYaw), 10.0#) * SgnR

        With Stiff
            .Fx = DFx / dx
            .Fy = DFy / dy
            .MYaw = DMz / Drz
        End With

    End Sub

    Private Sub TotFMoor(ByRef FLocl As Force, ByRef ShipLoc As ShipGlobal, Optional ByRef Moved As Boolean = True)

        Dim LineForce As Force
        Dim FSw, FSg, MY_Renamed As Single
        Dim MoorLine_Renamed As MoorLine

        FSg = 0#
        FSw = 0#
        MY_Renamed = 0#

        For Each MoorLine_Renamed In mcolMoorLines
            LineForce = MoorLine_Renamed.FMoorLocl(ShipLoc, Moved)
            FSg = FSg + LineForce.Fx
            FSw = FSw + LineForce.Fy
            MY_Renamed = MY_Renamed + LineForce.MYaw
        Next MoorLine_Renamed

        With FLocl
            .Fx = FSg
            .Fy = FSw
            .MYaw = MY_Renamed
        End With

    End Sub

    Private Sub FGlobFrmLocl(ByRef FGlob As Force, ByRef FLocl As Force, ByRef Heading As Single)

        Dim Alpha As Single

        Alpha = PI / 2.0# - Heading

        If FGlob Is Nothing Then
        Else
            With FGlob
                .Fx = System.Math.Cos(Alpha) * FLocl.Fx - System.Math.Sin(Alpha) * FLocl.Fy
                .Fy = System.Math.Sin(Alpha) * FLocl.Fx + System.Math.Cos(Alpha) * FLocl.Fy
                .MYaw = FLocl.MYaw
            End With
        End If

    End Sub

    Private Sub SLoclFrmGlob(ByRef SGlob As Force, ByRef SLocl As Force, ByRef Heading As Single)

        Dim Alpha As Single

        Alpha = PI / 2# - Heading
		
		With SLocl
			.Fx = System.Math.Cos(Alpha) ^ 2 * SGlob.Fx + System.Math.Sin(Alpha) ^ 2 * SGlob.Fy
			.Fy = System.Math.Sin(Alpha) ^ 2 * SGlob.Fx + System.Math.Cos(Alpha) ^ 2 * SGlob.Fy
			.MYaw = SGlob.MYaw
		End With
		
	End Sub
End Class