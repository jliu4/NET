Option Strict Off
Option Explicit On

' Version 8.1
' 2014, Copyright DTCEL, All Rights Reserved
' Jin Liu May 2014

Friend Class Riser

    ' private linkage to the Segments, AuxLines, BottomPackage,
    ' and StaticSolution objects
    Private mSegments As New Segments
    Private mAuxLines As New AuxLines
    Private mLatSupports As New LatSupports
    Private mBottomPackage As New BottomPackage
    Private mStaticSolution As New StaticSolution

    ' Riser properties
    Private mrstRiserType As VIVAMain.RiserTypes
    Private mbcsUpperBC, mbcsLowerBC As VIVAMain.BoundaryConditions
    Private msngUpperStiffK, msngLowerStiffK As Single
    Private msngTopTension, msngTopLocation As Single
    Private msngContentDensity, msngBuoyancyHead As Single
    Private msngFatigueConstB, msngFatigueConstA, msngSCF As Single
    Private msngFatigueCurveName As String
    Private mblnCalculateBendingStiffness As Boolean '04/28/2008
    Private mRiserTypes As Short
    Private mBoundaryConditions As Short
    Private mDamping As Single
    Private mHiRe As Short
    Private mMode As Short 'mode status:  0 - program calc natural freq and modes
    '              1 - user provide freq, program calc modes
    '              2 - user provide both

    Private mstrFreqFile As String
    Private mstrModeFile As String
    Private mstrCurvFile As String
    Private mRiserLocX As Single
    Private mRiserLocY As Single

    ' Riser output
    Public NumOutputModes As Short

    Public DOSOutText(MaxDOSOutLines) As String

    Public MaxA(NumVIVACorePoints / 7) As Single



    Public InitialFreqs(NumVIVACorePoints / 7) As Single ' TODO Jliu Initial frequencies
    Public ModeNumber(NumVIVACorePoints / 7) As Short ' Mode numbers
    Public CoupledFreqs(NumVIVACorePoints / 7) As Single ' Coupled frequencies

    Public MaxAoverD(NumVIVACorePoints / 7) As Single ' Max A/D
    Public MaxAmplitude(NumVIVACorePoints / 7) As Single
    Public AvgDragCoef(NumVIVACorePoints / 7) As Single ' Average drag coefficient

    Public MaxStress(NumVIVACorePoints / 7) As Single ' Maximum stress value
    Public MaxStrain(NumVIVACorePoints / 7) As Single ' Maximum strain value
    Public MaxBndMoment(NumVIVACorePoints / 7) As Single ' Maximum bending moment value
    Public MaxStressPos(NumVIVACorePoints / 7) As Single ' Max. stress position on riser

    Public XOM(NumVIVACorePoints / 7) As Single ' Used in fatigue calculation
    Public FatiguePeriod(NumVIVACorePoints / 7) As Single ' Fatigue period
    Public FatigueLife(NumVIVACorePoints / 7) As Single ' Fatigue life in years

    Public AoverD(NumVIVACorePoints / 7) As Single ' A/D by mode shape
    Public Amplitude(NumVIVACorePoints / 7) As Single
    Public DragCoef(NumVIVACorePoints / 7) As Single ' Drag coefficient
    Public BndMoment(NumVIVACorePoints / 7) As Single ' Bending moment
    Public Stress(NumVIVACorePoints / 7) As Single ' Bending stress
    Public Strain(NumVIVACorePoints / 7) As Single ' Bending stress
    Public FOverFs(NumVIVACorePoints / 7) As Single      ' f/fs
    Public FatLife(NumVIVACorePoints / 7) As Single ' fatigue life

    Public WaveReal(NumVIVACorePoints / 7) As Single ' Real part of the deflection
    Public WaveImag(NumVIVACorePoints / 7) As Single ' Imaginary part of the deflection
    Public WaveMag(NumVIVACorePoints / 7) As Single ' Absolute value of the deflection
    Public PhaseAngle(NumVIVACorePoints / 7) As Single ' Phase angle of the deflection

    Public SegMaxAoverD(MaxSegments, NumVIVACorePoints / 7) As Single ' Max A/D by segment
    Public SegDragCoef(MaxSegments, NumVIVACorePoints / 7) As Single ' Average drag coefficient by segment

    

    Public Property Damping() As Single

        Get
            Damping = mDamping
        End Get

        Set(ByVal Value As Single)
            If Value < 0 Then
                Err.Raise(Number:=vbObjectError + 32103, Description:="Structural damping cannot be negative")
            Else
                mDamping = Value
            End If
        End Set

    End Property

    Public Property FreqFile() As String

        Get
            FreqFile = mstrFreqFile
        End Get

        Set(ByVal Value As String)
            mstrFreqFile = Value
        End Set

    End Property

    Public Property ModeFile() As String

        Get
            ModeFile = mstrModeFile
        End Get

        Set(ByVal Value As String)
            mstrModeFile = Value
        End Set

    End Property

    Public Property CurvFile() As String

        Get
            CurvFile = mstrCurvFile
        End Get

        Set(ByVal Value As String)
            mstrCurvFile = Value
        End Set

    End Property

    Public Property HiRe() As HiReData

        Get
            HiRe = mHiRe
        End Get

        Set(ByVal Value As HiReData)
            mHiRe = Value
        End Set

    End Property

    Public Property Mode() As Short
        Get
            Mode = mMode
        End Get

        Set(ByVal Value As Short)
            mMode = Value
        End Set

    End Property

    Public Property RiserType() As VIVAMain.RiserTypes

        Get
            RiserType = mrstRiserType
        End Get

        Set(ByVal Value As VIVAMain.RiserTypes)
            mrstRiserType = Value
        End Set

    End Property


    Public Property UpperBC() As VIVAMain.BoundaryConditions

        Get
            UpperBC = mbcsUpperBC
        End Get

        Set(ByVal Value As VIVAMain.BoundaryConditions)
            mbcsUpperBC = Value
        End Set

    End Property


    Public Property UpperStiffK() As Single

        Get
            UpperStiffK = msngUpperStiffK
        End Get

        Set(ByVal Value As Single)
            msngUpperStiffK = Value
        End Set

    End Property

    Public ReadOnly Property UpperStiffC() As Single

        Get
            Dim EI As Single
            Dim LastSeg As Short

            LastSeg = mSegments.Count

            With mSegments.Item(LastSeg)
                If mAuxLines.AuxInEICalcs Then
                    EI = mAuxLines.EI + .EI
                Else
                    EI = .EI
                End If
            End With

            UpperStiffC = msngUpperStiffK / EI
        End Get

    End Property

    Public Property LowerBC() As VIVAMain.BoundaryConditions

        Get
            LowerBC = mbcsLowerBC
        End Get

        Set(ByVal Value As VIVAMain.BoundaryConditions)
            mbcsLowerBC = Value
        End Set

    End Property

    Public Property LowerStiffK() As Single

        Get
            LowerStiffK = msngLowerStiffK
        End Get

        Set(ByVal Value As Single)
            msngLowerStiffK = Value
        End Set

    End Property

    Public ReadOnly Property LowerStiffC() As Single

        Get
            Dim EI As Single

            With mSegments.Item(1)
                If mAuxLines.AuxInEICalcs Then
                    EI = mAuxLines.EI + .EI
                Else
                    EI = .EI
                End If
            End With

            LowerStiffC = msngLowerStiffK / EI
        End Get

    End Property

    Public Property TopTension() As Single

        Get
            TopTension = msngTopTension
        End Get

        Set(ByVal Value As Single)
            msngTopTension = Value
        End Set

    End Property

    Public Property TopLocation() As Single

        Get
            TopLocation = msngTopLocation
        End Get

        Set(ByVal Value As Single)
            msngTopLocation = Value
        End Set

    End Property

    Public Property ContentDensity() As Single

        Get
            ContentDensity = msngContentDensity
        End Get

        Set(ByVal Value As Single)
            msngContentDensity = Value
        End Set

    End Property

    Public Property BuoyancyHead() As Single

        Get
            BuoyancyHead = msngBuoyancyHead
        End Get

        Set(ByVal Value As Single)
            msngBuoyancyHead = Value
        End Set

    End Property

    Public ReadOnly Property TotalLength() As Single

        Get
            Dim seg As Segment

            TotalLength = 0.0#

            For Each seg In mSegments
                TotalLength = TotalLength + seg.JointLength * seg.NumJoints
            Next seg

            If mbcsLowerBC = VIVAMain.BoundaryConditions.Free Then TotalLength = TotalLength + mBottomPackage.BOP.Height + mBottomPackage.LMRP.Height
        End Get

    End Property

    Public ReadOnly Property TotalSegments() As Short

        Get
            TotalSegments = mSegments.Count

            If mbcsLowerBC = VIVAMain.BoundaryConditions.Free Then
                If mBottomPackage.BOP.Height > 0.0# Then TotalSegments = TotalSegments + 1
                If mBottomPackage.LMRP.Height > 0.0# Then TotalSegments = TotalSegments + 1
            End If
        End Get

    End Property

    Public ReadOnly Property SegNum(ByVal xr As Single, ByVal SegN2 As Short) As Short

        Get
            Dim seg As Segment
            Dim length As Single
            Dim FstSeg As Short

            length = TotalLength
            FstSeg = 0

            '   with bottom package
            If mbcsLowerBC = VIVAMain.BoundaryConditions.Free Then
                If mBottomPackage.BOP.Height > 0.0# Then
                    length = length - mBottomPackage.BOP.Height
                    FstSeg = FstSeg + 1
                    If length <= xr Then
                        SegNum = FstSeg
                        If length = xr Then
                            SegN2 = SegNum + 1
                        Else
                            SegN2 = 0
                        End If
                        Exit Property
                    End If
                End If

                If mBottomPackage.LMRP.Height > 0.0# Then
                    length = length - mBottomPackage.LMRP.Height
                    FstSeg = FstSeg + 1
                    If length <= xr Then
                        SegNum = FstSeg
                        If length = xr Then
                            SegN2 = SegNum + 1
                        Else
                            SegN2 = 0
                        End If
                        Exit Property
                    End If
                End If
            End If

            '   normal
            For Each seg In mSegments
                With seg
                    length = length - .JointLength * .NumJoints
                    If length <= xr Then
                        SegNum = .SegNum + FstSeg
                        If length = xr And SegNum < mSegments.Count + FstSeg Then
                            SegN2 = SegNum + 1
                        Else
                            SegN2 = 0
                        End If
                        Exit Property
                    End If
                End With
            Next seg

            SegNum = mSegments.Count + FstSeg
        End Get

    End Property

    Public ReadOnly Property Diameter(ByVal SegN As Short) As Single

        Get
            '   with bottom package
            If mbcsLowerBC = VIVAMain.BoundaryConditions.Free Then
                If mBottomPackage.BOP.Height > 0.0# Then
                    If SegN = 1 Then
                        Diameter = mBottomPackage.BOP.DragDiameter
                        Exit Property
                    Else
                        SegN = SegN - 1
                    End If
                End If

                If mBottomPackage.LMRP.Height > 0.0# Then
                    If SegN = 1 Then
                        Diameter = mBottomPackage.LMRP.DragDiameter
                        Exit Property
                    Else
                        SegN = SegN - 1
                    End If
                End If
            End If

            '   normal
            Diameter = mSegments.Item(SegN).BuoyModDia

        End Get

    End Property

    Public Property FatigueCurveName() As String

        Get
            FatigueCurveName = msngFatigueCurveName
        End Get

        Set(ByVal Value As String)
            msngFatigueCurveName = Value
        End Set

    End Property

    Public Property FatigueConstA() As Single

        Get
            FatigueConstA = msngFatigueConstA
        End Get

        Set(ByVal Value As Single)
            msngFatigueConstA = Value
        End Set

    End Property

    Public Property FatigueConstB() As Single

        Get
            FatigueConstB = msngFatigueConstB
        End Get

        Set(ByVal Value As Single)
            msngFatigueConstB = Value
        End Set

    End Property

    Public Property SCF() As Single

        Get
            SCF = msngSCF
        End Get

        Set(ByVal Value As Single)
            msngSCF = Value
        End Set

    End Property

    Public ReadOnly Property Segments() As Segments

        Get
            Segments = mSegments
        End Get

    End Property

    Public ReadOnly Property AuxLines() As AuxLines

        Get
            AuxLines = mAuxLines
        End Get

    End Property

    Public ReadOnly Property LatSupports() As LatSupports

        Get
            LatSupports = mLatSupports
        End Get

    End Property

    Public ReadOnly Property BottomPackage() As BottomPackage

        Get
            BottomPackage = mBottomPackage
        End Get

    End Property

    Public ReadOnly Property StaticSolution() As StaticSolution

        Get
            StaticSolution = mStaticSolution
        End Get

    End Property

    Public Property CalculateBendingStiffness() As Boolean

        Get
            CalculateBendingStiffness = mblnCalculateBendingStiffness
        End Get

        Set(ByVal Value As Boolean)
            mblnCalculateBendingStiffness = Value
        End Set

    End Property

    Private msngVectoAngle As Single
    Public Property VectoAngle() As Single
        Get
            Return msngVectoAngle
        End Get
        Set(ByVal value As Single)
            msngVectoAngle = value
        End Set
    End Property

    Public Sub New()

        MyBase.New()
        '   inintialize property variables
        mrstRiserType = VIVAMain.RiserTypes.Rigid
        mbcsUpperBC = VIVAMain.BoundaryConditions.Pinned
        mbcsLowerBC = VIVAMain.BoundaryConditions.Pinned
        msngTopTension = 0.0#
        msngContentDensity = 1025.0#
        msngBuoyancyHead = 0.0#
        msngSCF = 1.5
        msngFatigueCurveName = "API X"
        msngFatigueConstA = 1372700000.0#
        msngFatigueConstB = 4.38
        mblnCalculateBendingStiffness = False
        mDamping = 0.01
        mHiRe = 2


    End Sub

End Class