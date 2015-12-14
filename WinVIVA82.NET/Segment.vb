Option Strict Off
Option Explicit On

' Version 8.1
' 2014, Copyright DTCEL, All Rights Reserved
' Jin Liu May 2014

Friend Class Segment

    Private mintSegNum, mintNumJoints As Short
    Private msngJointLength As Single
    Private msngWallThickness, msngMainTubeOD, msngID As Single
    Private msngDryWeight, msngWetWeight As Single
    Private msngBuoyModDia, msngModulusOfElasticity As Single
    Private msngStrakesHeight As Single
    Private msngFairThick, msngFairChord As Single
    Private mchrSectionType As VIVAMain.ICHARValues


    Public Property SegNum() As Short

        Get
            SegNum = mintSegNum
        End Get

        Set(ByVal Value As Short)
            Static blnAlreadySet As Boolean

            If Not blnAlreadySet Then
                blnAlreadySet = True
                mintSegNum = Value
            Else
                Err.Raise(Number:=vbObjectError + 32106, Description:="Segment" & CStr(Value) & " already defined.")
            End If
        End Set

    End Property


    Public Property NumJoints() As Short

        Get
            NumJoints = mintNumJoints
        End Get

        Set(ByVal Value As Short)
            mintNumJoints = Value
        End Set

    End Property


    Public Property JointLength() As Single

        Get
            JointLength = msngJointLength
        End Get

        Set(ByVal Value As Single)
            msngJointLength = Value
        End Set

    End Property


    Public Property MainTubeOD() As Single

        Get
            MainTubeOD = msngMainTubeOD
        End Get

        Set(ByVal Value As Single)
            msngMainTubeOD = Value
            msngID = msngMainTubeOD - 2 * msngWallThickness
        End Set

    End Property


    Public Property WallThickness() As Single

        Get
            WallThickness = msngWallThickness
        End Get

        Set(ByVal Value As Single)
            msngWallThickness = Value
            msngID = msngMainTubeOD - 2 * msngWallThickness
        End Set

    End Property


    Public Property DryWeight() As Single

        Get
            DryWeight = msngDryWeight
        End Get

        Set(ByVal Value As Single)
            msngDryWeight = Value
        End Set

    End Property


    Public Property WetWeight() As Single

        Get
            WetWeight = msngWetWeight
        End Get

        Set(ByVal Value As Single)
            msngWetWeight = Value
        End Set

    End Property


    Public Property BuoyModDia() As Single

        Get
            BuoyModDia = msngBuoyModDia
            If msngMainTubeOD > BuoyModDia Then BuoyModDia = msngMainTubeOD
        End Get

        Set(ByVal Value As Single)
            msngBuoyModDia = Value
        End Set

    End Property


    Public Property SectionType() As VIVAMain.ICHARValues

        Get
            SectionType = mchrSectionType
        End Get

        Set(ByVal Value As VIVAMain.ICHARValues)
            mchrSectionType = Value
        End Set

    End Property


    Public Property StrakesHeight() As Single

        Get
            StrakesHeight = msngStrakesHeight
        End Get

        Set(ByVal Value As Single)
            msngStrakesHeight = Value
        End Set

    End Property


    Public Property FairThick() As Single

        Get
            FairThick = msngFairThick
        End Get

        Set(ByVal Value As Single)
            msngFairThick = Value
        End Set

    End Property


    Public Property FairChord() As Single

        Get
            FairChord = msngFairChord
        End Get

        Set(ByVal Value As Single)
            msngFairChord = Value
        End Set

    End Property


    Public Property ModulusOfElasticity() As Single

        Get
            ModulusOfElasticity = msngModulusOfElasticity
        End Get

        Set(ByVal Value As Single)
            msngModulusOfElasticity = Value
        End Set

    End Property


    Public ReadOnly Property MassPerLength() As Single

        Get
            MassPerLength = msngDryWeight / msngJointLength / Gm + IntXSection() * CurProj.Riser(CurProj.RiserId).ContentDensity
        End Get

    End Property


    Public ReadOnly Property AddedMassPerLength() As Single

        Get
            AddedMassPerLength = CurProj.Water.Density * msngBuoyModDia ^ 2 * Pi / 4.0#
        End Get

    End Property


    Public ReadOnly Property WetWeightPerLength(Optional ByVal TopAbvWL As Single = 0.0#) As Single

        Get
            '   4/5/2005: change jointlength to seglength in the calculations
            Dim psngSegLength As Single

            psngSegLength = msngJointLength * mintNumJoints

            If TopAbvWL > 0.0# Then
                If TopAbvWL > psngSegLength Then
                    WetWeightPerLength = msngDryWeight / msngJointLength + IntXSection() * CurProj.Riser(CurProj.RiserId).ContentDensity * Gm
                Else
                    WetWeightPerLength = (msngDryWeight / msngJointLength * TopAbvWL + msngWetWeight / msngJointLength * (psngSegLength - TopAbvWL)) / psngSegLength + IntXSection() * (CurProj.Riser(CurProj.RiserId).ContentDensity - CurProj.Water.Density * (psngSegLength - TopAbvWL) / psngSegLength) * Gm
                End If
            Else
                WetWeightPerLength = msngWetWeight / msngJointLength + IntXSection() * (CurProj.Riser(CurProj.RiserId).ContentDensity - CurProj.Water.Density) * Gm
            End If
        End Get

    End Property


    Public ReadOnly Property EI() As Single

        Get
            EI = msngModulusOfElasticity * (msngMainTubeOD ^ 4 - msngID ^ 4) * Pi / 64.0#
        End Get

    End Property


    Friend Function IntXSection() As Single

        IntXSection = Pi * msngID ^ 2 / 4.0#

    End Function

    Private msngBendingStiffness As Single
    Public Property BendingStiffness() As Single
        Get
            Return msngBendingStiffness
        End Get
        Set(ByVal value As Single)
            msngBendingStiffness = value
        End Set
    End Property

End Class