Option Strict Off
Option Explicit On
Friend Class EnvLoad

    ' environments

    ' properties
    ' ShipDraft:    ship draft (ft)
    ' ShipHead:     current ship heading (N clock-wise) (rad)

    ' EnvCur:       current environment
    ' Environments: list of metocean criteria

    ' CurrentFC:    current force coefficient
    ' DrftFrSpec:   get wave drift force spectrum
    ' WaveFC:       wave drift force coefficient
    ' Wave2FC:      wave drift force coefficient by wave velocity
    ' WindFC:       wind force coefficient

    ' FCurrGlob:    current force in global system
    ' FCurrLocl:    current force in local system
    ' FWaveGlob:    wave force in global system
    ' FWaveLocl:    wave force in local system
    ' FWindGlob:    wind force in global system
    ' FWindLocl:    wind force in local system
    ' FEnvGlob:     environment load in global system
    ' FEnvLocl:     environment load in local system

    ' methods
    ' EnvironmentAdd:   add a metocean into the list
    ' EnvironmentCount: total number of metocean in the list
    ' EnvironmentDelete:delete a metocean from the list

    Private msngShipDraft As Single
    Private msngShipHead As Single
    Private mblnWave2 As Boolean

    Private mcolWindFC As Collection
    Private mcolCurrentFC As Collection
    Private mcolWaveFC As Collection
    Private mcolWave2FC As Collection

    Private mclsEnvCur As Metocean
    Private mcolEnvironments As Collection
    Private mclsFWaveGlob As Force
    Private mclsFWaveLocl As Force
    Private mclsFCurrGlob As Force
    Private mclsFCurrLocl As Force
    Private mclsFWindGlob As Force
    Private mclsFWindLocl As Force
    Private mclsFEnvGlob As Force
    Private mclsFEnvLocl As Force

    Private Const FreqS As Single = 0.02
    Private Const FreqE As Single = 2.4
    Private Const FreqD As Single = 0.02

    Public Sub New()
        MyBase.New()
        mclsEnvCur = New Metocean
        mcolEnvironments = New Collection
        NoEnvironment()

        mcolWindFC = New Collection
        mcolCurrentFC = New Collection
        mcolWaveFC = New Collection
        mcolWave2FC = New Collection

        mclsFWaveGlob = New Force
        mclsFWaveLocl = New Force
        mclsFCurrGlob = New Force
        mclsFCurrLocl = New Force
        mclsFWindGlob = New Force
        mclsFWindLocl = New Force
        mclsFEnvGlob = New Force
        mclsFEnvLocl = New Force

        mblnWave2 = False
    End Sub


    ' properties

    Public Property ShipDraft() As Single
        Get

            ShipDraft = msngShipDraft

        End Get
        Set(ByVal Value As Single)

            msngShipDraft = Value

        End Set
    End Property


    Public Property ShipHead() As Single
        Get

            ShipHead = msngShipHead

        End Get
        Set(ByVal Value As Single)

            msngShipHead = Value

        End Set
    End Property

    Public ReadOnly Property EnvCur() As Metocean
        Get

            EnvCur = mclsEnvCur

        End Get
    End Property

    Public ReadOnly Property Environments(ByVal vntIndexKey As Object) As Metocean
        Get

            Environments = mcolEnvironments.Item(vntIndexKey)

        End Get
    End Property

    Public ReadOnly Property CurrentFC(ByVal Direction As Single, Optional ByVal Draft As Single = 0) As Force
        Get

            ' Input
            '   Direction:  current direction (local) (rad)
            '   Draft:      vessel draft (ft)

            'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
            If IsNothing(Draft) Then Draft = msngShipDraft

            CurrentFC = ForceCoef(Draft, Direction, mcolCurrentFC)

        End Get
    End Property

    Public ReadOnly Property DrftFrSpec(ByVal Frequency As Single, Optional ByVal Heading As Single = 0) As Force
        Get

            ' Input
            '   Frequency:  wave frequency (rad/sec)
            '   Heading:    vessel heading (global) (rad)

            Dim DrftFC As Force
            'UPGRADE_NOTE: MY was upgraded to MY_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
            Dim FSw, FSg, MY_Renamed As Single
            Dim EnvDir As Single
            Dim Sw1, Freq1, Freq2, Sw2 As Single

            'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
            If IsNothing(Heading) Then Heading = msngShipHead

            EnvDir = Heading - mclsEnvCur.Wave.Heading
            FSg = 0#
            FSw = 0#
            MY_Renamed = 0#

            For Freq1 = FreqS To FreqE Step FreqD
                Freq2 = Freq1 + Frequency / 2.0#
                Sw1 = mclsEnvCur.Wave.Spectrum(Freq1)
                Sw2 = mclsEnvCur.Wave.Spectrum(Freq2)
                DrftFC = DrftCoef(msngShipDraft, Freq2, EnvDir)

                With DrftFC
                    FSg = FSg + .Fx ^ 2 * Sw1 * Sw2 * FreqD
                    FSw = FSw + .Fy ^ 2 * Sw1 * Sw2 * FreqD
                    MY_Renamed = MY_Renamed + .MYaw ^ 2 * Sw1 * Sw2 * FreqD
                End With
                'UPGRADE_NOTE: Object DrftFC may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                DrftFC = Nothing
            Next

            DrftFrSpec = New Force

            With DrftFrSpec
                .Fx = FSg * 8.0#
                .Fy = FSw * 8.0#
                .MYaw = MY_Renamed * 8.0#
            End With

        End Get
    End Property

    Public ReadOnly Property WaveFC(ByVal Frequency As Single, ByVal Direction As Single, Optional ByVal Draft As Single = 0) As Force
        Get

            ' Input
            '   Frequency:  wave frequency (rad/sec)
            '   Direction:  wave direction (local) (rad)
            '   Draft:      vessel draft (ft)

            'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
            If IsNothing(Draft) Then Draft = msngShipDraft

            WaveFC = DrftCoef(Draft, Frequency, Direction)

        End Get
    End Property

    Public ReadOnly Property Wave2FC(ByVal Direction As Single, Optional ByVal Draft As Single = 0) As Force
        Get

            ' Input
            '   Direction:  wave direction (local) (rad)
            '   Draft:      vessel draft (ft)

            'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
            If IsNothing(Draft) Then Draft = msngShipDraft

            Wave2FC = ForceCoef(Draft, Direction, mcolWave2FC)

        End Get
    End Property

    Public ReadOnly Property WindFC(ByVal Direction As Single, Optional ByVal Draft As Single = 0) As Force
        Get

            ' Input
            '   Direction:  wind direction (local) (rad)
            '   Draft:      vessel draft (ft)

            'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
            If IsNothing(Draft) Then Draft = msngShipDraft

            WindFC = ForceCoef(Draft, Direction, mcolWindFC)

        End Get
    End Property

    Public ReadOnly Property FWindGlob(Optional ByVal Heading As Object = Nothing, Optional ByVal ReCalculate As Boolean = True) As Force
        Get

            ' Input
            '   Heading:        vessel heading (global) (rad)
            '   ReCalculate:    flag to refresh the memory

            Dim FCrG, FWdG, FWdL, FCrL As Force
            Dim FEnG, FWvG, FWvL, FEnL As Force

            If Not ReCalculate Then

                FWdG = mclsFWindGlob

                'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
            ElseIf IsNothing(Heading) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Heading. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Heading = msngShipHead
                FWdG = mclsFWindGlob
                FWdL = mclsFWindLocl
                FCrG = mclsFCurrGlob
                FCrL = mclsFCurrLocl
                FWvG = mclsFWaveGlob
                FWvL = mclsFWaveLocl
                FEnG = mclsFEnvGlob
                FEnL = mclsFEnvLocl

                'UPGRADE_WARNING: Couldn't resolve default property of object Heading. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call FEnv(FWdL, FWdG, FCrL, FCrG, FWvL, FWvG, FEnL, FEnG, Heading)
            Else
                FWdG = New Force
                FWdL = New Force
                FCrG = New Force
                FCrL = New Force
                FWvG = New Force
                FWvL = New Force
                FEnG = New Force
                FEnL = New Force

                'UPGRADE_WARNING: Couldn't resolve default property of object Heading. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call FEnv(FWdL, FWdG, FCrL, FCrG, FWvL, FWvG, FEnL, FEnG, Heading)

                'UPGRADE_NOTE: Object FWdL may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                FWdL = Nothing
                'UPGRADE_NOTE: Object FCrG may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                FCrG = Nothing
                'UPGRADE_NOTE: Object FCrL may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                FCrL = Nothing
                'UPGRADE_NOTE: Object FWvG may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                FWvG = Nothing
                'UPGRADE_NOTE: Object FWvL may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                FWvL = Nothing
                'UPGRADE_NOTE: Object FEnG may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                FEnG = Nothing
                'UPGRADE_NOTE: Object FEnL may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                FEnL = Nothing

            End If

            FWindGlob = FWdG

        End Get
    End Property

    Public ReadOnly Property FWindLocl(Optional ByVal Heading As Object = Nothing, Optional ByVal ReCalculate As Boolean = True) As Force
        Get

            ' Input
            '   Heading:        vessel heading (global) (rad)
            '   ReCalculate:    flag to refresh the memory

            Dim FCrG, FWdG, FWdL, FCrL As Force
            Dim FEnG, FWvG, FWvL, FEnL As Force


            If Not ReCalculate Then

                FWdL = mclsFWindLocl

                'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
            ElseIf IsNothing(Heading) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Heading. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Heading = msngShipHead
                FWdG = mclsFWindGlob
                FWdL = mclsFWindLocl
                FCrG = mclsFCurrGlob
                FCrL = mclsFCurrLocl
                FWvG = mclsFWaveGlob
                FWvL = mclsFWaveLocl
                FEnG = mclsFEnvGlob
                FEnL = mclsFEnvLocl

                'UPGRADE_WARNING: Couldn't resolve default property of object Heading. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call FEnv(FWdL, FWdG, FCrL, FCrG, FWvL, FWvG, FEnL, FEnG, Heading)
            Else
                FWdG = New Force
                FWdL = New Force
                FCrG = New Force
                FCrL = New Force
                FWvG = New Force
                FWvL = New Force
                FEnG = New Force
                FEnL = New Force

                'UPGRADE_WARNING: Couldn't resolve default property of object Heading. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call FEnv(FWdL, FWdG, FCrL, FCrG, FWvL, FWvG, FEnL, FEnG, Heading)

                'UPGRADE_NOTE: Object FWdG may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                FWdG = Nothing
                'UPGRADE_NOTE: Object FCrG may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                FCrG = Nothing
                'UPGRADE_NOTE: Object FCrL may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                FCrL = Nothing
                'UPGRADE_NOTE: Object FWvG may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                FWvG = Nothing
                'UPGRADE_NOTE: Object FWvL may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                FWvL = Nothing
                'UPGRADE_NOTE: Object FEnG may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                FEnG = Nothing
                'UPGRADE_NOTE: Object FEnL may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                FEnL = Nothing

            End If

            FWindLocl = FWdL

        End Get
    End Property

    Public ReadOnly Property FCurrGlob(Optional ByVal Heading As Object = Nothing, Optional ByVal ReCalculate As Boolean = True) As Force
        Get

            ' Input
            '   Heading:        vessel heading (global) (rad)
            '   ReCalculate:    flag to refresh the memory

            Dim FCrG, FWdG, FWdL, FCrL As Force
            Dim FEnG, FWvG, FWvL, FEnL As Force

            If Not ReCalculate Then

                FCrG = mclsFCurrGlob

                'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
            ElseIf IsNothing(Heading) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Heading. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Heading = msngShipHead
                FWdG = mclsFWindGlob
                FWdL = mclsFWindLocl
                FCrG = mclsFCurrGlob
                FCrL = mclsFCurrLocl
                FWvG = mclsFWaveGlob
                FWvL = mclsFWaveLocl
                FEnG = mclsFEnvGlob
                FEnL = mclsFEnvLocl

                Call FEnv(FWdL, FWdG, FCrL, FCrG, FWvL, FWvG, FEnL, FEnG, Heading)
            Else
                FWdG = New Force
                FWdL = New Force
                FCrG = New Force
                FCrL = New Force
                FWvG = New Force
                FWvL = New Force
                FEnG = New Force
                FEnL = New Force

                Call FEnv(FWdL, FWdG, FCrL, FCrG, FWvL, FWvG, FEnL, FEnG, Heading)

                FWdG = Nothing
                FWdL = Nothing
                FCrL = Nothing
                FWvG = Nothing
                FWvL = Nothing
                FEnG = Nothing
                FEnL = Nothing

            End If

            FCurrGlob = FCrG

        End Get
    End Property

    Public ReadOnly Property FCurrLocl(Optional ByVal Heading As Object = Nothing, Optional ByVal ReCalculate As Boolean = True) As Force
        Get

            ' Input
            '   Heading:        vessel heading (global) (rad)
            '   ReCalculate:    flag to refresh the memory

            Dim FCrG, FWdG, FWdL, FCrL As Force
            Dim FEnG, FWvG, FWvL, FEnL As Force

            If Not ReCalculate Then

                FCrL = mclsFCurrLocl

            ElseIf IsNothing(Heading) Then
                Heading = msngShipHead
                FWdG = mclsFWindGlob
                FWdL = mclsFWindLocl
                FCrG = mclsFCurrGlob
                FCrL = mclsFCurrLocl
                FWvG = mclsFWaveGlob
                FWvL = mclsFWaveLocl
                FEnG = mclsFEnvGlob
                FEnL = mclsFEnvLocl

                Call FEnv(FWdL, FWdG, FCrL, FCrG, FWvL, FWvG, FEnL, FEnG, Heading)
            Else
                FWdG = New Force
                FWdL = New Force
                FCrG = New Force
                FCrL = New Force
                FWvG = New Force
                FWvL = New Force
                FEnG = New Force
                FEnL = New Force

                Call FEnv(FWdL, FWdG, FCrL, FCrG, FWvL, FWvG, FEnL, FEnG, Heading)

                FWdG = Nothing
                FWdL = Nothing
                FCrG = Nothing
                FWvG = Nothing
                FWvL = Nothing
                FEnG = Nothing
                FEnL = Nothing

            End If

            FCurrLocl = FCrL

        End Get
    End Property

    Public ReadOnly Property FWaveGlob(Optional ByVal Heading As Object = Nothing, Optional ByVal ReCalculate As Boolean = True) As Force
        Get

            ' Input
            '   Heading:        vessel heading (global) (rad)
            '   ReCalculate:    flag to refresh the memory

            Dim FCrG, FWdG, FWdL, FCrL As Force
            Dim FEnG, FWvG, FWvL, FEnL As Force


            If Not ReCalculate Then

                FWvG = mclsFWaveGlob

            ElseIf IsNothing(Heading) Then
                Heading = msngShipHead
                FWdG = mclsFWindGlob
                FWdL = mclsFWindLocl
                FCrG = mclsFCurrGlob
                FCrL = mclsFCurrLocl
                FWvG = mclsFWaveGlob
                FWvL = mclsFWaveLocl
                FEnG = mclsFEnvGlob
                FEnL = mclsFEnvLocl

                Call FEnv(FWdL, FWdG, FCrL, FCrG, FWvL, FWvG, FEnL, FEnG, Heading)
            Else
                FWdG = New Force
                FWdL = New Force
                FCrG = New Force
                FCrL = New Force
                FWvG = New Force
                FWvL = New Force
                FEnG = New Force
                FEnL = New Force

                Call FEnv(FWdL, FWdG, FCrL, FCrG, FWvL, FWvG, FEnL, FEnG, Heading)

                FWdG = Nothing
                FWdL = Nothing
                FCrG = Nothing
                FCrL = Nothing
                FWvL = Nothing
                FEnG = Nothing
                FEnL = Nothing

            End If

            FWaveGlob = FWvG

        End Get
    End Property

    Public ReadOnly Property FWaveLocl(Optional ByVal Heading As Object = Nothing, Optional ByVal ReCalculate As Boolean = True) As Force
        Get

            ' Input
            '   Heading:        vessel heading (global) (rad)
            '   ReCalculate:    flag to refresh the memory

            Dim FCrG, FWdG, FWdL, FCrL As Force
            Dim FEnG, FWvG, FWvL, FEnL As Force


            If Not ReCalculate Then

                FWvL = mclsFWaveLocl

            ElseIf IsNothing(Heading) Then
                Heading = msngShipHead
                FWdG = mclsFWindGlob
                FWdL = mclsFWindLocl
                FCrG = mclsFCurrGlob
                FCrL = mclsFCurrLocl
                FWvG = mclsFWaveGlob
                FWvL = mclsFWaveLocl
                FEnG = mclsFEnvGlob
                FEnL = mclsFEnvLocl

                Call FEnv(FWdL, FWdG, FCrL, FCrG, FWvL, FWvG, FEnL, FEnG, Heading)
            Else
                FWdG = New Force
                FWdL = New Force
                FCrG = New Force
                FCrL = New Force
                FWvG = New Force
                FWvL = New Force
                FEnG = New Force
                FEnL = New Force

                Call FEnv(FWdL, FWdG, FCrL, FCrG, FWvL, FWvG, FEnL, FEnG, Heading)

                FWdG = Nothing
                FWdL = Nothing
                FCrG = Nothing
                FCrL = Nothing
                FWvG = Nothing
                FEnG = Nothing
                FEnL = Nothing

            End If

            FWaveLocl = FWvL

        End Get
    End Property

    Public ReadOnly Property FEnvGlob(Optional ByVal Heading As Object = Nothing, Optional ByVal ReCalculate As Boolean = True) As Force
        Get

            ' Input
            '   Heading:        vessel heading (global) (rad)
            '   ReCalculate:    flag to refresh the memory

            Dim FCrG, FWdG, FWdL, FCrL As Force
            Dim FEnG, FWvG, FWvL, FEnL As Force


            If Not ReCalculate Then

                FEnG = mclsFEnvGlob

            ElseIf IsNothing(Heading) Then
                Heading = msngShipHead
                FWdG = mclsFWindGlob
                FWdL = mclsFWindLocl
                FCrG = mclsFCurrGlob
                FCrL = mclsFCurrLocl
                FWvG = mclsFWaveGlob
                FWvL = mclsFWaveLocl
                FEnG = mclsFEnvGlob
                FEnL = mclsFEnvLocl

                Call FEnv(FWdL, FWdG, FCrL, FCrG, FWvL, FWvG, FEnL, FEnG, Heading)
            Else
                FWdG = New Force
                FWdL = New Force
                FCrG = New Force
                FCrL = New Force
                FWvG = New Force
                FWvL = New Force
                FEnG = New Force
                FEnL = New Force

                Call FEnv(FWdL, FWdG, FCrL, FCrG, FWvL, FWvG, FEnL, FEnG, Heading)

                FWdG = Nothing
                FWdL = Nothing
                FCrG = Nothing
                FCrL = Nothing
                FWvG = Nothing
                FWvL = Nothing
                FEnL = Nothing

            End If

            FEnvGlob = FEnG

        End Get
    End Property

    Public ReadOnly Property FEnvLocl(Optional ByVal Heading As Object = Nothing, Optional ByVal Vsurge As Single = 0#, Optional ByVal Vsway As Single = 0#, Optional ByVal ReCalculate As Boolean = True) As Force
        Get

            ' Input
            '   Heading:        vessel heading (global) (rad)
            '   ReCalculate:    flag to refresh the memory

            Dim FCrG, FWdG, FWdL, FCrL As Force
            Dim FEnG, FWvG, FWvL, FEnL As Force


            If Not ReCalculate Then

                FEnL = mclsFEnvLocl

            ElseIf IsNothing(Heading) Then
                Heading = msngShipHead
                FWdG = mclsFWindGlob
                FWdL = mclsFWindLocl
                FCrG = mclsFCurrGlob
                FCrL = mclsFCurrLocl
                FWvG = mclsFWaveGlob
                FWvL = mclsFWaveLocl
                FEnG = mclsFEnvGlob
                FEnL = mclsFEnvLocl

                Call FEnv(FWdL, FWdG, FCrL, FCrG, FWvL, FWvG, FEnL, FEnG, Heading, Vsurge, Vsway)
            Else
                FWdG = New Force
                FWdL = New Force
                FCrG = New Force
                FCrL = New Force
                FWvG = New Force
                FWvL = New Force
                FEnG = New Force
                FEnL = New Force

                Call FEnv(FWdL, FWdG, FCrL, FCrG, FWvL, FWvG, FEnL, FEnG, Heading, Vsurge, Vsway)

                FWdG = Nothing
                FWdL = Nothing
                FCrG = Nothing
                FCrL = Nothing
                FWvG = Nothing
                FWvL = Nothing
                FEnG = Nothing

            End If

            FEnvLocl = FEnL

        End Get
    End Property

    ' methods

    Public Sub EnvironmentAdd()

        Dim EnvNew As New Metocean

        mcolEnvironments.Add(EnvNew)

    End Sub

    Public Function EnvironmentCount() As Integer

        EnvironmentCount = mcolEnvironments.Count()

    End Function

    Public Sub EnvironmentDelete(ByRef vntIndexKey As Object)

        mcolEnvironments.Remove(vntIndexKey)

    End Sub

    Private Sub NoEnvironment()

        Dim envNone As New Metocean

        With envNone
            .Name = "No Environment"
            .Heading = 0#
            .Current.ProfileAdd(0#, 0#)
            With .Wave
                .Height = 0#
                .Period = 0#
            End With
            .Wind.Velocity = 0#
        End With

        With mclsEnvCur
            .Name = "No Environment"
            .Heading = 0#
            .Current.ProfileAdd(0#, 0#)
            With .Wave
                .Height = 0#
                .Period = 0#
            End With
            .Wind.Velocity = 0#
        End With

        mcolEnvironments.Add(envNone)

    End Sub

    Public Function InputFC(ByVal FileNum As Short) As Boolean

        Dim k, i, j, M As Short
        Dim NumDir, NumDraft, NumFreq As Short
        Dim NewFC As EnvForceCoef
        Dim NewCol As Collection
        Dim NewDrft As DrftForceCoef
        Dim IdStr As String
        Dim Freq, Draft, Coef As Single
        Dim TmpStr, TmpStr1 As String
        Dim Fields() As String

        InputFC = False

        On Error GoTo ErrorHandler

        Do While mcolWindFC.Count() > 0
            mcolWindFC.Remove((1))
        Loop
        Do While mcolWaveFC.Count() > 0
            mcolWaveFC.Remove((1))
        Loop
        Do While mcolCurrentFC.Count() > 0
            mcolCurrentFC.Remove((1))
        Loop

        For i = 1 To 2
            Input(FileNum, IdStr)
            Input(FileNum, NumDraft)
            Select Case IdStr
                Case "WIND", "Wind", "wind"
                    NewCol = mcolWindFC
                Case "CURR", "Curr", "curr"
                    NewCol = mcolCurrentFC
                    'Case CStr(Default)
                    'Exit Function
            End Select

            For j = 1 To NumDraft
                Input(FileNum, Draft)
                Input(FileNum, NumDir)

                NewFC = New EnvForceCoef
                NewFC.Draft = Draft

                For k = 1 To NumDir
                    Input(FileNum, Coef)
                    NewFC.FCx.ForceCoefAdd(Coef)
                Next k

                For k = 1 To NumDir
                    Input(FileNum, Coef)
                    NewFC.FCy.ForceCoefAdd(Coef)
                Next k

                For k = 1 To NumDir
                    Input(FileNum, Coef)
                    NewFC.FCm.ForceCoefAdd(Coef)
                Next k

                NewCol.Add(NewFC)
            Next j
        Next i

        Input(FileNum, IdStr)
        Select Case IdStr
            Case "WAVE", "Wave", "wave"
                mblnWave2 = False
                Input(FileNum, NumDraft)
                Input(FileNum, NumFreq)
            Case "WAVE2", "Wave2", "wave2"
                mblnWave2 = True
                NewCol = mcolWave2FC
                Input(FileNum, NumDraft)
                'Case CStr(Default)
                'Exit Function
        End Select

        For i = 1 To NumDraft
            Input(FileNum, Draft)
            Input(FileNum, NumDir)

            If mblnWave2 Then
                NewFC = New EnvForceCoef
                NewFC.Draft = Draft

                For k = 1 To NumDir
                    Input(FileNum, Coef)
                    NewFC.FCx.ForceCoefAdd(Coef)
                Next k

                For k = 1 To NumDir
                    Input(FileNum, Coef)
                    NewFC.FCy.ForceCoefAdd(Coef)
                Next k

                For k = 1 To NumDir
                    Input(FileNum, Coef)
                    NewFC.FCm.ForceCoefAdd(Coef)
                Next k

                NewCol.Add(NewFC)

            Else
                NewDrft = New DrftForceCoef
                NewDrft.Draft = Draft

                For j = 1 To 3
                    Input(FileNum, IdStr)
                    For k = 1 To NumFreq
                        Input(FileNum, Freq)

                        Select Case IdStr
                            Case "SURGE", "Surge", "surge"
                                NewDrft.DrftFCAdd(Freq)
                                For M = 1 To NumDir
                                    Input(FileNum, Coef)
                                    NewDrft.DrftFCItem(k).FCx.ForceCoefAdd(Coef)
                                Next M
                            Case "SWAY", "Sway", "sway"
                                For M = 1 To NumDir
                                    Input(FileNum, Coef)
                                    NewDrft.DrftFCItem(k).FCy.ForceCoefAdd(Coef)
                                Next M
                            Case "YAW", "Yaw", "yaw"
                                For M = 1 To NumDir
                                    Input(FileNum, Coef)
                                    NewDrft.DrftFCItem(k).FCm.ForceCoefAdd(Coef)
                                Next M
                                ' Case CStr(Default)
                                'Exit Function
                        End Select
                    Next k
                Next j

                mcolWaveFC.Add(NewDrft)
            End If
        Next i
        ' end

        InputFC = True
        Exit Function

ErrorHandler:
        Exit Function

    End Function

    Public Function InputEnv(ByVal FileNum As Short) As Boolean

        Dim i, j As Short
        Dim TmpStr, TmpStr2 As String
        Dim NumCur, NumEnv, NumEnvAv As Short
        Dim EnvName As String
        Dim EnvAv As Boolean
        Dim WindVel, WindDir As Single
        Dim WaveTp, WaveHt, WaveDir As Single
        Dim CurDepth, CurDir_Renamed, CurVel As Single

        Dim NewEnv As Metocean

        InputEnv = False

        On Error GoTo ErrorHandler

        Input(FileNum, NumEnv)
        For i = 1 To NumEnv
            Input(FileNum, EnvName)
            Input(FileNum, TmpStr)
            Input(FileNum, TmpStr2)
            WindDir = CSng(TmpStr2)
            WindVel = CSng(TmpStr)
            Input(FileNum, WaveHt)
            Input(FileNum, WaveTp)
            Input(FileNum, TmpStr)
            WaveDir = CSng(TmpStr)
            Input(FileNum, NumCur)
            Input(FileNum, TmpStr)
            CurDir_Renamed = CSng(TmpStr)

            '       determine whether the same criteria have been input
            NumEnvAv = mcolEnvironments.Count()
            EnvAv = False
            For j = 1 To NumEnvAv Step 1
                'UPGRADE_WARNING: Couldn't resolve default property of object mcolEnvironments(j).Name. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If EnvName = mcolEnvironments.Item(j).Name Then
                    EnvAv = True
                    Exit For
                End If
            Next j

            If EnvAv Then
                For j = 1 To NumCur
                    Input(FileNum, TmpStr2)
                    Input(FileNum, TmpStr)
                    CurDepth = CSng(TmpStr2)
                    CurVel = CSng(TmpStr)
                Next j
            Else
                NewEnv = New Metocean
                With NewEnv
                    .Name = EnvName
                    .Wind.Velocity = WindVel * Knots2Ftps
                    .Wind.Heading = WindDir * Degrees2Radians
                    .Wave.Height = WaveHt
                    .Wave.Period = WaveTp
                    .Wave.Heading = WaveDir * Degrees2Radians
                    .Current.Heading = CurDir_Renamed * Degrees2Radians
                End With

                For j = 1 To NumCur
                    Input(FileNum, TmpStr2)
                    Input(FileNum, TmpStr)
                    CurDepth = CSng(TmpStr2)
                    CurVel = CSng(TmpStr)
                    Call NewEnv.Current.ProfileAdd(CurDepth, CurVel * Knots2Ftps)
                Next j

                mcolEnvironments.Add(NewEnv)
            End If
        Next i

        InputEnv = True

ErrorHandler:

    End Function

    Public Function OutputEnv(ByVal FileNum As Short) As Boolean

        Dim i, j As Short
        Dim NumEnv, NumCur As Short

        OutputEnv = False

        NumEnv = mcolEnvironments.Count()
        WriteLine(FileNum, NumEnv)

        For i = 1 To NumEnv
            With mcolEnvironments.Item(i)
                WriteLine(FileNum, .Name)
                WriteLine(FileNum, Format(.Wind.VelCorr * Ftps2Knots, "0.000"), Format(.Wind.Heading * Radians2Degrees, "0.000"))
                WriteLine(FileNum, .Wave.Height, .Wave.Period, Format(.Wave.Heading * Radians2Degrees, "0.000"))

                NumCur = .Current.ProfileCount
                WriteLine(FileNum, NumCur, Format(.Current.Heading * Radians2Degrees, "0.000"))
            End With

            For j = 1 To NumCur
                With mcolEnvironments.Item(i).Current.Profile(j)
                    WriteLine(FileNum, Format(.Depth, "0.000"), Format(.Velocity * Ftps2Knots, "0.000"))
                End With
            Next j
        Next i

        OutputEnv = True

    End Function

    Private Function ForceCoef(ByVal Draft As Single, ByVal Direction As Single, ByRef FrCoef As Collection) As Force

        Dim N, i, Ns As Short
        Dim Rd As Single
        Dim Fy1, Fx1, Fm1 As Single
        Dim Fy2, Fx2, Fm2 As Single
        Dim FC As Force
        FC = New Force

        N = FrCoef.Count()

        If N = 0 Then
            With FC
                .Fx = 0#
                .Fy = 0#
                .MYaw = 0#
            End With
        ElseIf N = 1 Then
            With FC
                .Fx = FrCoef.Item(1).FCx.ForceCoef(Direction)
                .Fy = FrCoef.Item(1).FCy.ForceCoef(Direction)
                .MYaw = FrCoef.Item(1).FCm.ForceCoef(Direction)
            End With
        Else
            If Draft <= FrCoef.Item(1).Draft Then
                With FrCoef.Item(1)
                    Fx1 = .FCx.ForceCoef(Direction)
                    Fy1 = .FCy.ForceCoef(Direction)
                    Fm1 = .FCm.ForceCoef(Direction)
                End With
                With FrCoef.Item(2)
                    Fx2 = .FCx.ForceCoef(Direction)
                    Fy2 = .FCy.ForceCoef(Direction)
                    Fm2 = .FCm.ForceCoef(Direction)
                End With
                Rd = (Draft - FrCoef.Item(1).Draft) / (FrCoef.Item(2).Draft - FrCoef.Item(1).Draft)
            ElseIf Draft >= FrCoef.Item(N).Draft Then
                With FrCoef.Item(N - 1)
                    Fx1 = .FCx.ForceCoef(Direction)
                    Fy1 = .FCy.ForceCoef(Direction)
                    Fm1 = .FCm.ForceCoef(Direction)
                End With
                With FrCoef.Item(N)
                    Fx2 = .FCx.ForceCoef(Direction)
                    Fy2 = .FCy.ForceCoef(Direction)
                    Fm2 = .FCm.ForceCoef(Direction)
                End With
                Rd = (Draft - FrCoef.Item(N - 1).Draft) / (FrCoef.Item(N).Draft - FrCoef.Item(N - 1).Draft)
            Else
                For i = 2 To N
                    If Draft < FrCoef.Item(i).Draft Then
                        Ns = i
                        Exit For
                    End If
                Next
                With FrCoef.Item(Ns - 1)
                    Fx1 = .FCx.ForceCoef(Direction)
                    Fy1 = .FCy.ForceCoef(Direction)
                    Fm1 = .FCm.ForceCoef(Direction)
                End With
                With FrCoef.Item(Ns)
                    Fx2 = .FCx.ForceCoef(Direction)
                    Fy2 = .FCy.ForceCoef(Direction)
                    Fm2 = .FCm.ForceCoef(Direction)
                End With
                Rd = (Draft - FrCoef.Item(Ns - 1).Draft) / (FrCoef.Item(Ns).Draft - FrCoef.Item(Ns - 1).Draft)
            End If
            With FC
                .Fx = Fx1 + (Fx2 - Fx1) * Rd
                .Fy = Fy1 + (Fy2 - Fy1) * Rd
                .MYaw = Fm1 + (Fm2 - Fm1) * Rd
            End With
        End If

        ForceCoef = FC

    End Function

    Private Function DrftCoef(ByVal Draft As Single, ByVal Frequency As Single, ByVal Direction As Single) As Force

        Dim N, i, Ns As Short
        Dim Rd As Single
        Dim Fy1, Fx1, Fm1 As Single
        Dim Fy2, Fx2, Fm2 As Single
        Dim FC As Force

        N = mcolWaveFC.Count()

        If N = 0 Then
            FC = New Force
            With FC
                .Fx = 0#
                .Fy = 0#
                .MYaw = 0#
            End With
        ElseIf N = 1 Then
            FC = mcolWaveFC.Item(1).DrftFC(Frequency, Direction)
        Else
            If Draft <= mcolWaveFC.Item(1).Draft Then
                FC = mcolWaveFC.Item(1).DrftFC(Frequency, Direction)
                With FC
                    Fx1 = .Fx
                    Fy1 = .Fy
                    Fm1 = .MYaw
                End With
                FC = Nothing
                FC = mcolWaveFC.Item(2).DrftFC(Frequency, Direction)
                With FC
                    Fx2 = .Fx
                    Fy2 = .Fy
                    Fm2 = .MYaw
                End With
                Rd = (Draft - mcolWaveFC.Item(1).Draft) / (mcolWaveFC.Item(2).Draft - mcolWaveFC.Item(1).Draft)
            ElseIf Draft >= mcolWaveFC.Item(N).Draft Then
                FC = mcolWaveFC.Item(N - 1).DrftFC(Frequency, Direction)
                With FC
                    Fx1 = .Fx
                    Fy1 = .Fy
                    Fm1 = .MYaw
                End With
                FC = Nothing
                FC = mcolWaveFC.Item(N).DrftFC(Frequency, Direction)
                With FC
                    Fx2 = .Fx
                    Fy2 = .Fy
                    Fm2 = .MYaw
                End With
                Rd = (Draft - mcolWaveFC.Item(N - 1).Draft) / (mcolWaveFC.Item(N).Draft - mcolWaveFC.Item(N - 1).Draft)
            Else
                For i = 2 To N
                    If Draft <= mcolWaveFC.Item(i).Draft Then
                        Ns = i
                        Exit For
                    End If
                Next

                FC = mcolWaveFC.Item(Ns - 1).DrftFC(Frequency, Direction)
                With FC
                    Fx1 = .Fx
                    Fy1 = .Fy
                    Fm1 = .MYaw
                End With
                FC = Nothing
                FC = mcolWaveFC.Item(Ns).DrftFC(Frequency, Direction)
                With FC
                    Fx2 = .Fx
                    Fy2 = .Fy
                    Fm2 = .MYaw
                End With
                Rd = (Draft - mcolWaveFC.Item(Ns - 1).Draft) / (mcolWaveFC.Item(Ns).Draft - mcolWaveFC.Item(Ns - 1).Draft)
            End If
            With FC
                .Fx = Fx1 + (Fx2 - Fx1) * Rd
                .Fy = Fy1 + (Fy2 - Fy1) * Rd
                .MYaw = Fm1 + (Fm2 - Fm1) * Rd
            End With
        End If

        DrftCoef = FC

    End Function

    Private Sub FEnv(ByRef FWindLocl As Force, ByRef FWindGlob As Force, ByRef FCurrLocl As Force, ByRef FCurrGlob As Force, ByRef FWaveLocl As Force, ByRef FWaveGlob As Force, ByRef FEnvLocl As Force, ByRef FEnvGlob As Force, ByVal Heading As Single, Optional ByRef Vsurge As Single = 0#, Optional ByRef Vsway As Single = 0#)

        Dim EnvFC As Force
        Dim FSw, FSg, MY_Renamed As Single
        Dim SSw, SSg, SMY As Single
        Dim EnvDirCorr, EnvDir, MP As Single
        Dim Vy, Vx, Vcorr As Single

        Dim Freq, Sw As Single

        SSg = 0#
        SSw = 0#
        SMY = 0#

        '   windload
        With mclsEnvCur.Wind
            EnvDir = Heading - .Heading
            Vx = .VelCorr * System.Math.Cos(EnvDir) - Vsurge
            Vy = .VelCorr * System.Math.Sin(EnvDir) - Vsway
            Vcorr = System.Math.Sqrt(Vx ^ 2 + Vy ^ 2)
            EnvDirCorr = Atan(Vx, Vy, True)
        End With

        EnvFC = ForceCoef(msngShipDraft, EnvDirCorr, mcolWindFC)

        With EnvFC
            MP = Vcorr ^ 2
            FSg = .Fx * MP
            FSw = .Fy * MP
            MY_Renamed = .MYaw * MP
            SSg = SSg + FSg
            SSw = SSw + FSw
            SMY = SMY + MY_Renamed
        End With
        EnvFC = Nothing

        With FWindLocl
            .Fx = FSg
            .Fy = FSw
            .MYaw = MY_Renamed
        End With

        Call FGlobFrmLocl(FWindGlob, FWindLocl, Heading)

        '   current load
        With mclsEnvCur.Current
            EnvDir = Heading - .Heading
            Vx = .Profile(1).Velocity * System.Math.Cos(EnvDir) - Vsurge
            Vy = .Profile(1).Velocity * System.Math.Sin(EnvDir) - Vsway
            Vcorr = System.Math.Sqrt(Vx ^ 2 + Vy ^ 2)
            EnvDirCorr = Atan(Vx, Vy, True)
        End With

        EnvFC = ForceCoef(msngShipDraft, EnvDirCorr, mcolCurrentFC)

        With EnvFC
            MP = Vcorr ^ 2
            FSg = .Fx * MP
            FSw = .Fy * MP
            MY_Renamed = .MYaw * MP
            SSg = SSg + FSg
            SSw = SSw + FSw
            SMY = SMY + MY_Renamed
        End With
        EnvFC = Nothing

        With FCurrLocl
            .Fx = FSg
            .Fy = FSw
            .MYaw = MY_Renamed
        End With

        Call FGlobFrmLocl(FCurrGlob, FCurrLocl, Heading)

        '   wave load
        EnvDir = Heading - mclsEnvCur.Wave.Heading

        If mblnWave2 Then
            EnvFC = ForceCoef(msngShipDraft, EnvDir, mcolWave2FC)

            With EnvFC
                If mclsEnvCur.Wave.Period = 0 Then
                    MP = 0#
                Else
                    MP = (mclsEnvCur.Wave.Height / mclsEnvCur.Wave.Period) ^ 2
                End If
                FSg = .Fx * MP
                FSw = .Fy * MP
                MY_Renamed = .MYaw * MP
            End With
            EnvFC = Nothing

        Else
            FSg = 0#
            FSw = 0#
            MY_Renamed = 0#

            For Freq = FreqS To FreqE Step FreqD
                Sw = mclsEnvCur.Wave.Spectrum(Freq) * FreqD * 2.0#
                EnvFC = DrftCoef(msngShipDraft, Freq, EnvDir)

                With EnvFC
                    FSg = FSg + .Fx * Sw
                    FSw = FSw + .Fy * Sw
                    MY_Renamed = MY_Renamed + .MYaw * Sw
                End With
                EnvFC = Nothing
            Next
        End If

        SSg = SSg + FSg
        SSw = SSw + FSw
        SMY = SMY + MY_Renamed

        With FWaveLocl
            .Fx = FSg
            .Fy = FSw
            .MYaw = MY_Renamed
        End With

        Call FGlobFrmLocl(FWaveGlob, FWaveLocl, Heading)

        '   total load
        With FEnvLocl
            .Fx = SSg
            .Fy = SSw
            .MYaw = SMY
        End With

        Call FGlobFrmLocl(FEnvGlob, FEnvLocl, Heading)

    End Sub

    Private Sub FGlobFrmLocl(ByRef FGlob As Force, ByRef FLocl As Force, ByRef Heading As Single)

        Dim Alpha As Single

        Alpha = PI / 2# - Heading
		
		With FGlob
			.Fx = System.Math.Cos(Alpha) * FLocl.Fx - System.Math.Sin(Alpha) * FLocl.Fy
			.Fy = System.Math.Sin(Alpha) * FLocl.Fx + System.Math.Cos(Alpha) * FLocl.Fy
			.MYaw = FLocl.MYaw
		End With
		
	End Sub
End Class