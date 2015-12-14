Option Strict Off
Option Explicit On

' Version 8.1
' 2014, Copyright DTCEL, All Rights Reserved
' Jin Liu May 2014

Module modFileUtilities
    Const PropertiesPerRiser As Short = 34
    Const PropertiesPerSegment As Short = 14
    Const PropertiesPerAuxLine As Short = 4
    Const PropertiesPerLatSupport As Short = 3
    Const PropertiesPerSSPoint As Short = 10
    Const PropertiesPerSSPoint0 As Short = 7
    Const DISTANCE As Short = 0
    Const AMP As Short = 1
    Const BM As Short = 2
    Const STRESS As Short = 3
    Const CD As Short = 4
    Const FAT_LIFE As Short = 5

    Private CFAngle, CFLength, CFDim, CFForce As Single
    Private CFStress, CFStiff As Single
    Private CFDensity, CFContent As Single
    Private CFVel, CFChFr As Single
    Private CFViscosity As Single
    Private CFStiffness, CFDamping As Single
    Private CFEI As Single
    Private CFMom As Single

    Private CFLengthS, CFDimS, CFForceS As String
    Private CFStressS, CFStiffS As String
    Private CFDensityS, CFContentS As String
    Private CFVelS, CFChFrS As String
    Private CFViscosityS As String
    Private CFStiffnessS, CFDampingS As String
    Private CFEIS As String
    Private CFMomS As String
    Private category(0 To 5) As String

    Private Sub IniConvertFactor(ByRef CurUnits As VIVAMain.Units)

        If CurUnits = VIVAMain.Units.Metric Then
            CFLength = 1.0#
            CFLengthS = "(m)"
            CFDim = mm2M
            CFAngle = Deg2Rad
            CFForce = 1.0#
            CFForceS = ""
            CFStress = 1000000.0#
            CFStressS = "(MPa)"
            CFStiff = 1.0# / Deg2Rad
            CFDensity = 1.0#
            CFContent = 1.0#
            CFVel = 1.0#
            CFVelS = "(m/s)"
            CFChFr = 1.0#
            CFViscosity = 1.0#
            CFStiffness = 1.0#
            CFDamping = 1.0#
            CFEI = CFStress * CFLength ^ 4
            CFMom = 1.0#
            CFMomS = "(Nm)"
        Else
            CFLength = Ft2M
            CFLengthS = "(ft)"
            CFDim = In2Ft * Ft2M
            CFAngle = Deg2Rad
            CFForce = Lb2N
            CFStress = Lb2N / (In2Ft * Ft2M) ^ 2
            CFStressS = "(psi)"
            CFStiff = Lb2N * Ft2M / Deg2Rad
            CFDensity = Lb2Kg / Ft2M ^ 3
            CFContent = Lb2Kg / Gal2M3
            CFVel = Kn2MPS
            CFVelS = "(knts)"
            CFChFr = Lb2N / Ft2M
            CFViscosity = Ft2M ^ 2
            CFStiffness = Lb2N / Ft2M
            CFDamping = Lb2N / Ft2M
            CFEI = Lb2N * (In2Ft * Ft2M) ^ 2
            CFMom = 1.0# / (Lb2N * Ft2M)
            CFMomS = "(ft-lb)"
        End If

    End Sub


    Public Function GetVIVADirFromRegistry() As String

        Dim PathStr, AppPath, ExePath As String
        Dim hKey As Integer

        'initialization
        ExePath = My.Application.Info.DirectoryPath & "\"
        AppPath = "c:\WinVIVA8"

        PathStr = RegPathWinCurrentVersion() & "\JDM\WinVIVA6.5" 'JLIU TODO
        PathStr = "C:\WinVIVA8"

        If Not RegOpenKey(HKEY_LOCAL_MACHINE, PathStr, hKey) Then
            MsgBox("WinVIVA does not seem to be properly registered. Use " & ExePath & " as directory.", MsgBoxStyle.OkOnly, "WinVIVA - No Registry Entry")
            GetVIVADirFromRegistry = ExePath
        ElseIf Not RegQueryStringValue(hKey, "Install Directory", AppPath) Then
            MsgBox("WinVIVA does not seem to be properly registered. Use " & ExePath & " as directory.", MsgBoxStyle.OkOnly, "WinVIVA - No Registry Entry")
            GetVIVADirFromRegistry = ExePath
        Else
            GetVIVADirFromRegistry = AppPath
        End If

    End Function


    Public Function GetDirFromRegistry(ByVal RegistryName As String, Optional ByVal DefaultDir As String = "") As String

        Dim PathStr, AppPath As String
        Dim hKey As Integer
        Dim msgNoRegWarning, msgNoRegTitle As String

        'initializaiton
        AppPath = ""

        PathStr = RegPathWinCurrentVersion() & RegistryName
        msgNoRegWarning = RegistryName & " does not seem to be properly registered." & " The program assumes " & DefaultDir & _
            " the installation directory."
        msgNoRegTitle = "WinVIVA - No Registry Entry"

        If Not RegOpenKey(HKEY_LOCAL_MACHINE, PathStr, hKey) Then
            MsgBox(msgNoRegWarning, MsgBoxStyle.OkOnly, msgNoRegTitle)
            GetDirFromRegistry = DefaultDir
        ElseIf Not RegQueryStringValue(hKey, "", AppPath) Then
            MsgBox(msgNoRegWarning, MsgBoxStyle.OkOnly, msgNoRegTitle)
            GetDirFromRegistry = DefaultDir
        Else
            GetDirFromRegistry = GetPathName(AppPath)
        End If

    End Function


    Public Function NoSfxFileName(ByVal FileName As String) As String

        Dim DotPos As Short

        DotPos = InStr(FileName, ".")

        If DotPos = 0 Then
            NoSfxFileName = FileName
        Else
            NoSfxFileName = Left(FileName, DotPos - 1)
        End If

    End Function


    Public Function FNOnly(ByVal FQFN As String) As Object

        Dim i, FNLen As Short
        '   Return the filename from a fully-qualified filename string
        For i = Len(FQFN) To 1 Step -1
            If Mid(FQFN, i, 1) = "\" Then
                Exit For
            End If
        Next i

        FNLen = Len(FQFN) - i
        FNOnly = Right(FQFN, FNLen)

    End Function

    Public Function GetVersion(ByRef FileNum As Short) As Single
        On Error GoTo ErrHandler
        Dim KeyName, V1 As String
        'initialization
        KeyName = ""
        V1 = ""
        GetVersion = 0.0

        If EOF(FileNum) Then Exit Function

        Input(FileNum, KeyName)
        Input(FileNum, V1)

        If KeyName = "Version" Then

            GetVersion = CSng(V1)

            Exit Function
        End If
ErrHandler:
        If EOF(FileNum) Then Exit Function
        MsgBox("Error when reading input file. No Version Specified", MsgBoxStyle.OkOnly, " WinVIVA - Read Input File Error")
        Exit Function
    End Function

    Public Sub InputProject(ByRef FileNum As Short)

        On Error GoTo ErrHandler

        Dim VerNum As Single
        Dim V2, KeyName, V1, V3 As String
        Dim NumCurrentPairs, NumSegments As Short
        Dim NumLatSupports, NumAuxLines, NumSSPoints As Short
        Dim Thing As Object
        Dim NumFiles As Short

        Dim i, j As Long
        Dim tmp2, tmp, tmp1, tmp3 As Single

        'initialization
        KeyName = ""
        V1 = ""
        V2 = ""
        V3 = ""
        Thing = Nothing
        VerNum = 6.5
        CurProj.nRisers = 1
        CurProj.RiserId = 1
        Do While Not EOF(FileNum)
            Input(FileNum, KeyName)
            Input(FileNum, V1)
            'MsgBox(KeyName & ":" & V1)
            Select Case KeyName
                'project properties
                Case "Version"
                    VerNum = CSng(V1)

                    If VerNum <= 5.1 Then
                        CurProj.Riser(1).CalculateBendingStiffness = True
                        CurProj.Riser(1).VectoAngle = -1
                    End If

                Case "Title"
                    CurProj.Title = V1
                Case "Desc"
                    CurProj.Desc = V1
                Case "NumIterations"
                    CurProj.NumIterations = CShort(V1)
                Case "NumPoints"
                    CurProj.NumPoints = CShort(V1)
                Case "ShowDOSBox"
                    If V1 = "#TRUE#" Then V1 = "TRUE"
                    If V1 = "#FALSE#" Then V1 = "FALSE"
                    CurProj.ShowDOSBox = CBool(V1)
                Case "FrequencyResponse"
                    CurProj.FrequencyResponse = CFr(V1)
                Case "ProjUnits"
                    CurProj.Units = CShort(V1)
                    IniConvertFactor((CurProj.Units))
                Case "HighRe"
                    CurProj.Riser(1).HiRe = CShort(V1)
                Case "DampingMethod"
                    CurProj.DampingMethod = CShort(V1)

                    '           Water properties
                Case "Density"
                    CurProj.Water.Density = CSng(V1) * CFDensity
                Case "Damping"
                    CurProj.Riser(1).Damping = CSng(V1)
                Case "CurrentProfileName"
                    CurProj.Water.CurrentProfile.ProfileName = V1
                Case "CurrentPairs"
                    NumCurrentPairs = CShort(V1)

                    For i = 1 To NumCurrentPairs
                        Input(FileNum, KeyName)
                        Input(FileNum, V1)

                        If VerNum = 0.0 Then
                            Input(FileNum, KeyName)
                            Input(FileNum, V2)
                            V3 = "0."
                        Else
                            Input(FileNum, KeyName)
                            Input(FileNum, V2)
                            Input(FileNum, V3)
                        End If

                        If i <= MaxCurrentPair Then CurProj.Water.CurrentProfile.Add(CDepth:=CSng(V1) * CFLength, _
                            CVelU:=CSng(V2) * CFVel, CVelV:=CSng(V3) * CFVel)
                    Next i

                    'Riser properties
                Case "RiserType"
                    CurProj.Riser(1).RiserType = CRst(V1)
                Case "TopTension"
                    CurProj.Riser(1).TopTension = CSng(V1) * CFForce
                Case "TopLocation"
                    CurProj.Riser(1).TopLocation = CSng(V1) * CFLength
                Case "MudWeight"
                    CurProj.Riser(1).ContentDensity = CSng(V1) * CFContent
                Case "BuoyancyHead"
                    CurProj.Riser(1).BuoyancyHead = CSng(V1) * CFLength
                Case "UpperBC"
                    CurProj.Riser(1).UpperBC = CBcs(V1)
                Case "UpperStiffness"
                    CurProj.Riser(1).UpperStiffK = CSng(V1) * CFStiff
                Case "LowerBC"
                    CurProj.Riser(1).LowerBC = CBcs(V1)
                Case "LowerStiffness"
                    CurProj.Riser(1).LowerStiffK = CSng(V1) * CFStiff
                Case "FatigueCurveName"
                    CurProj.Riser(1).FatigueCurveName = V1
                Case "FatigueConstantA"
                    CurProj.Riser(1).FatigueConstA = CSng(V1) * CFStress
                Case "FatigueConstantB"
                    CurProj.Riser(1).FatigueConstB = CSng(V1)
                Case "StressConcentrationFactor"
                    CurProj.Riser(1).SCF = CSng(V1)

                    'Segments
                Case "CalculateBendingStiffness"
                    CurProj.Riser(1).CalculateBendingStiffness = CBool(V1)
                Case "VectoAngle"
                    CurProj.Riser(1).VectoAngle = CSng(V1)
                Case "Segments"
                    NumSegments = CShort(V1)

                    For i = 1 To NumSegments
                        For j = 1 To PropertiesPerSegment
                            Input(FileNum, KeyName)
                            Input(FileNum, V1)

                            If i <= MaxSegments Then
                                Select Case KeyName

                                    Case "SegNum"
                                        Thing = CurProj.Riser(1).Segments.Add
                                    Case "NumJoints"
                                        Thing.NumJoints = CShort(V1)
                                    Case "JointLength"
                                        Thing.JointLength = CSng(V1) * CFLength
                                    Case "MainTubeOD"
                                        Thing.MainTubeOD = CSng(V1) * CFDim
                                    Case "WallThickness"
                                        Thing.WallThickness = CSng(V1) * CFDim
                                    Case "DryWeight"
                                        Thing.DryWeight = CSng(V1) * CFForce
                                    Case "WetWeight"
                                        Thing.WetWeight = CSng(V1) * CFForce
                                    Case "BuoyancyModuleDiameter"
                                        Thing.BuoyModDia = CSng(V1) * CFDim
                                    Case "SectionType"
                                        Thing.SectionType = CInt(V1)
                                    Case "StrakesHeight"
                                        Thing.StrakesHeight = CSng(V1) * CFDim
                                    Case "FairThick"

                                        tmp = CSng(V1) * CFDim

                                        If tmp < Thing.MainTubeOD Or tmp < Thing.BuoyModDia Then
                                            Thing.FairThick = 0.0#
                                        Else
                                            Thing.FairThick = tmp
                                        End If
                                    Case "FairChord"

                                        tmp = CSng(V1) * CFDim
                                        If tmp < Thing.MainTubeOD * 1.0001 Or tmp < Thing.BuoyModDia * 1.0001 Then
                                            Thing.FairChord = 0.0#
                                        Else
                                            Thing.FairChord = tmp
                                        End If
                                    Case "ModulusOfElasticity"
                                        Thing.ModulusOfElasticity = CSng(V1) * CFStress
                                    Case "BendingStiffness"
                                        Thing.BendingStiffness = CSng(V1) * CFEI
                                End Select
                            End If
                        Next j

                        If Thing.StrakesHeight > 0.0# Or Thing.SectionType = VIVAMain.ICHARValues.TestedStrakes Then
                            If Thing.SectionType = VIVAMain.ICHARValues.TestedStrakes Then Thing.StrakesHeight = 0.0#
                            Thing.FairThick = 0.0#
                            Thing.FairChord = 0.0#
                        ElseIf Thing.FairThick = 0.0# Then
                            Thing.FairChord = 0.0#
                        ElseIf Thing.FairChord = 0.0# Then
                            Thing.FairThick = 0.0#
                        End If
                    Next i

                    '           AuxLines
                Case "AuxLines"
                    NumAuxLines = CShort(V1)
                    For i = 1 To NumAuxLines * PropertiesPerAuxLine
                        Input(FileNum, KeyName)
                        Input(FileNum, V1)
                        If i <= MaxAuxLines * PropertiesPerAuxLine Then
                            Select Case KeyName
                                '                       New AuxLine
                                Case "AuxLineNum"
                                    Thing = CurProj.Riser(1).AuxLines.Add
                                Case "ID"
                                    Thing.ID = CSng(V1) * CFDim
                                Case "OD"
                                    Thing.OD = CSng(V1) * CFDim
                                Case "ContentDensity"
                                    Thing.ContentDensity = CSng(V1) * CFContent
                            End Select
                        End If
                    Next i
                Case "AuxInEICalcs"
                    If V1 = "#TRUE#" Then V1 = "TRUE"
                    If V1 = "#FALSE#" Then V1 = "FALSE"
                    CurProj.Riser(1).AuxLines.AuxInEICalcs = CBool(V1)

                    '           LatSupports
                Case "LatSupports"
                    NumLatSupports = CShort(V1)
                    For i = 1 To NumLatSupports
                        If i <= MaxNumLatSupports Then
                            For j = 1 To PropertiesPerLatSupport
                                Input(FileNum, KeyName)
                                Input(FileNum, V1)
                                Select Case KeyName
                                    Case "Dist"
                                        tmp1 = CSng(V1) * CFLength
                                    Case "Stiffness"
                                        tmp2 = CSng(V1) * CFStiffness
                                    Case "Damping"
                                        tmp3 = CSng(V1) * CFDamping
                                End Select
                            Next j
                            CurProj.Riser(1).LatSupports.Add(tmp1, tmp2, tmp3)
                        End If
                    Next i

                    '           LMRP
                Case "LMRPHeight"
                    CurProj.Riser(1).BottomPackage.LMRP.Height = CSng(V1) * CFLength
                Case "LMRPDryWeight"
                    CurProj.Riser(1).BottomPackage.LMRP.DryWeight = CSng(V1) * CFForce
                Case "LMRPWetWeight"
                    CurProj.Riser(1).BottomPackage.LMRP.WetWeight = CSng(V1) * CFForce
                Case "LMRPDragDiameter"
                    CurProj.Riser(1).BottomPackage.LMRP.DragDiameter = CSng(V1) * CFDim

                    '           BOP
                Case "BOPHeight"
                    CurProj.Riser(1).BottomPackage.BOP.Height = CSng(V1) * CFLength
                Case "BOPDryWeight"
                    CurProj.Riser(1).BottomPackage.BOP.DryWeight = CSng(V1) * CFForce
                Case "BOPWetWeight"
                    CurProj.Riser(1).BottomPackage.BOP.WetWeight = CSng(V1) * CFForce
                Case "BOPDragDiameter"
                    CurProj.Riser(1).BottomPackage.BOP.DragDiameter = CSng(V1) * CFDim

                    '           Static Solution
                Case "StaticSolution"
                    NumSSPoints = CShort(V1)

                    If VerNum = 0.0 Then
                        For i = 1 To NumSSPoints * PropertiesPerSSPoint0
                            Input(FileNum, KeyName)
                            Input(FileNum, V1)
                            If i <= MaxSSPoints * PropertiesPerSSPoint Then
                                Select Case KeyName
                                    '                       New Solution Point
                                    Case "PointNum"
                                        Thing = CurProj.Riser(1).StaticSolution.Add
                                    Case "Distance"
                                        Thing.Distance = CSng(V1) * CFLength
                                    Case "AxialStaticTension"
                                        Thing.AxialStaticTension = CSng(V1) * CFForce
                                    Case "InPlaneAngle"
                                        Thing.InPlaneAngle = CSng(V1) * CFAngle
                                    Case "PointDepth"
                                        Thing.Depth = CSng(V1) * CFLength
                                    Case "HorCurrentVelocity"
                                        Thing.UCurrentVelocity = CSng(V1) * CFVel
                                        Thing.VCurrentVelocity = 0.0#
                                    Case "ChangeInStaticTension"
                                        Thing.ChangeInStaticTension = CSng(V1) * CFChFr
                                End Select
                            End If
                        Next i

                    Else
                        For i = 1 To NumSSPoints * PropertiesPerSSPoint
                            Input(FileNum, KeyName)
                            Input(FileNum, V1)
                            If i <= MaxSSPoints * PropertiesPerSSPoint Then
                                Select Case KeyName
                                    '                       New Solution Point
                                    Case "PointNum"
                                        Thing = CurProj.Riser(1).StaticSolution.Add
                                    Case "Distance"
                                        Thing.Distance = CSng(V1) * CFLength
                                    Case "AxialStaticTension"
                                        Thing.AxialStaticTension = CSng(V1) * CFForce
                                    Case "InPlaneAngle"
                                        Thing.InPlaneAngle = CSng(V1) * CFAngle
                                    Case "PointDepth"
                                        Thing.Depth = CSng(V1) * CFLength
                                    Case "UCurrentVelocity"
                                        Thing.UCurrentVelocity = CSng(V1) * CFVel
                                    Case "VCurrentVelocity"
                                        Thing.VCurrentVelocity = CSng(V1) * CFVel
                                    Case "ChangeInStaticTension"
                                        Thing.ChangeInStaticTension = CSng(V1) * CFChFr
                                End Select
                            End If
                        Next i
                    End If

                Case "Mode-def"
                    CurProj.Riser(1).Mode = CShort(V1)
                    Input(FileNum, V1)
                    CurProj.Water.Viscosity = CSng(V1) * CFViscosity
                Case "FreqFile"
                    CurProj.Riser(1).FreqFile = V1
                Case "ModeFile"
                    CurProj.Riser(1).ModeFile = V1
                Case "CurvFile"
                    CurProj.Riser(1).CurvFile = V1

            End Select
        Loop
        '   bottom package
        If CurProj.Riser(1).BottomPackage.LMRP.DryWeight = 0.0# Or CurProj.Riser(1).BottomPackage.LMRP.DragDiameter = 0.0# Then CurProj.Riser(1).BottomPackage.LMRP.Height = 0.0#
        If CurProj.Riser(1).BottomPackage.BOP.DryWeight = 0.0# Or CurProj.Riser(1).BottomPackage.BOP.DragDiameter = 0.0# Then CurProj.Riser(1).BottomPackage.BOP.Height = 0.0#
        Exit Sub

ErrHandler:
        If EOF(FileNum) Then Exit Sub
        MsgBox("Error when reading input file. The data may be imcomplete or corrupted", MsgBoxStyle.OkOnly, " WinVIVA - Read Input File Error")
        Exit Sub

    End Sub

    Public Sub InputProject8(ByRef FileNum As Short)

        On Error GoTo ErrHandler

        Dim VerNum As Single
        Dim V2, KeyName, V1, V3 As String
        Dim NumCurrentPairs, NumSegments As Short
        Dim NumLatSupports, NumAuxLines, NumSSPoints As Short
        Dim Thing As Object

        Dim NumFiles As Short

        Dim i, j, ir, ip As Long
        Dim tmp2, tmp, tmp1, tmp3 As Single

        'initialization
        KeyName = ""
        V1 = ""
        V2 = ""
        V3 = ""
        VerNum = 8.1
        Thing = Nothing
        ' MsgBox(CurProj.nRisers)
        Do While Not EOF(FileNum)
            Input(FileNum, KeyName)
            Input(FileNum, V1)
            'MsgBox(KeyName & ":" & V1)
            Select Case KeyName
                '           project properties
                Case "Version"
                    VerNum = CSng(V1)
                Case "Title"
                    CurProj.Title = V1
                Case "Desc"
                    CurProj.Desc = V1
                Case "ProjUnits"
                    CurProj.Units = CShort(V1)
                    IniConvertFactor((CurProj.Units))
                Case "DampingMethod"
                    CurProj.DampingMethod = CShort(V1)
                Case "NumIterations"
                    CurProj.NumIterations = CShort(V1)
                Case "NumPoints"
                    CurProj.NumPoints = CShort(V1)
                Case "ShowDOSBox"
                    If V1 = "#TRUE#" Then V1 = "TRUE"
                    If V1 = "#FALSE#" Then V1 = "FALSE"
                    CurProj.ShowDOSBox = CBool(V1)
                    '           Water properties
                Case "Density"
                    CurProj.Water.Density = CSng(V1) * CFDensity
                Case "Vicosity"
                    CurProj.Water.Viscosity = CSng(V1) * CFViscosity
                Case "FrequencyResponse"
                    CurProj.FrequencyResponse = CFr(V1)

                Case "CurrentProfileName"
                    CurProj.Water.CurrentProfile.ProfileName = V1
                Case "CurrentPairs"
                    NumCurrentPairs = CShort(V1)

                    For i = 1 To NumCurrentPairs
                        Input(FileNum, KeyName)
                        Input(FileNum, V1)

                        If VerNum = 0.0 Then
                            Input(FileNum, KeyName)
                            Input(FileNum, V2)
                            V3 = "0."
                        Else
                            Input(FileNum, KeyName)
                            Input(FileNum, V2)
                            Input(FileNum, V3)
                        End If

                        If i <= MaxCurrentPair Then CurProj.Water.CurrentProfile.Add(CDepth:=CSng(V1) * CFLength, _
                            CVelU:=CSng(V2) * CFVel, CVelV:=CSng(V3) * CFVel)
                    Next i
                Case "RiserFLocY"
                    CurProj.RiserFLocY = CSng(V1) * CFLength
                Case "RiserFLocZ"
                    CurProj.RiserFLocZ = CSng(V1) * CFLength
                Case "RiserRLocY"
                    CurProj.RiserRLocY = CSng(V1) * CFLength
                Case "RiserRLocZ"
                    CurProj.RiserRLocZ = CSng(V1) * CFLength
                Case "NumRisers"
                    CurProj.nRisers = CShort(V1)
           

                    For ir = 1 To CurProj.nRisers
                        ip = 0
                        CurProj.RiserId = ir

                        While ip < PropertiesPerRiser
                            ip = ip + 1
                            Input(FileNum, KeyName)

                            Input(FileNum, V1)
                            'MsgBox(KeyName & ":" & V1 & "-" & ip)
                            Select Case KeyName
                                Case "HighRe"
                                    CurProj.Riser(ir).HiRe = CShort(V1)
                                Case "Mode-def"
                                    CurProj.Riser(ir).Mode = CShort(V1)
                                Case "FreqFile"
                                    CurProj.Riser(ir).FreqFile = V1
                                Case "ModeFile"
                                    CurProj.Riser(ir).ModeFile = V1
                                Case "CurvFile"
                                    CurProj.Riser(ir).CurvFile = V1
                                Case "Damping"
                                    CurProj.Riser(ir).Damping = CSng(V1)
                                    'Riser properties
                                Case "RiserType"
                                    CurProj.Riser(ir).RiserType = CRst(V1)
                                Case "TopTension"
                                    CurProj.Riser(ir).TopTension = CSng(V1) * CFForce
                                Case "TopLocation"
                                    CurProj.Riser(ir).TopLocation = CSng(V1) * CFLength
                                Case "MudWeight"
                                    CurProj.Riser(ir).ContentDensity = CSng(V1) * CFContent
                                Case "BuoyancyHead"
                                    CurProj.Riser(ir).BuoyancyHead = CSng(V1) * CFLength
                                Case "UpperBC"
                                    CurProj.Riser(ir).UpperBC = CBcs(V1)
                                Case "UpperStiffness"
                                    CurProj.Riser(ir).UpperStiffK = CSng(V1) * CFStiff
                                Case "LowerBC"
                                    CurProj.Riser(ir).LowerBC = CBcs(V1)
                                Case "LowerStiffness"
                                    CurProj.Riser(ir).LowerStiffK = CSng(V1) * CFStiff
                                Case "FatigueCurveName"
                                    CurProj.Riser(ir).FatigueCurveName = V1
                                Case "FatigueConstantA"
                                    CurProj.Riser(ir).FatigueConstA = CSng(V1) * CFStress
                                Case "FatigueConstantB"
                                    CurProj.Riser(ir).FatigueConstB = CSng(V1)
                                Case "StressConcentrationFactor"
                                    CurProj.Riser(ir).SCF = CSng(V1)

                                    'Segments
                                Case "CalculateBendingStiffness"
                                    CurProj.Riser(ir).CalculateBendingStiffness = CBool(V1)
                                Case "VectoAngle"
                                    CurProj.Riser(ir).VectoAngle = CSng(V1)
                                Case "Segments"
                                    NumSegments = CShort(V1)

                                    For i = 1 To NumSegments
                                        For j = 1 To PropertiesPerSegment
                                            Input(FileNum, KeyName)
                                            Input(FileNum, V1)
                                            ' MsgBox("Seg:" & KeyName & ":" & V1)
                                            If i <= MaxSegments Then
                                                Select Case KeyName

                                                    Case "SegNum"
                                                        Thing = CurProj.Riser(ir).Segments.Add
                                                    Case "NumJoints"
                                                        Thing.NumJoints = CShort(V1)
                                                    Case "JointLength"
                                                        Thing.JointLength = CSng(V1) * CFLength
                                                    Case "MainTubeOD"
                                                        Thing.MainTubeOD = CSng(V1) * CFDim
                                                    Case "WallThickness"
                                                        Thing.WallThickness = CSng(V1) * CFDim
                                                    Case "DryWeight"
                                                        Thing.DryWeight = CSng(V1) * CFForce
                                                    Case "WetWeight"
                                                        Thing.WetWeight = CSng(V1) * CFForce
                                                    Case "BuoyancyModuleDiameter"
                                                        Thing.BuoyModDia = CSng(V1) * CFDim
                                                    Case "SectionType"
                                                        Thing.SectionType = CInt(V1)
                                                    Case "StrakesHeight"
                                                        Thing.StrakesHeight = CSng(V1) * CFDim
                                                    Case "FairThick"
                                                        'comment out 4/10/2015 JLIU TODO
                                                        tmp = CSng(V1) * CFDim
                                                        'If tmp < Thing.MainTubeOD Or tmp < Thing.BuoyModDia Then
                                                        'Thing.FairThick = 0.0#
                                                        'Else
                                                        Thing.FairThick = tmp
                                                        'End If
                                                    Case "FairChord"
                                                        'comment out 4/10/2015 JLIU TODO
                                                        tmp = CSng(V1) * CFDim
                                                        'If tmp < Thing.MainTubeOD * 1.0001 Or tmp < Thing.BuoyModDia * 1.0001 Then
                                                        'Thing.FairChord = 0.0#
                                                        'Else
                                                        Thing.FairChord = tmp
                                                        'End If
                                                    Case "ModulusOfElasticity"
                                                        Thing.ModulusOfElasticity = CSng(V1) * CFStress
                                                    Case "BendingStiffness"
                                                        Thing.BendingStiffness = CSng(V1) * CFEI
                                                End Select
                                            End If
                                        Next j
                                        'comment out 4/10/2015 JLIU TODO
                                        ' If Thing.StrakesHeight > 0.0# Or Thing.SectionType = VIVAMain.ICHARValues.TestedStrakes Then
                                        'If Thing.SectionType = VIVAMain.ICHARValues.TestedStrakes Then Thing.StrakesHeight = 0.0#
                                        'Thing.FairThick = 0.0#
                                        'Thing.FairChord = 0.0#
                                        'ElseIf Thing.FairThick = 0.0# Then
                                        'Thing.FairChord = 0.0#
                                        'ElseIf Thing.FairChord = 0.0# Then
                                        'Thing.FairThick = 0.0#
                                        'End If
                                    Next i

                                    '           AuxLines
                                Case "AuxLines"
                                    NumAuxLines = CShort(V1)
                                    For i = 1 To NumAuxLines * PropertiesPerAuxLine
                                        Input(FileNum, KeyName)
                                        Input(FileNum, V1)
                                        'MsgBox(FileNum & "," & KeyName & ":" & V1)
                                        If i <= MaxAuxLines * PropertiesPerAuxLine Then
                                            Select Case KeyName
                                                '                       New AuxLine
                                                Case "AuxLineNum"
                                                    Thing = CurProj.Riser(ir).AuxLines.Add
                                                Case "ID"
                                                    Thing.ID = CSng(V1) * CFDim
                                                Case "OD"
                                                    Thing.OD = CSng(V1) * CFDim
                                                Case "ContentDensity"
                                                    Thing.ContentDensity = CSng(V1) * CFContent
                                            End Select
                                        End If
                                    Next i
                                Case "AuxInEICalcs"
                                    If V1 = "#TRUE#" Then V1 = "TRUE"
                                    If V1 = "#FALSE#" Then V1 = "FALSE"
                                    CurProj.Riser(ir).AuxLines.AuxInEICalcs = CBool(V1)

                                    '           LatSupports
                                Case "LatSupports"
                                    NumLatSupports = CShort(V1)
                                    For i = 1 To NumLatSupports
                                        If i <= MaxNumLatSupports Then
                                            For j = 1 To PropertiesPerLatSupport
                                                Input(FileNum, KeyName)
                                                Input(FileNum, V1)
                                                Select Case KeyName
                                                    Case "Dist"
                                                        tmp1 = CSng(V1) * CFLength
                                                    Case "Stiffness"
                                                        tmp2 = CSng(V1) * CFStiffness
                                                    Case "Damping"
                                                        tmp3 = CSng(V1) * CFDamping
                                                End Select
                                            Next j
                                            CurProj.Riser(ir).LatSupports.Add(tmp1, tmp2, tmp3)
                                        End If
                                    Next i

                                    '           LMRP
                                Case "LMRPHeight"
                                    CurProj.Riser(ir).BottomPackage.LMRP.Height = CSng(V1) * CFLength
                                Case "LMRPDryWeight"
                                    CurProj.Riser(ir).BottomPackage.LMRP.DryWeight = CSng(V1) * CFForce
                                Case "LMRPWetWeight"
                                    CurProj.Riser(ir).BottomPackage.LMRP.WetWeight = CSng(V1) * CFForce
                                Case "LMRPDragDiameter"
                                    CurProj.Riser(ir).BottomPackage.LMRP.DragDiameter = CSng(V1) * CFDim

                                    '           BOP
                                Case "BOPHeight"
                                    CurProj.Riser(ir).BottomPackage.BOP.Height = CSng(V1) * CFLength
                                Case "BOPDryWeight"
                                    CurProj.Riser(ir).BottomPackage.BOP.DryWeight = CSng(V1) * CFForce
                                Case "BOPWetWeight"
                                    CurProj.Riser(ir).BottomPackage.BOP.WetWeight = CSng(V1) * CFForce
                                Case "BOPDragDiameter"
                                    CurProj.Riser(ir).BottomPackage.BOP.DragDiameter = CSng(V1) * CFDim

                                    '           Static Solution
                                Case "StaticSolution"
                                    NumSSPoints = CShort(V1)

                                    If VerNum = 0.0 Then
                                        For i = 1 To NumSSPoints * PropertiesPerSSPoint0
                                            Input(FileNum, KeyName)
                                            Input(FileNum, V1)
                                            If i <= MaxSSPoints * PropertiesPerSSPoint Then
                                                Select Case KeyName
                                                    '                       New Solution Point
                                                    Case "PointNum"
                                                        Thing = CurProj.Riser(ir).StaticSolution.Add
                                                    Case "Distance"
                                                        Thing.Distance = CSng(V1) * CFLength
                                                    Case "AxialStaticTension"
                                                        Thing.AxialStaticTension = CSng(V1) * CFForce
                                                    Case "InPlaneAngle"
                                                        Thing.InPlaneAngle = CSng(V1) * CFAngle
                                                    Case "PointDepth"
                                                        Thing.Depth = CSng(V1) * CFLength
                                                    Case "UCurrentVelocity"
                                                        Thing.UCurrentVelocity = CSng(V1) * CFVel

                                                        Thing.VCurrentVelocity = 0.0
                                                    Case "ChangeInStaticTension"
                                                        Thing.ChangeInStaticTension = CSng(V1) * CFChFr
                                                End Select
                                            End If
                                        Next i

                                    Else
                                        For i = 1 To NumSSPoints * PropertiesPerSSPoint
                                            Input(FileNum, KeyName)
                                            Input(FileNum, V1)
                                            If i <= MaxSSPoints * PropertiesPerSSPoint Then
                                                Select Case KeyName
                                                    '                       New Solution Point
                                                    Case "PointNum"
                                                        Thing = CurProj.Riser(ir).StaticSolution.Add
                                                    Case "Distance"
                                                        Thing.Distance = CSng(V1) * CFLength
                                                    Case "AxialStaticTension"
                                                        Thing.AxialStaticTension = CSng(V1) * CFForce
                                                    Case "InPlaneAngle"
                                                        Thing.InPlaneAngle = CSng(V1) * CFAngle
                                                    Case "Ydist"
                                                        Thing.Ydist = CSng(V1) * CFLength
                                                    Case "Zdist"
                                                        Thing.Zdist = CSng(V1) * CFLength
                                                    Case "PointDepth"
                                                        Thing.Depth = CSng(V1) * CFLength
                                                    Case "UCurrentVelocity"
                                                        Thing.UCurrentVelocity = CSng(V1) * CFVel
                                                    Case "VCurrentVelocity"
                                                        Thing.VCurrentVelocity = CSng(V1) * CFVel
                                                    Case "ChangeInStaticTension"
                                                        Thing.ChangeInStaticTension = CSng(V1) * CFChFr
                                                End Select
                                            End If
                                        Next i
                                    End If

                            End Select
                        End While
                    Next ir
            End Select
        Loop
        '   bottom package
        For ir = 1 To CurProj.nRisers
            If CurProj.Riser(ir).BottomPackage.LMRP.DryWeight = 0.0# Or CurProj.Riser(ir).BottomPackage.LMRP.DragDiameter = 0.0# Then CurProj.Riser(ir).BottomPackage.LMRP.Height = 0.0#
            If CurProj.Riser(ir).BottomPackage.BOP.DryWeight = 0.0# Or CurProj.Riser(ir).BottomPackage.BOP.DragDiameter = 0.0# Then CurProj.Riser(ir).BottomPackage.BOP.Height = 0.0#
        Next ir
        Exit Sub

ErrHandler:
        If EOF(FileNum) Then Exit Sub
        MsgBox("Error when reading input file. The data may be imcomplete or corrupted", MsgBoxStyle.OkOnly, " WinVIVA - Read Input File Error")
        Exit Sub

    End Sub


    Public Sub OutputProject(ByRef FileNum As Short)

        Dim Thing As Object
        Dim i, ir As Short

        '   Properties of the Project object
        With CurProj

            Write(FileNum, "Version")
            PrintLine(FileNum, Format(My.Application.Info.Version.Major, "0") & "." & Format(My.Application.Info.Version.Minor, "0"))
            WriteLine(FileNum, "Title", .Title)
            WriteLine(FileNum, "Desc", .Desc)

            WriteLine(FileNum, "ProjUnits", .Units)

            IniConvertFactor((.Units))

            WriteLine(FileNum, "DampingMethod", .DampingMethod)
            WriteLine(FileNum, "NumIterations", .NumIterations)
            WriteLine(FileNum, "NumPoints", .NumPoints)
            WriteLine(FileNum, "ShowDOSBox", CStr(.ShowDOSBox))
            WriteLine(FileNum, "FrequencyResponse", .FrequencyResponse)
            With CurProj.Water
                WriteLine(FileNum, "Density", .Density / CFDensity)
                WriteLine(FileNum, "Vicosity", .Viscosity / CFViscosity)
                WriteLine(FileNum, "CurrentProfileName", .CurrentProfile.ProfileName)
                WriteLine(FileNum, "CurrentPairs", .CurrentProfile.Count)
                For Each Thing In .CurrentProfile
                    WriteLine(FileNum, "Depth", Thing.Depth / CFLength)
                    WriteLine(FileNum, "Velocities", Thing.VelU / CFVel, Thing.VelV / CFVel) ' jjx    'Thing.Velocity / CFVel
                Next Thing
            End With
            WriteLine(FileNum, "RiserFLocY", .RiserFLocY / CFLength)
            WriteLine(FileNum, "RiserFLocZ", .RiserFLocZ / CFLength)
            WriteLine(FileNum, "RiserRLocY", .RiserRLocY / CFLength)
            WriteLine(FileNum, "RiserRLocZ", .RiserRLocZ / CFLength)
            WriteLine(FileNum, "NumRisers", .nRisers)
 
            For ir = 1 To .nRisers
                .RiserId = ir

                With CurProj.Riser(ir)
 
                    WriteLine(FileNum, "HighRe", .HiRe)
                    WriteLine(FileNum, "Mode-def", .Mode)
                    WriteLine(FileNum, "FreqFile", .FreqFile)
                    WriteLine(FileNum, "ModeFile", .ModeFile)
                    WriteLine(FileNum, "CurvFile", .CurvFile)
                    WriteLine(FileNum, "Damping", .Damping)

                    WriteLine(FileNum, "RiserType", .RiserType)
                    WriteLine(FileNum, "TopTension", .TopTension / CFForce)
                    WriteLine(FileNum, "TopLocation", .TopLocation / CFLength)
                    WriteLine(FileNum, "MudWeight", .ContentDensity / CFContent)
                    WriteLine(FileNum, "BuoyancyHead", .BuoyancyHead / CFLength)
                    WriteLine(FileNum, "UpperBC", .UpperBC)
                    WriteLine(FileNum, "UpperStiffness", .UpperStiffK / CFStiff)
                    WriteLine(FileNum, "LowerBC", .LowerBC)
                    WriteLine(FileNum, "LowerStiffness", .LowerStiffK / CFStiff)
                    WriteLine(FileNum, "FatigueCurveName", .FatigueCurveName)
                    WriteLine(FileNum, "FatigueConstantA", .FatigueConstA / CFStress)
                    WriteLine(FileNum, "FatigueConstantB", .FatigueConstB)
                    WriteLine(FileNum, "StressConcentrationFactor", .SCF)
                    '       Segments; write one line for each property
                    WriteLine(FileNum, "CalculateBendingStiffness", CStr(.CalculateBendingStiffness))
                    WriteLine(FileNum, "VectoAngle", .VectoAngle)
                    WriteLine(FileNum, "Segments", .Segments.Count)

                    For Each Thing In .Segments
                        With Thing
                            WriteLine(FileNum, "SegNum", .SegNum)
                            WriteLine(FileNum, "NumJoints", .NumJoints)
                            WriteLine(FileNum, "JointLength", .JointLength / CFLength)
                            WriteLine(FileNum, "MainTubeOD", .MainTubeOD / CFDim)
                            WriteLine(FileNum, "WallThickness", .WallThickness / CFDim)
                            WriteLine(FileNum, "DryWeight", .DryWeight / CFForce)
                            WriteLine(FileNum, "WetWeight", .WetWeight / CFForce)
                            WriteLine(FileNum, "BuoyancyModuleDiameter", .BuoyModDia / CFDim)
                            WriteLine(FileNum, "SectionType", .SectionType)
                            WriteLine(FileNum, "StrakesHeight", .StrakesHeight / CFDim)
                            WriteLine(FileNum, "FairThick", .FairThick / CFDim)
                            WriteLine(FileNum, "FairChord", .FairChord / CFDim)
                            WriteLine(FileNum, "ModulusOfElasticity", .ModulusOfElasticity / CFStress)
                            WriteLine(FileNum, "BendingStiffness", .bendingstiffness / CFEI)
                        End With
                    Next Thing

                    '       AuxLines;
                    WriteLine(FileNum, "AuxLines", .AuxLines.Count)
                    For Each Thing In .AuxLines
                        With Thing
                            WriteLine(FileNum, "AuxLineNum", .AuxLineNum)
                            WriteLine(FileNum, "ID", .ID / CFDim)
                            WriteLine(FileNum, "OD", .OD / CFDim)
                            WriteLine(FileNum, "ContentDensity", .ContentDensity / CFContent)
                        End With
                    Next Thing

                    WriteLine(FileNum, "AuxInEICalcs", CStr(.AuxLines.AuxInEICalcs))

                    '       LatSupports;
                    WriteLine(FileNum, "LatSupports", .LatSupports.Count)

                    For i = 1 To .LatSupports.Count
                        Thing = .LatSupports.Item(i)
                        With Thing
                            WriteLine(FileNum, "Dist", .Dist / CFLength)
                            WriteLine(FileNum, "Stiffness", .Stiffness / CFStiffness)
                            WriteLine(FileNum, "Damping", .Damping / CFDamping)
                        End With
                    Next i

                    '       LMRP;
                    With .BottomPackage.LMRP
                        WriteLine(FileNum, "LMRPHeight", .Height / CFLength)
                        WriteLine(FileNum, "LMRPDryWeight", .DryWeight / CFForce)
                        WriteLine(FileNum, "LMRPWetWeight", .WetWeight / CFForce)
                        WriteLine(FileNum, "LMRPDragDiameter", .DragDiameter / CFDim)
                    End With

                    '       BOP;
                    With .BottomPackage.BOP
                        WriteLine(FileNum, "BOPHeight", .Height / CFLength)
                        WriteLine(FileNum, "BOPDryWeight", .DryWeight / CFForce)
                        WriteLine(FileNum, "BOPWetWeight", .WetWeight / CFForce)
                        WriteLine(FileNum, "BOPDragDiameter", .DragDiameter / CFDim)
                    End With

                    '       SCR Static Solution
                    WriteLine(FileNum, "StaticSolution", .StaticSolution.Count)
                    For Each Thing In .StaticSolution
                        With Thing
                            WriteLine(FileNum, "PointNum", .PointNum)
                            WriteLine(FileNum, "Distance", .Distance / CFLength)
                            WriteLine(FileNum, "AxialStaticTension", .AxialStaticTension / CFForce)
                            WriteLine(FileNum, "InPlaneAngle", .InPlaneAngle / CFAngle)
                            WriteLine(FileNum, "Ydist", .Ydist / CFLength)
                            WriteLine(FileNum, "Zdist", .Zdist / CFLength)
                            WriteLine(FileNum, "PointDepth", .Depth / CFLength)
                            WriteLine(FileNum, "UCurrentVelocity", .UCurrentVelocity / CFVel)
                            WriteLine(FileNum, "VCurrentVelocity", .VCurrentVelocity / CFVel)
                            WriteLine(FileNum, "ChangeInStaticTension", .ChangeInStaticTension / CFChFr)
                        End With
                    Next Thing

                End With
            Next ir
            FileClose(FileNum)
        End With

    End Sub


    Public Function WriteDOSVIVAInput() As Boolean

        On Error GoTo ErrHandler

        Dim ICHOICE, NTYP, NumCurPairs As Short

        Dim BPID, BPEI As Single
        Dim AuxAddedMass, AuxMass, AuxWetWeight As Single
        Dim AuxEI As Single
        Dim SegLength, StackWeight As Single
        Dim AxialTension, TmpLAW As Single
        Dim AddedMass, Mass, WetWeight As Single
        Dim HTD, DragDiam, EI, LAW As Single
        Dim WaterDepth, TotalLength As Single
        Dim SegTopAbvWL, SegAuxWetWt As Single
        Dim SegMidWD, SegMidCurU, SegMidCurV, SegMidCurAngle As Single
        Dim SegType As VIVAMain.ICHARValues
        Dim IniCurDepth, SSDepth As Single

        Dim DOSFile As String
        Dim seg As Segment
        Dim SP As StaticSolutionPoint
        Dim CP As CurrentPair

        Dim i, FileNum, ir, ix As Short

        WriteDOSVIVAInput = False
    
        For ir = 1 To CurProj.nRisers
            ix = ir * CurProj.nRisers - ir
            With CurProj.Riser(ir)
                'conditions.in
                FileNum = FreeFile()
                FileOpen(FileNum, VIVADIR & OutputFiles(ix, conditions_in), OpenMode.Output, OpenAccess.Write)
                WriteLine(FileNum, .Mode)
                WriteLine(FileNum, CurProj.Water.Viscosity)
                If CurProj.Riser(ir).HiRe = HiReData.NoUse Then
                    PrintLine(FileNum, .HiRe, 1)
                Else
                    PrintLine(FileNum, .HiRe, 9)
                End If

                'riser boundary condition codes
                PrintLine(FileNum, CShort(.UpperBC), CShort(.LowerBC))
                'Up and/or lower rotational spring stiffness
                If .UpperBC = VIVAMain.BoundaryConditions.Spring Then
                    PrintLine(FileNum, .UpperStiffC)
                End If
                If .LowerBC = VIVAMain.BoundaryConditions.Spring Then
                    PrintLine(FileNum, .LowerStiffC)
                End If
                'frequency response code
                'PrintLine(FileNum, CShort(CurProj.FrequencyResponse))
                PrintLine(FileNum, CurProj.DampingMethod)
                FileClose(FileNum)
                'end conditions.in
                'rispre.in
                FileNum = FreeFile()
                FileOpen(FileNum, VIVADIR & OutputFiles(ix, rispre_in), OpenMode.Output, OpenAccess.Write)
                'output file name
                Select Case ix
                    Case 0
                        PrintLine(FileNum, "'RISOUT'") 'only one Risers
                    Case 1
                        PrintLine(FileNum, "'RISOUTF'") 'two risers Front
                    Case 2
                        PrintLine(FileNum, "'RISOUTR'") 'two risers rear
                End Select
                'number of segments
                PrintLine(FileNum, .TotalSegments)
                'bottom package
                If .LowerBC = VIVAMain.BoundaryConditions.Free Then
                    With .Segments.Item(1)
                        BPID = .MainTubeOD - .WallThickness * 2
                        BPEI = .EI / .MainTubeOD ^ 4
                    End With
                    StackWeight = 0.0#
                    With .BottomPackage.BOP
                        If .Height > 0.0# Then
                            PrintLine(FileNum, .Height)
                            PrintLine(FileNum, .MassPerLength)
                            PrintLine(FileNum, .AddedMassPerLength)
                            PrintLine(FileNum, .WetWeightPerLength)
                            PrintLine(FileNum, .DragDiameter)
                            PrintLine(FileNum, .DragDiameter, (.DragDiameter - BPID) / 2.0#)
                            BPEI = BPEI * .DragDiameter ^ 4
                            PrintLine(FileNum, BPEI)
                            PrintLine(FileNum, 1)
                            PrintLine(FileNum, 1)
                            StackWeight = StackWeight + .WetWeightPerLength * .Height
                        End If
                    End With
                    With .BottomPackage.LMRP
                        If .Height > 0.0# Then
                            PrintLine(FileNum, .Height)
                            PrintLine(FileNum, .MassPerLength)
                            PrintLine(FileNum, .AddedMassPerLength)
                            PrintLine(FileNum, .WetWeightPerLength)
                            PrintLine(FileNum, .DragDiameter)
                            PrintLine(FileNum, .DragDiameter, (.DragDiameter - BPID) / 2.0#)
                            BPEI = BPEI * .DragDiameter ^ 4
                            PrintLine(FileNum, BPEI)
                            PrintLine(FileNum, 1)
                            PrintLine(FileNum, 1)
                            StackWeight = StackWeight + .WetWeightPerLength * .Height
                        End If
                    End With
                End If

                'auxilliary contribution (doesn't change with distance)
                With .AuxLines
                    AuxMass = .MassPerLength
                    AuxAddedMass = .AddedMassPerLength
                    AuxWetWeight = .WetWeightPerLength
                    If .AuxInEICalcs Then
                        AuxEI = .EI
                    Else
                        AuxEI = 0.0#
                    End If
                End With

                'segment data
                LAW = .TopLocation
                WaterDepth = 0.0#
                For Each seg In .Segments
                    WaterDepth = WaterDepth + seg.JointLength * seg.NumJoints
                Next seg
                WaterDepth = WaterDepth - LAW
                TotalLength = 0.0#

                For Each seg In .Segments
                    With seg
                        SegLength = .JointLength * .NumJoints
                        TotalLength = TotalLength + SegLength
                        SegTopAbvWL = TotalLength - WaterDepth

                        If .SectionType >= ICHARValues.VetcoRiser0 And .SectionType <= ICHARValues.VetcoRiser150 Then
                            SegMidWD = -SegTopAbvWL - SegLength / 2.0
                            SegMidCurU = CurProj.Water.CurrentProfile.VofD(SegMidWD, 1)
                            SegMidCurV = CurProj.Water.CurrentProfile.VofD(SegMidWD, 2)
                            SegMidCurAngle = Math.Atan2(SegMidCurV, SegMidCurU) * 180 / Math.PI
                            SegMidCurAngle = SegMidCurAngle - CurProj.Riser(ir).VectoAngle
                            If SegMidCurAngle < 0.0 Then SegMidCurAngle = SegMidCurAngle + 360.0
                            If SegMidCurAngle > 180 Then SegMidCurAngle -= 180

                            If SegMidCurAngle >= 0 And SegMidCurAngle < 15 Then
                                SegType = ICHARValues.VetcoRiser0
                            ElseIf SegMidCurAngle >= 15 And SegMidCurAngle < 45 Then
                                SegType = ICHARValues.VetcoRiser30
                            ElseIf SegMidCurAngle >= 45 And SegMidCurAngle < 75 Then
                                SegType = ICHARValues.VetcoRiser60
                            ElseIf SegMidCurAngle >= 75 And SegMidCurAngle < 105 Then
                                SegType = ICHARValues.VetcoRiser90
                            ElseIf SegMidCurAngle >= 105 And SegMidCurAngle < 135 Then
                                SegType = ICHARValues.VetcoRiser120
                            ElseIf SegMidCurAngle >= 135 And SegMidCurAngle < 175 Then
                                SegType = ICHARValues.VetcoRiser90
                            Else
                                SegType = ICHARValues.VetcoRiser0
                            End If
                        End If

                        DragDiam = .BuoyModDia
                        If DragDiam < .MainTubeOD Then DragDiam = .MainTubeOD
                        Mass = .MassPerLength + AuxMass
                        AddedMass = .AddedMassPerLength + AuxAddedMass
                        If SegTopAbvWL >= SegLength Then
                            SegAuxWetWt = AuxMass * Gm
                        ElseIf SegTopAbvWL <= 0.0# Then
                            SegAuxWetWt = AuxWetWeight
                        Else
                            SegAuxWetWt = (AuxMass * Gm * SegTopAbvWL + AuxWetWeight * (SegLength - SegTopAbvWL)) / SegLength
                        End If
                        WetWeight = .WetWeightPerLength(SegTopAbvWL) + SegAuxWetWt

                        If CurProj.Riser(ir).CalculateBendingStiffness Then
                            EI = .EI + AuxEI
                        Else
                            EI = .BendingStiffness + AuxEI
                        End If

                        If .StrakesHeight > 0 Then
                            NTYP = 7
                            HTD = .StrakesHeight / DragDiam
                        ElseIf .FairChord > DragDiam And .FairThick >= DragDiam Then
                            NTYP = 8
                        Else
                            NTYP = 1
                        End If

                        PrintLine(FileNum, SegLength)
                        PrintLine(FileNum, Mass)
                        PrintLine(FileNum, AddedMass)
                        PrintLine(FileNum, WetWeight)
                        PrintLine(FileNum, DragDiam)
                        PrintLine(FileNum, .MainTubeOD, .WallThickness)
                        PrintLine(FileNum, EI)
                        PrintLine(FileNum, NTYP)
                        If NTYP = 7 Then PrintLine(FileNum, HTD)
                        If NTYP = 8 Then
                            PrintLine(FileNum, .FairThick, .FairChord)
                        End If

                        If .SectionType >= ICHARValues.VetcoRiser0 And .SectionType <= ICHARValues.VetcoRiser150 Then
                            PrintLine(FileNum, SegType)
                        Else
                            PrintLine(FileNum, .SectionType)
                        End If
                    End With
                Next seg

                'riser type
                If .RiserType = VIVAMain.RiserTypes.Rigid Then
                    ICHOICE = 1
                Else
                    ICHOICE = 2
                End If
                PrintLine(FileNum, ICHOICE)

                'end rispre.in
                FileClose(FileNum)

                'calculate totalweight for top tension adjustment
                AxialTension = .TopTension
                If .RiserType <> VIVAMain.RiserTypes.SCRorLazyWave Then
                    LAW = .TopLocation
                    TmpLAW = LAW
                    For i = .Segments.Count To 1 Step -1
                        With .Segments.Item(i)
                            Mass = .MassPerLength + AuxMass
                            WetWeight = .WetWeightPerLength + AuxWetWeight
                            SegLength = .JointLength * .NumJoints
                            If TmpLAW > 0.0# Then
                                If TmpLAW > SegLength Then
                                    AxialTension = AxialTension - Mass * Gm * SegLength
                                    TmpLAW = TmpLAW - SegLength
                                Else
                                    AxialTension = AxialTension - Mass * Gm * TmpLAW - WetWeight * (SegLength - LAW)
                                    TmpLAW = 0.0#
                                End If
                            Else
                                AxialTension = AxialTension - WetWeight * SegLength
                            End If
                            If AxialTension < 0.0# Then Exit For
                        End With
                    Next
                    If AxialTension >= 0.0# Then AxialTension = AxialTension - StackWeight
                End If

                If CurProj.BatchProcess = False Then 'only check if not batch process

                    If AxialTension < 0.0# Then
                        Select Case MsgBox("Warning: There is compression of about " & VB6.Format(AxialTension, "#,##0") _
                            & " N. along the riser string. Do you still want to continue VIV analysis before checking the Top Tension?", _
                            MsgBoxStyle.YesNo, "WinVIVA - Top Tension")
                            Case MsgBoxResult.Yes
                                'go ahead
                            Case MsgBoxResult.No
                                Exit Function
                        End Select
                    ElseIf .LowerBC = VIVAMain.BoundaryConditions.Free Then
                        If AxialTension > 0.0# And .RiserType <> VIVAMain.RiserTypes.SCRorLazyWave Then
                            Select Case MsgBox("Warning: There is a Bottom Tension of about " & VB6.Format(AxialTension, "#,##0") _
                                & " N. " & "Do you still want to continue VIV analysis " & " before checking the Top Tension?", _
                                MsgBoxStyle.YesNo, "WinVIVA - Top Tension")
                                Case MsgBoxResult.Yes
                                    'go ahead
                                Case MsgBoxResult.No
                                    Exit Function
                            End Select
                        End If
                    ElseIf .UpperBC = VIVAMain.BoundaryConditions.Free Then
                        If .TopTension <> 0.0# Then
                            Select Case MsgBox("Do you want to set Riser Top Tension to zero " & "for a free-standing riser and continue?", _
                                MsgBoxStyle.YesNo, "WinVIVA - Top Tension")
                                Case MsgBoxResult.Yes
                                    .TopTension = 0.0#
                                Case MsgBoxResult.No
                                    Exit Function
                            End Select
                        End If
                    End If
                End If
                'vivo-n.in
                DOSFile = VIVADIR & OutputFiles(ix, vivo_n_in) 'TODO JLIU
                If .RiserType = VIVAMain.RiserTypes.SCRorLazyWave Then
                    AxialTension = .StaticSolution.Item(1).AxialStaticTension
                    If .TopTension <> AxialTension Then
                        Select Case MsgBox("Do you want to set Riser Top Tension to the Axial " _
                            & "Tension value of the first SCR point and continue?", MsgBoxStyle.YesNo, "WinVIVA - Top Tension")
                            Case MsgBoxResult.Yes
                                .TopTension = AxialTension
                            Case MsgBoxResult.No
                                Exit Function
                        End Select
                    End If
                    SSDepth = .StaticSolution.Item(1).Depth
                    If .TopLocation <> -SSDepth Then
                        Select Case MsgBox("Do you want to set Riser Top Location to the Vertical " _
                            & "Distance value of the first SCR point and continue?", MsgBoxStyle.YesNo, "WinVIVA - Top Tension")
                            Case MsgBoxResult.Yes
                                .TopLocation = -SSDepth
                            Case MsgBoxResult.No
                                Exit Function
                        End Select
                    End If

                    FileNum = FreeFile()
                    FileOpen(FileNum, DOSFile, OpenMode.Output, OpenAccess.Write)

                    SSDepth = .StaticSolution.Item(.StaticSolution.Count).Depth
                    PrintLine(FileNum, .StaticSolution.Count, SSDepth)

                    For Each SP In .StaticSolution
                        With SP
                            Print(FileNum, .Distance, TAB)
                            Print(FileNum, .AxialStaticTension, TAB)
                            Print(FileNum, .ChangeInStaticTension, TAB)
                            Print(FileNum, SSDepth - .Depth, TAB)
                            Print(FileNum, .Ydist, TAB)
                            Print(FileNum, .Zdist, TAB)
                            Print(FileNum, CSng(.UCurrentVelocity * System.Math.Sin(.InPlaneAngle)), TAB) 'TODO JLIU
                            Print(FileNum, .VCurrentVelocity, TAB)
                            Print(FileNum, 0, TAB)
                            PrintLine(FileNum, 0) 'dummies for URiser and VRiser

                            If .AxialStaticTension < 0.0# Then
                                Select Case MsgBox("Warning: There is compression along the " _
                                    & "riser string. Do you still want to continue " & "VIV analysis before checking the SCR input?", _
                                    MsgBoxStyle.YesNo, "WinVIVA - Top Tension")
                                    Case MsgBoxResult.Yes
                                        'go ahead
                                    Case MsgBoxResult.No
                                        FileClose(FileNum)
                                        Exit Function
                                End Select
                            End If
                        End With
                    Next SP

                    'end vivo-n.in
                    FileClose(FileNum)
                Else
                    If Len(Dir(DOSFile)) > 0 Then Kill(DOSFile)
                End If

                'preparing risfat.in
                FileNum = FreeFile()
                FileOpen(FileNum, VIVADIR & OutputFiles(ix, risfat_in), OpenMode.Output, OpenAccess.Write)

                PrintLine(FileNum, 1)
                PrintLine(FileNum, .FatigueConstA, .FatigueConstB)
                PrintLine(FileNum, .SCF)

                'end risfat.in
                FileClose(FileNum)

            End With

            'preparing risdyn-n.in
            FileNum = FreeFile()
            FileOpen(FileNum, VIVADIR & OutputFiles(ix, risdyn_n_in), OpenMode.Output, OpenAccess.Write)
            'output file name
            Select Case ix
                Case 0
                    PrintLine(FileNum, "'RISOUT'")
                    PrintLine(FileNum, "'DYNOUT'")
                Case 1
                    PrintLine(FileNum, "'RISOUTF'")
                    PrintLine(FileNum, "'DYNOUTF'")
                Case 2
                    PrintLine(FileNum, "'RISOUTR'")
                    PrintLine(FileNum, "'DYNOUTR'")
            End Select


            'set the name of the output file this will produce
            PrintLine(FileNum, CurProj.Riser(ir).TopTension)
            PrintLine(FileNum, CurProj.NumPoints)
            PrintLine(FileNum, CurProj.NumIterations)

            'current profile
            With CurProj.Water
                NumCurPairs = .CurrentProfile.Count
                IniCurDepth = .CurrentProfile.Item(1).Depth
                If -LAW > IniCurDepth Then
                    i = 1
                    For Each CP In .CurrentProfile
                        If -LAW >= CP.Depth Then
                            i = i + 1
                        Else
                            If i = CP.PairNum Then
                                PrintLine(FileNum, NumCurPairs - i + 2)
                                PrintLine(FileNum, -LAW, CurProj.Water.CurrentProfile.VofD(-LAW, 1), CurProj.Water.CurrentProfile.VofD(-LAW, 2))
                                PrintLine(FileNum, CP.Depth, CP.VelU, CP.VelV)
                            Else
                                PrintLine(FileNum, CP.Depth, CP.VelU, CP.VelV)
                            End If
                        End If
                    Next CP
                Else
                    If -LAW = IniCurDepth Then
                        PrintLine(FileNum, NumCurPairs)
                    Else
                        PrintLine(FileNum, NumCurPairs + 2)
                        PrintLine(FileNum, -LAW, 0.0#, 0.0#)
                        PrintLine(FileNum, IniCurDepth - 0.01, 0.0#, 0.0#)
                    End If
                    For Each CP In .CurrentProfile
                        If CP.Depth >= 0.0# Then
                            PrintLine(FileNum, CP.Depth, CP.VelU, CP.VelV)
                        Else
                            PrintLine(FileNum, CP.Depth, 0.0#, 0.0#)
                        End If
                    Next CP
                End If

                PrintLine(FileNum, .Density)

            End With
            PrintLine(FileNum, CurProj.Riser(ir).Damping)
            PrintLine(FileNum, CurProj.Riser(ir).LatSupports.Count) 'number of intermediate support springs

            For i = 1 To CurProj.Riser(ir).LatSupports.Count
                With CurProj.Riser(ir).LatSupports.Item(i)
                    PrintLine(FileNum, .Dist, .Stiffness, .Damping)
                End With
            Next
            If ir = 1 Then
                PrintLine(FileNum, CurProj.RiserFLocY, CurProj.RiserFLocZ)
            Else
                PrintLine(FileNum, CurProj.RiserRLocY, CurProj.RiserRLocZ)
            End If
            'end of risdyn-n.in
            FileClose(FileNum)

        Next

        WriteDOSVIVAInput = True
        Exit Function

ErrHandler:
        WriteDOSVIVAInput = False
        FileClose(FileNum)
        MsgBox("Error occurred when writing  DOS VIVA input files!", MsgBoxResult.Yes, "WinVIVA - Write DOS VIVA Input Error")
        Exit Function

    End Function
  

    Public Function VIVAFilesExist() As Boolean
        Dim i As Short
        Dim Fn As String
        ' See if the files that VIVACORE will need exist in the specified
        ' directory.  VIVACORE will check for these, but the error message
        ' issued if they do not exist is lost in the background shell process
        ' as it runs.
        VIVAFilesExist = True
        With VIVACoreFiles
            For i = 1 To .NumFiles
                Fn = .FileNames(i - 1)
                If Dir(VIVADIR & Fn, FileAttribute.Normal) = "" Then
                    MsgBox("The file " & Fn & " does not exist in the working directory " & VIVADIR & ".  Execution cancelled.", MsgBoxStyle.OkOnly, "WinVIVA - File Missing")
                    VIVAFilesExist = False
                    Exit For
                End If
            Next i
        End With
    End Function
 
    Public Function VIVACoreOutputsExist() As Boolean
        ' See if VIVACORE produced the expected files; if not, assume an error

        Dim i, ir, ix As Short

        VIVACoreOutputsExist = True

        For ir = 1 To CurProj.nRisers
            ix = CurProj.nRisers * ir - ir
            With CurProj.Riser(ir)
                For i = NumInputFiles To NumOutputFiles - 1

                    If Dir(VIVADIR & OutputFiles(ix, i), FileAttribute.Normal) = "" Then
                        MsgBox("The file " & OutputFiles(ix, i) & " does not exist in the working directory " & VIVADIR & ".  File Retrieve.", MsgBoxStyle.OkOnly, "WinVIVA - File Missing")
                        VIVACoreOutputsExist = False
                    End If
                Next i

            End With
        Next

    End Function

  
    Public Sub CopyVIVAInputs()

        ' Copy all Input files to the current proj's data directory
        Dim i, ir, ix As Short
        Dim Fn As String
        On Error Resume Next
    
        For ir = 1 To CurProj.nRisers
            ix = CurProj.nRisers * ir - ir

            For i = 0 To NumInputFiles - 1
                Fn = OutputFiles(ix, i)
                If Fn <> "" Then
                    fso.DeleteFile(CurProj.DataDirectory & Fn)
                End If
            Next i

            For i = 0 To NumInputFiles - 1
                Fn = OutputFiles(ix, i)
                If Fn <> "" Then
                    FileCopy(VIVADIR & Fn, CurProj.DataDirectory & Fn)
                End If
            Next i
        Next
    End Sub


    Public Sub MoveVIVAOutputs()
        ' Move all output files to the current proj's data directory

        Dim i, ir, ix As Short
        Dim Fn As String
        On Error Resume Next
      

        For ir = 1 To CurProj.nRisers
            ix = CurProj.nRisers * ir - ir
            'delete old output before move in new one
            For i = 0 To NumOutputFiles - 1
                Fn = OutputFiles(ix, i)
                If Fn <> "" Then
                    fso.DeleteFile(CurProj.DataDirectory & Fn)
                End If
            Next i

            For i = 0 To NumOutputFiles - 1
                Fn = OutputFiles(ix, i)
                If Fn <> "" Then
                    fso.MoveFile(VIVADIR & Fn, CurProj.DataDirectory & Fn)
                End If
            Next i

        Next

    End Sub


    Public Sub RemovePreviousVIVAOutputs()
        ' Move all output files to the current proj's data directory

        Dim i, ir, ix As Short
        Dim Fn As String
    
        For ir = 1 To CurProj.nRisers
            ix = CurProj.nRisers * ir - ir
            On Error Resume Next

            'delete old output in the executing directory
            For i = 0 To NumOutputFiles - 1
                Fn = OutputFiles(ix, i)
                fso.DeleteFile(VIVADIR & Fn)
            Next i

        Next
    End Sub


    Public Sub CopyImportedFiles()
        ' copy imported external files to VIVADIR directory
        On Error GoTo ErrHandler

        Dim ir, ix As Short
        Dim Fn As String

        For ir = 1 To CurProj.nRisers
            ix = CurProj.nRisers * ir - ir

            With CurProj.Riser(ir)
                If .Mode = 1 Then
                    FileCopy(.FreqFile, VIVADIR & OutputFiles(ix, freq_in))
                ElseIf .Mode = 2 Then
                    FileCopy(.FreqFile, VIVADIR & OutputFiles(ix, freq_in))
                    FileCopy(.ModeFile, VIVADIR & OutputFiles(ix, modes_us_in))
                    FileCopy(.CurvFile, VIVADIR & OutputFiles(ix, curv_us_in))
                End If
            End With
        Next ir
        Exit Sub

ErrHandler:
        MsgBox("Error when copying file " & Fn, MsgBoxStyle.OkOnly, "WinVIVA - Copy File Error")
        Exit Sub

    End Sub


    Public Function ReadVIVACoreOutputFiles() As Boolean

        On Error GoTo ErrHandler

        Dim TypelessFN As String
        Dim FileNum, ir, ix As Short

        ReadVIVACoreOutputFiles = False
        For ir = 1 To CurProj.nRisers
            ix = CurProj.nRisers * ir - ir
            'With CurProj.Riser(ir)

            'fat1.out
            FileNum = FreeFile()
            FileOpen(FileNum, CurProj.DataDirectory & OutputFiles(ix, fat1_h3_out), OpenMode.Input, OpenAccess.Read)
            If EOF(FileNum) Then
                MsgBox("File " & CurProj.DataDirectory & OutputFiles(ix, fat1_out) & " is empty. The rel evant data have not been updated!", MsgBoxStyle.OkOnly, "WinVIVA - No Data Available")
            Else
                ReadOutput8(FileNum, ir, fat1_h3_out, ix)
            End If
            FileClose(FileNum)

            '   "summ1.out" file : general data
            FileNum = FreeFile()
            FileOpen(FileNum, CurProj.DataDirectory & OutputFiles(ix, summ1_out), OpenMode.Input, OpenAccess.Read)

            If EOF(FileNum) Then
                MsgBox("File " & CurProj.DataDirectory & OutputFiles(ix, summ1_out) & " is empty. The relevant data have not been updated!", MsgBoxStyle.OkOnly, "WinVIVA - No Data Available")

            Else
                ReadOutput1(FileNum, ir, summ1_out, ix)
            End If
            FileClose(FileNum)

    
            FileNum = FreeFile()
            '"summ1_mm.out" file : general data
            FileOpen(FileNum, CurProj.DataDirectory & OutputFiles(ix, summ1_mm_out), OpenMode.Input, OpenAccess.Read)

            If EOF(FileNum) Then
                MsgBox("File " & OutputFiles(ix, summ1_mm_out) & " is empty. The relevant data have not been updated!", MsgBoxStyle.OkOnly, "WinVIVA - No Data Available")

            Else
                ReadOutput11(FileNum, ir, summ1_mm_out, ix)
            End If
            FileClose(FileNum)
      
        Next ir
        ReadVIVACoreOutputFiles = True
        Exit Function

ErrHandler:
        ReadVIVACoreOutputFiles = False

        Exit Function

    End Function
   
    Public Function ImportVIVACoreOutputFilesToExcel(ByVal results As Short, ByVal riserId As Short) As Boolean

        On Error GoTo ErrHandler

        Dim FileNum, j, k, ix As Short
        Dim oxApp As Excel.Application, oxBook As Excel.Workbook, oxSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim dlgSave As System.Windows.Forms.SaveFileDialog = _
        New System.Windows.Forms.SaveFileDialog

        Dim resultsNm(0 To 3) As String

        'should the user cancel the dialog box, exit
        On Error GoTo ErrHandler

        dlgSave.OverwritePrompt = True
        dlgSave.Filter = "All Files (*.*)|*.*|Excel Files (*.xls)|*.xls"
        dlgSave.FileName = "*.xls"

        dlgSave.ShowDialog()
        IniConvertFactor(CurProj.Units)
        ImportVIVACoreOutputFilesToExcel = False
        resultsNm(1) = "MM"
        resultsNm(2) = "AllByMode"
        resultsNm(3) = "AllByCategory"

        category(DISTANCE) = "Distance" & CFLengthS
        category(AMP) = "Amp" & CFLengthS
        category(BM) = "BM" & CFMomS
        category(STRESS) = "Stress" & CFStressS
        category(CD) = "CD"
        category(FAT_LIFE) = "Fatigue(yr)"

        oxApp = New Excel.Application
        oxBook = oxApp.Workbooks.Add(misValue)

        ix = CurProj.nRisers * riserId - riserId

        With CurProj.Riser(riserId)
            If results > 1 Then
                If results = 2 Then
                    For j = 1 To .NumOutputModes
                        k = .ModeNumber(j)
                        oxSheet = oxBook.Worksheets.Add

                        oxSheet.Name = k
                        oxSheet.Cells(1, 1) = "Initial Freq"
                        oxSheet.Cells(1, 2) = .InitialFreqs(k)
                        oxSheet.Cells(2, 1) = "Coupled Freq"
                        oxSheet.Cells(2, 2) = .CoupledFreqs(k)
                    Next j
                Else
                    For j = 1 To 5
                        oxSheet = oxBook.Worksheets.Add
                        oxSheet.Name = category(j)
                    Next j
                End If

                '"OUT.OUT" file: mode shape
                FileNum = FreeFile()
                FileOpen(FileNum, CurProj.DataDirectory & OutputFiles(ix, out_out), OpenMode.Input, OpenAccess.Read)

                If EOF(FileNum) Then
                    MsgBox("File " & OutputFiles(ix, out_out) & " is empty. The relevant data have not been updated!", MsgBoxStyle.OkOnly, "WinVIVA - No Data Available")
                Else
                    ImportOutput2(FileNum, oxBook, AMP + 1, results, riserId, out_out, ix)
                End If
                FileClose(FileNum)

                '"BEND.OUT" file: bending moment and stress
                FileNum = FreeFile()
                FileOpen(FileNum, CurProj.DataDirectory & OutputFiles(ix, bend_out), OpenMode.Input, OpenAccess.Read)

                If EOF(FileNum) Then
                    MsgBox("File " & OutputFiles(ix, bend_out) & " is empty. The relevant data have not been updated!", MsgBoxStyle.OkOnly, "WinVIVA - No Data Available")
                Else
                    ImportOutput4(FileNum, oxBook, BM + 1, results, riserId, bend_out, ix)
                End If
                FileClose(FileNum)

                '"CDRAG.OUT" file: drag coeffecients
                FileNum = FreeFile()
                FileOpen(FileNum, CurProj.DataDirectory & OutputFiles(ix, cdrag_out), OpenMode.Input, OpenAccess.Read)

                If EOF(FileNum) Then
                    MsgBox("File " & CurProj.DataDirectory & OutputFiles(ix, cdrag_out) & " is empty. The relevant data have not been updated!", MsgBoxStyle.OkOnly, "WinVIVA - No Data Available")
                Else
                    ImportOutput6(FileNum, oxBook, CD + 1, results, riserId, cdrag_out, ix)
                End If
                FileClose(FileNum)

                '"FAT-MONO_h3.OUT" file: drag coeffecients
                FileNum = FreeFile()
                FileOpen(FileNum, CurProj.DataDirectory & OutputFiles(ix, fat_mono_h3_out), OpenMode.Input, OpenAccess.Read)

                If EOF(FileNum) Then
                    MsgBox("File " & OutputFiles(ix, fat_mono_h3_out) & " is empty. The relevant data have not been updated!", MsgBoxStyle.OkOnly, "WinVIVA - No Data Available")
                Else
                    ImportOutput12(FileNum, oxBook, FAT_LIFE + 1, results, riserId, fat_mono_h3_out, ix)
                End If
                FileClose(FileNum)

            End If

            oxSheet = oxBook.Worksheets.Add
            oxSheet.Name = "MM"

            '"OUT_MM.OUT" file: mode shape (multi-frequency)
            FileNum = FreeFile()
            FileOpen(FileNum, CurProj.DataDirectory & OutputFiles(ix, out_mm_out), OpenMode.Input, OpenAccess.Read)
            If EOF(FileNum) Then
                MsgBox("File " & OutputFiles(ix, out_mm_out) & " is empty. The relevant data have not been updated!", MsgBoxStyle.OkOnly, "WinVIVA - No Data Available")
            Else
                ImportOutput3(FileNum, oxBook, AMP + 1, riserId, out_mm_out, ix)
            End If
            FileClose(FileNum)

            '"BEND_MM.OUT" file: bending moment and stress (multi-frequency)
            FileNum = FreeFile()
            FileOpen(FileNum, CurProj.DataDirectory & OutputFiles(ix, bend_mm_out), OpenMode.Input, OpenAccess.Read)
            If EOF(FileNum) Then
                MsgBox("File " & CurProj.DataDirectory & OutputFiles(ix, bend_mm_out) & " is empty. The relevant data have not been updated!", MsgBoxStyle.OkOnly, "WinVIVA - No Data Available")
            Else
                ImportOutput5(FileNum, oxBook, BM + 1, bend_mm_out, ix)
            End If
            FileClose(FileNum)

            '"CD_MM.OUT" file: drag coeffecients (multi-frequency)
            FileNum = FreeFile()
            FileOpen(FileNum, CurProj.DataDirectory & OutputFiles(ix, cdrag_mm_out), OpenMode.Input, OpenAccess.Read)
            If EOF(FileNum) Then
                MsgBox("File " & CurProj.DataDirectory & OutputFiles(ix, cdrag_mm_out) & " is empty. The relevant data have not been updated!", MsgBoxStyle.OkOnly, "WinVIVA - No Data Available")
            Else
                ImportOutput7(FileNum, oxBook, CD + 1, riserId, cdrag_mm_out, ix)
            End If
            FileClose(FileNum)

            '"FAT_MULT_h3.OUT" file: drag coeffecients (multi-frequency)
            FileNum = FreeFile()
            FileOpen(FileNum, CurProj.DataDirectory & OutputFiles(ix, fat_multi_h3_out), OpenMode.Input, OpenAccess.Read)
            If EOF(FileNum) Then
                MsgBox("File " & OutputFiles(ix, fat_multi_h3_out) & " is empty. The relevant data have not been updated!", MsgBoxStyle.OkOnly, "WinVIVA - No Data Available")
            Else
                ImportOutput13(FileNum, oxBook, FAT_LIFE + 1, riserId, fat_multi_h3_out, ix)
            End If
            FileClose(FileNum)

            ' oxBook.SaveAs(CurProj.DataDirectory & CurProj.Title & resultsNm(results) & ".xls", Excel.XlFileFormat.xlWorkbookNormal)
            If dlgSave.FileName IsNot Nothing Then
                oxBook.SaveAs(dlgSave.FileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, _
                    Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
                oxBook.Close(True, misValue, misValue)

                oxBook = Nothing
                oxApp.Quit()
            End If
        End With
        ImportVIVACoreOutputFilesToExcel = True
        Exit Function

ErrHandler:
        ImportVIVACoreOutputFilesToExcel = False
        Exit Function
    End Function

    Private Sub ReadOutput8(ByVal FileNum As Short, ByVal riserId As Short, ByVal whichFile As Short, ByVal whichIndex As Short)

        'read "FAT1_h3.OUT" file: fatigue data
        On Error GoTo ErrHandler

        Dim j, k, ir As Short
        Dim dum As Object
        Dim tmpInt As Short
        Dim tmpSingle, tmpSingle2 As Single
        Dim tmpSingle3 As Object

        'initialization
        tmpSingle3 = Nothing
        With CurProj.Riser(riserId)
            ' first line: fatigue constants and concentration factor
            If whichFile = fat1_out Then
                Input(FileNum, tmpSingle)
                Input(FileNum, tmpSingle2)
                Input(FileNum, tmpSingle3)
                .FatigueConstA = tmpSingle
                .FatigueConstB = tmpSingle2
                .SCF = tmpSingle3
            End If
            Input(FileNum, tmpInt)
            .NumOutputModes = tmpInt
            For j = 1 To .NumOutputModes

                If EOF(FileNum) Then Exit Sub
                Input(FileNum, .ModeNumber(j))
                k = .ModeNumber(j)
                Input(FileNum, tmpSingle)
                Input(FileNum, .FatigueLife(k))
                Input(FileNum, .MaxStressPos(k))
                .FatiguePeriod(k) = 1.0# / tmpSingle
                If .MaxStressPos(k) > .TotalLength Then
                    .MaxStressPos(k) = .TotalLength
                ElseIf .MaxStressPos(k) < 0.0# Then
                    .MaxStressPos(k) = 0.0#
                End If
            Next j

            If EOF(FileNum) Then Exit Sub
            'read in data for multi-mode
            Input(FileNum, .FatigueLife(0))
            Input(FileNum, .MaxStressPos(0))
        End With

        Exit Sub

ErrHandler:
        MsgBox("Error occurred when reading " & OutputFiles(whichIndex, whichFile) & " file. The data may be imcomplete or corrupted", MsgBoxStyle.OkOnly, " WinVIVA - Read Output File Error")
        Exit Sub

    End Sub


    Private Sub ReadOutput1(ByVal FileNum As Short, ByVal riserId As Short, ByVal whichFile As Short, ByVal whichIndex As Short)
        'read "summ1.out" file
        On Error GoTo ErrHandler

        Dim i, k As Short
        Dim dum As Object

        'initialization
        dum = Nothing
        With CurProj.Riser(riserId)
            '       frequency
            For i = 1 To .NumOutputModes
                If EOF(FileNum) Then Exit Sub
                k = .ModeNumber(i)
                Input(FileNum, dum)
                Input(FileNum, .InitialFreqs(k))
                Input(FileNum, .CoupledFreqs(k))
                'Input(FileNum, dum)
                Input(FileNum, .MaxA(k))
            Next i
        End With
        Exit Sub
ErrHandler:
        MsgBox("Error occurred when reading " & OutputFiles(whichIndex, whichFile) & " file. The data may be imcomplete or corrupted", MsgBoxStyle.OkOnly, " WinVIVA - Read Output File Error")
        Exit Sub

    End Sub
    Private Sub ReadOutput11(ByVal FileNum As Short, ByVal riserId As Short, ByVal whichFile As Short, ByVal whichindex As Short)
        'read "summ1_mm.out" file
        Dim dum As Object
        Dim temp(100) As Single
        dum = Nothing
        Dim i As Short
        On Error GoTo ErrHandler
        Input(FileNum, dum)
        i = 1
        Do Until EOF(FileNum)
            Input(FileNum, temp(i))
            i = i + 1
            'MsgBox(MaxA(0))
        Loop
        'MsgBox(temp(i - 1))
        With CurProj.Riser(riserId)
            .MaxA(0) = temp(i - 2)
        End With
        Exit Sub
ErrHandler:
        MsgBox("Error occurred when reading " & OutputFiles(whichindex, whichFile) & " file. The data may be imcomplete or corrupted", MsgBoxStyle.OkOnly, " WinVIVA - Read Output File Error")
        Exit Sub

    End Sub


    Private Sub ReadOutput2(ByVal FileNum As Short, ByVal riserId As Short, ByVal whichFile As Short, ByVal whichIndex As Short)

        'read "OUT.DAT" file: mode shape

        On Error GoTo ErrHandler

        Dim j, i, k As Short
        Dim seg1, NumSegments, seg2 As Short
        Dim AoD1, dia1, dia2, AoD2, distancei As Single

        With CurProj.Riser(riserId)
            NumSegments = .TotalSegments
            For j = 1 To .NumOutputModes
                k = .ModeNumber(j)
                .MaxAoverD(k) = 0.0#
                For i = 1 To NumSegments
                    .SegMaxAoverD(i, k) = 0.0#
                Next

                For i = 1 To CurProj.NumPoints
                    If EOF(FileNum) Then Exit Sub
                    Input(FileNum, distancei)
                    Input(FileNum, .Amplitude(k))

                    seg1 = .SegNum(distancei, seg2)
                    dia1 = .Diameter(seg1)
                    If dia1 = 0.0# Then
                        AoD1 = 0.0#
                    Else
                        AoD1 = .Amplitude(k) / dia1
                        If AoD1 > .SegMaxAoverD(seg1, k) Then .SegMaxAoverD(seg1, k) = AoD1

                        If seg2 = 0 Then
                            AoD2 = 0.0#
                        Else
                            dia2 = .Diameter(seg2)
                            If dia2 = 0.0# Then
                                AoD2 = 0.0#
                            Else
                                AoD2 = .WaveMag(k) / dia2
                                If AoD2 > .SegMaxAoverD(seg2, k) Then .SegMaxAoverD(seg1, k) = AoD2
                            End If
                        End If
                        If AoD1 >= AoD2 Then
                            .AoverD(k) = AoD1
                        Else
                            .AoverD(k) = AoD2
                        End If
                        If .AoverD(k) > .MaxAoverD(k) Then .MaxAoverD(k) = .AoverD(k)
                    End If
                Next i
            Next j
        End With

        Exit Sub

ErrHandler:
        MsgBox("Error occurred when reading " & OutputFiles(whichIndex, whichFile) & " file. The data may be imcomplete or corrupted", MsgBoxStyle.OkOnly, " WinVIVA - Read Output File Error")
        Exit Sub

    End Sub
    Private Sub ImportOutput2(ByVal FileNum As Short, ByRef oxBook As Excel.Workbook, ByVal col As Short, ByVal results As Short, ByVal riserId As Short, ByVal whichFile As Short, ByVal whichindex As Short)

        'read "OUT.DAT" file: mode shape

        On Error GoTo ErrHandler
        Dim oxSheet As Excel.Worksheet
        Dim j, i, k As Short
        Dim distancei, amplitudei As Single
        Dim SheetName As String
        Dim seg1, NumSegments, seg2 As Short
        Dim AoD1, dia1, dia2, AoD2 As Single
        With CurProj.Riser(riserId)
            NumSegments = .TotalSegments
            '   bending moment and stress
            If results = 2 Then
                For j = 1 To .NumOutputModes
                    k = .ModeNumber(j)
                    .MaxAoverD(k) = 0.0#
                    For i = 1 To NumSegments
                        .SegMaxAoverD(i, k) = 0.0#
                    Next i
                    SheetName = k
                    oxSheet = oxBook.Worksheets(SheetName)
                    oxSheet.Cells(4, col) = category(AMP)
                    For i = 1 To CurProj.NumPoints
                        If EOF(FileNum) Then Exit Sub
                        Input(FileNum, distancei)
                        Input(FileNum, amplitudei)
                        oxSheet.Cells(i + 4, col) = amplitudei / CFLength
                        seg1 = .SegNum(distancei, seg2)
                        dia1 = .Diameter(seg1)
                        If dia1 = 0.0# Then
                            AoD1 = 0.0#
                        Else
                            AoD1 = amplitudei / dia1
                            If AoD1 > .SegMaxAoverD(seg1, k) Then .SegMaxAoverD(seg1, k) = AoD1

                            If seg2 = 0 Then
                                AoD2 = 0.0#
                            Else
                                dia2 = .Diameter(seg2)
                                If dia2 = 0.0# Then
                                    AoD2 = 0.0#
                                Else
                                    AoD2 = .WaveMag(k) / dia2
                                    If AoD2 > .SegMaxAoverD(seg2, k) Then .SegMaxAoverD(seg1, k) = AoD2
                                End If
                            End If
                            If AoD1 >= AoD2 Then
                                .AoverD(k) = AoD1
                            Else
                                .AoverD(k) = AoD2
                            End If
                            If .AoverD(k) > .MaxAoverD(k) Then .MaxAoverD(k) = .AoverD(k)
                        End If
                    Next i
                    oxSheet.Range(oxSheet.Cells(2, col), oxSheet.Cells(CurProj.NumPoints + 4, col)).NumberFormat = "0.00"

                Next j
            Else
                oxSheet = oxBook.Worksheets(category(AMP))

                oxSheet.Cells(1, 1) = category(DISTANCE)
                For j = 1 To .NumOutputModes
                    k = .ModeNumber(j)
                    .MaxAoverD(k) = 0.0#
                    For i = 1 To NumSegments
                        .SegMaxAoverD(i, k) = 0.0#
                    Next i
                    oxSheet.Cells(1, j + 1) = k
                    For i = 1 To CurProj.NumPoints
                        If EOF(FileNum) Then Exit Sub
                        Input(FileNum, distancei)
                        Input(FileNum, amplitudei)
                        If j = 1 Then
                            oxSheet.Cells(i + 1, 1) = distancei / CFLength
                        End If
                        oxSheet.Cells(i + 1, j + 1) = amplitudei / CFLength
                    Next i
                Next j
                oxSheet.Range(oxSheet.Cells(2, 1), oxSheet.Cells(CurProj.NumPoints + 1, .NumOutputModes + 1)).NumberFormat = "0.00"
            End If
        End With
        Exit Sub
ErrHandler:
        MsgBox("Error occurred when reading " & OutputFiles(whichindex, whichFile) & " file. The data may be imcomplete or corrupted", MsgBoxStyle.OkOnly, " WinVIVA - Read Output File Error")
        Exit Sub

    End Sub

    Private Sub ImportOutput3(ByVal FileNum As Short, ByRef oxBook As Excel.Workbook, ByVal col As Short, ByVal riserId As Short, ByVal whichFile As Short, ByVal whichIndex As Short)

        'read "OUT_MM.DAT" file: amplitude
        On Error GoTo ErrHandler
        Dim oxSheet As Excel.Worksheet
        Dim j, i, k As Short
        Dim distancei, amplitudei As Single

        Dim seg1, NumSegments, seg2 As Short
        Dim AoD1, dia1, dia2, AoD2 As Single
        With CurProj.Riser(riserId)
            NumSegments = .TotalSegments
            '   
            .MaxAoverD(0) = 0.0#
            For i = 1 To NumSegments
                .SegMaxAoverD(i, 0) = 0.0#
            Next i
            oxSheet = oxBook.Worksheets("MM")
            oxSheet.Cells(4, col) = category(AMP)
            For i = 1 To CurProj.NumPoints
                If EOF(FileNum) Then Exit Sub
                Input(FileNum, distancei)
                Input(FileNum, amplitudei)
                oxSheet.Cells(i + 4, col) = amplitudei / CFLength

            Next i
            oxSheet.Range(oxSheet.Cells(4, col), oxSheet.Cells(CurProj.NumPoints + 4, col)).NumberFormat = "0.000"
        End With
        Exit Sub

ErrHandler:
        MsgBox("Error occurred when reading " & OutputFiles(whichIndex, whichFile) & " file. The data may be imcomplete or corrupted", MsgBoxStyle.OkOnly, " WinVIVA - Read Output File Error")
        Exit Sub

    End Sub

    Private Sub ImportOutput4(ByVal FileNum As Short, ByRef oxBook As Excel.Workbook, ByVal col As Short, ByVal results As Short, ByVal riserId As Short, ByVal whichFile As Short, ByVal whichIndex As Short)
        'read "BEND.DAT" file: mode shape
        On Error GoTo ErrHandler
        Dim oxSheet1, oxSheet As Excel.Worksheet
        Dim j, i, k As Short
        Dim sheetName As String
        Dim distancei, bndMomenti, stressi As Single

        '   bending moment and stress
        With CurProj.Riser(riserId)
            If results = 2 Then

                For j = 1 To .NumOutputModes
                    k = .ModeNumber(j)
                    sheetName = k
                    oxSheet = oxBook.Worksheets(sheetName)

                    oxSheet.Name = k
                    oxSheet.Cells(4, 1) = category(DISTANCE)
                    oxSheet.Cells(4, col) = category(BM)
                    oxSheet.Cells(4, col + 1) = category(STRESS)
                    For i = 1 To CurProj.NumPoints
                        'Input #FileNum, distancei, stressi

                        'If EOF(FileNum) Then Exit Sub
                        Input(FileNum, distancei)
                        Input(FileNum, bndMomenti)
                        Input(FileNum, stressi)
                        oxSheet.Cells(i + 4, 1) = distancei / CFLength
                        oxSheet.Cells(i + 4, col) = bndMomenti * CFMom
                        oxSheet.Cells(i + 4, col + 1) = stressi / CFStress
                    Next i
                    oxSheet.Range(oxSheet.Cells(2, 1), oxSheet.Cells(CurProj.NumPoints + 4, 1)).NumberFormat = "0.00"
                    oxSheet.Range(oxSheet.Cells(2, col), oxSheet.Cells(CurProj.NumPoints + 4, col)).NumberFormat = "0.00E-00"
                    oxSheet.Range(oxSheet.Cells(2, col + 1), oxSheet.Cells(CurProj.NumPoints + 4, col + 1)).NumberFormat = "0.00E-00"
                Next j
            Else
                oxSheet = oxBook.Worksheets(category(BM))
                oxSheet1 = oxBook.Worksheets(category(STRESS))
                oxSheet.Cells(1, 1) = category(DISTANCE)
                oxSheet1.Cells(1, 1) = oxSheet.Cells(1, 1)

                For j = 1 To .NumOutputModes
                    k = .ModeNumber(j)
                    oxSheet.Cells(1, j + 1) = k
                    oxSheet1.Cells(1, j + 1) = oxSheet.Cells(1, j + 1)
                    For i = 1 To CurProj.NumPoints
                        'If EOF(FileNum) Then Exit Sub
                        Input(FileNum, distancei)
                        Input(FileNum, bndMomenti)
                        Input(FileNum, stressi)

                        If j = 1 Then
                            oxSheet.Cells(i + 1, 1) = distancei / CFLength
                            oxSheet1.Cells(i + 1, 1) = oxSheet.Cells(i + 1, 1)
                        End If
                        oxSheet.Cells(i + 1, j + 1) = bndMomenti * CFMom
                        oxSheet1.Cells(i + 1, j + 1) = stressi / CFStress
                    Next i
                Next j
                oxSheet.Range(oxSheet.Cells(2, 1), oxSheet.Cells(CurProj.NumPoints + 1, 1)).NumberFormat = "0.00"
                oxSheet.Range(oxSheet.Cells(2, 2), oxSheet.Cells(CurProj.NumPoints + 1, .NumOutputModes + 1)).NumberFormat = "0.00E-00"
                oxSheet1.Range(oxSheet1.Cells(2, 1), oxSheet1.Cells(CurProj.NumPoints + 1, 1)).NumberFormat = "0.00"
                oxSheet1.Range(oxSheet1.Cells(2, 2), oxSheet1.Cells(CurProj.NumPoints + 1, .NumOutputModes + 1)).NumberFormat = "0.00E-00"
            End If
        End With

        Exit Sub

ErrHandler:
        MsgBox("Error occurred when reading " & OutputFiles(whichIndex, whichFile) & "file. The data may be imcomplete or corrupted", MsgBoxStyle.OkOnly, " WinVIVA - Read Output File Error")
        Exit Sub

    End Sub

    Private Sub ImportOutput5(ByVal FileNum As Short, ByRef oxBook As Excel.Workbook, ByVal col As Short, ByVal whichFile As Short, ByVal whichIndex As Short)

        'read "BEND_MM.DAT" file: mode shape (multi-frequency)
        On Error GoTo ErrHandler
        Dim i As Short
        Dim oxSheet As Excel.Worksheet
        Dim distancei, bndMomenti, stressi As Single
        oxSheet = oxBook.Worksheets("MM")
        oxSheet.Cells(4, 1) = category(DISTANCE)
        oxSheet.Cells(4, col) = category(BM)
        oxSheet.Cells(4, col + 1) = category(STRESS)
        '   multi-mode response
        For i = 1 To CurProj.NumPoints
            'If EOF(FileNum) Then Exit Sub
            Input(FileNum, distancei)
            Input(FileNum, bndMomenti)
            Input(FileNum, stressi)
            oxSheet.Cells(i + 4, 1) = distancei / CFLength
            oxSheet.Cells(i + 4, col) = bndMomenti * CFMom
            oxSheet.Cells(i + 4, col + 1) = stressi / CFStress

        Next i
        oxSheet.Range(oxSheet.Cells(4, 1), oxSheet.Cells(CurProj.NumPoints + 4, 1)).NumberFormat = "0.00"
        oxSheet.Range(oxSheet.Cells(4, col), oxSheet.Cells(CurProj.NumPoints + 4, col)).NumberFormat = "0.00E-00"
        oxSheet.Range(oxSheet.Cells(4, col + 1), oxSheet.Cells(CurProj.NumPoints + 4, col + 1)).NumberFormat = "0.00E-00"
        Exit Sub

ErrHandler:
        MsgBox("Error occurred when reading " & OutputFiles(whichFile, whichIndex) & " file. The data may be imcomplete or corrupted", MsgBoxStyle.OkOnly, " WinVIVA - Read Output File Error")
        Exit Sub
    End Sub

    Private Sub ImportOutput6(ByVal FileNum As Short, ByRef oxBook As Excel.Workbook, ByVal col As Short, ByVal results As Short, ByVal riserId As Short, ByVal whichFile As Short, ByVal whichIndex As Short)
        'read "CDRAG.DAT" file: drag coeffecients
        On Error GoTo ErrHandler

        Dim k, i, j, dum As Short
        Dim NumSegments As Short
        Dim seg2, seg0, seg1, SegNodeNum As Short
        Dim tmpSingle, SegSum, distancei As Single
        Dim SheetName As String
        Dim oxSheet As Excel.Worksheet
        With CurProj.Riser(riserId)
            '   average drag coefficient for each mode and each segment
            If results = 2 Then
                For j = 1 To .NumOutputModes
                    k = .ModeNumber(j)
                    SheetName = k
                    oxSheet = oxBook.Worksheets(SheetName)
                    tmpSingle = 0.0#

                    oxSheet.Cells(4, col) = category(CD)
                    For i = 1 To CurProj.NumPoints
                        'If EOF(FileNum) Then Exit Sub
                        Input(FileNum, distancei)
                        Input(FileNum, .DragCoef(k))
                        oxSheet.Cells(i + 4, col) = .DragCoef(k)
                        tmpSingle = tmpSingle + .DragCoef(k)

                        If i = 1 Then
                            SegNodeNum = 1
                            seg1 = .SegNum(distancei, seg2)
                            SegSum = .DragCoef(k)
                        Else
                            seg0 = seg1
                            seg1 = .SegNum(distancei, seg2)
                            If seg1 <> seg0 Then
                                SegNodeNum = 1
                                SegSum = .DragCoef(k)
                            Else
                                SegNodeNum = SegNodeNum + 1
                                SegSum = SegSum + .DragCoef(k)
                            End If
                        End If

                        .SegDragCoef(seg1, k) = SegSum / SegNodeNum
                    Next i

                    .AvgDragCoef(k) = tmpSingle / CurProj.NumPoints
                    oxSheet.Range(oxSheet.Cells(2, col), oxSheet.Cells(CurProj.NumPoints + 4, col)).NumberFormat = "0.000"
                Next j
            Else
                oxSheet = oxBook.Worksheets(category(CD))
                oxSheet.Cells(1, 1) = category(DISTANCE)

                For j = 1 To .NumOutputModes
                    k = .ModeNumber(j)
                    oxSheet.Cells(1, j + 1) = k
                    For i = 1 To CurProj.NumPoints
                        If EOF(FileNum) Then Exit Sub
                        Input(FileNum, distancei)
                        Input(FileNum, .DragCoef(k))
                        If j = 1 Then
                            oxSheet.Cells(i + 1, 1) = distancei / CFLength
                        End If
                        oxSheet.Cells(i + 1, j + 1) = .DragCoef(k)
                    Next i
                Next j
                oxSheet.Range(oxSheet.Cells(2, 1), oxSheet.Cells(CurProj.NumPoints + 1, .NumOutputModes + 1)).NumberFormat = "0.000"
            End If
        End With
        '   end of "CD.DAT"
        Exit Sub

ErrHandler:
        MsgBox("Error occurred when reading " & OutputFiles(whichFile, whichIndex) & " file. The data may be imcomplete or corrupted", MsgBoxStyle.OkOnly, " WinVIVA - Read Output File Error")
        Exit Sub

    End Sub

    Private Sub ImportOutput7(ByVal FileNum As Short, ByRef oxBook As Excel.Workbook, ByVal col As Short, ByVal riserId As Short, ByVal whichFile As Short, ByVal whichIndex As Short)
        'read "CD_MM.DAT" file: drag coeffecients
        On Error GoTo ErrHandler

        Dim i, j, dum As Short
        Dim NumSegments As Short
        Dim seg2, seg0, seg1, SegNodeNum As Short
        Dim tmpSingle, SegSum, distancei As Single
        Dim oxSheet As Excel.Worksheet
        '   average drag coefficient for each mode and each segment
        oxSheet = oxBook.Worksheets("MM")
        tmpSingle = 0.0#

        oxSheet.Cells(4, col) = category(CD)
        For i = 1 To CurProj.NumPoints
            'If EOF(FileNum) Then Exit Sub
            Input(FileNum, distancei)
            Input(FileNum, CurProj.Riser(riserId).DragCoef(0))
            oxSheet.Cells(i + 4, col) = CurProj.Riser(riserId).DragCoef(0)
            tmpSingle = tmpSingle + CurProj.Riser(riserId).DragCoef(0)

            If i = 1 Then
                SegNodeNum = 1
                seg1 = CurProj.Riser(riserId).SegNum(distancei, seg2)
                SegSum = CurProj.Riser(riserId).DragCoef(0)
            Else
                seg0 = seg1
                seg1 = CurProj.Riser(riserId).SegNum(distancei, seg2)
                If seg1 <> seg0 Then
                    SegNodeNum = 1
                    SegSum = CurProj.Riser(riserId).DragCoef(0)
                Else
                    SegNodeNum = SegNodeNum + 1
                    SegSum = SegSum + CurProj.Riser(riserId).DragCoef(0)
                End If
            End If

            CurProj.Riser(riserId).SegDragCoef(seg1, 0) = SegSum / SegNodeNum
        Next i

        CurProj.Riser(riserId).AvgDragCoef(0) = tmpSingle / CurProj.NumPoints
        oxSheet.Range(oxSheet.Cells(2, col), oxSheet.Cells(CurProj.NumPoints + 4, col)).NumberFormat = "0.000"

        Exit Sub

ErrHandler:
        MsgBox("Error occurred when reading " & OutputFiles(whichIndex, whichFile) & " file. The data may be imcomplete or corrupted", MsgBoxStyle.OkOnly, " WinVIVA - Read Output File Error")
        Exit Sub

    End Sub

    Private Sub ReadOutput9(ByVal FileNum As Short)

        'read "DYN-N.DAT" file: velocity

        On Error GoTo ErrHandler

        Dim j, i, k As Short
        Dim dum As Object
        Dim tmpVelU, tmpVelV As Single
        Dim tmpLine As String
        Dim tmpDiameter, VelAi, ReynoldsNumberi As Single
        Dim NumPts As Short

        'initialization
        dum = Nothing

            'extract velocity etc.
            '    Open .DataDirectory & NoSfxFileName(.FileName) & ".vel" For Output Access Write As #filenum
            '       number of actual output points
            Input(FileNum, dum)
            Input(FileNum, dum)
            Input(FileNum, NumPts)
            Input(FileNum, dum)

            tmpLine = LineInput(FileNum) 'ignore one line
            '    Write #filenum, "i", "dist", "velocity", "reynolds", "diameter"
            '       velocity arrays
            For i = 1 To NumPts
                'If EOF(FileNum) Then Exit Sub
                Input(FileNum, dum)
                Input(FileNum, dum)
                Input(FileNum, dum)
                Input(FileNum, dum)
                Input(FileNum, VelAi)
                Input(FileNum, dum)
                Input(FileNum, tmpDiameter)
                Input(FileNum, dum)
                Input(FileNum, dum)
                Input(FileNum, dum)
                ReynoldsNumberi = VelAi / CurProj.Water.Viscosity * tmpDiameter
                '      Write #filenum, i, Distance(i), VelA(i), ReynoldsNumber(i), tmpDiameter

                Input(FileNum, dum)
                Input(FileNum, dum)
                Input(FileNum, dum)
                Input(FileNum, dum)
                Input(FileNum, dum)
                Input(FileNum, dum)

                'tmpLine = LineInput(FileNum)
                'tmpLine = LineInput(FileNum)
            Next i

            '   Close #filenum
            '       end of "DYN-N.DAT" file

        Exit Sub

ErrHandler:
        MsgBox("Error occurred when reading dyn-n.out file. The data may be imcomplete or corrupted", MsgBoxStyle.OkOnly, " WinVIVA - Read Output File Error")
        Exit Sub

    End Sub

    Private Sub ImportOutput12(ByVal FileNum As Short, ByRef oxBook As Excel.Workbook, ByVal col As Short, ByVal results As Short, ByVal riserId As Short, ByVal whichFile As Short, ByVal whichIndex As Short)

        'read "FAT-MONO.DAT" file: fatigue life
        On Error GoTo ErrHandler

        Dim j, i, k As Short
        Dim distancei, dum As Single
        Dim SheetName As String
        Dim oxSheet As Excel.Worksheet
        '   fatigue life
        Input(FileNum, dum)
        Input(FileNum, dum)

        If results = 2 Then
            For j = 1 To CurProj.Riser(riserId).NumOutputModes
                k = CurProj.Riser(riserId).ModeNumber(j)
                SheetName = k
                oxSheet = oxBook.Worksheets(SheetName)

                oxSheet.Cells(4, col) = category(FAT_LIFE)

                For i = 1 To CurProj.NumPoints
                    '       If EOF(FileNum) Then Exit Sub
                    Input(FileNum, dum)
                    Input(FileNum, CurProj.Riser(riserId).FatLife(k))
                    oxSheet.Cells(i + 4, col) = CurProj.Riser(riserId).FatLife(k)

                Next i
                oxSheet.Range(oxSheet.Cells(2, col), oxSheet.Cells(CurProj.NumPoints + 4, col)).NumberFormat = "0.00E-00"

            Next j
        Else
            oxSheet = oxBook.Worksheets(category(FAT_LIFE))
            oxSheet.Cells(1, 1) = category(DISTANCE)

            For j = 1 To CurProj.Riser(riserId).NumOutputModes
                k = CurProj.Riser(riserId).ModeNumber(j)
                oxSheet.Cells(1, j + 1) = k
                For i = 1 To CurProj.NumPoints
                    '  If EOF(FileNum) Then Exit Sub
                    Input(FileNum, distancei)
                    Input(FileNum, CurProj.Riser(riserId).FatLife(k))
                    If j = 1 Then
                        oxSheet.Cells(i + 1, 1) = distancei / CFLength
                    End If
                    oxSheet.Cells(i + 1, j + 1) = CurProj.Riser(riserId).FatLife(k)
                Next i
            Next j
            oxSheet.Range(oxSheet.Cells(2, 1), oxSheet.Cells(CurProj.NumPoints + 1, 1)).NumberFormat = "0.00"
            oxSheet.Range(oxSheet.Cells(2, 2), oxSheet.Cells(CurProj.NumPoints + 1, CurProj.Riser(riserId).NumOutputModes + 1)).NumberFormat = "0.00E-00"
        End If

        Exit Sub

ErrHandler:
        MsgBox("Error occurred when reading " & OutputFiles(whichIndex, whichFile) & " file. The data may be imcomplete or corrupted", MsgBoxStyle.OkOnly, " WinVIVA - Read Output File Error")
        Exit Sub

    End Sub
    Private Sub ImportOutput13(ByVal FileNum As Short, ByRef oxBook As Excel.Workbook, ByVal col As Short, ByVal riserId As Short, ByVal whichFile As Short, ByVal whichIndex As Short)

        'read "FAT_MULTI.DAT" file: fatigue life (multi-frequency)
        On Error GoTo ErrHandler

        Dim j, i, k As Short
        Dim dum As Single
        Dim oxSheet As Excel.Worksheet
        '   fatigue life
        oxSheet = oxBook.Worksheets("MM")
        With oxSheet
            .Cells(4, col) = category(FAT_LIFE)

            For i = 1 To CurProj.NumPoints
                'If EOF(FileNum) Then Exit Sub
                Input(FileNum, dum)
                Input(FileNum, CurProj.Riser(riserId).FatLife(0))
                .Cells(i + 4, col) = CurProj.Riser(riserId).FatLife(0)

            Next i
            .Range(.Cells(4, col), .Cells(CurProj.NumPoints + 4, col)).NumberFormat = "0.00E-00"
        End With
        '   end of "FAT_MULTI.DAT" file
        Exit Sub

ErrHandler:
        MsgBox("Error occurred when reading " & OutputFiles(whichIndex, whichFile) & " file. The data may be imcomplete or corrupted", MsgBoxStyle.OkOnly, " WinVIVA - Read Output File Error")
        Exit Sub

    End Sub



    Private Sub ReadOutput14(ByVal FileNum As Short, ByVal riserId As Short, ByVal whichFile As Short, ByVal whichIndex As Short)
        'read "STARIN.DAT" file: Bending Strain
        On Error GoTo ErrHandler

        Dim j, i, k As Short
        Dim dum, distancei As Single
        With CurProj.Riser(riserId)
            'bending strain
            For j = 1 To .NumOutputModes
                k = .ModeNumber(j)
                .MaxStrain(k) = 0
                For i = 1 To CurProj.NumPoints
                    If EOF(FileNum) Then Exit Sub
                    Input(FileNum, distancei)
                    Input(FileNum, .Strain(k))
                    If System.Math.Abs(.Strain(k)) > .MaxStrain(k) Then .MaxStrain(k) = System.Math.Abs(.Strain(k))
                Next i
            Next j
        End With
        'end of "STARIN.OUT" file
        Exit Sub

ErrHandler:
        MsgBox("Error occurred when reading " & OutputFiles(whichIndex, whichFile) & " file. The data may be imcomplete or corrupted", MsgBoxStyle.OkOnly, " WinVIVA - Read Output File Error")
        Exit Sub

    End Sub


    Private Sub ReadOutput15(ByVal FileNum As Short, ByVal riserId As Short, ByVal whichFile As Short, ByVal whichIndex As Short)

        'read "STRAIN_MM.DAT" file: BENDING STRAIN (multi-frequency)
        On Error GoTo ErrHandler

        Dim i As Short
        Dim dum As Single
        With CurProj.Riser(riserId)
            'multi-mode response
            .MaxStrain(0) = 0.0#
            For i = 1 To CurProj.NumPoints
                If EOF(FileNum) Then Exit Sub
                Input(FileNum, dum)
                Input(FileNum, .Strain(0))
                If System.Math.Abs(.Strain(0)) > .MaxStrain(0) Then .MaxStrain(0) = System.Math.Abs(.Strain(0))
            Next i
            'end of "STRAIN_MM.DAT" file
        End With
        Exit Sub

ErrHandler:
        MsgBox("Error occurred when reading " & OutputFiles(whichIndex, whichFile) & " file. The data may be imcomplete or corrupted", MsgBoxStyle.OkOnly, " WinVIVA - Read Output File Error")
        Exit Sub

    End Sub


    Public Function ReadDOSText(ByVal FileSfx As String) As Short

        Dim NumTextLines As Short

        FileOpen(5, NoSfxFileName(CurProj.FQFileName) & "." & FileSfx, OpenMode.Input, OpenAccess.Read)

        '   read the DOS file
        'NumTextLines = ReadDOSOutput(5, DOSOutText)
        '   did we find anything?  If not, inform the user
        If NumTextLines <= 0 Then
            MsgBox(FileSfx & " file is empty. The relevant data have not been updated!", MsgBoxStyle.OkOnly, "WinVIVA - No Data Available")
        End If
        ReadDOSText = NumTextLines

        FileClose(5)

    End Function


    Public Function ReadDOSOutput(ByVal FileNum As Short, ByRef OutText As Object) As Short
        ' read DOS output file

        '   counting the number of lines
        ReadDOSOutput = 0
        Do While Not EOF(FileNum)
            ReadDOSOutput = ReadDOSOutput + 1
            OutText(ReadDOSOutput) = LineInput(FileNum)
            If ReadDOSOutput = MaxDOSOutLines Then
                MsgBox("The size of output file is too big. Please use other viewer " & " to view the full contents.", MsgBoxStyle.OkOnly, " WinVIVA - DOS Output File")
                Exit Do
            End If
        Loop

    End Function


    Public Sub GetDirNFileName(ByVal NewFullName As String, ByRef NewDir As String, ByRef NewFile As String)

        Dim DirChrPos, Pos, NameLen As Short
        Dim Valid As Boolean

        NameLen = Len(NewFullName)
        DirChrPos = NameLen
        Valid = True
        For Pos = NameLen To 1 Step -1
            Select Case Mid(NewFullName, Pos, 1)
                Case "\"
                    DirChrPos = Pos
                    Exit For
                Case "/", ":", "*", "?", "<", ">", "|", Chr(34)
                    Valid = False
            End Select
        Next Pos

        If Valid Then
            NewFile = Right(NewFullName, NameLen - DirChrPos)
        Else
            NewFile = ""
        End If
        NewDir = Left(NewFullName, DirChrPos)

    End Sub


    Public Sub ReadInWinVivaIni()

        On Error GoTo ErrorHandler

        Dim i As Short
        Dim tmpString As String
        Dim File As String

        'initializaiton
        tmpString = Nothing

        File = VIVADIR & "Winviva82.ini"
        FileOpen(10, File, OpenMode.Input, OpenAccess.Read)
        Input(10, tmpString)
        CurProj.DataDirectory = tmpString
        Do While Not (EOF(10))

            Input(10, tmpString)
            PreviousFiles.AddPreFile(Trim(tmpString))
        Loop
        FileClose(10)
        Exit Sub

ErrorHandler:
        FileClose(10)

        Exit Sub

    End Sub


    Public Sub OutputWinVivaIni()

        On Error GoTo ErrorHandler

        Dim i As Short
        Dim File As String
        File = VIVADIR & "Winviva82.ini"
        FileOpen(10, File, OpenMode.Output)

        WriteLine(10, CurProj.DataDirectory)
        '     Debug.Print "data dir = " & CurProj.DataDirectory
        For i = PreviousFiles.CountPreFile To 1 Step -1
            WriteLine(10, PreviousFiles.PreFile(i))
        Next i

        FileClose(10)
        Exit Sub

ErrorHandler:
        FileClose(10)
        MsgBox("Error when writing to 'WinViva8.ini' file.", MsgBoxStyle.OkOnly, " WinVIVA - Iutput/Output File Error")
        Exit Sub

    End Sub


    Public Sub ClearPlotTableResults()

        Dim i, ir As Short
        For ir = 1 To CurProj.nRisers
            With CurProj.Riser(ir)
                For i = 1 To NumVIVACorePoints / 7
                    .ModeNumber(i) = 0
                    .CoupledFreqs(i) = 0
                    .MaxAoverD(i) = 0
                    .AvgDragCoef(i) = 0
                    .MaxStress(i) = 0
                    .MaxBndMoment(i) = 0
                    .MaxStressPos(i) = 0
                    .FatiguePeriod(i) = 0
                    .FatigueLife(i) = 0
                Next

            End With
        Next

    End Sub


    Function RandomFileName(ByVal inpDir As String, ByVal ext As String) As String
        ' Pass in the name of a directory, build a random file name, make sure the file does not exist
        ' in the directory, and then return the full path (directory + file name) to the new file name.
        Dim directory As String
        Dim slash As String
        Dim checkFileName As String
        Dim fileName As String
        Dim i As Integer
        Dim num As Integer
        Dim ch As String

        directory = inpDir  ' We don't want to change the passed-in parm, so make a local copy of the value
        If InStr(directory, "/") = 0 And InStr(directory, "\") = 0 Then
            RandomFileName = ""  ' No slashes. Unknown file system setup - return empty string
            Exit Function  ' ======== EXIT ============
        End If
        If InStr(directory, "/") <> 0 And InStr(directory, "\") <> 0 Then
            RandomFileName = ""  ' Both slashes. Unknown file system setup - return empty string
            Exit Function  ' ======== EXIT ============
        End If
        If InStr(directory, "/") <> 0 Then slash = "/" Else slash = "\"
        ' Make sure the directory exists first
        If Right(directory, 1) <> slash Then directory = directory & slash ' Format dir so it ends in a slash
        On Error Resume Next
        checkFileName = Dir$(directory & "*.*")  ' Look for something in the directory
        On Error GoTo 0
        If Err.Number <> 0 Then  ' Will give a "path not found" error if the path doesn't exist
            Err.Number = 0
            RandomFileName = ""
            Exit Function  ' ======== EXIT ============
        End If
        Randomize()
        fileName = ""  ' Get us into the while loop
        While fileName = ""
            For i = 1 To 8  ' Random file name will have 8 characters
                If i = 1 Then num = Rnd() * 52 Else num = Rnd() * 62 ' starts w/ a letter, but other chars can be nums
                Select Case num
                    Case 0 To 25 : fileName = fileName & Chr(Asc("A") + num)  ' 0...25 = upper case letter
                    Case 26 To 51 : fileName = fileName & Chr(Asc("a") + num - 26)  ' 26...52 = lower case letter
                    Case 52 To 61 : fileName = fileName & CStr(num - 52)  ' 52...61 = number
                End Select
            Next
            checkFileName = Dir$(directory & fileName & "." & ext)  ' See if the file exists
            If checkFileName <> "" Then fileName = "" ' If found, continue the while loop
        End While
        ' At this point, the random file name doesn't exist in the dir - return full path to the new file
        RandomFileName = directory & fileName & "." & ext
    End Function

End Module