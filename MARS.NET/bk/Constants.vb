Option Strict Off
Option Explicit On
Module MARSConstants
	'--------------  conditional compilation flag -----------------
	
	Public Const DTCEL As Boolean = True
	
	'--------------------------------------------------------------
	
	Public Cancelled As Boolean ' for progress bar cancel action

    Public Const MaxNumLines As Short = 24
    Public Const MaxNumSeg As Short = 24
    Public Const MaxNumSubSeg As Short = 200
    Public Const MaxNumWells As Short = 30
    Public Const MaxNumIter As Short = 100

    Public Const MaxNumPipes As Short = 10
    Public Const MaxNumPipeNodes As Short = 11
    Public Const MaxNumGrids As Short = 201

    Public Const MaxNumPreFiles As Short = 4
    Public Const FileNumRes As Short = 99

    ' Program constants
    Public Const mX As Short = 1
    'UPGRADE_NOTE: MY was upgraded to MY_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Public Const MY_Renamed As Short = 2
    Public Const mZ As Short = 3

    Public Const ObjectFileName As String = "ObjFile.dat"
    Public Const CatenaryFile As String = "catline.dat"

    Public Const WarningColor As Short = 12
    Public Const DefaultColor As Short = 0

    ' Physical constants
    Public Const PI As Double = 3.14159265358979
    Public Const Degrees2Radians As Double = PI / 180.0#
    Public Const Radians2Degrees As Double = 180.0# / PI

    Public Const Knots2Ftps As Double = 1.6878098571
    Public Const Ftps2Knots As Double = 0.5924838013

    Public Const RegistryName As String = "\App Paths\mars.exe"
    Public Const MarsDirDefault As String = "C:\Program Files\MARS"
    Public Const msgTitle As String = "MARS"
    Public Const IniFile As String = "mars.ini"
    Public Const OnBoardFile As String = "onboard.dat"
    Public Const DerpDir As String = "C:\Riser Analysis\"
    Public Const DerpFile As String = "mars.off"

    Public Const PipeFile As String = "pipelines.inp"
    Public Const SeaBedFile As String = "seabed.inp"

    ' "Baroness"        ' default vessel
    Public Const PontHeight As Double = 28.9
    Public Const PontLength As Double = 357
    Public Const PontWidth As Double = 28.9
    Public Const PLength As Double = 49
    Public Const PDistance As Double = 61
    Public Const hBraceRadius As Double = 8.8
    Public Const DeckElev As Double = 120

    '   '  The following are constants to be used with a pontoon
    '   '  "Mar700"
    '    Public Const PontHeight As Double = 31.98
    '    Public Const PontLength As Double = 390.32
    '    Public Const PontWidth As Double = 39.36
    '    Public Const PLength As Double = 22.96
    '    Public Const PDistance As Double = 180.4
    '    Public Const hBraceRadius As Double = 3.28 * 3
    '    Public Const DeckElev As Double = 3.28 * 34.75

    Public MarsDir As String
    Public VesselPointsFromFile As Boolean
    Public MaxDynamicMotion As Double

    Public CurProj As Project
    Public Defaults As MarsIni
    Public CurVessel As Vessel
    Public CurField As Wells
    Public NumTimeSteps As Short
    Public TimeStep As Short
    Public MaxTimeSteps As Short
    Public TransX() As Object
    Public TransY() As Object
    Public TransYaw() As Object
    Public TransTension() As Double

    Public Enum SegType '2.1.4
        Wire = 1
        Chain = 2
        PresetWire = 3
        PresetChain = 4
        PresetPoly = 5
        PresetBuoy = 6
    End Enum

    Public IsMetricUnit As Boolean
    Public StressFactor, LFactor, FrcFactor, DiameterFactor As Double
    Public VelFactor As Double
    Public FrcUnit, LUnit, VelUnit, DiameterUnit As String

    Public Sub RefreshUnitLabels(ByRef frm As System.Windows.Forms.Form)
        Dim i As Short
        'UPGRADE_ISSUE: Control lblLengthUnit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        For i = 0 To frm.lblLengthUnit.Count - 1
            If IsMetricUnit Then
                'UPGRADE_ISSUE: Control lblLengthUnit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
                frm.lblLengthUnit(i).Caption = "m"
            Else
                'UPGRADE_ISSUE: Control lblLengthUnit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
                frm.lblLengthUnit(i).Caption = "ft"
            End If
        Next i
        'UPGRADE_ISSUE: Control lblForceUnit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        For i = 0 To frm.lblForceUnit.Count - 1
            If IsMetricUnit Then
                'UPGRADE_ISSUE: Control lblForceUnit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
                frm.lblForceUnit(i).Caption = "KN"
            Else
                'UPGRADE_ISSUE: Control lblForceUnit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
                frm.lblForceUnit(i).Caption = "kips"
            End If
        Next i
        'UPGRADE_ISSUE: Control lblVelUnit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        For i = 0 To frm.lblVelUnit.Count - 1
            If IsMetricUnit Then
                'UPGRADE_ISSUE: Control lblVelUnit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
                frm.lblVelUnit(i).Caption = "m/s"
            Else
                'UPGRADE_ISSUE: Control lblVelUnit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
                frm.lblVelUnit(i).Caption = "kn"
            End If
        Next i
    End Sub

    Public Sub InitProject()

        CurProj = New Project
        Defaults = New MarsIni

        ' MarsDir = GetDirFromRegistry(RegistryName) & "\"
        ' MarsDir = "c:\program files\mars\"

        Const ForReading As Short = 1
        Const ForWriting As Short = 2
        Const ForAppending As Short = 3
        Const TristateUseDefault As Short = -2
        Const TristateTrue As Short = -1
        Const TristateFalse As Short = 0
        Dim TS, FS, f, MyFile As Object
        Dim S, ss As String
        Dim i As Short

        FS = CreateObject("Scripting.FileSystemObject")
        'UPGRADE_WARNING: Couldn't resolve default property of object FS.FolderExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Not FS.FolderExists(My.Application.Info.DirectoryPath & "\Data") Then
            'UPGRADE_WARNING: Couldn't resolve default property of object FS.CreateFolder. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call FS.CreateFolder(My.Application.Info.DirectoryPath & "\Data")
        End If


        ' check if data subdir exist
        MarsDir = My.Application.Info.DirectoryPath & "\Data\"

    End Sub
	'------------------------------------------
End Module